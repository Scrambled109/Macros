#!/usr/bin/env python3
"""Best-effort VBA source extractor for SolidWorks .swp compound files.

This has no third-party dependencies. It reads OLE Compound File Binary (CFB)
streams and searches VBA module streams for MS-OVBA compressed source. Export
from the SolidWorks VBA editor remains the authoritative method, especially for
forms, references, and password-protected projects.
"""

from __future__ import annotations

import argparse
import math
import struct
from pathlib import Path


FREE = 0xFFFFFFFF
END = 0xFFFFFFFE


class CompoundFile:
    def __init__(self, path: Path):
        self.data = path.read_bytes()
        if self.data[:8] != bytes.fromhex("D0CF11E0A1B11AE1"):
            raise ValueError(f"Not an OLE compound file: {path}")
        self.sector_size = 1 << self._u16(30)
        self.mini_sector_size = 1 << self._u16(32)
        self.mini_cutoff = self._u32(56)
        self.dir_start = self._u32(48)

        difat = list(struct.unpack_from("<109I", self.data, 76))
        next_difat = self._u32(68)
        for _ in range(self._u32(72)):
            sector = self._sector(next_difat)
            values = struct.unpack(f"<{self.sector_size // 4}I", sector)
            difat.extend(values[:-1])
            next_difat = values[-1]
        fat_sectors = [value for value in difat if value < 0xFFFFFFFA]
        fat_bytes = b"".join(self._sector(value) for value in fat_sectors)
        self.fat = struct.unpack(f"<{len(fat_bytes) // 4}I", fat_bytes)

        directory = self._read_chain(self.dir_start)
        self.entries = []
        for offset in range(0, len(directory), 128):
            entry = directory[offset : offset + 128]
            if len(entry) < 128:
                break
            name_len = struct.unpack_from("<H", entry, 64)[0]
            name = entry[: max(0, name_len - 2)].decode("utf-16le", "replace")
            self.entries.append(
                {
                    "name": name,
                    "type": entry[66],
                    "start": struct.unpack_from("<I", entry, 116)[0],
                    "size": struct.unpack_from("<Q", entry, 120)[0],
                }
            )

        root = next((entry for entry in self.entries if entry["type"] == 5), None)
        self.mini_stream = self._read_chain(root["start"])[: root["size"]] if root else b""
        mini_fat_start = self._u32(60)
        mini_fat_count = self._u32(64)
        if mini_fat_count and mini_fat_start not in (FREE, END):
            raw = self._read_chain(mini_fat_start)[: mini_fat_count * self.sector_size]
            self.mini_fat = struct.unpack(f"<{len(raw) // 4}I", raw)
        else:
            self.mini_fat = ()

    def _u16(self, offset: int) -> int:
        return struct.unpack_from("<H", self.data, offset)[0]

    def _u32(self, offset: int) -> int:
        return struct.unpack_from("<I", self.data, offset)[0]

    def _sector(self, number: int) -> bytes:
        start = (number + 1) * self.sector_size
        return self.data[start : start + self.sector_size]

    def _chain(self, start: int, table: tuple[int, ...]) -> list[int]:
        result, seen = [], set()
        current = start
        while current not in (FREE, END) and current < len(table) and current not in seen:
            seen.add(current)
            result.append(current)
            current = table[current]
        return result

    def _read_chain(self, start: int) -> bytes:
        return b"".join(self._sector(item) for item in self._chain(start, self.fat))

    def read_stream(self, entry: dict[str, int | str]) -> bytes:
        size = int(entry["size"])
        start = int(entry["start"])
        if size < self.mini_cutoff and self.mini_fat:
            chunks = []
            for item in self._chain(start, self.mini_fat):
                offset = item * self.mini_sector_size
                chunks.append(self.mini_stream[offset : offset + self.mini_sector_size])
            return b"".join(chunks)[:size]
        return self._read_chain(start)[:size]

    def streams(self):
        for entry in self.entries:
            if entry["type"] == 2 and entry["name"]:
                yield str(entry["name"]), self.read_stream(entry)


def decompress_vba(data: bytes) -> bytes:
    """Decompress an MS-OVBA compressed container beginning with 0x01."""
    if not data or data[0] != 1:
        raise ValueError("Missing compressed-container signature")
    cursor = 1
    output = bytearray()
    while cursor + 2 <= len(data):
        header = struct.unpack_from("<H", data, cursor)[0]
        if ((header >> 12) & 0x7) != 0x3:
            break
        chunk_size = (header & 0x0FFF) + 3
        compressed = bool(header & 0x8000)
        chunk_end = min(len(data), cursor + chunk_size)
        cursor += 2
        chunk_start = len(output)
        if not compressed:
            output.extend(data[cursor:chunk_end])
            cursor = chunk_end
            continue
        while cursor < chunk_end:
            flags = data[cursor]
            cursor += 1
            for bit in range(8):
                if cursor >= chunk_end:
                    break
                if not flags & (1 << bit):
                    output.append(data[cursor])
                    cursor += 1
                    continue
                if cursor + 2 > chunk_end:
                    raise ValueError("Truncated copy token")
                token = struct.unpack_from("<H", data, cursor)[0]
                cursor += 2
                position = len(output) - chunk_start
                bit_count = max(4, math.ceil(math.log2(max(position, 1))))
                length_mask = 0xFFFF >> bit_count
                length = (token & length_mask) + 3
                offset = (token >> (16 - bit_count)) + 1
                if offset > position:
                    raise ValueError("Invalid copy-token offset")
                for _ in range(length):
                    output.append(output[-offset])
        cursor = chunk_end
    return bytes(output)


def source_candidates(stream: bytes):
    for offset, value in enumerate(stream):
        if value != 1:
            continue
        try:
            source = decompress_vba(stream[offset:])
        except (ValueError, IndexError, struct.error):
            continue
        text = source.decode("cp1252", "replace")
        if "Attribute VB_" in text or "Option Explicit" in text:
            yield offset, text


def safe_name(value: str) -> str:
    cleaned = "".join(char if char.isalnum() or char in "._-" else "_" for char in value)
    return cleaned.strip("._") or "module"


def extract(path: Path, output: Path) -> int:
    compound = CompoundFile(path)
    output.mkdir(parents=True, exist_ok=True)
    found = 0
    seen_source = set()
    for stream_name, stream in compound.streams():
        for _, text in source_candidates(stream):
            marker = text.find("Attribute VB_Name")
            source = text[marker:] if marker >= 0 else text
            if source in seen_source:
                continue
            seen_source.add(source)
            match = next(
                (line for line in source.splitlines() if line.startswith("Attribute VB_Name")),
                "",
            )
            module_name = match.split('"')[1] if '"' in match else stream_name
            target = output / f"{safe_name(module_name)}.bas"
            clean_source = "\n".join(
                line.rstrip() for line in source.rstrip("\x00\r\n").splitlines()
            )
            target.write_text(clean_source + "\n", encoding="utf-8")
            print(f"{path}: extracted {module_name} -> {target}")
            found += 1
            break
    return found


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("swp", nargs="+", type=Path, help="SolidWorks .swp file(s)")
    parser.add_argument("--output-root", type=Path, default=Path("macros/source"))
    args = parser.parse_args()
    total = 0
    for path in args.swp:
        total += extract(path, args.output_root / safe_name(path.stem))
    print(f"Extracted {total} module(s).")
    return 0 if total else 1


if __name__ == "__main__":
    raise SystemExit(main())
