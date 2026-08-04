from __future__ import annotations

import re
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
CONVERTER = REPO_ROOT / "solidworks" / "cad-batch-converter"
COLOR_TO_LAYER = REPO_ROOT / "autocad" / "dxf-orchestrator" / "ColortoLayer.lsp"
MODULE_NAMES = (
    "Config.bas",
    "Utilities.bas",
    "AutoCAD_Filter.bas",
    "Main.bas",
    "SolidWorks_Import.bas",
    "NativeSketch.bas",
    "TextMarking.bas",
    "TextStamp.bas",
)

# Test injection hook used only by the source audit's own validation.
MODULE_SOURCES: dict[str, str] | None = None
COLOR_TO_LAYER_SOURCE: str | None = None

PROC_START = re.compile(
    r"^(?:Public |Private |Friend )?"
    r"(?:Sub|Function|Property(?: Get| Let| Set)?)\b",
    re.IGNORECASE,
)
PROC_END = re.compile(r"^End (?:Sub|Function|Property)\b", re.IGNORECASE)
TYPE_START = re.compile(
    r"^(?:Public |Private )?(?:Type|Enum)\b", re.IGNORECASE
)
TYPE_END = re.compile(r"^End (?:Type|Enum)\b", re.IGNORECASE)


def _sources() -> dict[str, str]:
    if MODULE_SOURCES is not None:
        return MODULE_SOURCES
    return {
        name: (CONVERTER / name).read_text(encoding="utf-8")
        for name in MODULE_NAMES
    }


def _color_to_layer_source() -> str:
    if COLOR_TO_LAYER_SOURCE is not None:
        return COLOR_TO_LAYER_SOURCE
    return COLOR_TO_LAYER.read_text(encoding="utf-8")


def _code_only(line: str) -> str:
    """Remove VBA strings and comments without treating apostrophes in strings as comments."""
    out: list[str] = []
    in_string = False
    i = 0
    while i < len(line):
        char = line[i]
        if char == '"':
            if in_string and i + 1 < len(line) and line[i + 1] == '"':
                i += 2
                continue
            in_string = not in_string
            out.append(" ")
        elif char == "'" and not in_string:
            break
        else:
            out.append(" " if in_string else char)
        i += 1
    return "".join(out).strip()


def _late_top_level_code(source: str) -> list[tuple[int, str]]:
    """Return declarations or other top-level code found after the first procedure."""
    seen_procedure = False
    in_procedure = False
    in_type = False
    offenders: list[tuple[int, str]] = []

    for line_number, raw in enumerate(source.splitlines(), start=1):
        code = _code_only(raw)
        if not code:
            continue

        if in_procedure:
            if PROC_END.match(code):
                in_procedure = False
            continue

        if in_type:
            if TYPE_END.match(code):
                in_type = False
            continue

        if PROC_START.match(code):
            seen_procedure = True
            in_procedure = True
            continue

        if TYPE_START.match(code):
            if seen_procedure:
                offenders.append((line_number, code))
            in_type = True
            continue

        if seen_procedure:
            offenders.append((line_number, code))

    return offenders


class CadBatchVbaSourceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.sources = _sources()

    def test_all_expected_modules_are_present(self) -> None:
        self.assertEqual(set(MODULE_NAMES), set(self.sources))

    def test_module_declarations_precede_every_procedure(self) -> None:
        problems = {
            name: _late_top_level_code(source)
            for name, source in self.sources.items()
            if _late_top_level_code(source)
        }
        self.assertEqual(
            {},
            problems,
            "VBA module declarations and Public Type blocks must be above every "
            "Sub/Function; otherwise other modules report 'User-defined type not defined'.",
        )

    def test_required_custom_types_exist_in_config(self) -> None:
        config = self.sources["Config.bas"]
        for type_name in ("TFileResult", "TSegment", "TTextMark"):
            self.assertRegex(
                config,
                rf"(?mi)^Public Type {type_name}\s*$",
                f"{type_name} must be a Public Type in Config.bas",
            )

    def test_no_early_bound_cad_types_remain(self) -> None:
        pattern = re.compile(
            r"\bAs\s+(?:SldWorks\.|SolidWorks\.|AutoCAD\.|Acad[A-Z])",
            re.IGNORECASE,
        )
        hits = {
            name: [
                (number, line.strip())
                for number, line in enumerate(source.splitlines(), start=1)
                if pattern.search(_code_only(line))
            ]
            for name, source in self.sources.items()
        }
        self.assertFalse(
            any(hits.values()),
            f"CAD COM objects must stay late-bound so no type-library reference is required: {hits}",
        )

    def test_solidworks_has_one_direct_main_entry_point(self) -> None:
        entries = []
        for name, source in self.sources.items():
            for match in re.finditer(r"(?mi)^\s*Public Sub main\s*\(\s*\)", source):
                entries.append((name, source[: match.start()].count("\n") + 1))
        self.assertEqual(1, len(entries), entries)
        self.assertEqual("Main.bas", entries[0][0])

    def test_public_procedure_names_are_unique(self) -> None:
        declarations: dict[str, list[str]] = {}
        pattern = re.compile(
            r"(?mi)^\s*Public\s+(?:Sub|Function|Property(?:\s+(?:Get|Let|Set))?)"
            r"\s+([A-Za-z_]\w*)"
        )
        for module_name, source in self.sources.items():
            for name in pattern.findall(source):
                declarations.setdefault(name.casefold(), []).append(module_name)
        duplicates = {
            name: modules for name, modules in declarations.items() if len(modules) > 1
        }
        self.assertEqual({}, duplicates, f"Ambiguous public VBA procedure names: {duplicates}")

    def test_converter_layer_contract_matches_orchestrator(self) -> None:
        config = self.sources["Config.bas"]
        mappings = dict(
            re.findall(
                r'\((\d+)\s*\.\s*"([^"]+)"\)', _color_to_layer_source()
            )
        )
        constants = dict(
            re.findall(
                r'(?mi)^Public Const (\w+) As String = (?:_\s*)?\n?\s*"([^"]+)"',
                config,
            )
        )
        profile_layers = {
            name.strip() for name in constants["PROFILE_LAYERS"].split(",")
        }
        marking_layers = {
            name.strip() for name in constants["TEXT_LAYER"].split(",")
        }

        self.assertEqual(mappings["1"], constants["TARGET_LAYER"])
        self.assertEqual(mappings["5"], constants["INSIDE_CUT_LAYER"])
        self.assertEqual({mappings["1"], mappings["5"]}, profile_layers)
        self.assertTrue({mappings["3"], mappings["6"], mappings["7"]} <= marking_layers)
        self.assertNotIn(mappings["2"], profile_layers | marking_layers)
        self.assertIn("NormalizedLayerName", self.sources["Utilities.bas"])
        self.assertIn(
            "ModelSpaceLayerSummary(doc)", self.sources["AutoCAD_Filter.bas"]
        )

    def test_autocad_object_property_assignment_uses_set(self) -> None:
        text_stamp = self.sources["TextStamp.bas"]
        self.assertIsNone(
            re.search(
                r"(?mi)^\s*(?!Set\b)doc\.ActiveLayer\s*=\s*scratch\s*$",
                text_stamp,
            )
        )
        self.assertRegex(
            text_stamp,
            r"(?mi)^\s*Set doc\.ActiveLayer\s*=\s*scratch\s*$",
        )


if __name__ == "__main__":
    unittest.main()
