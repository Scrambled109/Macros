# Job Assistant EXE launch and quarantine troubleshooting

## Permission error or disappearing EXE

An executable that exists after the build but disappears when it is launched
was most likely quarantined by Windows Security or organizational endpoint
protection. The white GUI application icon is not itself an error: Windows
uses a different default icon for a windowed application than for the two
console companions.

The current build packages the GUI as a folder-based application to avoid the
older self-extracting EXE's unpack-to-temporary-directory behavior. Copy the
whole `dist\Engineering Job Assistant` directory, including `_internal` and
both companion EXEs. Copying only `Engineering Job Assistant.exe` produces an
incomplete application.

## Safe checks

1. Check **Windows Security > Virus & threat protection > Protection history**
   and the organization's Cylance or other endpoint-security console.
2. Record the event time, detection name, original path, repository commit,
   and SHA-256 hash of the build before taking further action.
3. Give those details to IT/security and ask them to verify the build and use
   the organization's approved signing, deployment, or allow-list process.
4. Rebuild on the approved Windows build machine, then test the complete folder
   from an IT-approved local location. Network shares may have stricter launch
   policies than local folders.

To calculate evidence without launching the application:

```bat
certutil -hashfile "dist\Engineering Job Assistant\Engineering Job Assistant.exe" SHA256
```

Do not disable endpoint protection, add a personal exclusion, change file
permissions to bypass a policy, or repeatedly restore and launch a quarantined
file. A PyInstaller build is not code-signed merely because it built
successfully; production distribution should use the organization's signing
and release process.

## Distinguishing an incomplete copy

If the EXE remains present but reports a missing DLL or fails immediately,
confirm that `_internal` is beside it and recopy the entire application folder.
If the GUI opens but a conversion reports a missing companion, confirm that
`Engineering BOM Converter.exe` and
`Engineering Production Comparison.exe` are also beside the GUI EXE.
