# AutoBOM bounding-box macros

`AutoBOMProperties.swp` and `AUTOBOMACTUAL.swp` are retained as runnable
artifacts because both were already in use and the repository does not yet
record which name is canonical. Their corresponding `.bas` exports are under
`reference/`.

## Static debugging applied to the reference source

The reviewed reference modules now:

- exit cleanly when `GetComponents` returns no array, rather than calling
  `LBound`/`UBound` on `Empty`;
- mark a bounding box as newly inserted only if the API actually returned a
  feature; and
- count and report a failed `Save3` as skipped instead of reporting it as a
  successful processed part.

These fixes affect reference source only. They must be applied and compiled in
the VBA projects before the `.swp` behavior changes. Both macros remain
high-impact: they overwrite properties and silently save part files.
