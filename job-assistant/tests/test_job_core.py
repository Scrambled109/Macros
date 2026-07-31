import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from job_core import (  # noqa: E402
    MANIFEST_VERSION,
    JobError,
    PromotionItem,
    acknowledge_override,
    command_bom,
    command_comparison,
    command_dxf,
    dashboard_warnings,
    discover_drawings,
    load_manifest,
    load_settings,
    migrate_manifest,
    parse_comparison_summary,
    plan_promotions,
    prepare_dxf_workspace,
    promote_files,
    save_manifest,
    save_settings,
    setup_job,
    stage_checks,
    suggest_job_number,
)


class CoreTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name) / "01234 Customer" / "SA ENGINEERING PROCESS"
        self.model = self.root / "odd MODEL name"
        self.cut = self.root / "incoming cuts"
        self.model.mkdir(parents=True)
        self.cut.mkdir()

    def tearDown(self):
        self.temp.cleanup()

    def manifest(self):
        path = setup_job(self.root, self.model, self.cut, "01234", "Test", "A")
        return load_manifest(path), path

    def test_setup_uses_selected_paths_and_isolated_workspace(self):
        manifest, path = self.manifest()
        self.assertEqual(Path(manifest["paths"]["model_3d"]), self.model)
        self.assertEqual(path.parent.name, "_JOB_ASSISTANT")
        self.assertTrue(Path(manifest["workspace"]["staging"]).is_dir())
        self.assertEqual(suggest_job_number(self.root), "01234")

    def test_setup_never_replaces_existing_manifest_history(self):
        manifest, path = self.manifest()
        manifest["events"].append({"type": "keep_me"})
        save_manifest(manifest, path)
        with self.assertRaisesRegex(JobError, "Open it instead"):
            setup_job(self.root, self.model, self.cut, "99999", "Replace", "Z")
        self.assertEqual(load_manifest(path)["events"][-1]["type"], "keep_me")

    def test_atomic_save_leaves_no_temporary_file(self):
        manifest, path = self.manifest()
        save_manifest(manifest, path)
        self.assertEqual(
            json.loads(path.read_text())["manifest_version"], MANIFEST_VERSION
        )
        self.assertEqual(list(path.parent.glob("*.tmp")), [])

    def test_v1_manifest_migrates_missing_keys(self):
        old = {
            "manifest_version": 1,
            "root": str(self.root),
            "job": {"number": "1", "name": "x", "revision": "A"},
            "paths": {"cad_models": str(self.model), "cut_files": str(self.cut)},
            "stages": {},
            "events": [],
        }
        migrated = migrate_manifest(old)
        self.assertEqual(migrated["manifest_version"], MANIFEST_VERSION)
        self.assertIn("comparison", migrated)
        self.assertEqual(migrated["paths"]["model_3d"], str(self.model))
        self.assertEqual(migrated["events"][-1]["type"], "manifest_migrated")

    def test_future_manifest_has_useful_error(self):
        with self.assertRaisesRegex(JobError, "supports up to"):
            migrate_manifest({"manifest_version": 999})

    def test_incomplete_manifest_has_useful_error(self):
        with self.assertRaisesRegex(JobError, "job number is missing"):
            migrate_manifest(
                {
                    "manifest_version": MANIFEST_VERSION,
                    "root": str(self.root),
                    "job": {"revision": "A"},
                }
            )

    def test_settings_round_trip_and_defaults(self):
        target = Path(self.temp.name) / "local" / "settings.json"
        save_settings(
            {"parts_list_template": r"\\server\share\Parts List.xlsx"}, target
        )
        loaded = load_settings(target)
        self.assertEqual(
            loaded["parts_list_template"], r"\\server\share\Parts List.xlsx"
        )
        self.assertIn("solidworks_executable", loaded)

    def test_drawing_discovery_defaults_and_safe_numbered_copy(self):
        dxf = self.cut / "Plate A.dxf"
        dwg = self.cut / "Shape B.DWG"
        txt = self.cut / "note.txt"
        dxf.write_text("dxf")
        dwg.write_text("dwg")
        txt.write_text("note")
        candidates = discover_drawings(self.cut)
        self.assertEqual(
            [(c.path.name, c.selected) for c in candidates],
            [("Plate A.dxf", True), ("Shape B.DWG", False)],
        )
        manifest, _ = self.manifest()
        run = prepare_dxf_workspace(manifest, [dxf])
        self.assertEqual((run / "001" / dxf.name).read_text(), "dxf")
        self.assertTrue(dxf.exists())

        second_run = prepare_dxf_workspace(manifest, [dxf])
        self.assertNotEqual(run, second_run)
        self.assertTrue(second_run.name.startswith(run.name))

    def test_warnings_are_overridable_and_audited(self):
        manifest, _ = self.manifest()
        checks = stage_checks(manifest, "dxf")
        self.assertTrue(any(c.code == "previous_stage_incomplete" for c in checks))
        acknowledge_override(manifest, "dxf", checks)
        event = manifest["events"][-1]
        self.assertEqual(event["type"], "warning_overridden")
        self.assertTrue(event["user"])

    def test_promotion_defaults_conflict_to_skip_then_backs_up_replace(self):
        manifest, _ = self.manifest()
        staged = Path(manifest["workspace"]["staging"]) / "part.dwg"
        staged.write_text("new")
        production = self.root / "production"
        production.mkdir()
        existing = production / staged.name
        existing.write_text("old")
        planned = plan_promotions([staged], production)
        self.assertTrue(planned[0].conflict)
        self.assertEqual(planned[0].action, "skip")
        result = promote_files(
            manifest, [PromotionItem(staged, existing, True, "backup_replace")]
        )[0]
        self.assertEqual(result["status"], "replaced")
        self.assertEqual(existing.read_text(), "new")
        self.assertEqual(Path(result["backup"]).read_text(), "old")
        report = Path(result["report"])
        self.assertTrue(report.is_file())
        report_data = json.loads(report.read_text())
        self.assertEqual(report_data["counts"]["replaced"], 1)
        self.assertEqual(report_data["revision"], "A")

    def test_promotion_rejects_file_outside_staging(self):
        manifest, _ = self.manifest()
        outside = self.root / "not-staged.dwg"
        outside.write_text("do not promote")
        production = self.root / "production"
        result = promote_files(
            manifest,
            [PromotionItem(outside, production / outside.name, False, "copy")],
        )[0]
        self.assertEqual(result["status"], "failed")
        self.assertIn("inside the assistant Staging folder", result["error"])
        self.assertFalse((production / outside.name).exists())

    def test_comparison_summary(self):
        reports = Path(self.temp.name) / "reports"
        reports.mkdir()
        (reports / "errors_requiring_action.csv").write_text("part,status\nP1,error\n")
        (reports / "source_data_issues.csv").write_text("issue\nwarning\n")
        summary = parse_comparison_summary(reports)
        self.assertEqual(summary["status"], "action_required")
        self.assertEqual((summary["errors"], summary["warnings"]), (1, 1))

    def test_dashboard_warnings_summarize_actionable_comparison(self):
        manifest, _ = self.manifest()
        manifest["comparison"] = {"status": "action_required", "errors": 4}
        self.assertIn(
            "Production comparison has 4 actionable discrepancy item(s).",
            dashboard_warnings(manifest),
        )

    def test_versioned_comparison_summary_in_timestamped_folder(self):
        reports = Path(self.temp.name) / "reports"
        run = reports / "Part_Comparison_2026-01-02_030405"
        run.mkdir(parents=True)
        payload = {
            "schema_version": 1,
            "outcome": "action_required",
            "counts": {
                "errors": 2,
                "missing_core": 1,
                "not_checked": 3,
                "source_issues": 4,
            },
            "reports": {
                "excel": str(run / "production_part_comparison.xlsx"),
                "html": str(run / "comparison_report.html"),
                "folder": str(run),
            },
        }
        (run / "comparison_summary.json").write_text(json.dumps(payload))
        summary = parse_comparison_summary(reports)
        self.assertEqual(summary["status"], "action_required")
        self.assertEqual((summary["errors"], summary["warnings"]), (3, 7))
        self.assertEqual(Path(summary["folder"]), run)

    def test_cad_batch_source_uses_runtime_paths_not_job_constants(self):
        config = (
            Path(__file__).resolve().parents[2]
            / "solidworks/cad-batch-converter/Config.bas"
        ).read_text(encoding="utf-8")
        self.assertIn("Environ$(environmentName)", config)
        self.assertIn("GetSetting(SETTINGS_APP, SETTINGS_SECTION", config)
        self.assertNotIn("U:\\Engineering\\CAD Services\\Current Jobs", config)

    def test_external_commands_are_argument_lists_with_spaces(self):
        repo = Path(r"Z:\Shared Macros")
        bom = command_bom(
            "python.exe",
            repo,
            Path("source copy.xlsx"),
            Path("out.xlsx"),
            Path("template.xlsx"),
        )
        self.assertEqual(bom[-3:], ["source copy.xlsx", "out.xlsx", "template.xlsx"])
        compare = command_comparison(
            "python.exe",
            repo,
            Path("Nest Data"),
            Path("parts.csv"),
            Path("sw.csv"),
            Path("Reports"),
        )
        self.assertIn("Nest Data", compare)

        dxf = command_dxf(
            repo,
            Path(r"C:\CAD Suite\accoreconsole.exe"),
            Path(r"C:\CAD Suite\acad.exe"),
        )
        self.assertEqual(
            dxf[-4:],
            [
                "-AcadConsolePath",
                r"C:\CAD Suite\accoreconsole.exe",
                "-AcadGuiPath",
                r"C:\CAD Suite\acad.exe",
            ],
        )

        packaged = command_bom(
            "ignored-python.exe",
            repo,
            Path("source.xlsx"),
            Path("out.xlsx"),
            Path("template.xlsx"),
            Path("Engineering BOM Converter.exe"),
        )
        self.assertEqual(packaged[0], "Engineering BOM Converter.exe")
        self.assertNotIn("bom_converter.py", " ".join(packaged))


if __name__ == "__main__":
    unittest.main()
