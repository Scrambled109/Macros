from __future__ import annotations

import ast
import unittest
from pathlib import Path


SOURCE = Path(__file__).resolve().parents[1] / "job_assistant.py"


def _function(tree: ast.AST, name: str) -> ast.FunctionDef:
    for node in ast.walk(tree):
        if isinstance(node, ast.FunctionDef) and node.name == name:
            return node
    raise AssertionError(f"Missing function: {name}")


def _qualified_name(node: ast.AST) -> str:
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute):
        parent = _qualified_name(node.value)
        return f"{parent}.{node.attr}" if parent else node.attr
    return ""


class JobAssistantSourceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.tree = ast.parse(SOURCE.read_text(encoding="utf-8"))

    def test_comparison_launch_is_nonblocking(self) -> None:
        launch = _function(self.tree, "run_comparison")
        calls = {
            _qualified_name(node.func)
            for node in ast.walk(launch)
            if isinstance(node, ast.Call)
        }
        self.assertIn("subprocess.Popen", calls)
        self.assertNotIn("subprocess.run", calls)
        self.assertNotIn("process.wait", calls)
        self.assertNotIn("process.communicate", calls)
        self.assertIn("self._poll_comparison_process", calls)

    def test_comparison_completion_is_polled_by_tk_event_loop(self) -> None:
        poll = _function(self.tree, "_poll_comparison_process")
        calls = {
            _qualified_name(node.func)
            for node in ast.walk(poll)
            if isinstance(node, ast.Call)
        }
        self.assertIn("process.poll", calls)
        self.assertIn("self.after", calls)
        self.assertIn("self.post_background_notice", calls)
        self.assertNotIn("messagebox.showinfo", calls)
        self.assertNotIn("messagebox.showerror", calls)

    def test_dashboard_only_contains_assistant_owned_stages(self) -> None:
        core_source = SOURCE.with_name("job_core.py").read_text(encoding="utf-8")
        core_tree = ast.parse(core_source)
        stages = next(
            node
            for node in core_tree.body
            if isinstance(node, ast.Assign)
            and any(isinstance(target, ast.Name) and target.id == "STAGES" for target in node.targets)
        )
        values = ast.literal_eval(stages.value)
        self.assertEqual(
            [key for key, _label in values],
            ["bom", "dxf", "plate_model", "autobom", "comparison"],
        )

    def test_launcher_uses_shared_source_checkout(self) -> None:
        launcher = SOURCE.with_name("Launch Job Assistant.bat").read_text(
            encoding="utf-8"
        )
        self.assertIn("py job-assistant\\job_assistant.py", launcher)
        self.assertIn("py -m pip install -r requirements.txt", launcher)
        self.assertNotIn("Engineering Job Assistant.exe", launcher)


if __name__ == "__main__":
    unittest.main()
