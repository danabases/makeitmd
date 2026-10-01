from __future__ import annotations

import contextlib
import importlib.util
import io
import os
import sys
import tempfile
import types
import unittest
import urllib.error
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]

# The tests exercise file selection and output behavior without PDF libraries
# or an Ollama server. Restore sys.modules after loading the converter.
fitz_stub = types.ModuleType("fitz")
fitz_stub.TOOLS = types.SimpleNamespace(mupdf_display_errors=lambda _: None)
pymupdf4llm_stub = types.ModuleType("pymupdf4llm")

spec = importlib.util.spec_from_file_location(
    "convert_files_to_markdown", ROOT / "convert_files_to_markdown.py"
)
converter = importlib.util.module_from_spec(spec)
with mock.patch.dict(
    sys.modules, {"fitz": fitz_stub, "pymupdf4llm": pymupdf4llm_stub}
):
    spec.loader.exec_module(converter)


class ConverterTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / "source"
        self.source.mkdir()
        self.output = self.root / "output"
        converter._ollama_ready = False

    def run_converter(self, *options: str) -> None:
        argv = [
            "convert_files_to_markdown.py",
            str(self.source),
            "--o",
            str(self.output),
            *options,
        ]
        with mock.patch.object(sys, "argv", argv), contextlib.redirect_stdout(io.StringIO()):
            converter.main()

    def test_output_name_and_skip_use_requested_destination(self) -> None:
        source = self.source / "Meeting Notes.txt"
        source.write_text("first", encoding="utf-8")
        (self.source / "Meeting Notes.md").write_text("unrelated", encoding="utf-8")

        self.run_converter()
        destination = self.output / "Meeting Notes.txt.md"
        self.assertEqual(destination.read_text(encoding="utf-8"), "first\n")

        source.write_text("second", encoding="utf-8")
        self.run_converter("--skip")
        self.assertEqual(destination.read_text(encoding="utf-8"), "first\n")
        newer = destination.stat().st_mtime + 5
        os.utime(source, (newer, newer))
        self.run_converter()
        self.assertEqual(destination.read_text(encoding="utf-8"), "second\n")

    def test_mixed_images_do_not_move_sources_or_keep_temp_png(self) -> None:
        for name in ("photo.jpg", "photo.png", "photo.heic"):
            (self.source / name).write_bytes(b"sample")
        before = {path.name for path in self.source.iterdir()}
        transcribed_paths = []

        def fake_heic_to_png(_source: Path, destination: Path) -> None:
            destination.write_bytes(b"temporary png")

        def fake_transcribe(path: Path) -> str:
            transcribed_paths.append(path)
            return path.suffix

        with (
            mock.patch.object(converter, "ensure_ollama_ready"),
            mock.patch.object(converter, "heic_to_png", side_effect=fake_heic_to_png),
            mock.patch.object(converter, "vision_transcribe_image", side_effect=fake_transcribe),
        ):
            self.run_converter()

        self.assertEqual({path.name for path in self.source.iterdir()}, before)
        self.assertFalse((self.source / "jpg").exists())
        self.assertEqual(
            {path.name for path in self.output.iterdir()},
            {"photo.jpg.md", "photo.png.md", "photo.heic.md"},
        )
        temp_png = [path for path in transcribed_paths if path.parent != self.source]
        self.assertEqual(len(temp_png), 1)
        self.assertFalse(temp_png[0].exists())

    def test_heic_temp_png_is_removed_on_failure(self) -> None:
        source = self.source / "photo.heic"
        source.write_bytes(b"sample")
        temp_paths = []

        def fake_heic_to_png(_source: Path, destination: Path) -> None:
            destination.write_bytes(b"temporary png")
            temp_paths.append(destination)

        with (
            mock.patch.object(converter, "ensure_ollama_ready"),
            mock.patch.object(converter, "heic_to_png", side_effect=fake_heic_to_png),
            mock.patch.object(
                converter,
                "vision_transcribe_image",
                side_effect=RuntimeError("transcription failed"),
            ),
        ):
            with self.assertRaisesRegex(RuntimeError, "transcription failed"):
                converter.image_to_markdown(source)
        self.assertEqual(len(temp_paths), 1)
        self.assertFalse(temp_paths[0].exists())
        self.assertFalse((self.source / "photo.png").exists())

    def test_type_filter_never_moves_jpg(self) -> None:
        for name in ("photo.jpg", "photo.png", "report.pdf"):
            (self.source / name).write_bytes(b"sample")
        with mock.patch.object(converter, "pdf_to_markdown", return_value="pdf\n"):
            self.run_converter("--t", "pdf")
        self.assertTrue((self.source / "photo.jpg").is_file())
        self.assertFalse((self.source / "jpg").exists())
        self.assertEqual(
            {path.name for path in self.output.iterdir()}, {"report.pdf.md"}
        )

    def test_file_paths_and_ambiguity(self) -> None:
        from_source = self.source / "report.txt"
        from_source.write_text("source", encoding="utf-8")
        working = self.root / "working"
        working.mkdir()
        from_working = working / "report.txt"
        from_working.write_text("working", encoding="utf-8")

        with mock.patch.object(Path, "cwd", return_value=working):
            self.assertEqual(
                converter.resolve_single_file(self.source, str(from_source), True),
                from_source,
            )
            with self.assertRaisesRegex(ValueError, "Use an absolute path"):
                converter.resolve_single_file(self.source, "report.txt", True)
            from_source.unlink()
            self.assertEqual(
                converter.resolve_single_file(self.source, "report.txt", True),
                from_working,
            )

    def test_destination_collision_fails_before_conversion(self) -> None:
        first = self.root / "one" / "same.txt"
        second = self.root / "two" / "same.txt"
        first.parent.mkdir()
        second.parent.mkdir()
        first.write_text("one", encoding="utf-8")
        second.write_text("two", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "same Markdown destination"):
            converter.validate_unique_destinations(
                [first, second], self.source, self.output
            )
        self.assertFalse(self.output.exists())

    def test_ollama_requires_running_service_and_model(self) -> None:
        with (
            mock.patch.object(
                converter,
                "ollama_tags",
                side_effect=urllib.error.URLError("connection refused"),
            ),
            mock.patch.object(converter.subprocess, "Popen") as start,
            mock.patch.object(converter.subprocess, "run") as install,
        ):
            with self.assertRaisesRegex(RuntimeError, "ollama serve"):
                converter.ensure_ollama_ready()
            start.assert_not_called()
            install.assert_not_called()

        with mock.patch.object(converter, "ollama_tags", return_value={"models": []}):
            with self.assertRaisesRegex(RuntimeError, "ollama pull"):
                converter.ensure_ollama_ready()

        with mock.patch.object(
            converter,
            "ollama_tags",
            return_value={"models": [{"name": "minicpm-v:latest"}]},
        ):
            converter.ensure_ollama_ready()
            self.assertTrue(converter._ollama_ready)

    def test_short_switches_and_legacy_aliases(self) -> None:
        with mock.patch.object(sys, "argv", ["converter.py", str(self.source)]):
            args = converter.parse_args()
        self.assertEqual(args.source_dir, self.source)
        self.assertTrue(args.scan_subfolders)
        self.assertTrue(args.overwrite_existing)
        self.assertEqual(args.vision_model, "minicpm-v")

        with mock.patch.object(
            sys, "argv", ["converter.py", str(self.source), "--all"]
        ):
            self.assertTrue(converter.parse_args().scan_subfolders)

        with mock.patch.object(
            sys,
            "argv",
            [
                "converter.py", str(self.source), "--o", str(self.output),
                "--f", "report.txt", "--top", "--skip", "--t", "txt",
                "--m", "qwen2.5vl:7b",
            ],
        ):
            args = converter.parse_args()
        self.assertEqual(args.output_dir, self.output)
        self.assertEqual(args.file, "report.txt")
        self.assertFalse(args.scan_subfolders)
        self.assertFalse(args.overwrite_existing)
        self.assertEqual(args.extensions, {".txt"})
        self.assertEqual(args.vision_model, "qwen2.5vl:7b")

        with mock.patch.object(
            sys,
            "argv",
            [
                "converter.py", "--source-dir", str(self.source),
                "--output-dir", str(self.output), "--file", "report.txt",
                "--no-recursive", "--skip-existing", "--types", "txt",
                "--vision-model", "qwen2.5vl:7b",
            ],
        ):
            legacy = converter.parse_args()
        self.assertEqual(legacy.source_dir, self.source)
        self.assertEqual(legacy.output_dir, args.output_dir)
        self.assertEqual(legacy.file, args.file)
        self.assertEqual(legacy.scan_subfolders, args.scan_subfolders)
        self.assertEqual(legacy.overwrite_existing, args.overwrite_existing)
        self.assertEqual(legacy.extensions, args.extensions)
        self.assertEqual(legacy.vision_model, args.vision_model)

    def test_help_shows_short_switches_only(self) -> None:
        help_text = io.StringIO()
        with (
            mock.patch.object(sys, "argv", ["converter.py", "--help"]),
            contextlib.redirect_stdout(help_text),
            self.assertRaises(SystemExit) as raised,
        ):
            converter.parse_args()
        self.assertEqual(raised.exception.code, 0)
        output = help_text.getvalue()
        for switch in ("--o", "--f", "--all", "--top", "--skip", "--t", "--m"):
            self.assertIn(switch, output)
        for old_switch in ("--source-dir", "--output-dir", "--skip-existing", "--vision-model"):
            self.assertNotIn(old_switch, output)

    def test_failed_image_continues_and_exits_nonzero(self) -> None:
        (self.source / "a.png").write_bytes(b"sample")
        (self.source / "b.txt").write_text("converted", encoding="utf-8")
        with mock.patch.object(
            converter, "ensure_ollama_ready", side_effect=RuntimeError("Ollama missing")
        ):
            with self.assertRaises(SystemExit) as raised:
                self.run_converter()
        self.assertEqual(raised.exception.code, 1)
        self.assertEqual(
            (self.output / "b.txt.md").read_text(encoding="utf-8"), "converted\n"
        )


if __name__ == "__main__":
    unittest.main()
