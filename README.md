# File to Markdown Converter

Use `convert_files_to_markdown.py` to convert files in a folder to Markdown. Give the source folder as the first argument. The script scans subfolders by default.

Supported files: `.pdf`, `.docx`, `.doc`, `.xlsx`, `.txt`, `.eml`, `.jpg`, `.jpeg`, `.png`, `.heic`, and `.heif`.

## Install

Install Python 3.10 or newer. From the project folder, create a local `.venv` and install the packages in `requirements.txt`.

macOS or Linux (Bash or zsh):

```bash
python3 -m venv .venv
./.venv/bin/python -m pip install -r requirements.txt
```

Windows (PowerShell):

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

The requirements cover PDF, DOCX, XLSX, and image conversion. Converting legacy `.doc` files also requires Microsoft Word on Windows. Plain text and email files need no additional software.

For scanned PDFs, install Tesseract. The script looks for it at `C:\Program Files\Tesseract-OCR\tesseract.exe`. If yours is elsewhere, change `TESSERACT_EXE` near the top of the script. Pages without embedded text may be empty when Tesseract is unavailable.

## Image transcription with Ollama

[Install Ollama](https://ollama.com/download). If it is not already running, start it in one terminal:

```sh
ollama serve
```

In another terminal, pull the default vision model:

```sh
ollama pull minicpm-v
```

### Choose a vision model

| Model | When to use it |
| --- | --- |
| [`minicpm-v`](https://ollama.com/library/minicpm-v) (default) | Start here for routine images and faster batches, especially on a CPU. |
| [`qwen2.5vl:7b`](https://ollama.com/library/qwen2.5vl:7b) | Try it for dense text, forms, tables, or complex layouts when the default misses detail. It may be slower on a CPU. |

To use `qwen2.5vl:7b`, pull it and select it when converting (Windows example):

```powershell
ollama pull qwen2.5vl:7b
.\.venv\Scripts\python.exe convert_files_to_markdown.py ./documents --m qwen2.5vl:7b
```

Pass another installed vision model with `--m MODEL`.

> *Before converting images, run Ollama and pull the selected model yourself. `qwen2.5vl:7b` requires Ollama 0.7.0 or newer. Allow about 6 GB of disk space per model. Large images may take several minutes on a CPU; check a few results before a long batch.*

## Convert files

Run commands from the project folder. Replace `./documents` with your source folder.

macOS or Linux:

```bash
./.venv/bin/python convert_files_to_markdown.py ./documents
```

Windows (PowerShell):

```powershell
.\.venv\Scripts\python.exe convert_files_to_markdown.py ./documents
```

Add `--o` to write into a separate folder, or `--f` to convert one file. For example, on Windows:

```powershell
.\.venv\Scripts\python.exe convert_files_to_markdown.py ./documents --o ./markdown
.\.venv\Scripts\python.exe convert_files_to_markdown.py ./documents --f report.docx
```

On macOS or Linux, replace `.\.venv\Scripts\python.exe` in the examples with `./.venv/bin/python`.

## Switches

| Switch | Effect |
| --- | --- |
| `SOURCE` | Folder to scan. Required. |
| `--o PATH` | Write Markdown to a separate folder. Default: beside each source file. |
| `--f FILE` | Convert one file by filename or path. Use an absolute path if a relative path is ambiguous. |
| `--t pdf docx` | Convert only the listed file types. |
| `--top` | Scan only files directly in `SOURCE`. |
| `--all` | Include subfolders (the default). |
| `--skip` | Skip files that already have a Markdown output, regardless of modification time. |
| `--m MODEL` | Use a different installed Ollama vision model for images. |

## Output

By default, Markdown is written beside each source file. With `--o`, the script mirrors the source folder structure in the output folder.

The output name keeps the full source filename and adds `.md`. For example, `Meeting Notes.docx` becomes `Meeting Notes.docx.md`.

- On later runs, the script converts a file again only when its source is newer than its Markdown output.
- If two files would write to the same output path, the script stops before writing.
- If a file fails, the script reports it, continues with other files, and exits with a nonzero status.

Source images stay in place. HEIC/HEIF conversion uses a temporary PNG that is removed afterward.
