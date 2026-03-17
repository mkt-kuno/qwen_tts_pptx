# AGENTS.md
Repository guidance for coding agents working in this project.

## Project Summary
- Name: PowerPoint TTS Movie Builder.
- Goal: turn a `.pptx` deck into narrated MP4s by extracting tagged notes, generating WAV narration with Qwen TTS, exporting slide images through PowerPoint COM, and muxing audio/video with `ffmpeg`.
- Language: Python 3.12.
- Primary platform: Windows, because slide export depends on Microsoft PowerPoint COM.
- Main entrypoints: `app.pipeline.build_movie`, `app.models.export_voice`, `app.gui`.

## Repository Layout
- `app/common/`: shared dataclasses, note parsing, language resolution, workspace path helpers.
- `app/models/`: voice asset lookup and prompt cache export.
- `app/synthesis/`: Qwen model loading, prompt creation, and audio generation.
- `app/slides/`: PowerPoint COM slide export.
- `app/video/`: WAV concatenation and `ffmpeg` video assembly.
- `app/gui/`: Tkinter desktop UI.
- `requirements/`: pip requirement fragments for CPU and CUDA.
- `conda*.yml`: Conda environment definitions.
- `example/`: sample reference audio/text assets.
- `work/`, `output/`: generated artifacts; do not treat as source folders.

## Rules Files
- No `.cursorrules` file exists.
- No `.cursor/rules/` directory exists.
- No `.github/copilot-instructions.md` file exists.
- Follow this file and the existing code patterns in the touched modules.

## Environment Setup
- Preferred stable environment: `conda env create -f conda.env.yml` then `conda activate qwentts`.
- Alternative environment: `conda.cuda.yml`, which creates `qwentts-cuda`.
- Python is pinned to 3.12 in `.python-version` and the Conda files.
- Conda files include `cffi`, `pillow`, and `charset-normalizer`; if `_cffi_backend` or `PIL.Image` is missing, recreate the environment or reinstall the missing package.

## Build / Run Commands
Prefer the Conda Python executable. Examples assume:

```powershell
C:\Users\mkt-kuno\miniconda3\envs\qwentts\python.exe
```

- Full CLI movie build with a named voice:
  - `python -m app.pipeline.build_movie <deck>.pptx --voice-name kuwano --languages JA EN ZH`
- Full CLI movie build with explicit reference files:
  - `python -m app.pipeline.build_movie <deck>.pptx --ref-audio example/kuwano.wav --ref-text example/kuwano.txt --languages JA EN ZH`
- Prompt cache export from CLI:
  - `python -m app.models.export_voice --voice-name kuwano`
- Tkinter GUI:
  - `python -m app.gui`

## Verification Commands
- Bytecode-compile the app tree:
  - `python -m compileall app`
- Import the GUI module:
  - `python -c "import app.gui.main; print('gui import ok')"`
- Import Qwen runtime only:
  - `python -c "from qwen_tts import Qwen3TTSModel; print('qwen_tts ok')"`
- Verify Pillow import if model import fails:
  - `python -c "import PIL.Image; print('pillow ok')"`

## Lint / Format / Test Status
- No checked-in lint config was found: no `pyproject.toml`, `ruff.toml`, `mypy.ini`, or `pytest.ini` are present.
- `ruff` is not part of the checked environment.
- `pytest` is not part of the checked environment.
- There is no `tests/` directory today.
- Practical verification baseline is `python -m compileall app` plus focused imports and smoke checks for the changed area.

## Test Guidance
- There is currently no automated test suite to run.
- If tests are added later, prefer `pytest`.
- Run all tests later with:
  - `python -m pytest`
- Run one test file later with:
  - `python -m pytest tests/test_notes.py`
- Run one test later with:
  - `python -m pytest tests/test_notes.py::test_extract_tagged_text`
- If you add pytest-based tests, also update this file and the environment dependencies.

## Manual Verification Guidance
- Parser-only changes: exercise note parsing helpers against a local `.pptx` deck.
- Voice asset changes: test both `--voice-name` lookup and explicit `--ref-audio/--ref-text` inputs.
- GUI-only changes: import `app.gui.main` and, if feasible, open the Tkinter window manually.
- GUI behavior note: the GUI requires explicit `Ref Audio` and `Ref Text`; it no longer uses `.pt` prompt cache files.
- Pipeline or synthesis changes: run `python -m compileall app`; test `from qwen_tts import Qwen3TTSModel` before trying the full pipeline.
- ffmpeg or PowerPoint integration changes: prefer smoke tests over large refactors; do not delete unrelated `work/` or `output/` artifacts.

## Code Style
Match the existing repository style closely.

### Imports
- Start Python files with `from __future__ import annotations`.
- Group imports in this order: standard library, third-party packages, local `app.*` imports.
- Prefer explicit imports over wildcard imports.
- Keep multiline imports readable.
- In GUI code, lazy imports inside worker methods are acceptable for heavy runtime dependencies.

### Formatting
- Follow normal PEP 8 formatting.
- Use 4-space indentation.
- Keep blocks small and single-purpose.
- Leave blank lines between top-level definitions.
- Prefer simple f-strings over older formatting styles.

### Types
- Add type hints to public functions and most helpers.
- Prefer built-in generics such as `list[str]`, `dict[str, Path]`, and `tuple[int, int, int]`.
- Use `Path` instead of raw path strings in Python APIs.
- Keep `Any` limited to GUI, queue, or thread boundaries where libraries make precise typing awkward.

### Naming
- Modules: lowercase, descriptive, underscore-separated if needed.
- Functions: snake_case verbs such as `build_movie`, `export_voice_prompt`, `resolve_voice_asset`.
- Classes: PascalCase nouns such as `WorkspacePaths`, `LanguageSpec`, `MovieBuilderGui`.
- Public language tags stay uppercase: `JA`, `EN`, `ZH`.

### Paths and Data
- Centralize workspace layout through `WorkspacePaths`.
- Keep generated intermediates under `work/` and final outputs under `output/`.
- Preserve naming conventions: slides `pageN.png`, audio `pageN.wav`, videos `ja.mp4`, `en.mp4`, `zh.mp4`.

### Error Handling
- Fail fast with specific exceptions: `FileNotFoundError` for missing files, `ValueError` for invalid inputs, `RuntimeError` for pipeline or external-tool failures.
- Include the relevant path, tag, or command context in error messages.
- Raise exceptions in reusable functions; reserve `sys.exit()` for CLI-only termination paths.
- In the GUI, log failures first, then surface them via dialogs.

### Logging and Progress
- Use module-level loggers: `logger = logging.getLogger(__name__)`.
- Prefer `logging` over `print()` for runtime diagnostics.
- Log major transitions: voice resolution, device detection, model import/load, slide export, prompt creation, audio generation, and final outputs.
- Keep optional progress callbacks separate from core business logic.
- The GUI mirrors logs both to its log window and to standard output; preserve that behavior.

### External Tools and Processes
- Use `subprocess.run(..., check=True)` for `ffmpeg`, PowerShell, and helper probes.
- Pass subprocess arguments as lists, not shell-joined strings, inside Python code.
- Preserve Windows compatibility for PowerPoint COM export.
- `ffmpeg` and `sox` availability can affect runtime behavior; log missing tool issues clearly.

### GUI Guidance
- Keep the GUI thin; business logic belongs in reusable modules under `app.pipeline`, `app.models`, `app.synthesis`, `app.slides`, and `app.video`.
- Never run long tasks on the Tkinter main thread.
- Use worker threads plus queue-based status and log handoff.
- Validate user inputs before starting background work.
- Do not reintroduce prompt-cache export into the GUI unless requirements change.

## Change Scope Guidance
- Prefer small, composable functions.
- Reuse existing helpers before adding abstractions.
- Avoid changing public interfaces unless an API change is intentional and you update all callers.
- Avoid changing output layout or file naming without a clear reason.
- Do not destructively clean unrelated user files or generated artifacts.

## Agent Checklist
- Read the touched module and nearby helpers before editing.
- Preserve existing interfaces unless an API change is intentional.
- Run `python -m compileall app` at minimum after code changes.
- Run a targeted import or smoke check for the code you changed.
- Update `README.md` or this file when commands, workflows, dependencies, or GUI behavior change.
