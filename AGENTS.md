# AGENTS.md

Repository guidance for coding agents working in this project.

## Scope and Goal

- Goal: convert tagged PowerPoint decks (`.pptx`) into narrated MP4 videos.
- Primary user flow is a compact 3-step pipeline exposed in both CLI and Gradio UI.
- Python target is 3.12 (`.python-version` is `3.12.13`).
- Keep modules small, explicit, and boundary-focused.

## Architecture at a Glance

- `app/pipeline/steps.py`: canonical step APIs used by CLI and UI.
- `app/web/gradio_app.py`: thin UI orchestrator; avoid business logic here.
- `app/common/`: pure helpers (`notes`, `languages`, `paths`).
- `app/synthesis/`: Qwen TTS runtime integration and WAV generation.
- `app/slides/`: Windows PowerPoint COM export implementation.
- `app/video/`: ffmpeg-based assembly utilities.
- `app/models/`: voice asset resolution and prompt export helper CLI.

## Step Contracts (Do Not Blur)

- Step 1: image export (`.pptx -> work/slides/pageN.png`) on Windows only.
- Step 2: audio synthesis (`.pptx + ref wav/txt -> work/audio/*/pageN.wav`).
- Step 3: video build (`PNG + WAV -> output/*.mp4`).
- Step 2 cache is in-memory/process-local (cleared on restart).
- Step 2 cache key includes text hash, language assignment, sampling params, dtype, model/device, and reference hashes.
- In JP/EN/ZH runs, Step 2 retries once for anomalous audio (too long/short/silent/invalid).

## Platform and Runtime Constraints

- Step 1 requires Windows + PowerPoint COM; do not claim Linux support for Step 1.
- Step 2/3 can run on CPU or CUDA, but CUDA may fall back to CPU at runtime.
- ffmpeg must be available in PATH for video assembly.
- Prefer `pathlib.Path` across the codebase.

## Environment Setup Commands

```powershell
# CPU
conda env create -f conda.env.yml
conda activate qwentts
# CUDA
conda env create -f conda.cuda.yml
conda activate qwentts-cuda
# Optional (CUDA only)
pip install --no-cache-dir --no-build-isolation flash-attn
```

## Build / Run Commands

```powershell
# Gradio UI
python -m app.web
# Full pipeline
python -m app.pipeline.build_movie <deck>.pptx --ref-audio example/kuwano.wav --ref-text example/kuwano.txt --languages EN JP ZH ES IT FR
# Export voice prompt
python -m app.models.export_voice --ref-audio <voice>.wav --ref-text <voice>.txt --model-size 1.7B
# Legacy helper
python -m app.synthesis.generate_audio <deck>.pptx --ref-audio <voice>.wav --ref-text <voice>.txt --languages EN JP ZH ES IT FR
```

## Verification Commands (Fast Checks)

```powershell
python -m compileall app
python -c "import app.web.gradio_app; print('web import ok')"
python -c "from qwen_tts import Qwen3TTSModel; print('qwen_tts ok')"
```

## Lint and Format Commands

There is no committed linter config file (`pyproject.toml`, `setup.cfg`, etc.) currently.
If `ruff` is available in your environment, use these defaults:

```powershell
python -m ruff check app
python -m ruff format --check app
```

If you format, keep diffs minimal and avoid opportunistic refactors.

## Test Commands

Current repository status: no committed `tests/` suite yet.
When tests are present and `pytest` is installed, use:

```powershell
# All tests
python -m pytest -q
# Single test file
python -m pytest tests/test_example.py -q
# Single test function (important)
python -m pytest tests/test_example.py::test_case_name -q
# Match by expression
python -m pytest -k "cache and step2" -q
# unittest single test equivalent
python -m unittest tests.test_example.TestClass.test_method
```

## Code Style and Conventions

### Imports

- Use standard grouping: stdlib, third-party, local imports.
- Keep imports explicit; avoid wildcard imports.
- Use `from __future__ import annotations` in module headers (existing pattern).

### Formatting

- Follow PEP 8 style and existing project formatting.
- Keep lines readable; avoid dense one-liners.
- Prefer small helper functions over deeply nested blocks.
- Keep changes focused; avoid broad formatting-only edits.

### Types

- Use typed signatures broadly (`Path`, `list[str]`, etc.).
- Prefer dataclasses for structured step I/O and immutable specs.
- Keep return types explicit on public functions.

### Naming

- `snake_case` for functions/variables.
- `PascalCase` for classes/dataclasses.
- `UPPER_SNAKE_CASE` for module constants.
- Use descriptive names (`resolved_device`, `slide_padding_sec`, etc.).

### Error Handling

- Raise specific exceptions (`ValueError`, `FileNotFoundError`, `RuntimeError`) with actionable messages.
- Raise from shared modules; do not terminate process there.
- Do not use `sys.exit()` outside CLI entrypoints.
- Validate file existence and input assumptions early.

### Logging

- Use `logging`, not `print`, in application code.
- Keep log messages concise and operationally useful.
- Include contextual fields (slide number, language, path) where relevant.

### Boundaries and Responsibilities

- Keep UI thin; put logic in pipeline/common/synthesis/video modules.
- Do not move framework-specific concerns into core domain modules.
- Preserve step boundaries and existing data flow unless explicitly requested.

### Filesystem and Paths

- Use `Path` operations instead of string path math.
- Create parent directories with `mkdir(parents=True, exist_ok=True)` when writing outputs.
- Keep generated artifacts under `work/` and `output/`.

## Change Discipline for Agents

- Prefer minimal, targeted diffs.
- Reuse existing helpers before adding new abstractions.
- Avoid silent behavior changes; document major behavior changes in `README.md`.
- For runtime-sensitive changes (TTS params, caching, ffmpeg behavior), run verification commands.
- If adding tests later, include at least one focused single-test runnable example in docs.

## Cursor / Copilot Rules

- No `.cursor/rules/`, `.cursorrules`, or `.github/copilot-instructions.md` files were found.
- Therefore, this `AGENTS.md` is the authoritative agent instruction file in-repo.
