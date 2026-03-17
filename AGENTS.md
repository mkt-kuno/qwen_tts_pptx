# AGENTS.md

Repository guidance for coding agents.

## Project summary

- Goal: convert tagged PowerPoint decks into narrated MP4 videos.
- UI: Gradio (`app/web/gradio_app.py`).
- Language: Python 3.12.
- Main workflow: three compact steps in `app/pipeline/steps.py`.

## Source layout

- `app/common/`: pure helpers (notes, languages, paths)
- `app/synthesis/`: Qwen TTS runtime helpers
- `app/slides/`: Windows PowerPoint COM export
- `app/video/`: ffmpeg assembly utilities
- `app/pipeline/steps.py`: step APIs
- `app/web/`: Gradio UI

## Step contracts

- Step 1: image export (`.pptx -> pageN.png`) on Windows only.
- Step 2: audio synthesis (`.pptx + ref wav/txt -> pageN.wav`) with SHA256 cache.
- Step 3: video build (`PNG + WAV -> ja/en/zh MP4 + multilingual MP4`).

Keep these boundaries stable. UI code must stay thin and call step APIs.

## Environment

- CPU: `conda.env.yml` (`qwentts`)
- CUDA: `conda.cuda.yml` (`qwentts-cuda`)

## Commands

- Gradio app:
  - `python -m app.web`
- Full CLI pipeline:
  - `python -m app.pipeline.build_movie <deck>.pptx --ref-audio example/kuwano.wav --ref-text example/kuwano.txt --languages JA EN ZH`

## Verification

- `python -m compileall app`
- `python -c "import app.web.gradio_app; print('web import ok')"`
- `python -c "from qwen_tts import Qwen3TTSModel; print('qwen_tts ok')"`

## Style

- Keep files small and explicit.
- Prefer dataclasses and typed function signatures for step I/O.
- Avoid framework logic in core modules.
- Use logging, not print.
- Raise exceptions from shared modules; do not `sys.exit()` outside CLI entrypoints.
