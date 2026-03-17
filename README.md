# PowerPoint TTS Movie Builder

PowerPoint TTS Movie Builder converts a tagged `.pptx` deck into narrated videos.
The app now uses a Gradio web UI and a compact three-step workflow.

## Workflow

1. **Image Export (Windows only)**: `.pptx -> pageN.png`
2. **Audio Synthesis**: `.pptx + ref wav/txt -> pageN.wav` (with per-page cache)
3. **Video Build**: `PNG + WAV -> MP4`

Outputs include:

- Per-language MP4 (`ja.mp4`, `en.mp4`, `zh.mp4`)
- Multilingual MP4 (`multilingual.mp4`, one video track + multiple audio tracks)

`PPTX` is required for script extraction in Step 2.

## Project layout

- `app/common/`: language specs, note parsing, workspace paths
- `app/synthesis/`: Qwen model loading and WAV synthesis helpers
- `app/slides/`: PowerPoint COM slide export (Windows)
- `app/video/`: ffmpeg-based video assembly
- `app/pipeline/steps.py`: compact step APIs used by both CLI and web UI
- `app/web/gradio_app.py`: Gradio app

## Environment

- `conda.env.yml`: CPU-first environment (`qwentts`)
- `conda.cuda.yml`: CUDA environment (`qwentts-cuda`)

Install one environment:

```powershell
conda env create -f conda.env.yml
conda activate qwentts
```

For CUDA users, FlashAttention is optional. The CUDA environment intentionally skips it for maximum setup compatibility. If you want faster inference and your system supports building it, install it manually after environment creation:

```powershell
pip install --no-cache-dir --no-build-isolation flash-attn
```

## Run web UI

```powershell
python -m app.web
```

This starts the Gradio app with tabs for Step 1, Step 2, and Step 3.

## Colab (T4) quick start

Use `colab_run.ipynb` for a Gradio-first Colab workflow.

- Open `colab_run.ipynb` in Colab
- Set runtime to GPU (T4)
- Run all cells
- Open the public Gradio URL and use the same three-step UI

## Run full pipeline from CLI

```powershell
python -m app.pipeline.build_movie <deck>.pptx --ref-audio example/kuwano.wav --ref-text example/kuwano.txt --languages JA EN ZH
```

## Cache behavior (Step 2)

The audio step stores per-page cache metadata JSON under `work/cache/audio/`.
WAV regeneration is skipped when the cache key matches:

- script SHA256
- language tag
- model/device/dtype settings
- reference audio SHA256
- reference text SHA256

Use `--force-regenerate` in CLI (or checkbox in UI) to bypass cache.

## Linux / Colab note

Step 1 uses PowerPoint COM and is Windows-only.
On Linux/Colab, upload a `slides.zip` containing `page1.png`, `page2.png`, ... and an `audio.zip` containing `ja/`, `en/`, `zh/` WAV folders in Step 3.

## Verification

```powershell
python -m compileall app
python -c "import app.web.gradio_app; print('web import ok')"
python -c "from qwen_tts import Qwen3TTSModel; print('qwen_tts ok')"
```
