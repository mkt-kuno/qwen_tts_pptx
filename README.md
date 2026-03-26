# PowerPoint TTS Movie Builder

PowerPoint TTS Movie Builder converts a tagged `.pptx` deck into narrated videos.
The app now uses a Gradio web UI and a compact three-step workflow.

## Workflow

1. **Image Export (Windows only)**: `.pptx -> pageN.png`
2. **Audio Synthesis**: `.pptx + ref wav/txt -> pageN.wav` (with in-memory cache)
3. **Video Build**: `PNG + WAV -> MP4`

Outputs include:

- Per-language MP4 (`en.mp4`, `jp.mp4`, `zh.mp4`, `es.mp4`, `it.mp4`, `fr.mp4`) — each with a video track sized to that language's audio duration, an ISO 639-3 language tag on the audio track, and embedded chapter markers at each slide boundary.
- Multilingual MP4 (`multilingual.mp4`, one video track + multiple audio tracks) — video track sized to the longest language per slide, with embedded chapter markers.

In `multilingual.mp4`, EN is muxed as the first audio track and marked as default.

Supported note tags (ISO 639-1): `EN`, `JP`, `ZH`, `ES`, `IT`, `FR`

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
python -m app.pipeline.build_movie <deck>.pptx --ref-audio example/kuwano.wav --ref-text example/kuwano.txt --languages EN JP ZH ES IT FR
```

## Fill missing multilingual note tags from JP (CLI)

Use local Ollama (`translategemma:4b`) to translate missing note blocks for `EN/ZH/ES/IT/FR` from `[JP]... [JP]` text on each slide.
The CLI first requests all missing tags in one JSON response while passing existing non-JP note blocks as terminology references, then falls back per tag only for missing/invalid outputs.

```powershell
python -m app.pipeline.translate_notes_multi <deck>.pptx
```

The command writes `<deck>_multi.pptx` in the same directory and overwrites it if it already exists.

## Cache behavior (Step 2)

Step 2 uses an in-process audio cache (memory only).
Cache entries are reused only while the app/process is running and are cleared on restart.
WAV regeneration is skipped when the cache key matches:

- script SHA256
- language + Qwen language assignment
- `temperature` / `top_p` / `top_k` / `repetition_penalty`
- model/device/dtype settings
- reference audio SHA256
- reference text SHA256

Use `--force-regenerate` in CLI (or checkbox in UI) to bypass cache.

When JP/EN/ZH are all active, Step 2 also retries once if a generated WAV is detected as anomalous (too long, too short, or effectively silent).

## Linux / Colab note

Step 1 uses PowerPoint COM and is Windows-only.
On Linux/Colab, upload a `slides.zip` containing `page1.png`, `page2.png`, ... and an `audio.zip` containing `en/`, `jp/`, `zh/`, `es/`, `it/`, `fr/` WAV folders in Step 3.

## Verification

```powershell
python -m compileall app
python -c "import app.web.gradio_app; print('web import ok')"
python -c "from qwen_tts import Qwen3TTSModel; print('qwen_tts ok')"
```
