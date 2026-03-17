# PowerPoint TTS Movie Builder

PowerPoint TTS Movie Builder turns a tagged `.pptx` deck into narrated MP4 files.
It extracts note text by language tag, generates WAV narration with Qwen TTS,
exports slide images through PowerPoint COM, and muxes the result with `ffmpeg`.

## Project layout

- `app/common/`: shared dataclasses, note parsing, language resolution, and workspace helpers
- `app/models/`: voice asset lookup and prompt cache export
- `app/synthesis/`: Qwen model loading, prompt creation, and audio generation
- `app/slides/`: PowerPoint COM slide export
- `app/video/`: WAV concatenation and MP4 assembly
- `app/gui/`: Tkinter desktop app
- `requirements/`: pip requirement fragments for CPU and CUDA environments

The end-to-end pipeline entrypoint lives in `app/pipeline/build_movie.py`.

## Expected inputs

- A PowerPoint deck with note tags such as `[JA]... [JA]`, `[EN]... [EN]`, and `[ZH]... [ZH]`
- Reference voice assets in `wav + txt` form, such as `kuwano.wav` and `kuwano.txt`

## Environment setup

Python 3.12 is pinned in `.python-version` and the checked-in Conda files.

- `conda.env.yml`: CPU-first environment
- `conda.cuda.yml`: CUDA environment

Create the environment you want with:

```powershell
conda env create -f conda.env.yml
conda activate qwentts
```

Or, for CUDA:

```powershell
conda env create -f conda.cuda.yml
conda activate qwentts-cuda
```

The environments include `cffi`, `pillow`, and `charset-normalizer`. If `_cffi_backend` or `PIL.Image` is missing, recreate the environment or reinstall the missing package.

## CLI usage

Build narrated videos with a named voice:

```powershell
python -m app.pipeline.build_movie <deck>.pptx --voice-name kuwano --languages JA EN ZH
```

Build with explicit reference files:

```powershell
python -m app.pipeline.build_movie <deck>.pptx --ref-audio example/kuwano.wav --ref-text example/kuwano.txt --languages JA EN ZH
```

Export a reusable prompt cache:

```powershell
python -m app.models.export_voice --voice-name kuwano
```

Outputs are written to:

- `work/slides/pageN.png`
- `work/audio/ja/pageN.wav`
- `work/audio/en/pageN.wav`
- `work/audio/zh/pageN.wav`
- `output/ja.mp4`
- `output/en.mp4`
- `output/zh.mp4`

## GUI usage

Launch the Tkinter app with:

```powershell
python -m app.gui
```

The GUI auto-fills the first `.pptx` in the project root and the first matching `wav + txt` pair when it finds them.

Typical workflow:

1. Choose the `PPTX` deck.
2. Set `Ref Audio` and `Ref Text` for the voice you want to clone.
3. Pick languages, model size, device, slide size, padding, and FPS.
4. Click `Build Videos`.
5. Use `Open Output Folder` to open `output/`.

Leave `Device` blank for auto-detection, or enter values like `cuda:0` or `cpu`. If CUDA probing or model load fails, the app falls back to CPU automatically.

The GUI runs heavy work on a background thread and builds the voice prompt directly from the selected reference audio and transcript instead of using `.pt` prompt cache files.

## Verification

There is no checked-in test suite today. The practical verification baseline is:

```powershell
python -m compileall app
python -c "import app.gui.main; print('gui import ok')"
python -c "from qwen_tts import Qwen3TTSModel; print('qwen_tts ok')"
```

If model import fails, verify Pillow separately:

```powershell
python -c "import PIL.Image; print('pillow ok')"
```

## Notes

- PowerPoint slide export depends on Microsoft PowerPoint COM, so the main workflow targets Windows.
- `python-pptx`, `opencv`, and `pydub` are not required for the current pipeline.
