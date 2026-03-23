from __future__ import annotations

import logging
import os
import queue
import sys
import threading
from collections.abc import Iterator
from pathlib import Path

import gradio as gr

from app.common.paths import WorkspacePaths
from app.pipeline.steps import (
    AudioSynthesisResult,
    Step2ProgressUpdate,
    export_zip_from_dir,
    import_audio_zip,
    import_slide_zip,
    step1_export_images,
    step2_synthesize_audio,
    step3_build_videos,
)


LOG_FORMAT = "%(asctime)s %(levelname)s %(message)s"
logger = logging.getLogger(__name__)
LANGUAGE_CHOICES = ["EN", "JP", "ZH", "ES", "IT", "FR"]
STEP3_OUTPUT_CHOICES = ["MULTI", *LANGUAGE_CHOICES]


def _project_root_from_pptx(pptx_path: Path) -> Path:
    return pptx_path.parent.resolve()


def _to_path(file_path: str | None, label: str) -> Path:
    if not file_path:
        raise ValueError(f"{label} is required.")
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"{label} not found: {path}")
    return path


def run_step1(
    pptx_file: str | None,
    slide_width: int,
    slide_height: int,
) -> tuple[str, str | None]:
    pptx_path = _to_path(pptx_file, "PPTX")
    project_root = _project_root_from_pptx(pptx_path)
    result = step1_export_images(
        pptx_path=pptx_path,
        project_root=project_root,
        slide_width=slide_width,
        slide_height=slide_height,
    )
    paths = WorkspacePaths.from_root(project_root)
    zip_path = export_zip_from_dir(
        source_dir=result.slides_dir,
        output_zip=paths.output / "slides.zip",
    )
    return (
        f"Exported {result.slide_count} slides to {result.slides_dir}",
        str(zip_path),
    )


def run_step2(
    pptx_file: str | None,
    ref_audio_file: str | None,
    ref_text_file: str | None,
    languages: list[str],
    model_size: str,
    device: str,
    dtype: str,
    force_regenerate: bool,
) -> Iterator[tuple[str, str | None]]:
    pptx_path = _to_path(pptx_file, "PPTX")
    ref_audio = _to_path(ref_audio_file, "Ref Audio")
    ref_text = _to_path(ref_text_file, "Ref Text")
    if not languages:
        raise ValueError("Select at least one language.")

    project_root = _project_root_from_pptx(pptx_path)

    progress_updates: queue.Queue[Step2ProgressUpdate] = queue.Queue()
    result_holder: dict[str, AudioSynthesisResult] = {}
    error_holder: dict[str, Exception] = {}

    def _on_progress(update: Step2ProgressUpdate) -> None:
        progress_updates.put(update)

    def _worker() -> None:
        try:
            result_holder["result"] = step2_synthesize_audio(
                pptx_path=pptx_path,
                project_root=project_root,
                ref_audio=ref_audio,
                ref_text=ref_text,
                languages=languages,
                model_size=model_size,
                device=device.strip() or None,
                dtype=dtype,
                force_regenerate=force_regenerate,
                progress_callback=_on_progress,
            )
        except Exception as exc:  # pragma: no cover - forwarded to UI
            error_holder["error"] = exc

    worker = threading.Thread(target=_worker, daemon=True)
    worker.start()

    yield ("Preparing Step 2 audio synthesis...", None)
    while worker.is_alive() or not progress_updates.empty():
        try:
            update = progress_updates.get(timeout=0.2)
        except queue.Empty:
            continue
        yield (_format_step2_progress_status(update), None)

    worker.join()
    error = error_holder.get("error")
    if error is not None:
        raise error

    result = result_holder.get("result")
    if result is None:
        raise RuntimeError("Step 2 completed without returning a result.")

    paths = WorkspacePaths.from_root(project_root)
    zip_path = export_zip_from_dir(
        source_dir=paths.audio,
        output_zip=paths.output / "audio.zip",
    )
    final_total = result.slide_count * len(languages)
    final_status = (
        f"{{{final_total}/{final_total}}} Audio synthesis complete: "
        f"slides={result.slide_count}, "
        f"generated={result.generated_count}, "
        f"cache_hits={result.cache_hit_count}"
    )
    yield (
        final_status,
        str(zip_path),
    )


def _format_step2_progress_status(update: Step2ProgressUpdate) -> str:
    return (
        f"{{{update.completed_units}/{update.total_units}}} "
        f"Processing slide {update.slide_number} [{update.language_tag}]"
    )


def run_step3(
    pptx_file: str | None,
    slides_zip_file: str | None,
    audio_zip_file: str | None,
    languages: list[str],
    output_formats: list[str],
    slide_padding_sec: float,
    fps: int,
) -> tuple[str, list[str]]:
    pptx_path = _to_path(pptx_file, "PPTX")
    if not languages:
        raise ValueError("Select at least one language.")
    if not output_formats:
        raise ValueError("Select at least one output format.")

    include_multilingual = "MULTI" in output_formats
    output_language_tags = [tag for tag in output_formats if tag != "MULTI"]
    invalid_output_languages = sorted(
        tag for tag in output_language_tags if tag not in LANGUAGE_CHOICES
    )
    if invalid_output_languages:
        raise ValueError(
            "Unsupported output language tags: " + ", ".join(invalid_output_languages)
        )

    project_root = _project_root_from_pptx(pptx_path)

    if slides_zip_file:
        import_slide_zip(zip_path=Path(slides_zip_file), project_root=project_root)
    if audio_zip_file:
        import_audio_zip(zip_path=Path(audio_zip_file), project_root=project_root)

    result = step3_build_videos(
        project_root=project_root,
        languages=languages,
        slide_padding_sec=slide_padding_sec,
        fps=fps,
        output_language_tags=output_language_tags,
        include_multilingual=include_multilingual,
    )
    outputs = [*result.per_language_outputs]
    if result.multilingual_output is not None:
        outputs.append(result.multilingual_output)
    return (
        f"Built {len(outputs)} videos in {WorkspacePaths.from_root(project_root).output}",
        [str(path) for path in outputs],
    )


def build_app(*, default_device: str = "", default_dtype: str = "auto") -> gr.Blocks:
    step1_enabled = sys.platform == "win32"
    with gr.Blocks(title="PowerPoint TTS Movie Builder") as demo:
        gr.Markdown(
            "# PowerPoint TTS Movie Builder\n"
            "PPTX is required for script extraction. Use the three-step workflow."
        )

        pptx_input = gr.File(label="PPTX", file_types=[".pptx"], type="filepath")

        with gr.Tab("Step 1: Image Export (Windows)", visible=step1_enabled):
            with gr.Row():
                slide_width = gr.Number(label="Slide Width", value=2560, precision=0)
                slide_height = gr.Number(label="Slide Height", value=1440, precision=0)
            step1_run = gr.Button("Run Step 1")
            step1_status = gr.Textbox(label="Status")
            step1_zip = gr.File(label="Slides ZIP")

        with gr.Tab("Step 2: Audio Synthesis"):
            ref_audio = gr.File(label="Ref Audio", type="filepath")
            ref_text = gr.File(label="Ref Text", file_types=[".txt"], type="filepath")
            with gr.Row():
                languages = gr.CheckboxGroup(
                    choices=LANGUAGE_CHOICES,
                    value=LANGUAGE_CHOICES,
                    label="Languages",
                )
                model_size = gr.Dropdown(
                    choices=["0.6B", "1.7B"], value="1.7B", label="Model Size"
                )
            with gr.Row():
                device = gr.Textbox(label="Device (blank=auto)", value=default_device)
                dtype = gr.Dropdown(
                    choices=["auto", "float16", "bfloat16", "float32"],
                    value=default_dtype,
                    label="DType",
                )
            force_regenerate = gr.Checkbox(
                label="Force regenerate all WAVs (ignore cache)", value=False
            )
            step2_run = gr.Button("Run Step 2")
            step2_status = gr.Textbox(label="Status")
            step2_zip = gr.File(label="Audio ZIP")

        with gr.Tab("Step 3: Video Build"):
            gr.Markdown(
                "Optional: upload ZIPs to import data. If omitted, existing work/ data is used."
            )
            slides_zip = gr.File(label="Slides ZIP (pageN.png)", type="filepath")
            audio_zip = gr.File(
                label="Audio ZIP (en/jp/zh/es/it/fr folders)", type="filepath"
            )
            output_formats = gr.CheckboxGroup(
                choices=STEP3_OUTPUT_CHOICES,
                value=["MULTI"],
                label="Output Videos",
            )
            with gr.Row():
                step3_padding = gr.Number(label="Padding Sec", value=1.5)
                step3_fps = gr.Number(label="FPS", value=5, precision=0)
            step3_run = gr.Button("Run Step 3")
            step3_status = gr.Textbox(label="Status")
            step3_outputs = gr.Files(label="MP4 Outputs")

        step1_run.click(
            fn=run_step1,
            inputs=[pptx_input, slide_width, slide_height],
            outputs=[step1_status, step1_zip],
        )
        step2_run.click(
            fn=run_step2,
            inputs=[
                pptx_input,
                ref_audio,
                ref_text,
                languages,
                model_size,
                device,
                dtype,
                force_regenerate,
            ],
            outputs=[step2_status, step2_zip],
        )
        step3_run.click(
            fn=run_step3,
            inputs=[
                pptx_input,
                slides_zip,
                audio_zip,
                languages,
                output_formats,
                step3_padding,
                step3_fps,
            ],
            outputs=[step3_status, step3_outputs],
        )

    return demo


def launch_for_colab() -> None:
    logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
    app = build_app(default_device="cuda:0", default_dtype="float16")
    app.queue(default_concurrency_limit=1)
    app.launch(share=True, debug=True)


def main() -> None:
    logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
    if os.environ.get("COLAB_RELEASE_TAG"):
        launch_for_colab()
        return
    app = build_app()
    app.queue(default_concurrency_limit=1)
    app.launch()
