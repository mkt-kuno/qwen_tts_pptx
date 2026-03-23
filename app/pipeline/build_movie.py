from __future__ import annotations

import argparse
import logging
from pathlib import Path

from app.pipeline.steps import (
    step1_export_images,
    step2_synthesize_audio,
    step3_build_videos,
)


logger = logging.getLogger(__name__)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run step1->step2->step3 and build MP4 outputs from a PPTX deck."
    )
    parser.add_argument("pptx", type=Path, help="Input .pptx path")
    parser.add_argument(
        "--ref-audio", type=Path, required=True, help="Reference WAV path"
    )
    parser.add_argument(
        "--ref-text", type=Path, required=True, help="Reference transcript path"
    )
    parser.add_argument(
        "--languages",
        nargs="+",
        default=["EN", "JP", "ZH", "ES", "IT", "FR"],
    )
    parser.add_argument("--model-size", choices=["0.6B", "1.7B"], default="1.7B")
    parser.add_argument("--device", type=str, default=None)
    parser.add_argument(
        "--dtype",
        type=str,
        default="auto",
        help="Torch dtype for model load: auto, bfloat16, float16, or float32",
    )
    parser.add_argument("--slide-width", type=int, default=2560)
    parser.add_argument("--slide-height", type=int, default=1440)
    parser.add_argument("--slide-padding-sec", type=float, default=1.5)
    parser.add_argument("--fps", type=int, default=5)
    parser.add_argument("--force-regenerate", action="store_true")
    return parser


def main() -> None:
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s"
    )
    args = build_arg_parser().parse_args()
    project_root = args.pptx.parent.resolve()

    step1_result = step1_export_images(
        pptx_path=args.pptx,
        project_root=project_root,
        slide_width=args.slide_width,
        slide_height=args.slide_height,
    )
    logger.info("Step 1 done: slide_count=%d", step1_result.slide_count)

    step2_result = step2_synthesize_audio(
        pptx_path=args.pptx,
        project_root=project_root,
        ref_audio=args.ref_audio,
        ref_text=args.ref_text,
        languages=args.languages,
        model_size=args.model_size,
        device=args.device,
        dtype=args.dtype,
        force_regenerate=args.force_regenerate,
    )
    logger.info(
        "Step 2 done: generated=%d cache_hits=%d",
        step2_result.generated_count,
        step2_result.cache_hit_count,
    )

    step3_result = step3_build_videos(
        project_root=project_root,
        languages=args.languages,
        slide_padding_sec=args.slide_padding_sec,
        fps=args.fps,
    )
    outputs = [*step3_result.per_language_outputs]
    if step3_result.multilingual_output is not None:
        outputs.append(step3_result.multilingual_output)

    for output in outputs:
        logger.info("Created %s", output)


if __name__ == "__main__":
    main()
