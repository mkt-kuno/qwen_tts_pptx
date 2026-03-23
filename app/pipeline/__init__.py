from __future__ import annotations

from typing import Any


def step1_export_images(*args: Any, **kwargs: Any):
    from app.pipeline.steps import step1_export_images as _step1_export_images

    return _step1_export_images(*args, **kwargs)


def step2_synthesize_audio(*args: Any, **kwargs: Any):
    from app.pipeline.steps import step2_synthesize_audio as _step2_synthesize_audio

    return _step2_synthesize_audio(*args, **kwargs)


def step3_build_videos(*args: Any, **kwargs: Any):
    from app.pipeline.steps import step3_build_videos as _step3_build_videos

    return _step3_build_videos(*args, **kwargs)


__all__ = ["step1_export_images", "step2_synthesize_audio", "step3_build_videos"]
