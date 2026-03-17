from __future__ import annotations

import logging
import subprocess
import sys
import wave
from dataclasses import dataclass
from pathlib import Path

import soundfile as sf
import torch


logger = logging.getLogger(__name__)

CUDA_READY_PROBE = "\n".join(
    [
        "import sys",
        "import torch",
        "device = sys.argv[1]",
        "if not torch.cuda.is_available():",
        "    print('0')",
        "    raise SystemExit(0)",
        "index = 0 if ':' not in device else int(device.split(':', 1)[1])",
        "torch.cuda.set_device(index)",
        "torch.cuda.mem_get_info(index)",
        "print('1')",
    ]
)

MODEL_NAMES: dict[str, str] = {
    "0.6B": "Qwen/Qwen3-TTS-12Hz-0.6B-Base",
    "1.7B": "Qwen/Qwen3-TTS-12Hz-1.7B-Base",
}

DTYPE_ALIASES: dict[str, torch.dtype] = {
    "bf16": torch.bfloat16,
    "bfloat16": torch.bfloat16,
    "fp16": torch.float16,
    "float16": torch.float16,
    "fp32": torch.float32,
    "float32": torch.float32,
}

CUDA_FALLBACK_MESSAGES = (
    "CUDA error: operation not supported",
    "Found no NVIDIA driver",
    "CUDA driver version is insufficient",
    "no kernel image is available for execution on the device",
)


@dataclass(frozen=True)
class LoadedModel:
    model: object
    device: str


def detect_device(requested: str | None) -> str:
    if requested:
        requested = requested.strip()
    if requested:
        if _is_cuda_device(requested):
            logger.info("Validating requested device %s", requested)
            if _probe_cuda_available(requested):
                logger.info("Using requested device %s", requested)
                return requested
            logger.warning(
                "Requested device %s is unavailable or could not initialize. Falling back to CPU.",
                requested,
            )
            return "cpu"
        logger.info("Using requested device %s", requested)
        return requested
    logger.info("No device requested, probing GPU availability")
    if _probe_cuda_available("cuda:0"):
        logger.info("GPU probe succeeded, selecting cuda:0")
        return "cuda:0"
    logger.warning("GPU not available. Falling back to CPU.")
    return "cpu"


def _is_cuda_device(device: str) -> bool:
    return device.lower().startswith("cuda")


def _probe_cuda_available(device: str) -> bool:
    logger.info("Starting isolated torch GPU probe for %s", device)
    try:
        probe = subprocess.run(
            [
                sys.executable,
                "-c",
                CUDA_READY_PROBE,
                device,
            ],
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )
    except subprocess.TimeoutExpired:
        logger.warning("GPU probe timed out for %s, using CPU.", device)
        return False
    if probe.returncode != 0:
        stderr = probe.stderr.strip()
        if stderr:
            logger.warning("GPU probe failed for %s, using CPU: %s", device, stderr)
        else:
            logger.warning("GPU probe failed for %s, using CPU.", device)
        return False
    stdout = probe.stdout.strip()
    logger.info("GPU probe for %s completed with output '%s'", device, stdout)
    return stdout.endswith("1")


def resolve_dtype(requested: str | None, device: str) -> torch.dtype:
    choice = (requested or "auto").strip().lower()
    if choice in ("", "auto"):
        return torch.bfloat16 if device.startswith("cuda") else torch.float32
    try:
        return DTYPE_ALIASES[choice]
    except KeyError as exc:
        supported = ", ".join(["auto", *sorted(DTYPE_ALIASES)])
        raise ValueError(
            f"Unsupported torch dtype '{requested}'. Use one of: {supported}"
        ) from exc


def _dtype_name(dtype: torch.dtype) -> str:
    for name, value in DTYPE_ALIASES.items():
        if value == dtype and not name.startswith("fp"):
            return name
    return str(dtype)


def _load_pretrained_model(
    model_class, model_name: str, device: str, dtype: str | None
):
    resolved_dtype = resolve_dtype(dtype, device)
    logger.info(
        "Loading %s on %s (%s)", model_name, device, _dtype_name(resolved_dtype)
    )
    return model_class.from_pretrained(
        model_name,
        device_map=device,
        dtype=resolved_dtype,
    )


def _should_retry_on_cpu(device: str, exc: RuntimeError) -> bool:
    if not _is_cuda_device(device):
        return False
    message = str(exc)
    return any(fragment in message for fragment in CUDA_FALLBACK_MESSAGES)


def load_model(model_size: str, device: str, dtype: str | None = None) -> LoadedModel:
    logger.info("Importing qwen_tts runtime")
    from qwen_tts import Qwen3TTSModel

    model_name = MODEL_NAMES[model_size]
    try:
        model = _load_pretrained_model(
            Qwen3TTSModel,
            model_name=model_name,
            device=device,
            dtype=dtype,
        )
        logger.info("Loaded Qwen TTS runtime for %s", model_name)
        return LoadedModel(model=model, device=device)
    except RuntimeError as exc:
        if not _should_retry_on_cpu(device, exc):
            raise
        logger.warning(
            "Loading %s on %s failed with a CUDA runtime error: %s",
            model_name,
            device,
            exc,
        )
        logger.warning("Retrying %s on cpu (float32).", model_name)
        model = _load_pretrained_model(
            Qwen3TTSModel,
            model_name=model_name,
            device="cpu",
            dtype="float32",
        )
        logger.info(
            "Loaded Qwen TTS runtime for %s on cpu after CUDA fallback", model_name
        )
        return LoadedModel(model=model, device="cpu")


def create_voice_clone_prompt(model, ref_audio: Path, ref_text: str):
    logger.info("Creating prompt from %s", ref_audio)
    prompt = model.create_voice_clone_prompt(
        ref_audio=str(ref_audio), ref_text=ref_text
    )
    logger.info("Prompt creation finished for %s", ref_audio)
    return prompt


def save_prompt_cache(
    model,
    voice_name: str,
    model_size: str,
    ref_audio: Path,
    ref_text: str,
    output_path: Path,
):
    prompt = create_voice_clone_prompt(
        model=model, ref_audio=ref_audio, ref_text=ref_text
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    torch.save(
        {
            "voice_name": voice_name,
            "model_size": model_size,
            "ref_audio": str(ref_audio),
            "ref_text": ref_text,
            "prompt": prompt,
        },
        output_path,
    )
    return prompt


def load_prompt_cache(cache_path: Path):
    bundle = torch.load(cache_path, map_location="cpu", weights_only=False)
    return bundle["prompt"]


def write_silence_wav(
    path: Path, duration_sec: float = 0.5, sample_rate: int = 24000
) -> None:
    frames = int(duration_sec * sample_rate)
    with wave.open(str(path), "wb") as wav_file:
        wav_file.setnchannels(1)
        wav_file.setsampwidth(2)
        wav_file.setframerate(sample_rate)
        wav_file.writeframes(b"\x00\x00" * frames)


def synthesize_to_file(
    model, prompt, text: str, language: str, output_path: Path
) -> None:
    if not text.strip():
        write_silence_wav(output_path)
        return

    try:
        wavs, sample_rate = model.generate_voice_clone(
            text=text,
            language=language,
            voice_clone_prompt=prompt,
        )
    except torch.cuda.OutOfMemoryError:
        logger.error("GPU out of memory. Try --model-size 0.6B")
        sys.exit(1)
    sf.write(str(output_path), wavs[0], sample_rate)
