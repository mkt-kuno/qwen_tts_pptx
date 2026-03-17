from __future__ import annotations

import subprocess
import textwrap
from pathlib import Path


POWERSHELL_SCRIPT = textwrap.dedent(
    """
    param(
        [Parameter(Mandatory=$true)][string]$InputPath,
        [Parameter(Mandatory=$true)][string]$OutputDir,
        [Parameter(Mandatory=$true)][int]$Width,
        [Parameter(Mandatory=$true)][int]$Height
    )

    $ErrorActionPreference = "Stop"
    $powerpoint = $null
    $presentation = $null

    try {
        New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
        $powerpoint = New-Object -ComObject PowerPoint.Application
        $presentation = $powerpoint.Presentations.Open($InputPath, $true, $false, $false)

        foreach ($slide in $presentation.Slides) {
            $index = [int]$slide.SlideIndex
            $target = Join-Path $OutputDir ("page{0}.png" -f $index)
            $slide.Export($target, "PNG", $Width, $Height)
        }

        [Console]::WriteLine($presentation.Slides.Count)
    }
    finally {
        if ($presentation -ne $null) {
            $presentation.Close()
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation)
        }
        if ($powerpoint -ne $null) {
            $powerpoint.Quit()
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerpoint)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    """
).strip()


def export_slides(
    pptx_path: Path,
    output_dir: Path,
    temp_dir: Path,
    width: int,
    height: int,
) -> int:
    if not pptx_path.exists():
        raise FileNotFoundError(f"PPTX not found: {pptx_path}")

    output_dir.mkdir(parents=True, exist_ok=True)
    temp_dir.mkdir(parents=True, exist_ok=True)
    cleanup_previous_slides(output_dir)

    script_path = temp_dir / "export_slides.ps1"
    script_path.write_text(POWERSHELL_SCRIPT, encoding="utf-8-sig")

    completed = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_path),
            "-InputPath",
            str(pptx_path.resolve()),
            "-OutputDir",
            str(output_dir.resolve()),
            "-Width",
            str(width),
            "-Height",
            str(height),
        ],
        check=True,
        capture_output=True,
        text=True,
    )

    stdout = completed.stdout.strip().splitlines()
    if not stdout:
        raise RuntimeError("PowerPoint export did not return a slide count.")
    return int(stdout[-1].strip())


def cleanup_previous_slides(output_dir: Path) -> None:
    for path in output_dir.glob("page*.png"):
        path.unlink()
