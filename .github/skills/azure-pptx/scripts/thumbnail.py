#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.14"
# dependencies = []
# ///

"""Render PPTX slides as JPEG files with LibreOffice and Poppler."""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import tempfile
from pathlib import Path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert each slide in a PPTX file to a JPEG image."
    )
    parser.add_argument("presentation", type=Path, help="Input .pptx file")
    parser.add_argument(
        "--outdir",
        type=Path,
        help="Output directory (default: <presentation>-slides)",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=150,
        help="JPEG resolution in DPI (default: 150)",
    )
    parser.add_argument(
        "--prefix",
        default="slide",
        help="Output filename prefix (default: slide)",
    )
    parser.add_argument("--first", type=int, help="First slide number, starting at 1")
    parser.add_argument("--last", type=int, help="Last slide number, starting at 1")
    parser.add_argument(
        "--timeout",
        type=int,
        default=120,
        help="Timeout in seconds for each external command (default: 120)",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Replace existing JPEG files with the same prefix",
    )
    return parser


def find_libreoffice() -> str:
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if executable is None:
        raise ValueError(
            "LibreOffice was not found; install it and expose soffice or "
            "libreoffice on PATH"
        )
    return executable


def find_pdftoppm() -> str:
    executable = shutil.which("pdftoppm")
    if executable is None:
        raise ValueError(
            "pdftoppm was not found; install Poppler and expose pdftoppm on PATH"
        )
    return executable


def validate_args(args: argparse.Namespace) -> tuple[Path, Path]:
    presentation = args.presentation.expanduser()
    if not presentation.is_file():
        raise ValueError(f"input file does not exist: {presentation}")
    if presentation.suffix.lower() != ".pptx":
        raise ValueError("input extension must be .pptx")
    if args.dpi <= 0:
        raise ValueError("--dpi must be greater than zero")
    if args.timeout <= 0:
        raise ValueError("--timeout must be greater than zero")
    if args.first is not None and args.first < 1:
        raise ValueError("--first must be at least 1")
    if args.last is not None and args.last < 1:
        raise ValueError("--last must be at least 1")
    if args.first is not None and args.last is not None and args.first > args.last:
        raise ValueError("--first must not be greater than --last")
    if not args.prefix or Path(args.prefix).name != args.prefix:
        raise ValueError("--prefix must be a filename component, not a path")

    output_dir = (
        args.outdir.expanduser()
        if args.outdir is not None
        else presentation.parent / f"{presentation.stem}-slides"
    )
    return presentation, output_dir


def command_output(process: subprocess.CompletedProcess[str]) -> str:
    sections = []
    if process.stdout.strip():
        sections.append(f"stdout:\n{process.stdout.strip()}")
    if process.stderr.strip():
        sections.append(f"stderr:\n{process.stderr.strip()}")
    return "\n".join(sections) or "no command output"


def run_command(command: list[str], timeout: int) -> subprocess.CompletedProcess[str]:
    process = subprocess.run(
        command,
        capture_output=True,
        check=False,
        text=True,
        timeout=timeout,
    )
    if process.returncode != 0:
        raise ValueError(
            f"command failed with exit code {process.returncode}: "
            f"{' '.join(command)}\n{command_output(process)}"
        )
    return process


def page_number(path: Path) -> int:
    match = re.search(r"-(\d+)\.jpg$", path.name)
    return int(match.group(1)) if match else 0


def is_requested_page(path: Path, first: int | None, last: int | None) -> bool:
    number = page_number(path)
    return (first is None or number >= first) and (last is None or number <= last)


def remove_existing_images(
    output_dir: Path,
    prefix: str,
    force: bool,
    first: int | None,
    last: int | None,
) -> None:
    existing = [
        image
        for image in output_dir.glob(f"{prefix}-*.jpg")
        if is_requested_page(image, first, last)
    ]
    if existing and not force:
        raise ValueError(
            f"{len(existing)} matching JPEG file(s) already exist in {output_dir}; "
            "use --force to replace them"
        )
    for image in existing:
        image.unlink()


def render_slides(args: argparse.Namespace) -> list[Path]:
    presentation, output_dir = validate_args(args)
    libreoffice = find_libreoffice()
    pdftoppm = find_pdftoppm()

    output_dir.mkdir(parents=True, exist_ok=True)
    remove_existing_images(
        output_dir,
        args.prefix,
        args.force,
        args.first,
        args.last,
    )

    with tempfile.TemporaryDirectory(prefix="azure-pptx-render-") as temp_dir:
        temporary_root = Path(temp_dir)
        pdf_dir = temporary_root / "pdf"
        profile_dir = temporary_root / "libreoffice-profile"
        pdf_dir.mkdir()
        profile_dir.mkdir()

        conversion = run_command(
            [
                libreoffice,
                "--headless",
                f"-env:UserInstallation={profile_dir.as_uri()}",
                "--convert-to",
                "pdf:impress_pdf_Export",
                "--outdir",
                str(pdf_dir),
                str(presentation.resolve()),
            ],
            args.timeout,
        )

        pdf_path = pdf_dir / f"{presentation.stem}.pdf"
        if not pdf_path.is_file() or pdf_path.stat().st_size == 0:
            raise ValueError(
                "LibreOffice did not produce a non-empty PDF\n"
                f"{command_output(conversion)}"
            )

        command = [
            pdftoppm,
            "-jpeg",
            "-r",
            str(args.dpi),
            "-sep",
            "-",
            "-forcenum",
        ]
        if args.first is not None:
            command.extend(["-f", str(args.first)])
        if args.last is not None:
            command.extend(["-l", str(args.last)])
        command.extend(
            [
                str(pdf_path),
                str(output_dir.resolve() / args.prefix),
            ]
        )
        run_command(command, args.timeout)

    images = sorted(
        (
            image
            for image in output_dir.glob(f"{args.prefix}-*.jpg")
            if is_requested_page(image, args.first, args.last)
        ),
        key=page_number,
    )
    if not images:
        raise ValueError("pdftoppm did not produce any JPEG files")
    if any(image.stat().st_size == 0 for image in images):
        raise ValueError("pdftoppm produced an empty JPEG file")
    return images


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        images = render_slides(args)
    except (OSError, subprocess.TimeoutExpired, ValueError) as exc:
        parser.exit(1, f"error: {exc}\n")

    for image in images:
        print(image)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
