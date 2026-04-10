"""Command-line entry point for wordtomd."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from wordtomd import __version__
from wordtomd.converter import DocxConverter


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="wordtomd",
        description="Convert .docx Word documents to Markdown (.md).",
    )
    parser.add_argument(
        "input",
        type=Path,
        help="Path to the input .docx file.",
    )
    parser.add_argument(
        "output",
        type=Path,
        nargs="?",
        default=None,
        help="Path to the output .md file. Defaults to <input>.md in the same directory.",
    )
    parser.add_argument(
        "--image-dir",
        metavar="NAME",
        default=None,
        help="Name of the subdirectory for extracted images (default: <output_stem>_images).",
    )
    parser.add_argument(
        "--no-images",
        action="store_true",
        default=False,
        help="Skip image extraction entirely.",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        default=False,
        help="Print progress information to stderr.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )

    args = parser.parse_args()

    # Validate input
    if not args.input.exists():
        print(f"Error: input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    if args.input.suffix.lower() != ".docx":
        print(f"Error: input must be a .docx file, got: {args.input.suffix}", file=sys.stderr)
        sys.exit(2)

    output = args.output or args.input.with_suffix(".md")

    DocxConverter(
        input_path=args.input,
        output_path=output,
        image_dir=args.image_dir,
        extract_images=not args.no_images,
        verbose=args.verbose,
    ).convert()

    print(f"Converted: {args.input} -> {output}")


if __name__ == "__main__":
    main()
