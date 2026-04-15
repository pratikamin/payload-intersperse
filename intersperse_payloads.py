#!/usr/bin/env python3
"""
intersperse_payloads.py

Randomly intersperse payload lines from a plain text file into DOCX and CSV files.

For CSV targets:  each payload line is inserted as a new row (placed in the first
                  column; remaining columns are left empty).

For DOCX targets: each payload line is inserted as a paragraph at a random position.

Usage:
    python intersperse_payloads.py \\
        --input-folder /path/to/files \\
        --payload-file /path/to/payloads.txt \\
        --output-folder /path/to/output \\
        [--count N] \\
        [--seed 42]

    # Generate a sample payload text file to get started:
    python intersperse_payloads.py --generate-sample-payload sample_payloads.txt
"""

import argparse
import csv
import os
import random
import shutil
from pathlib import Path
from typing import List, Optional

try:
    from docx import Document
    from docx.oxml import OxmlElement
except ImportError:
    raise SystemExit("Missing dependency: pip install python-docx")


# ---------------------------------------------------------------------------
# Argument parsing
# ---------------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(
        description="Randomly intersperse payload lines into DOCX and CSV files.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--input-folder",
        help="Folder containing .docx and .csv files to modify.",
    )
    parser.add_argument(
        "--payload-file",
        help="Plain text file with one payload per line.",
    )
    parser.add_argument(
        "--output-folder",
        default="./output",
        help="Destination folder for modified files (default: ./output).",
    )
    parser.add_argument(
        "--count",
        type=int,
        default=None,
        help="Number of payload lines to inject per file. "
             "Defaults to a random value between 3 and 10.",
    )
    parser.add_argument(
        "--seed",
        type=int,
        default=None,
        help="Random seed for reproducible output.",
    )
    parser.add_argument(
        "--generate-sample-payload",
        metavar="OUTPUT_TXT",
        help="Write a sample payload text file to the given path and exit.",
    )
    return parser.parse_args()


# ---------------------------------------------------------------------------
# Payload utilities
# ---------------------------------------------------------------------------

def load_payload_file(path: str) -> List[str]:
    """Load a plain text file as a list of payloads.

    Blocks of text separated by one or more blank lines are treated as
    individual payload entries, so a file with no blank lines is a single
    payload regardless of how many lines it contains.
    """
    with open(path, encoding="utf-8") as fh:
        content = fh.read()

    blocks = [b.strip() for b in content.split("\n\n") if b.strip()]
    return blocks


def pick_payloads(payloads: List[str], count: Optional[int], rng: random.Random) -> List[str]:
    """Choose *count* payload entries (with replacement) to inject."""
    n = count if count is not None else rng.randint(3, 10)
    return rng.choices(payloads, k=n)


# ---------------------------------------------------------------------------
# CSV processing
# ---------------------------------------------------------------------------

def intersperse_csv(
    input_path: str,
    output_path: str,
    payloads: List[str],
    count: Optional[int],
    rng: random.Random,
) -> None:
    """Insert payload lines as rows at random positions within a CSV file."""
    with open(input_path, newline="", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        fieldnames: List[str] = list(reader.fieldnames or [])
        data_rows = list(reader)

    selected = pick_payloads(payloads, count, rng)

    # Build injected rows: payload text in first column, rest empty.
    injected = []
    for text in selected:
        row = {col: "" for col in fieldnames}
        if fieldnames:
            row[fieldnames[0]] = text
        injected.append(row)

    result = list(data_rows)
    for row in injected:
        pos = rng.randint(0, len(result))
        result.insert(pos, row)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(result)

    print(f"  [CSV]  Injected {len(selected)} rows  -> {output_path}")


# ---------------------------------------------------------------------------
# DOCX processing
# ---------------------------------------------------------------------------

def _make_paragraph_element(text: str):
    """Build a minimal w:p XML element containing the given text."""
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    p.append(r)
    return p


def intersperse_docx(
    input_path: str,
    output_path: str,
    payloads: List[str],
    count: Optional[int],
    rng: random.Random,
) -> None:
    """Insert payload lines as paragraphs at random positions within a DOCX file."""
    doc = Document(input_path)

    num_paras = len(doc.paragraphs)
    if num_paras == 0:
        print(f"  [DOCX] No paragraphs found, copying unchanged -> {output_path}")
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        shutil.copy2(input_path, output_path)
        return

    selected = pick_payloads(payloads, count, rng)

    # Process from highest position to lowest so earlier indices stay stable.
    positions = sorted(
        [rng.randint(0, num_paras - 1) for _ in range(len(selected))],
        reverse=True,
    )

    for text, pos in zip(selected, positions):
        target_para_elem = doc.paragraphs[pos]._element
        new_p = _make_paragraph_element(text)
        target_para_elem.addprevious(new_p)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc.save(output_path)
    print(f"  [DOCX] Injected {len(selected)} paragraphs -> {output_path}")


# ---------------------------------------------------------------------------
# Sample payload generator
# ---------------------------------------------------------------------------

SAMPLE_PAYLOADS = [
    "Unauthorised login attempt detected from external IP",
    "Sensitive file accessed outside business hours",
    "Unexpected outbound connection to unknown endpoint",
    "MFA bypass policy exception granted temporarily",
    "Security group rule opened port 22 to 0.0.0.0/0",
]


def generate_sample_payload(output_path: str) -> None:
    with open(output_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(SAMPLE_PAYLOADS) + "\n")
    print(f"Sample payload file written to: {output_path}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    args = parse_args()

    if args.generate_sample_payload:
        generate_sample_payload(args.generate_sample_payload)
        return

    if not args.input_folder or not args.payload_file:
        raise SystemExit(
            "Error: --input-folder and --payload-file are required.\n"
            "Run with --help for usage details."
        )

    rng = random.Random(args.seed)

    input_folder = Path(args.input_folder)
    output_folder = Path(args.output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)

    print(f"Loading payloads from: {args.payload_file}")
    payloads = load_payload_file(args.payload_file)
    if not payloads:
        raise SystemExit("Error: payload file is empty.")
    print(f"  {len(payloads)} payload line(s) loaded.\n")

    all_files = sorted(input_folder.iterdir())
    csv_files = [f for f in all_files if f.suffix.lower() == ".csv"]
    docx_files = [f for f in all_files if f.suffix.lower() == ".docx"]

    print(
        f"Found {len(docx_files)} DOCX and {len(csv_files)} CSV file(s) "
        f"in: {input_folder}\n"
    )

    errors = []

    for f in csv_files:
        out = output_folder / f.name
        try:
            intersperse_csv(str(f), str(out), payloads, args.count, rng)
        except Exception as exc:
            errors.append((f.name, exc))
            print(f"  [CSV]  ERROR processing {f.name}: {exc}")

    for f in docx_files:
        out = output_folder / f.name
        try:
            intersperse_docx(str(f), str(out), payloads, args.count, rng)
        except Exception as exc:
            errors.append((f.name, exc))
            print(f"  [DOCX] ERROR processing {f.name}: {exc}")

    print(f"\nDone. Output written to: {output_folder}")
    if errors:
        print(f"  {len(errors)} file(s) had errors (see above).")


if __name__ == "__main__":
    main()
