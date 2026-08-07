# /// script
# requires-python = ">=3.13"
# dependencies = [
#     "gspread>=6.2.1",
#     "gspread-formatting>=1.2.1",
#     "pypdf>=6.14.2",
#     "pyperclip>=1.11.0",
# ]
# ///
import re
import logging
import sys
import csv
import time

import pyperclip

from pypdf import PdfReader
from argparse import ArgumentParser
from pathlib import Path
from itertools import islice

logger = logging.getLogger(__name__)

TIMESTAMP = time.strftime("%Y%m%d-%H%M-%S")


def extract_data(
    p: Path,
    region_rename_csv,
    pagelim=3,
):
    """Extract turf data from the pdf

    This function takes in a path and reads the associated pdf.
    The generated pdfs usually store turf information on the first page, so
    to save time the function only reads that page.

    Specific pages can instead be set using the `pages` keyword.
    Setting pages=-1 will instead check EVERY page, which greatly increases
    processing time.

    Args:
        p:      path to the pdf input
        data:  optional list to add turf information to.
        pages:  pages in the pdf to check for turf information.


    Returns:
        data:  list of lists in the form [[full_name, name, location_type, ward,
               list_number, turf_number, num_doors]]
    """
    tr = re.compile(
        r"(?P<code>\d+-\d+)\s*(?P<turf>Turf \d+)\s*\d+\s*(?P<num_doors>\d+)"
    )
    pr = re.compile(
        r"^Turf Packet Summary.*?([^_\s]*_+[^_]*_+([^_]*?)(City|C|Village|villiage|Vge|V|Town|T])?_+(\d+[a-z]{0,1}).*)",
        re.IGNORECASE,
    )
    dr = re.compile(r".*(Village|Villiage|Vge|City|Town).*", re.IGNORECASE)
    logger.debug(f"{p.name}: extracting text from all pages")
    data = list()
    if region_rename_csv:
        logger.info(f"using region name overrides from {region_rename_csv.name}")
        with open(region_rename_csv, mode="r") as f:
            regions = {r[0]: r[1] for r in islice(csv.reader(f), 1, None) if r}
    else:
        regions = None
    if pagelim == -1:
        pages = enumerate(PdfReader(p).pages)
    else:
        pages = enumerate(PdfReader(p).pages[:pagelim])
    for i, page in pages:
        text = page.extract_text(extraction_mode="layout")
        logger.debug(f"{p.name}: first line\n{text.split('\n')[0]}")
        # metadata exists only on first page
        if i == 0:
            meta = list(next(pr.finditer(text)).groups())
            if regions and meta[0] in regions.keys():
                meta = list(next(pr.finditer(regions[meta[0]])).groups())
            assert len(meta) == 4
            logger.debug(f"{p.name}: metadata is {meta}")
            # normalize district type
            dists = {
                "city": "C",
                "town": "T",
                "village": "V",
                "villiage": "V",
                "c": "C",
                "t": "T",
                "v": "V",
                "Vge": "V",
                None: None,
            }
            meta[2] = (
                dists[meta[2].lower()]
                if meta[2]
                else list(next(dr.finditer(text)).groups())[0]
            )
            meta[3] = meta[3].zfill(4)  # left pad ward to 4 digits
        data.extend(meta + list(m.groups()) for m in tr.finditer(text))
        logger.debug(
            f"{p.name}: page {i} - {sum(1 for _ in tr.finditer(text))} matches"
        )
    return data


def post_process(data, district_rename_csv, priorities_csv):
    """Canonize names and apply post-processing for duplicate detection"""
    if district_rename_csv:
        with open(district_rename_csv, mode="r") as f:
            misnames = {r[0]: r[1] for r in islice(csv.reader(f), 1, None) if r}
    else:
        misnames = None
    if priorities_csv:
        with open(priorities_csv, mode="r") as f:
            prios = {r[0]: r[1] for r in islice(csv.reader(f), 1, None) if r}
    else:
        prios = None
    canonical = [
        misnames[r[1]] if misnames and r[1] in misnames.keys() else r[1]
        for r in data[1:]
    ]
    ward_turf = [f"{canonical[i]} - {r[2]} - {r[3]}" for i, r in enumerate(data[1:])]
    dup_finder = [f"{ward_turf[i]} - {r[5]}" for i, r in enumerate(data[1:])]
    seen = set()
    duplicates = set()
    for r in dup_finder:
        if r in seen:
            duplicates.add(r)
        else:
            seen.add(r)
    is_dup = ["DUPLICATE" if r in duplicates else None for r in dup_finder]
    new_cols = [
        ["misname_rename"] + canonical,
        ["priority"]
        + [
            prios[ward_turf[i]] if prios and ward_turf[i] in prios.keys() else None
            for i, r in enumerate(data[1:])
        ],
        ["ward_from_turf"] + ward_turf,
        ["list_number"] + [r[4] for r in data[1:]],
        ["turf_number"] + [r[5] for r in data[1:]],
        ["door_count"] + [r[6] for r in data[1:]],
        ["duplicate_finder"] + dup_finder,
        ["is_duplicate"] + is_dup,
    ]
    assert all(len(c) == len(data) for c in new_cols), "Column lengths must match"
    # transpose body to append new columns, then transpose back to rows to re-apply header
    return list(zip(*(list(zip(*data)) + new_cols)))


def update_sheet(url, data):
    """Update google sheet input"""
    import gspread  # avoid unnecessary authentication
    from gspread_formatting import (
        ConditionalFormatRule,
        GridRange,
        cellFormat,
        get_conditional_format_rules,
        BooleanRule,
        BooleanCondition,
        Color,
    )

    logger.debug(f"spreadsheet url: {url}")
    logger.debug("checking authorization")
    sh = gspread.oauth().open_by_url(url)
    logger.debug("uploading spreadsheet")
    w = sh.add_worksheet(
        title=f"turf_list_{TIMESTAMP}.csv",
        rows=len(data),
        cols=len(data[0]),
    )
    w.update(range_name="A1", values=data)
    logger.debug("uploaded spreadsheet")
    logger.debug("highlighting duplicates...")
    rules = get_conditional_format_rules(w)
    rules.append(
        ConditionalFormatRule(
            ranges=[GridRange.from_a1_range("A1:Z1000", w)],
            booleanRule=BooleanRule(
                condition=BooleanCondition("CUSTOM_FORMULA", ['=$O1="DUPLICATE"']),
                format=cellFormat(backgroundColor=Color(1, 1, 0)),
            ),
        )
    )
    rules.save()
    logger.info("formatting complete!")
    return


def main():
    parser = ArgumentParser()
    parser.add_argument(
        "input",
        nargs="*",
        type=Path,
        default=[Path.cwd()],
        help="Input turf pdf(s).",
    )
    parser.add_argument(
        "-r",
        "--remote",
        type=str,
        default=None,
        help="URL for google spreadsheet to update",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output file name. \nDefault: turf_list_YYYYmmdd-HHMM-SS.csv",
        default=Path(f"./turf_list_{TIMESTAMP}.csv"),
    )
    parser.add_argument(
        "-c",
        "--copy",
        action="store_true",
        help="Copy combined turf data output to system clipboard. \nDefault: False",
    )
    parser.add_argument(
        "--stdout",
        action="store_true",
        help="Send output to stdout instead of to file.\nDefault: false",
    )
    parser.add_argument(
        "--pagelim",
        default=3,
        type=int,
        help="Process up to this page in the pdf.\nDefault: 3",
    )
    parser.add_argument(
        "--district-rename",
        type=Path,
        default=None,
        help='To rename civil districts, specify a CSV file with contents in the format of: old_name,new_name. For example: "StevensPt","Stevens Point"',
    )
    parser.add_argument(
        "--region-rename",
        type=Path,
        default=None,
        help="To rename the entire region, specify a CSV file with contents in the format of: old_name,new_name. This is useful when the name is missing pieces, like ward number or civil district type",
    )
    parser.add_argument(
        "--priorities",
        type=Path,
        default=None,
        help="To assign priorities to civil districts, specify a CSV file ith contents in the format of: region,priority. This is useful for automatically sorting by canvassing priority.",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="count",
        default=0,
        help="Increase verbosity (e.g. -v, -vv).",
    )
    args = parser.parse_args()

    if args.verbose == 0:
        level = logging.WARNING
        logger.info("Running script...")
    elif args.verbose == 1:
        level = logging.INFO
        logger.info("Starting script in verbose mode...")
    else:
        level = logging.DEBUG
        logger.info("Starting script in debug mode...")

    logging.basicConfig(level=level, format="%(levelname)s: %(message)s")
    logging.getLogger("pypdf").setLevel(logging.ERROR)
    logging.getLogger("ezsheets").setLevel(logging.ERROR)

    # check if stdin is interactive
    # allows users to pipe input into the script
    if not sys.stdin.isatty():
        files = [
            Path(p.rstrip())
            for p in sys.stdin
            if not Path(p.rstrip()).is_dir() and p.rstrip().endswith(".pdf")
        ]
        for d in (Path(d.rstrip()) for d in sys.stdin if Path(d.rstrip()).is_dir()):
            files.extend(d.glob("*.pdf"))
    elif args.input:
        files = [p for p in args.input if not p.is_dir() and p.suffix == ".pdf"]
        for d in (d for d in args.input if d.is_dir()):
            files.extend(d.glob("*.pdf"))

    if len(files) == 0:
        parser.print_help()
        sys.exit(0)

    logger.info(
        f"running script with {len(files)} input {'file' if len(files) == 1 else 'files'}"
    )
    logger.debug(f"input filenames:\n{'\n'.join('\t' + str(f) for f in files)}")

    if not args.stdout:
        logger.debug(f"output: {str(args.output)}")
    else:
        logger.debug("output: stdout")

    logger.info("exporting turf data...")
    data = [
        [
            "region_name_raw",
            "civil_district_name",
            "civil_district_type",
            "ward_number",
            "list_number",
            "turf_number",
            "door_count",
        ]
    ]
    for f in files:
        data.extend(extract_data(f, args.region_rename, pagelim=args.pagelim))
    logger.info(f"found {len(data[1:])} turfs to process")
    data = post_process(data, args.district_rename, args.priorities)
    if args.stdout:
        csv.writer(sys.stdout).writerows(data)
    else:
        with open(args.output, "w") as output:
            csv.writer(output).writerows(data)
    if args.copy:
        pyperclip.copy(data)
    if args.remote:
        update_sheet(args.remote, data)


if __name__ == "__main__":
    main()
