import logging
import argparse
import pathlib
from gemguide import gemtodocx

def init_argparse() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description = "Generate GEM study guide"
    )

    parser.add_argument('source', help='Input Excel workbook')
    parser.add_argument('output', help='Output document name')
    parser.add_argument('-d', '--docx', help='Generate DOCX output (default)', action='store_true')
    parser.add_argument('-p', '--pdf', help='Generate PDF output', action='store_true')
    parser.add_argument("-v", "--version", action="version",version = f"{parser.prog} version october 2021")

    return parser

def makeAbsolute(fn : str) -> pathlib.Path:
    path = pathlib.Path(fn)
    cur = pathlib.Path.cwd()
    folder = path.parents[0]
    if folder == '.':
        path = cur / path

    return path

if __name__ == '__main__':
    log = logging.getLogger(__name__)
    logging.basicConfig(
		format = '%(asctime)s %(levelname)-8s %(message)s',
		level = logging.INFO,
		datefmt = '%Y-%m-%d %H:%M:%S')

    parser = init_argparse()
    args = parser.parse_args()

    fn = makeAbsolute(args.source)
    out = makeAbsolute(args.output)

    if out == fn:
        log.error('Input and output filenames should be different')
        quit()

    log.info(f'Generate GEM studyguide pages from {args.source}')

    if args.docx:
        log.info('Generating .docx')
        gemtodocx.convert2docx(fn, out.with_suffix('.docx'))

    if args.pdf:
        log.info('Generating .pdf')
        gemtodocx.convert2pdf(fn, out.with_suffix('.pdf'))