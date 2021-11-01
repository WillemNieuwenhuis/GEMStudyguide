import logging
import argparse
from gemguide import gemtodocx

def init_argparse() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description = "Generate GEM study guide"
    )

    parser.add_argument('source', help='Input Excel workbook')
    parser.add_argument('output', help='Output document name')
    parser.add_argument("-v", "--version", action="version",version = f"{parser.prog} version october 2021")

    return parser


if __name__ == '__main__':
    log = logging.getLogger(__name__)
    logging.basicConfig(
		format = '%(asctime)s %(levelname)-8s %(message)s',
		level = logging.INFO,
		datefmt = '%Y-%m-%d %H:%M:%S')

    parser = init_argparse()
    args = parser.parse_args()

    fn = args.source
    out = args.output

    if out == fn:
        log.error('Input and output filenames should be different')
        quit()

    log.info(f'Generate GEM studyguide pages from {fn}')

    gemtodocx.convert(fn, out)