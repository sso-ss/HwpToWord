"""Isolated native HWP reader and explicit UTF-16 recovery support."""

import argparse
import codecs
from pathlib import Path
import sys
import xml.etree.ElementTree as ET

from hwp3_parser import Hwp3Parser, SIGNATURE
from hwp_limits import HwpLimits, add_limit_arguments


class TextRecovery:
    def __init__(self):
        self.chunks = 0
        self.replacements = []
        self.handler_name = "hwp_recovery_" + str(id(self))
        codecs.register_error(self.handler_name, self.replace_error)

    def replace_error(self, error):
        if not isinstance(error, UnicodeDecodeError) or error.encoding != "utf-16-le":
            raise error
        self.replacements.append({
            "chunk_index": self.chunks,
            "byte_offset_in_chunk": error.start,
            "invalid_bytes_hex": error.object[error.start:error.end].hex(),
            "reason": error.reason,
            "replacement": "U+FFFD",
        })
        return "\ufffd", error.end

    def decode(self, data):
        self.chunks += 1
        return data.decode("utf-16le", errors=self.handler_name)


def parse_source(source, recover_text=False, limits=None, asset_directory=None):
    limits = limits if limits is not None else HwpLimits()
    limits.check_input(Path(source))
    with open(source, "rb") as stream:
        signature = stream.read(30)
    if signature == SIGNATURE:
        return Hwp3Parser().parse(source)
    if signature[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise ValueError("Not a supported binary HWP 3 or HWP 5 document; HWPX is unsupported.")
    # HWP 3 decoding can run without the third-party OLE reader installed.
    from hwp_parser import HwpParser
    return HwpParser(recover_text=recover_text, limits=limits, asset_directory=asset_directory).parse(source)


def main():

    arguments = argparse.ArgumentParser(description="Native HWP 3 / HWP 5 reader")
    arguments.add_argument("source")
    arguments.add_argument("--recover-text", action="store_true")
    arguments.add_argument("--asset-directory", type=Path, help=argparse.SUPPRESS)
    add_limit_arguments(arguments)
    options = arguments.parse_args()
    try:
        root, warnings = parse_source(options.source, recover_text=options.recover_text,
                                      limits=HwpLimits.from_arguments(options),
                                      asset_directory=options.asset_directory)
        ET.ElementTree(root).write(sys.stdout.buffer, encoding="utf-8")
    except (ValueError, OSError) as error:
        arguments.exit(1, str(error) + "\n")


if __name__ == "__main__":
    main()
