"""Process-local adapter for explicitly requested HWP text recovery."""

import codecs
import json
from pathlib import Path
import runpy
import sys


RECOVERY_PREFIX = "HWP_TEXT_RECOVERY:"


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


def main():
    import hwp5.dataio

    recovery = TextRecovery()
    hwp5.dataio.decode_utf16le_with_hypua = recovery.decode
    reader = Path(sys.executable).parent / "hwp5proc"
    sys.argv = [str(reader), "xml", "--embedbin", sys.argv[1]]
    try:
        runpy.run_path(str(reader), run_name="__main__")
    finally:
        print(RECOVERY_PREFIX + json.dumps(recovery.replacements), file=sys.stderr)


if __name__ == "__main__":
    main()