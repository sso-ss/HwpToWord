"""Shared, explicit resource budgets for the HWP 5 reader and converter."""

import argparse
from dataclasses import asdict, dataclass


MIB = 1024 * 1024


def positive_integer(value):
    try:
        number = int(value)
    except (ValueError, TypeError):
        raise argparse.ArgumentTypeError("must be a positive integer") from None
    if number <= 0:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return number


@dataclass(frozen=True)
class HwpLimits:
    max_input_mb: int = 50
    max_stream_mb: int = 64
    max_expanded_mb: int = 128
    max_records: int = 200000

    def __post_init__(self):
        for name, value in asdict(self).items():
            if type(value) is not int or value <= 0:
                raise ValueError(f"{name} must be a positive integer.")

    def check_input(self, source):
        if source.stat().st_size > self.max_input_mb * MIB:
            raise ValueError(f"Input exceeds the {self.max_input_mb} MiB size limit; "
                             "use --max-input-mb to raise the HWP 5 input budget.")

    def arguments(self):
        return [part for name, value in asdict(self).items()
                for part in ("--" + name.replace("_", "-"), str(value))]

    @classmethod
    def from_arguments(cls, arguments):
        return cls(**{name: getattr(arguments, name) for name in cls.__dataclass_fields__})


def add_limit_arguments(parser):
    defaults = HwpLimits()
    for flag, description in (
        ("max_input_mb", "Input file budget in MiB (HWP 3 still has a 50 MiB ceiling)"),
        ("max_stream_mb", "HWP 5 stored/expanded stream budget in MiB"),
        ("max_expanded_mb", "HWP 5 total expanded stream budget in MiB"),
        ("max_records", "HWP 5 record budget per stream"),
    ):
        default = getattr(defaults, flag)
        parser.add_argument("--" + flag.replace("_", "-"), type=positive_integer,
                            default=default, help=f"{description}; default: {default}.")
