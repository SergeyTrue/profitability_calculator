from pathlib import Path
from typing import Iterable
from abc import ABC, abstractmethod


class NoExcelFiles(Exception):
    pass


class IDataGateway(ABC):
    @abstractmethod
    def process(self, *args, **kwargs):
        pass


def correct_files_suffix(files: Iterable[Path]):
    for file in files:
        new = file.parent / (file.stem + file.suffix.lower())
        file.rename(new)
        yield new

