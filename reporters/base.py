from __future__ import annotations

from abc import ABC, abstractmethod

from core.models import ComparisonResult


class BaseReporter(ABC):
    name: str
    output_extension: str
    supports_rich_text: bool = False

    @abstractmethod
    def generate(self, result: ComparisonResult, output_path: str) -> str:
        raise NotImplementedError
