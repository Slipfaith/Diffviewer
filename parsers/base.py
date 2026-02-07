from __future__ import annotations

from abc import ABC, abstractmethod

from core.models import ParsedDocument


class BaseParser(ABC):
    name: str
    supported_extensions: list[str]
    format_description: str

    @abstractmethod
    def can_handle(self, filepath: str) -> bool:
        raise NotImplementedError

    @abstractmethod
    def parse(self, filepath: str) -> ParsedDocument:
        raise NotImplementedError

    @abstractmethod
    def validate(self, filepath: str) -> list[str]:
        raise NotImplementedError
