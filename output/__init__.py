from abc import ABC, abstractmethod


class IOutput(ABC):
    @abstractmethod
    def process(self, messages):
        pass
