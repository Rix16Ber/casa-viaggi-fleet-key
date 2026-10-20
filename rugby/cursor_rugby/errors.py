"""Errori che fermano un passaggio prima di lasciare il lavoro a metà."""


class StepError(RuntimeError):
    def __init__(self, step: str, message: str):
        self.step = step
        self.message = message
        super().__init__(f"Passaggio {step}: {message}")
