from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable


OperationFactory = Callable[[dict[str, Any] | None], "Operation"]
WidgetFactory = Callable[..., Any]


class OperationValidationError(ValueError):
    pass


@dataclass(slots=True)
class OperationResult:
    status: str
    message: str
    details: dict[str, Any] = field(default_factory=dict)


class Operation:
    name = "Operation"
    required_capabilities: set[str] = set()
    params_schema: dict[str, Any] = {}

    def __init__(self, params: dict[str, Any] | None = None) -> None:
        self.params = params or {}

    def validate(self, engine) -> None:
        self.ensure_supported(engine)

    def run(self, engine) -> OperationResult:
        raise NotImplementedError

    def describe(self) -> str:
        return self.name

    def ensure_supported(self, engine) -> None:
        capabilities = getattr(engine, "capabilities", set())
        missing = sorted(self.required_capabilities - set(capabilities))
        if missing:
            raise OperationValidationError(
                f"{self.name} requires capabilities that the current engine does not provide: "
                f"{', '.join(missing)}"
            )


@dataclass(slots=True)
class RegisteredOperation:
    display_name: str
    operation_cls: type[Operation]
    widget_factory: WidgetFactory | None = None

    def create(self, params: dict[str, Any] | None = None) -> Operation:
        return self.operation_cls(params)


class OperationRegistry:
    _registry: dict[str, RegisteredOperation] = {}

    @classmethod
    def register(
        cls,
        display_name: str,
        widget_factory: WidgetFactory | None = None,
    ) -> Callable[[type[Operation]], type[Operation]]:
        def decorator(operation_cls: type[Operation]) -> type[Operation]:
            if display_name in cls._registry:
                raise ValueError(f"Operation '{display_name}' is already registered.")
            cls._registry[display_name] = RegisteredOperation(
                display_name=display_name,
                operation_cls=operation_cls,
                widget_factory=widget_factory,
            )
            return operation_cls

        return decorator

    @classmethod
    def create(cls, display_name: str, params: dict[str, Any] | None = None) -> Operation:
        return cls.get(display_name).create(params)

    @classmethod
    def get(cls, display_name: str) -> RegisteredOperation:
        try:
            return cls._registry[display_name]
        except KeyError as exc:
            raise KeyError(f"Unknown operation '{display_name}'.") from exc

    @classmethod
    def list_operations(cls) -> list[str]:
        return sorted(cls._registry)

    @classmethod
    def get_widget_factory(cls, display_name: str) -> WidgetFactory | None:
        return cls.get(display_name).widget_factory
