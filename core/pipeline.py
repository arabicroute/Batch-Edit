from __future__ import annotations

import traceback
from dataclasses import dataclass
from typing import Callable

from PyQt5 import QtCore

from core.operations import Operation, OperationResult, OperationValidationError


@dataclass(slots=True)
class PipelineEntryResult:
    operation_name: str
    operation_summary: str
    result: OperationResult


class Pipeline:
    def __init__(
        self,
        operations: list[Operation] | None = None,
        continue_on_error: bool = False,
    ) -> None:
        self.operations = operations or []
        self.continue_on_error = continue_on_error

    def add(self, operation: Operation) -> None:
        self.operations.append(operation)

    def remove(self, index: int) -> Operation:
        return self.operations.pop(index)

    def move(self, source_index: int, target_index: int) -> None:
        operation = self.operations.pop(source_index)
        self.operations.insert(target_index, operation)

    def run(
        self,
        engine,
        progress_callback: Callable[[int, Operation, OperationResult], None] | None = None,
    ) -> list[OperationResult]:
        results: list[OperationResult] = []
        for index, operation in enumerate(self.operations):
            try:
                operation.validate(engine)
                result = operation.run(engine)
            except OperationValidationError as exc:
                result = OperationResult(status="error", message=str(exc))
            except Exception as exc:  # noqa: BLE001
                result = OperationResult(
                    status="error",
                    message=str(exc),
                    details={"traceback": traceback.format_exc()},
                )
            results.append(result)
            if progress_callback is not None:
                progress_callback(index, operation, result)
            if result.status == "error" and not self.continue_on_error:
                break
        return results


class PipelineWorker(QtCore.QThread):
    operation_started = QtCore.pyqtSignal(int, object)
    operation_finished = QtCore.pyqtSignal(int, object, object)
    pipeline_finished = QtCore.pyqtSignal(object)
    pipeline_failed = QtCore.pyqtSignal(str)

    def __init__(self, pipeline: Pipeline, engine) -> None:
        super().__init__()
        self.pipeline = pipeline
        self.engine = engine

    def run(self) -> None:
        try:
            results: list[PipelineEntryResult] = []

            def on_progress(index: int, operation: Operation, result: OperationResult) -> None:
                entry = PipelineEntryResult(
                    operation_name=operation.name,
                    operation_summary=operation.describe(),
                    result=result,
                )
                results.append(entry)
                self.operation_finished.emit(index, operation, result)

            original_operations = list(self.pipeline.operations)

            def wrapped_progress(index: int, operation: Operation, result: OperationResult) -> None:
                on_progress(index, operation, result)

            for index, operation in enumerate(original_operations):
                self.operation_started.emit(index, operation)
                temp_pipeline = Pipeline(
                    operations=[operation],
                    continue_on_error=self.pipeline.continue_on_error,
                )
                step_results = temp_pipeline.run(self.engine)
                if step_results:
                    wrapped_progress(index, operation, step_results[0])
                if step_results and step_results[0].status == "error" and not self.pipeline.continue_on_error:
                    break

            self.pipeline_finished.emit(results)
        except Exception as exc:  # noqa: BLE001
            self.pipeline_failed.emit(f"{exc}\n\n{traceback.format_exc()}")


class PipelineExecutionWorker(QtCore.QThread):
    operation_started = QtCore.pyqtSignal(int, object)
    operation_finished = QtCore.pyqtSignal(int, object, object)
    pipeline_finished = QtCore.pyqtSignal(object)
    pipeline_failed = QtCore.pyqtSignal(object)

    def __init__(
        self,
        pipeline: Pipeline,
        engine_factory: Callable[[], object],
        input_path: str,
        output_path: str,
    ) -> None:
        super().__init__()
        self.pipeline = pipeline
        self.engine_factory = engine_factory
        self.input_path = input_path
        self.output_path = output_path

    def run(self) -> None:
        engine = None
        try:
            results: list[PipelineEntryResult] = []
            try:
                engine = self.engine_factory()
            except Exception as exc:  # noqa: BLE001
                self.pipeline_failed.emit(
                    {
                        "category": "engine",
                        "stage": "create",
                        "message": str(exc),
                        "traceback": traceback.format_exc(),
                    }
                )
                return

            try:
                engine.open(self.input_path)
            except Exception as exc:  # noqa: BLE001
                self.pipeline_failed.emit(
                    {
                        "category": "engine",
                        "stage": "open",
                        "message": str(exc),
                        "traceback": traceback.format_exc(),
                        "input_path": self.input_path,
                    }
                )
                return

            for index, operation in enumerate(self.pipeline.operations):
                self.operation_started.emit(index, operation)
                step_pipeline = Pipeline(
                    operations=[operation],
                    continue_on_error=self.pipeline.continue_on_error,
                )
                step_results = step_pipeline.run(engine)
                if step_results:
                    entry = PipelineEntryResult(
                        operation_name=operation.name,
                        operation_summary=operation.describe(),
                        result=step_results[0],
                    )
                    results.append(entry)
                    self.operation_finished.emit(index, operation, step_results[0])
                    if (
                        step_results[0].status == "error"
                        and not self.pipeline.continue_on_error
                    ):
                        self.pipeline_finished.emit(
                            {
                                "results": results,
                                "saved": False,
                                "output_path": self.output_path,
                            }
                        )
                        return

            try:
                engine.save(self.output_path)
            except Exception as exc:  # noqa: BLE001
                self.pipeline_failed.emit(
                    {
                        "category": "engine",
                        "stage": "save",
                        "message": str(exc),
                        "traceback": traceback.format_exc(),
                        "output_path": self.output_path,
                        "results": results,
                    }
                )
                return
            self.pipeline_finished.emit(
                {
                    "results": results,
                    "saved": True,
                    "output_path": self.output_path,
                }
            )
        except Exception as exc:  # noqa: BLE001
            self.pipeline_failed.emit(
                {
                    "category": "unexpected",
                    "stage": "run",
                    "message": str(exc),
                    "traceback": traceback.format_exc(),
                }
            )
        finally:
            if engine is not None:
                try:
                    engine.close()
                except Exception:  # noqa: BLE001
                    pass
