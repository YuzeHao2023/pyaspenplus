from abc import ABC, abstractmethod
from typing import Tuple
from .types_ import StreamSpecification, ColumnInputSpecification, ColumnOutputSpecification, ProductSpecification

class BaseAspenDistillationAPI(ABC):
    @abstractmethod
    def set_input_stream_specification(self, stream_specification: StreamSpecification) -> None:
        raise NotImplementedError

    @abstractmethod
    def get_output_stream_specifications(self) -> Tuple[StreamSpecification, StreamSpecification]:
        raise NotImplementedError

    @abstractmethod
    def get_simulated_column_properties(self, column_input_specification: ColumnInputSpecification) -> ColumnOutputSpecification:
        raise NotImplementedError

    @abstractmethod
    def set_column_specification(self, column_input_specification: ColumnInputSpecification) -> None:
        raise NotImplementedError

    @abstractmethod
    def solve_flowsheet(self) -> None:
        raise NotImplementedError

    @abstractmethod
    def get_column_cost(self, stream_specification: StreamSpecification, column_input_specification: ColumnInputSpecification,
                        column_output_specification: ColumnOutputSpecification) -> float:
        raise NotImplementedError

    @abstractmethod
    def get_stream_value(self, stream, product_specification: ProductSpecification) -> float:
        raise NotImplementedError

    @abstractmethod
    def stream_is_product_or_outlet(self, stream: StreamSpecification, product_specification: ProductSpecification):
        raise NotImplementedError
