from dataclasses import dataclass
from typing import Tuple

@dataclass
class PerCompoundProperty:
    ethane: float = 0.0
    propane: float = 0.0
    isobutane: float = 0.0
    n_butane: float = 0.0
    isopentane: float = 0.0
    n_pentane: float = 0.0

    def __iter__(self):
        return iter((self.ethane, self.propane, self.isobutane, self.n_butane, self.isopentane, self.n_pentane))

    def as_tuple(self) -> Tuple[float, float, float, float, float, float]:
        return (self.ethane, self.propane, self.isobutane, self.n_butane, self.isopentane, self.n_pentane)

@dataclass
class StreamSpecification:
    temperature: float
    pressure: float
    molar_flows: PerCompoundProperty

    @property
    def total_molar_flow(self) -> float:
        return sum(self.molar_flows.as_tuple())

@dataclass
class ColumnInputSpecification:
    n_stages: int
    feed_stage_location: int
    reflux_ratio: float
    reboil_ratio: float
    condensor_pressure: float

@dataclass
class ColumnOutputSpecification:
    condenser_duty: float
    reboiler_duty: float
    vapor_flow_per_stage: tuple
    temperature_per_stage: tuple
    molar_weight_per_stage: tuple

@dataclass
class ProductSpecification:
    purity: float
    # 可以扩展：价格、destination 等
