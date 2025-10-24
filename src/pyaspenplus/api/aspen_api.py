"""
Aspen API adapter for pyaspenplus using COM-backed Simulation.
This module maps higher-level API calls (used by the rest of the package) to the COM wrapper.
"""
from typing import Tuple
from pyaspenplus.api.com_simulation import Simulation

from pyaspenplus.api.api_base import BaseAspenDistillationAPI
from pyaspenplus.api.types_ import StreamSpecification, ColumnInputSpecification, \
    ColumnOutputSpecification, ProductSpecification, PerCompoundProperty


class AspenAPI(BaseAspenDistillationAPI):
    def __init__(self, max_solve_iterations: int = 100,
                 flowsheet_path: str = "HydrocarbonMixture.bkp"):
        self._flowsheet: Simulation = Simulation(VISIBILITY=False,
                                                 SUPPRESS=True,
                                                 max_iterations=max_solve_iterations,
                                                 flowsheet_path=flowsheet_path)
        self._feed_name: str = "S1"
        self._tops_name: str = "S2"
        self._bottoms_name: str = "S3"
        # mapping names -> Aspen component ids; adjust to match your flowsheet
        self._name_to_aspen_name = PerCompoundProperty(ethane="ETHANE",
                                                       propane="PROPANE",
                                                       isobutane="I-BUTANE",
                                                       n_butane="N-BUTANE",
                                                       isopentane="I-PENTAN",
                                                       n_pentane="N-PENTAN")

    def set_input_stream_specification(self, stream_specification: StreamSpecification) -> None:
        self._flowsheet.STRM_Temperature(self._feed_name, stream_specification.temperature)
        self._flowsheet.STRM_Pressure(self._feed_name, stream_specification.pressure)
        # set component flows (molar flows). Ensure the component ids match your Aspen components
        self._flowsheet.STRM_Flowrate(self._feed_name, self._name_to_aspen_name.ethane, stream_specification.molar_flows.ethane)
        self._flowsheet.STRM_Flowrate(self._feed_name, self._name_to_aspen_name.propane, stream_specification.molar_flows.propane)
        self._flowsheet.STRM_Flowrate(self._feed_name, self._name_to_aspen_name.isobutane, stream_specification.molar_flows.isobutane)
        self._flowsheet.STRM_Flowrate(self._feed_name, self._name_to_aspen_name.n_butane, stream_specification.molar_flows.n_butane)
        self._flowsheet.STRM_Flowrate(self._feed_name, self._name_to_aspen_name.isopentane, stream_specification.molar_flows.isopentane)
        self._flowsheet.STRM_Flowrate(self._feed_name, self._name_to_aspen_name.n_pentane, stream_specification.molar_flows.n_pentane)

    def get_output_stream_specifications(self) -> Tuple[StreamSpecification, StreamSpecification]:
        tops_temperature = self._flowsheet.STRM_Get_Temperature(self._tops_name)
        tops_pressure = self._flowsheet.STRM_Get_Pressure(self._tops_name)
        tops_ethane = self._flowsheet.STRM_Get_Outputs(self._tops_name, self._name_to_aspen_name.ethane)
        tops_propane = self._flowsheet.STRM_Get_Outputs(self._tops_name, self._name_to_aspen_name.propane)
        tops_isobutane = self._flowsheet.STRM_Get_Outputs(self._tops_name, self._name_to_aspen_name.isobutane)
        tops_n_butane = self._flowsheet.STRM_Get_Outputs(self._tops_name, self._name_to_aspen_name.n_butane)
        tops_isopentane = self._flowsheet.STRM_Get_Outputs(self._tops_name, self._name_to_aspen_name.isopentane)
        tops_n_pentane = self._flowsheet.STRM_Get_Outputs(self._tops_name, self._name_to_aspen_name.n_pentane)

        tops_specifications = StreamSpecification(temperature=tops_temperature, pressure=tops_pressure,
                                                  molar_flows=PerCompoundProperty(ethane=tops_ethane,
                                                                                  propane=tops_propane,
                                                                                  isobutane=tops_isobutane,
                                                                                  n_butane=tops_n_butane,
                                                                                  isopentane=tops_isopentane,
                                                                                  n_pentane=tops_n_pentane))

        bots_temperature = self._flowsheet.STRM_Get_Temperature(self._bottoms_name)
        bots_pressure = self._flowsheet.STRM_Get_Pressure(self._bottoms_name)
        bots_ethane = self._flowsheet.STRM_Get_Outputs(self._bottoms_name, self._name_to_aspen_name.ethane)
        bots_propane = self._flowsheet.STRM_Get_Outputs(self._bottoms_name, self._name_to_aspen_name.propane)
        bots_isobutane = self._flowsheet.STRM_Get_Outputs(self._bottoms_name, self._name_to_aspen_name.isobutane)
        bots_n_butane = self._flowsheet.STRM_Get_Outputs(self._bottoms_name, self._name_to_aspen_name.n_butane)
        bots_isopentane = self._flowsheet.STRM_Get_Outputs(self._bottoms_name, self._name_to_aspen_name.isopentane)
        bots_n_pentane = self._flowsheet.STRM_Get_Outputs(self._bottoms_name, self._name_to_aspen_name.n_pentane)

        bots_specifications = StreamSpecification(temperature=bots_temperature, pressure=bots_pressure,
                                                  molar_flows=PerCompoundProperty(ethane=bots_ethane,
                                                                                  propane=bots_propane,
                                                                                  isobutane=bots_isobutane,
                                                                                  n_butane=bots_n_butane,
                                                                                  isopentane=bots_isopentane,
                                                                                  n_pentane=bots_n_pentane))

        return tops_specifications, bots_specifications

    def get_simulated_column_properties(self, column_input_specification: ColumnInputSpecification) -> ColumnOutputSpecification:
        D_Cond_Duty = self._flowsheet.BLK_Get_Condenser_Duty()
        D_Reb_Duty = self._flowsheet.BLK_Get_Reboiler_Duty()
        vap_flows = self._flowsheet.BLK_Get_Column_Stage_Vapor_Flows(column_input_specification.n_stages)
        stage_temp = self._flowsheet.BLK_Get_Column_Stage_Temperatures(column_input_specification.n_stages)
        stage_mw = self._flowsheet.BLK_Get_Column_Stage_Molar_Weights(column_input_specification.n_stages)

        return ColumnOutputSpecification(condenser_duty=D_Cond_Duty,
                                         reboiler_duty=D_Reb_Duty,
                                         vapor_flow_per_stage=vap_flows,
                                         temperature_per_stage=stage_temp,
                                         molar_weight_per_stage=stage_mw)

    def set_column_specification(self, column_input_specification: ColumnInputSpecification) -> None:
        self._flowsheet.BLK_NumberOfStages(column_input_specification.n_stages)
        self._flowsheet.BLK_FeedLocation(column_input_specification.feed_stage_location, self._feed_name)
        self._flowsheet.BLK_Pressure(column_input_specification.condensor_pressure)
        self._flowsheet.BLK_RefluxRatio(column_input_specification.reflux_ratio)
        self._flowsheet.BLK_ReboilerRatio(column_input_specification.reboil_ratio)

    def solve_flowsheet(self) -> None:
        self._flowsheet.Run()

    def get_column_cost(self, stream_specification: StreamSpecification, column_input_specification: ColumnInputSpecification,
                        column_output_specification: ColumnOutputSpecification) -> float:
        # If your flowsheet provides CAL_... macros they can be called via Simulation (not implemented here).
        # Provide a placeholder that returns 0 if no special cost routine.
        try:
            t_reboiler = column_output_specification.temperature_per_stage[-1]
            t_condenser = column_output_specification.temperature_per_stage[0]
            invest = getattr(self._flowsheet, "CAL_InvestmentCost", lambda *a, **k: 0.0)(
                stream_specification.pressure,
                column_input_specification.n_stages,
                column_output_specification.condenser_duty,
                t_reboiler,
                column_output_specification.reboiler_duty,
                t_condenser,
                column_output_specification.vapor_flow_per_stage,
                column_output_specification.molar_weight_per_stage,
                column_output_specification.temperature_per_stage)
            operating = getattr(self._flowsheet, "CAL_Annual_OperatingCost", lambda *a, **k: 0.0)(
                column_output_specification.reboiler_duty, column_output_specification.condenser_duty)
            return invest + operating
        except Exception:
            return 0.0

    def get_stream_value(self, stream, product_specification) -> float:
        stream_value, _ = getattr(self._flowsheet, "CAL_stream_value", lambda *a, **k: (0.0, None))(stream, product_specification.purity)
        return float(stream_value)

    def stream_is_product_or_outlet(self, stream: StreamSpecification,
                                    product_specification: ProductSpecification) -> Tuple[bool, bool]:
        total_flow = sum(stream.molar_flows)
        if total_flow > 0.001:
            is_purity, _ = getattr(self._flowsheet, "CAL_purity_check", lambda *a, **k: ([0]*6, None))(stream, product_specification.purity)
        else:
            is_purity = [0] * 6

        is_product = bool(any(is_purity))
        is_outlet = is_product or (total_flow < 0.001)
        return is_product, is_outlet
