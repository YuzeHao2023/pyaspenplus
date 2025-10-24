"""
示例：参数扫描（parameter sweep） + Aspen 执行 + 导出 CSV

要求：
- 在 Windows 上运行
- 安装 Aspen 并确保 COM 接口可用（常见 ProgID: "Apwn.Document" 或 "AspenPlus.Application" 等）
- 安装 pywin32

用法：
    python examples/run_sweep_aspen.py --flowsheet "C:\\path\\to\\HydrocarbonMixture.bkp" --out results.csv
"""
import argparse
import csv
import itertools
import os
import sys
from pyaspenplus.api.aspen_api import AspenAPI
from pyaspenplus.api.types_ import StreamSpecification, PerCompoundProperty, ColumnInputSpecification

def build_param_grid(temps, reflux_values):
    return [{"temperature": t, "reflux": r} for t, r in itertools.product(temps, reflux_values)]

def run_sweep(flowsheet_path: str, output_csv: str):
    # Instantiate AspenAPI which wraps COM-backed Simulation
    api = AspenAPI(flowsheet_path=flowsheet_path)

    temps = [290.0, 300.0, 320.0]  # example temps (K)
    refluxes = [0.8, 1.0, 1.5]     # example reflux ratios
    grid = build_param_grid(temps, refluxes)

    # Column input skeleton (we vary reflux in the sweep)
    base_column_input = ColumnInputSpecification(n_stages=50, feed_stage_location=25,
                                                 reflux_ratio=1.0, reboil_ratio=1.0, condensor_pressure=17.4)

    rows = []
    for idx, point in enumerate(grid):
        temp = point["temperature"]
        reflux = point["reflux"]
        # Build feed stream spec (adapt the molar flows to match your flowsheet)
        feed = StreamSpecification(temperature=temp, pressure=17.4,
                                   molar_flows=PerCompoundProperty(ethane=0.017,
                                                                   propane=1.110,
                                                                   isobutane=1.198,
                                                                   n_butane=0.516,
                                                                   isopentane=0.334,
                                                                   n_pentane=0.173))
        try:
            api.set_input_stream_specification(feed)
            # set column with varying reflux
            column_input = ColumnInputSpecification(n_stages=base_column_input.n_stages,
                                                    feed_stage_location=base_column_input.feed_stage_location,
                                                    reflux_ratio=reflux,
                                                    reboil_ratio=base_column_input.reboil_ratio,
                                                    condensor_pressure=base_column_input.condensor_pressure)
            api.set_column_specification(column_input)
            api.solve_flowsheet()

            tops, bots = api.get_output_stream_specifications()
            col_props = api.get_simulated_column_properties(column_input)

            row = {
                "idx": idx,
                "temp_set": temp,
                "reflux_set": reflux,
                "tops_temp": tops.temperature,
                "tops_pressure": tops.pressure,
                "tops_ethane": tops.molar_flows.ethane,
                "tops_propane": tops.molar_flows.propane,
                "bots_temp": bots.temperature,
                "bots_pressure": bots.pressure,
                "bots_ethane": bots.molar_flows.ethane,
                "bots_propane": bots.molar_flows.propane,
                "cond_duty": col_props.condenser_duty,
                "reb_duty": col_props.reboiler_duty,
            }
        except Exception as exc:
            row = {"idx": idx, "error": str(exc)}
        rows.append(row)

    # write CSV
    keys = sorted(set().union(*(r.keys() for r in rows)))
    with open(output_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=keys)
        writer.writeheader()
        for r in rows:
            writer.writerow(r)
    print(f"Finished sweep, wrote {len(rows)} rows to {output_csv}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--flowsheet", required=True, help="Path to .bkp flowsheet")
    parser.add_argument("--out", default="sweep_results.csv", help="Output CSV path")
    args = parser.parse_args()
    flowsheet = os.path.abspath(args.flowsheet)
    if not os.path.exists(flowsheet):
        print("Flowsheet not found:", flowsheet)
        sys.exit(2)
    run_sweep(flowsheet, args.out)
