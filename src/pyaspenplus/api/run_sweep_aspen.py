"""
Example parameter sweep for Aspen using the AspenAPI adapter.

Run on Windows with Aspen installed and pywin32 available:
    python -m pyaspenplus.api.run_sweep_aspen

Adjust flowsheet path below to point to your .bkp file.
"""
import csv
import itertools
import os
from pyaspenplus.api.aspen_api import AspenAPI
from pyaspenplus.api.types_ import StreamSpecification, PerCompoundProperty, ColumnInputSpecification

def build_param_grid(temps, cats):
    return [{"TEMPERATURE": t, "CAT_EQ": c} for t, c in itertools.product(temps, cats)]

def run_sweep(flowsheet_bkp: str, output_csv: str = "sweep_results_aspen.csv"):
    api = AspenAPI(flowsheet_path=flowsheet_bkp)
    temps = [290.0, 300.0, 320.0]
    cats = [0.05, 0.1, 0.2]
    grid = build_param_grid(temps, cats)

    column_input = ColumnInputSpecification(n_stages=50, feed_stage_location=25,
                                            reflux_ratio=1.0, reboil_ratio=1.0, condensor_pressure=17.4)

    rows = []
    for idx, params in enumerate(grid):
        temp = float(params["TEMPERATURE"])
        cat = float(params["CAT_EQ"])
        feed = StreamSpecification(temperature=temp, pressure=17.4,
                                   molar_flows=PerCompoundProperty(ethane=0.017,
                                                                   propane=1.110,
                                                                   isobutane=1.198,
                                                                   n_butane=0.516,
                                                                   isopentane=0.334,
                                                                   n_pentane=0.173))
        try:
            api.set_input_stream_specification(feed)
            api.set_column_specification(column_input)
            api.solve_flowsheet()
            tops, bots = api.get_output_stream_specifications()
            row = {
                "idx": idx,
                "temperature_set": temp,
                "cat_eq_set": cat,
                "tops_temp": tops.temperature,
                "tops_pressure": tops.pressure,
                "tops_ethane": tops.molar_flows.ethane,
                "tops_propane": tops.molar_flows.propane,
                "bots_temp": bots.temperature,
                "bots_pressure": bots.pressure,
                "bots_ethane": bots.molar_flows.ethane,
                "bots_propane": bots.molar_flows.propane,
            }
        except Exception as exc:
            row = {"idx": idx, "error": str(exc)}
        rows.append(row)

    keys = sorted(set().union(*(r.keys() for r in rows)))
    with open(output_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=keys)
        writer.writeheader()
        for r in rows:
            writer.writerow(r)
    print(f"Finished sweep, wrote {len(rows)} rows to {output_csv}")

if __name__ == "__main__":
    # 默认位置：请改成你的 .bkp 的绝对路径
    flowsheet = os.path.abspath(os.path.join(os.path.dirname(__file__), "../AspenSimulation/HydrocarbonMixture.bkp"))
    run_sweep(flowsheet_bkp=flowsheet, output_csv="sweep_results_aspen.csv")
