"""
COM-backed Simulation wrapper for Aspen (pyaspenplus).

Usage:
  sim = Simulation(VISIBILITY=False, SUPPRESS=True, max_iterations=100, flowsheet_path="...bkp")

Notes:
- Requires Windows + Aspen installed + pywin32.
- Node names in Aspen Tree vary by version; methods include multiple fallbacks and descriptive errors.
"""
from typing import Tuple, Optional
import time
import os
import traceback

try:
    import win32com.client as win32  # pywin32
except Exception as e:
    raise RuntimeError("pywin32 is required for Aspen COM integration. Install pywin32 on Windows.") from e


class Simulation:
    def __init__(self, VISIBILITY: bool = False, SUPPRESS: bool = True,
                 max_iterations: int = 100, flowsheet_path: Optional[str] = None):
        try:
            self.AspenSimulation = win32.gencache.EnsureDispatch("Apwn.Document")
        except Exception:
            try:
                self.AspenSimulation = win32.Dispatch("Apwn.Document")
            except Exception as exc:
                raise RuntimeError(f"Failed to dispatch Aspen COM object: {exc}") from exc

        self.max_iterations = max_iterations
        self.VISIBILITY = VISIBILITY
        self.SUPPRESS = SUPPRESS
        self.duration = None
        self.converged = None

        # Try to set visibility if attribute exists
        try:
            if hasattr(self.AspenSimulation, "Visible"):
                self.AspenSimulation.Visible = bool(VISIBILITY)
            elif hasattr(self.AspenSimulation, "VisibleApp"):
                self.AspenSimulation.VisibleApp = bool(VISIBILITY)
        except Exception:
            pass

        if flowsheet_path:
            self._open_flowsheet(flowsheet_path)

    def _open_flowsheet(self, path: str):
        if not os.path.exists(path):
            raise FileNotFoundError(f"Flowsheet file not found: {path}")
        try:
            if hasattr(self.AspenSimulation, "InitFromArchive"):
                self.AspenSimulation.InitFromArchive(path)
            elif hasattr(self.AspenSimulation, "InitFromFile"):
                self.AspenSimulation.InitFromFile(path)
            elif hasattr(self.AspenSimulation, "Open"):
                self.AspenSimulation.Open(path)
            else:
                raise RuntimeError("No known method to open flowsheet archive on this Aspen COM object.")
        except Exception as exc:
            raise RuntimeError(f"Failed to open flowsheet '{path}': {exc}") from exc

    # Convenience accessors (many Aspen versions use Tree.Elements("Data").Elements("Streams"/"Blocks"))
    @property
    def STRM(self):
        return self.AspenSimulation.Tree.Elements("Data").Elements("Streams")

    @property
    def BLK(self):
        return self.AspenSimulation.Tree.Elements("Data").Elements("Blocks")

    # Stream setters/getters
    def STRM_Temperature(self, Name: str, Temp: float):
        try:
            self.STRM.Elements(Name).Elements("Input").Elements("TEMP").Elements("MIXED").Value = float(Temp)
        except Exception:
            try:
                self.STRM.Elements(Name).Elements("Input").Elements("TEMP").Value = float(Temp)
            except Exception as exc:
                raise RuntimeError(f"Failed to set stream temperature for '{Name}': {exc}")

    def STRM_Pressure(self, Name: str, Pressure: float):
        try:
            self.STRM.Elements(Name).Elements("Input").Elements("PRES").Elements("MIXED").Value = float(Pressure)
        except Exception:
            try:
                self.STRM.Elements(Name).Elements("Input").Elements("PRES").Value = float(Pressure)
            except Exception as exc:
                raise RuntimeError(f"Failed to set stream pressure for '{Name}': {exc}")

    def STRM_Flowrate(self, Name: str, Chemical: str, Flowrate: float):
        try:
            # try common locations
            try:
                self.STRM.Elements(Name).Elements("Input").Elements("MOLEFLMX").Elements(Chemical).Value = float(Flowrate)
                return
            except Exception:
                pass
            try:
                self.STRM.Elements(Name).Elements("Input").Elements("MOLFRAC").Elements(Chemical).Value = float(Flowrate)
                return
            except Exception:
                pass
            # fallback generic
            self.STRM.Elements(Name).Elements("Input").Elements(Chemical).Value = float(Flowrate)
        except Exception as exc:
            raise RuntimeError(f"Failed to set stream flowrate for '{Name}' component '{Chemical}': {exc}")

    def STRM_Get_Outputs(self, Name: str, Chemical: str) -> float:
        try:
            try:
                return float(self.STRM.Elements(Name).Elements("Output").Elements("MOLEFLMX").Elements(Chemical).Elements("MIXED").Value)
            except Exception:
                pass
            try:
                return float(self.STRM.Elements(Name).Elements("Output").Elements("MASSFLOW3").Elements(Chemical).Value)
            except Exception:
                pass
            try:
                return float(self.STRM.Elements(Name).Elements("Output").Elements(Chemical).Value)
            except Exception:
                pass
            raise RuntimeError(f"Component '{Chemical}' not found in outputs for stream '{Name}'.")
        except Exception as exc:
            raise RuntimeError(f"Failed to get stream outputs for '{Name}', component '{Chemical}': {exc}")

    def STRM_Get_Temperature(self, Name: str) -> float:
        try:
            try:
                return float(self.STRM.Elements(Name).Elements("Output").Elements("TEMP_OUT").Elements("MIXED").Value)
            except Exception:
                pass
            try:
                return float(self.STRM.Elements(Name).Elements("Output").Elements("TEMP").Elements("MIXED").Value)
            except Exception:
                pass
            return float(self.STRM.Elements(Name).Elements("Output").Elements("TEMP").Value)
        except Exception as exc:
            raise RuntimeError(f"Failed to read stream temperature for '{Name}': {exc}")

    def STRM_Get_Pressure(self, Name: str) -> float:
        try:
            try:
                return float(self.STRM.Elements(Name).Elements("Output").Elements("PRES_OUT").Elements("MIXED").Value)
            except Exception:
                pass
            try:
                return float(self.STRM.Elements(Name).Elements("Output").Elements("PRES").Elements("MIXED").Value)
            except Exception:
                pass
            return float(self.STRM.Elements(Name).Elements("Output").Elements("PRES").Value)
        except Exception as exc:
            raise RuntimeError(f"Failed to read stream pressure for '{Name}': {exc}")

    # Block setters/getters (default block name B1; adjust in flowsheet or code if different)
    def BLK_NumberOfStages(self, nstages: int):
        try:
            self.BLK.Elements("B1").Elements("Input").Elements("NSTAGE").Value = int(nstages)
        except Exception as exc:
            raise RuntimeError(f"Failed to set number of stages: {exc}")

    def BLK_FeedLocation(self, Feed_Location: int, Feed_Name: str = "S1"):
        try:
            try:
                self.BLK.Elements("B1").Elements("Input").Elements("FEED_STAGE").Elements(Feed_Name).Value = int(Feed_Location)
                return
            except Exception:
                pass
            self.BLK.Elements("B1").Elements("Input").Elements("FEED_STAGE").Value = int(Feed_Location)
        except Exception as exc:
            raise RuntimeError(f"Failed to set feed location: {exc}")

    def BLK_Pressure(self, Pressure: float):
        try:
            self.BLK.Elements("B1").Elements("Input").Elements("PRES1").Value = float(Pressure)
        except Exception:
            try:
                self.BLK.Elements("B1").Elements("Input").Elements("PRES").Value = float(Pressure)
            except Exception as exc:
                raise RuntimeError(f"Failed to set block pressure: {exc}")

    def BLK_RefluxRatio(self, RfxR: float):
        try:
            try:
                self.BLK.Elements("B1").Elements("Input").Elements("BASIS_RR").Value = float(RfxR)
                return
            except Exception:
                pass
            self.BLK.Elements("B1").Elements("Input").Elements("REFLUX").Value = float(RfxR)
        except Exception as exc:
            raise RuntimeError(f"Failed to set reflux ratio: {exc}")

    def BLK_ReboilerRatio(self, RblR: float):
        try:
            try:
                self.BLK.Elements("B1").Elements("Input").Elements("BASIS_BR").Value = float(RblR)
                return
            except Exception:
                pass
            self.BLK.Elements("B1").Elements("Input").Elements("REBOILER").Value = float(RblR)
        except Exception as exc:
            raise RuntimeError(f"Failed to set reboiler ratio: {exc}")

    def BLK_Get_Condenser_Duty(self) -> float:
        try:
            return float(self.BLK.Elements("B1").Elements("Output").Elements("COND_DUTY").Elements("MIXED").Value)
        except Exception:
            try:
                return float(self.BLK.Elements("B1").Elements("Output").Elements("CONDENER_DUTY").Value)
            except Exception:
                raise RuntimeError("Failed to read condenser duty for B1 (check node name for your Aspen version).")

    def BLK_Get_Reboiler_Duty(self) -> float:
        try:
            return float(self.BLK.Elements("B1").Elements("Output").Elements("REB_DUTY").Elements("MIXED").Value)
        except Exception:
            try:
                return float(self.BLK.Elements("B1").Elements("Output").Elements("REBOILER_DUTY").Value)
            except Exception:
                raise RuntimeError("Failed to read reboiler duty for B1 (check node name).")

    def BLK_Get_Column_Stage_Molar_Weights(self, N_stages: int) -> Tuple[float, ...]:
        try:
            arr = []
            for idx in range(1, N_stages + 1):
                try:
                    val = self.BLK.Elements("B1").Elements("Output").Elements("STAGE").Elements("MW").Elements(str(idx)).Value
                    arr.append(float(val))
                    continue
                except Exception:
                    pass
                try:
                    node_path = fr"\Data\Blocks\B1\Output\STAGE\MW\{idx}"
                    val = self.AspenSimulation.Tree.FindNode(node_path).Value
                    arr.append(float(val))
                except Exception:
                    arr.append(0.0)
            return tuple(arr)
        except Exception as exc:
            raise RuntimeError(f"Failed to get per-stage molar weights: {exc}")

    def BLK_Get_Column_Stage_Temperatures(self, N_stages: int) -> Tuple[float, ...]:
        try:
            arr = []
            for idx in range(1, N_stages + 1):
                try:
                    val = self.BLK.Elements("B1").Elements("Output").Elements("STAGE").Elements("TEMP").Elements(str(idx)).Value
                    arr.append(float(val))
                    continue
                except Exception:
                    pass
                try:
                    node_path = fr"\Data\Blocks\B1\Output\STAGE\TEMP\{idx}"
                    val = self.AspenSimulation.Tree.FindNode(node_path).Value
                    arr.append(float(val))
                except Exception:
                    arr.append(0.0)
            return tuple(arr)
        except Exception as exc:
            raise RuntimeError(f"Failed to get per-stage temperatures: {exc}")

    def BLK_Get_Column_Stage_Vapor_Flows(self, N_stages: int) -> Tuple[float, ...]:
        try:
            arr = []
            for idx in range(1, N_stages + 1):
                try:
                    val = self.BLK.Elements("B1").Elements("Output").Elements("STAGE").Elements("VAPOR").Elements(str(idx)).Value
                    arr.append(float(val))
                    continue
                except Exception:
                    pass
                try:
                    node_path = fr"\Data\Blocks\B1\Output\STAGE\VAPOR\{idx}"
                    val = self.AspenSimulation.Tree.FindNode(node_path).Value
                    arr.append(float(val))
                except Exception:
                    arr.append(0.0)
            return tuple(arr)
        except Exception as exc:
            raise RuntimeError(f"Failed to get per-stage vapor flows: {exc}")

    def Run(self):
        start = time.time()
        try:
            if hasattr(self.AspenSimulation, "Engine") and hasattr(self.AspenSimulation.Engine, "Run2"):
                self.AspenSimulation.Engine.Run2()
            elif hasattr(self.AspenSimulation, "Engine") and hasattr(self.AspenSimulation.Engine, "Run"):
                self.AspenSimulation.Engine.Run()
            elif hasattr(self.AspenSimulation, "Run2"):
                self.AspenSimulation.Run2()
            elif hasattr(self.AspenSimulation, "Run"):
                self.AspenSimulation.Run()
            else:
                raise RuntimeError("No Run entrypoint on Aspen COM object.")
            self.duration = time.time() - start
            try:
                conv = getattr(self.AspenSimulation, "Converged", None)
                self.converged = bool(conv) if conv is not None else True
            except Exception:
                self.converged = True
        except Exception as exc:
            self.duration = time.time() - start
            self.converged = False
            tb = traceback.format_exc()
            raise RuntimeError(f"Aspen run failed: {exc}\n{tb}")
