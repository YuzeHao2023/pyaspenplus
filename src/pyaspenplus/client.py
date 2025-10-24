"""
AspenPlusClient: lightweight, backend-abstracted client for interacting with Aspen Plus.

Backends:
- 'com': uses pywin32 to talk to Aspen's COM API (Windows only). You must provide a valid progid or use default.
- 'mock': in-memory fake backend used for development / testing.

Note:
The actual COM ProgID for Aspen Plus may differ depending on version. Supply `progid` if necessary.
"""
from __future__ import annotations
from contextlib import contextmanager
from typing import Optional, Dict, List
import os

from .exceptions import AspenPlusError
from .models import Stream
from . import utils

# Optional import for Windows COM
try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None  # type: ignore

class BaseBackend:
    def connect(self):
        raise NotImplementedError

    def open_case(self, path: str):
        raise NotImplementedError

    def run(self):
        raise NotImplementedError

    def get_streams(self) -> List[Stream]:
        raise NotImplementedError

    def set_stream(self, name: str, stream: Stream):
        raise NotImplementedError

    def save(self, path: Optional[str] = None):
        raise NotImplementedError

    def close(self):
        raise NotImplementedError

class MockBackend(BaseBackend):
    """A simple backend used for development and CI without Aspen installed."""
    def __init__(self):
        self._streams: Dict[str, Stream] = {}
        self._open = False
        self._case = None

    def connect(self):
        self._open = True
        return self

    def open_case(self, path: str):
        self._case = path
        # populate with an example stream
        self._streams = {
            "F1": Stream(name="F1", flow=100.0, temperature=300.0, pressure=101325.0, composition={"H2O":1.0}),
            "F2": Stream(name="F2", flow=50.0, temperature=310.0, pressure=101325.0, composition={"Ethanol":1.0}),
        }

    def run(self):
        # pretend we ran; maybe change some stream values
        for s in self._streams.values():
            s.flow *= 1.0  # no-op

    def get_streams(self):
        return list(self._streams.values())

    def set_stream(self, name: str, stream: Stream):
        self._streams[name] = stream

    def save(self, path: Optional[str] = None):
        return self._case or path

    def close(self):
        self._open = False

class COMBackend(BaseBackend):
    """
    A COM backend that delegates to pyaspenplus.api.com_simulation.Simulation for Aspen interactions.

    Notes:
    - This implementation expects src/pyaspenplus/api/com_simulation.py to provide a Simulation class
      with methods: _open_flowsheet/_open? (we use InitFromArchive via Simulation constructor in aspen_api),
      STRM_Temperature/STRM_Pressure/STRM_Flowrate, Run, STRM_Get_* and Block methods as applicable.
    - This avoids duplicating low-level COM traversal code in client.
    """
    def __init__(self, progid: str = None, flowsheet_path: Optional[str] = None):
        if win32com is None:
            raise AspenPlusError("pywin32 (win32com) is required for COM backend.")
        # defer import until needed (so package can be imported on non-Windows for tests)
        from .api.com_simulation import Simulation  # local import
        self.progid = progid
        self.simulation: Optional[Simulation] = None
        self.flowsheet_path = flowsheet_path

    def connect(self):
        utils.ensure_windows()
        # instantiate the Simulation wrapper; it will attempt to dispatch the Aspen COM object
        try:
            from .api.com_simulation import Simulation
            # The Simulation constructor will call EnsureDispatch("Apwn.Document")
            self.simulation = Simulation(VISIBILITY=False, SUPPRESS=True, flowsheet_path=self.flowsheet_path)
            return self
        except Exception as exc:
            raise AspenPlusError(f"Failed to create Aspen Simulation COM object: {exc}")

    def open_case(self, path: str):
        if self.simulation is None:
            raise AspenPlusError("COM backend not connected.")
        # Ensure flowsheet opened
        self.simulation._open_flowsheet(path)

    def run(self):
        if self.simulation is None:
            raise AspenPlusError("COM backend not connected.")
        try:
            self.simulation.Run()
        except Exception as exc:
            raise AspenPlusError(f"Failed to run simulation via COM: {exc}")

    def get_streams(self):
        if self.simulation is None:
            raise AspenPlusError("COM backend not connected.")
        # Build Stream list by reading a few known streams (best-effort). Users should customize per flowsheet.
        res = []
        # Example: attempt to read a few known stream names; if your flowsheet differs, modify this method.
        for name in ("F1", "F2", "S1", "S2", "S3"):
            try:
                temp = None
                pres = None
                flow = None
                try:
                    temp = self.simulation.STRM_Get_Temperature(name)
                except Exception:
                    temp = None
                try:
                    pres = self.simulation.STRM_Get_Pressure(name)
                except Exception:
                    pres = None
                # Try to read ethane as example composition; adapt as needed
                comp = {}
                try:
                    eth = self.simulation.STRM_Get_Outputs(name, "ETHANE")
                    comp["ETHANE"] = eth
                except Exception:
                    comp = {}
                # flow: sum of components if available (this is heuristic)
                try:
                    flow = sum(comp.values()) if comp else None
                except Exception:
                    flow = None
                if temp is not None or pres is not None or flow is not None or comp:
                    res.append(Stream(name=name, flow=flow or 0.0, temperature=temp, pressure=pres, composition=comp))
            except Exception:
                continue
        return res

    def set_stream(self, name: str, stream: Stream):
        if self.simulation is None:
            raise AspenPlusError("COM backend not connected.")
        try:
            if stream.temperature is not None:
                self.simulation.STRM_Temperature(name, float(stream.temperature))
            if stream.pressure is not None:
                self.simulation.STRM_Pressure(name, float(stream.pressure))
            if stream.composition:
                for comp_name, val in stream.composition.items():
                    # direct set; ensure comp_name matches Aspen component id in your flowsheet
                    try:
                        self.simulation.STRM_Flowrate(name, comp_name, float(val))
                    except Exception:
                        # ignore missing comps
                        pass
        except Exception as exc:
            raise AspenPlusError(f"Failed to set stream via COM: {exc}")

    def save(self, path: Optional[str] = None):
        # Try to call Save/SaveAs on underlying Aspen document if exposed (not guaranteed)
        try:
            if self.simulation and hasattr(self.simulation.AspenSimulation, "SaveAs") and path:
                self.simulation.AspenSimulation.SaveAs(path)
                return path
            elif self.simulation and hasattr(self.simulation.AspenSimulation, "Save"):
                self.simulation.AspenSimulation.Save()
                return None
            else:
                return path
        except Exception as exc:
            raise AspenPlusError(f"Failed to save case via COM: {exc}")

    def close(self):
        try:
            if self.simulation and hasattr(self.simulation.AspenSimulation, "Close"):
                try:
                    self.simulation.AspenSimulation.Close()
                except Exception:
                    pass
        finally:
            self.simulation = None

class AspenPlusClient:
    def __init__(self, backend: str = "com", progid: Optional[str] = None, flowsheet_path: Optional[str] = None):
        """
        backend: 'com' or 'mock'
        progid: Optional COM ProgID for Aspen (only used for com backend)
        flowsheet_path: optional default flowsheet file path for COM backend
        """
        self.backend_name = backend
        self.backend = None
        if backend == "mock":
            self.backend = MockBackend()
        elif backend == "com":
            self.backend = COMBackend(progid=progid, flowsheet_path=flowsheet_path)
        else:
            raise AspenPlusError("Unknown backend: choose 'com' or 'mock'")

    @contextmanager
    def connect(self):
        """Context manager to connect and automatically close."""
        try:
            self.backend.connect()
            yield self
        finally:
            try:
                self.backend.close()
            except Exception:
                pass

    # Convenience wrapper methods
    def open_case(self, path: str):
        return self.backend.open_case(path)

    def run(self):
        return self.backend.run()

    def get_streams(self):
        return self.backend.get_streams()

    def set_stream(self, name: str, stream: Stream):
        return self.backend.set_stream(name, stream)

    def save(self, path: Optional[str] = None):
        return self.backend.save(path)

    def close(self):
        return self.backend.close()
