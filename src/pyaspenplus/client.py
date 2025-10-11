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
    """A COM backend that uses pywin32 to control Aspen Plus (Windows only).
    The actual method names/attributes depend on the Aspen COM API. This implementation
    provides a template — you must adapt attribute/method names to your Aspen version.
    """
    def __init__(self, progid: str = "AspenPlus.Application"):
        if win32com is None:
            raise AspenPlusError("pywin32 (win32com) is required for COM backend.")
        self.progid = progid
        self.app = None
        self.doc = None

    def connect(self):
        utils.ensure_windows()
        try:
            self.app = win32com.client.Dispatch(self.progid)
            # The actual object model differs by Aspen versions; doc below is placeholder
            self.doc = getattr(self.app, "ActiveDocument", None)
            return self
        except Exception as exc:
            raise AspenPlusError(f"Failed to dispatch COM object '{self.progid}': {exc}")

    def open_case(self, path: str):
        if not os.path.exists(path):
            raise AspenPlusError(f"Case file not found: {path}")
        # The actual COM method for opening a case depends on Aspen's object model.
        try:
            # Placeholder call - adapt to real API of your Aspen installation
            if hasattr(self.app, "OpenCase"):
                self.app.OpenCase(path)
            else:
                # Try common attribute names - adapt as needed
                if hasattr(self.app, "Open"):
                    self.app.Open(path)
                else:
                    raise AspenPlusError("COM backend: Open method not found. Adjust code for your Aspen version.")
        except Exception as exc:
            raise AspenPlusError(f"Failed to open case via COM: {exc}")

    def run(self):
        try:
            if hasattr(self.app, "Run"):
                self.app.Run()
            elif self.doc and hasattr(self.doc, "Run"):
                self.doc.Run()
            else:
                raise AspenPlusError("COM backend: Run method not found. Adjust code for your Aspen version.")
        except Exception as exc:
            raise AspenPlusError(f"Failed to run simulation via COM: {exc}")

    def get_streams(self):
        # Translate from Aspen COM object model to Stream dataclass.
        # This is highly dependent on Aspen COM API. This function should be adapted.
        streams = []
        try:
            # Example pseudo-code: iterate through material streams collection
            streams_collection = getattr(self.doc, "MaterialStreams", None)
            if streams_collection is None:
                # try alternative attribute names
                streams_collection = getattr(self.doc, "Streams", None)
            if streams_collection is None:
                raise AspenPlusError("COM backend: could not find streams collection on document.")

            for i in range(1, streams_collection.Count + 1):
                item = streams_collection.Item(i)
                name = getattr(item, "Name", f"Stream_{i}")
                flow = getattr(item, "TotalFlow", 0.0)
                temp = getattr(item, "Temperature", None)
                pres = getattr(item, "Pressure", None)
                # composition mapping will depend on object model
                comp = {}
                # try to read composition if present
                if hasattr(item, "Composition"):
                    comp_obj = getattr(item, "Composition")
                    # pseudo-code to iterate components
                    try:
                        for j in range(1, comp_obj.Count + 1):
                            comp_item = comp_obj.Item(j)
                            comp[comp_item.Name] = getattr(comp_item, "MoleFraction", 0.0)
                    except Exception:
                        comp = {}
                streams.append(Stream(name=name, flow=flow, temperature=temp, pressure=pres, composition=comp))
            return streams
        except AspenPlusError:
            raise
        except Exception as exc:
            raise AspenPlusError(f"Failed to read streams via COM backend: {exc}")

    def set_stream(self, name: str, stream: Stream):
        # Adapt to Aspen COM API: find stream by name and set properties
        try:
            streams_collection = getattr(self.doc, "MaterialStreams", None) or getattr(self.doc, "Streams", None)
            if streams_collection is None:
                raise AspenPlusError("COM backend: streams collection not found.")
            # find stream by name
            found = None
            for i in range(1, streams_collection.Count + 1):
                item = streams_collection.Item(i)
                if getattr(item, "Name", "") == name:
                    found = item
                    break
            if found is None:
                raise AspenPlusError(f"Stream {name} not found.")
            # set properties - actual attribute names will vary
            if hasattr(found, "TotalFlow"):
                found.TotalFlow = stream.flow
            if stream.temperature is not None and hasattr(found, "Temperature"):
                found.Temperature = stream.temperature
            if stream.pressure is not None and hasattr(found, "Pressure"):
                found.Pressure = stream.pressure
            # composition setting: needs mapping to Aspen model
        except Exception as exc:
            raise AspenPlusError(f"Failed to set stream via COM: {exc}")

    def save(self, path: Optional[str] = None):
        try:
            if path:
                if hasattr(self.doc, "SaveAs"):
                    self.doc.SaveAs(path)
                else:
                    raise AspenPlusError("COM backend: SaveAs not available.")
            else:
                if hasattr(self.doc, "Save"):
                    self.doc.Save()
            return path
        except Exception as exc:
            raise AspenPlusError(f"Failed to save case via COM: {exc}")

    def close(self):
        try:
            if self.doc and hasattr(self.doc, "Close"):
                self.doc.Close()
        except Exception:
            pass
        finally:
            self.doc = None
            self.app = None

class AspenPlusClient:
    def __init__(self, backend: str = "com", progid: Optional[str] = None):
        """
        backend: 'com' or 'mock'
        progid: Optional COM ProgID for Aspen (only used for com backend)
        """
        self.backend_name = backend
        self.backend = None
        if backend == "mock":
            self.backend = MockBackend()
        elif backend == "com":
            self.backend = COMBackend(progid=progid or "AspenPlus.Application")
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
