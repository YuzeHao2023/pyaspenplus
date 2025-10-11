import platform

def ensure_windows():
    if platform.system() != "Windows":
        raise RuntimeError("COM backend requires Windows and Aspen Plus installed.")
