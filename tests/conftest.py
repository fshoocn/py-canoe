import sys
import pytest


def is_canoe_available() -> bool:
    """Check if CANoe COM interface is available."""
    if sys.platform != "win32":
        return False

    try:
        import win32com.client
        win32com.client.Dispatch("CANoe.Application")
        return True
    except Exception:
        return False


skip_if_no_canoe = pytest.mark.skipif(not is_canoe_available(), reason="CANoe requires Windows and COM interface")
