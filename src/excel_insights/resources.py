"""Best-effort local memory checks; no external services or telemetry."""
import sys
from pathlib import Path


def available_memory() -> int | None:
    try:
        if sys.platform == "win32":
            import ctypes

            class MemoryStatus(ctypes.Structure):
                _fields_ = [("length", ctypes.c_uint32), ("load", ctypes.c_uint32),
                            *[(name, ctypes.c_uint64) for name in (
                                "total_physical", "available_physical", "total_pagefile", "available_pagefile",
                                "total_virtual", "available_virtual", "available_extended")]]

            status = MemoryStatus()
            status.length = ctypes.sizeof(status)
            if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
                return status.available_physical
        elif sys.platform.startswith("linux"):
            values = dict(line.split(":", 1) for line in Path("/proc/meminfo").read_text().splitlines())
            available = int(values["MemAvailable"].split()[0]) * 1024
            maximum = Path("/sys/fs/cgroup/memory.max")
            current = Path("/sys/fs/cgroup/memory.current")
            if maximum.exists() and current.exists():
                limit = maximum.read_text().strip()
                if limit != "max":
                    available = min(available, max(0, int(limit) - int(current.read_text())))
            return available
    except (OSError, ValueError, KeyError, AttributeError):
        pass
    return None


def ensure_import_memory(extra_bytes: int = 0):
    available = available_memory()
    if available is not None and available < extra_bytes + 256 * 1024 * 1024:
        import gc
        gc.collect()
        available = available_memory()
    if available is not None and available < extra_bytes + 256 * 1024 * 1024:
        raise MemoryError("可用記憶體不足以完成匯入。請只選需要的工作表、分批匯入，或關閉其他程式後重試；資料未截斷匯入。")
