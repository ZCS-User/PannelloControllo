from __future__ import annotations
import pyvisa
from typing import Optional

class ACSource:
    """Driver minimale per sorgente AC via VISA."""

    def __init__(self, resource: str, timeout_ms: int = 2000):
        self.resource = resource
        self.rm: Optional[pyvisa.ResourceManager] = None
        self.inst = None
        self.timeout_ms = timeout_ms
        self._connect()

    def _connect(self):
        self.rm = pyvisa.ResourceManager()
        self.inst = self.rm.open_resource(self.resource)
        self.inst.timeout = self.timeout_ms

    def configure(self, vrms: float, freq: float, phases: str = "mono") -> bool:
        self.inst.write('SYST:FUNC ONE' if phases.lower().startswith('mono') else 'SYST:FUNC THREE')
        self.inst.write(f'VOLT {vrms}')
        self.inst.write(f'FREQ {freq}')
        return True

    def turn_on(self) -> bool:
        self.inst.write('OUTP ON')
        return True

    def turn_off(self) -> bool:
        self.inst.write('OUTP OFF')
        return True

    def close(self):
        try:
            if self.inst is not None:
                self.inst.close()
        finally:
            if self.rm is not None:
                try: self.rm.close()
                except: pass
