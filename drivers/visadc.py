# drivers/visadc.py
from __future__ import annotations
import pyvisa
from typing import Optional


class DCSource:
    """Driver minimale per alimentatori DC (profilo solare) via VISA.
     Adatta i comandi SCPI ai tuoi strumenti.
     """
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

    # --- Configurazione curva I-V (esempio tipico da solar simulator) ---
    def set_iv(self, voc: float, isc: float, ff: float = 0.9) -> bool:
        if not (0 < ff < 1):
            raise ValueError("ff deve essere tra 0 e 1")
        self.inst.write('SOL:USER:VOC '+str(round(voc, 2)))
        self.inst.write('SOL:USER:VMP '+str(round(voc * ff, 2)))
        self.inst.write('SOL:USER:ISC '+str(round(isc, 2)))
        self.inst.write('SOL:USER:IMP '+str(round(isc * ff, 2)))
        return True

    def turn_on(self) -> bool:
        self.inst.write('OUTP 1')
        return True

    def turn_off(self) -> bool:
        self.inst.write('OUTP 0')
        return True

    def measure(self) -> dict:
        """Esempio: misura V/I/P se supportato (adatta ai tuoi comandi)."""
        try:
            v = float(self.inst.query('MEAS:VOLT?'))
            i = float(self.inst.query('MEAS:CURR?'))
            p = v * i
            return {"V": v, "I": i, "P": p}
        except Exception:
            return {}

    def close(self):
        try:
            if self.inst is not None:
                self.inst.close()
        finally:
            if self.rm is not None:
                try:
                    self.rm.close()
                except:
                    pass

    # --------- Aggiunte per config ITECH ---------
    def identify(self) -> str:
        """Ritorna la stringa *IDN? dell'alimentatore."""
        try:
            return self.inst.query("*IDN?").strip()
        except Exception as e:
            return f"UNKNOWN ({e})"

    def configure_itech_solar(self, curve_mode: str = "DEF_C") -> dict:
        '''
        Configura modalità 'SOL' e 'SOL:MODE <curve_mode>' (es. DEF_C) sugli ITECH.
        Restituisce un dict con IDN e modi letti da query.
        '''
        idn = self.identify()
         # imposta FUNZIONE SOLARE
        self.inst.write("FUNC:MODE SOL")
        func_mode = self.inst.query("FUNC:MODE?").strip()
        # imposta curva/algoritmo solare (es. DEF_C)
        self.inst.write(f"SOL:MODE {curve_mode}")
        sol_mode = self.inst.query("SOL:MODE?").strip()
        # echo stile tuo snippet
        print(f">>> {idn}")
        print(f">>> {func_mode}")
        print(f">>> {sol_mode}")
        return {"idn": idn, "func_mode": func_mode, "solar_mode": sol_mode}