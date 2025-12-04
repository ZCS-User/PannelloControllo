from dataclasses import dataclass
from typing import Dict, Optional, List, Union
from .visadc import DCSource
from .visaac import ACSource
from .modbus_inv import Inverter
import time

@dataclass
class InverterNode:
    name: str          # es. "INV1"
    role: str          # "master" | "slave"
    alimentatore: str  # "Nessuno" | "DC1" | ...
    driver: Inverter

class Instruments:
    """Facciata unica per DC/AC e più Inverter."""

    def __init__(self, dc_map: Dict[str, str] = None, ac_addr: Optional[str] = None,
                 inv_cfgs: List[dict] = None, protocol: Optional[str] = None):
        self.dc = {name: DCSource(addr) for name, addr in (dc_map or {}).items()}
        self.ac = ACSource(ac_addr) if ac_addr else None

        self.inverters: Dict[str, InverterNode] = {}
        if inv_cfgs:
            for idx, cfg in enumerate(inv_cfgs, start=1):
                name = cfg.get("name", f"INV{idx}")
                role = "slave" if cfg.get("is_slave") else "master"
                alimentatore = cfg.get("alimentatore", "Nessuno")
                proto = protocol or cfg.get("proto") or "TCP"
                try:
                    if proto in ("TCP", "AzzurroHUB"):
                        drv = Inverter(proto=proto, ip=cfg["address"], slave=int(cfg["modbus"]))
                    elif proto == "RTU":
                        drv = Inverter(proto="RTU", com=cfg["address"], slave=int(cfg["modbus"]))
                    else:
                        raise ValueError(f"Protocollo non supportato: {proto}")
                    self.inverters[name] = InverterNode(
                        name=name, role=role, alimentatore=alimentatore, driver=drv
                    )
                except Exception as e:
                    print(
                        f"[WARN] Inverter '{name}' non disponibile ({cfg.get('address')}): {e} — continuo con gli altri.")

    # ---- DC ----

    def dc_measure(self, name): return self.dc[name].measure()
    def dc_set_iv(self, name, voc, isc, ff=1.0):
        drv = self.dc.get(name);
        if not drv:
            print(f"[WARN] DC {name} non disponibile")
            return False
        return drv.set_iv(voc, isc, ff)
    def dc_on(self, name):
        drv = self.dc.get(name)
        if not drv:
            print(f"[WARN] DC {name} non disponibile")
            return False
        return drv.turn_on()
    def dc_off(self, name):
        drv = self.dc.get(name)
        if not drv:
            print(f"[WARN] DC {name} non disponibile")
            return False
        return drv.turn_off()
    def dc_safe_quench_and_off(self, name, target_v=200.0, target_i=1.0):
        drv = self.dc.get(name)
        if not drv:
            print(f"[WARN] DC {name} non disponibile")
            return False
        drv.set_iv(target_v, target_i)
        time.sleep(5)
        return drv.turn_off()
    # ========== Rilascia i client Modbus (libera COM/IP) ==========
    def inv_disconnect_all(self):
        for name, node in self.inverters.items():
            try:
                node.driver.close()
            except Exception as e:
                print(f"[WARN] close inverter {name}: {e}")

    # ---- AC ----
    def ac_set(self, vrms, freq, phases="mono"): return self.ac.configure(vrms, freq, phases) if self.ac else None
    def ac_on(self): return self.ac.turn_on() if self.ac else None
    def ac_off(self): return self.ac.turn_off() if self.ac else None

    # ---- Inverter ----
    def inv_read(self, inv_name: str, reg: Union[int,str], count=1):
        return self.inverters[inv_name].driver.read(reg, count)

    def inv_write(self, inv_name: str, reg: Union[int,str], values: Union[int,List[int]], scale=1):
        return self.inverters[inv_name].driver.write(reg, values, scale=scale)

    def inv_names(self, role: Optional[str] = None) -> List[str]:
        if role is None: return list(self.inverters.keys())
        return [n for n,node in self.inverters.items() if node.role==role]

    def inv_broadcast_write(self, reg, values, scale=1, role: Optional[str] = None):
        return {n: self.inverters[n].driver.write(reg, values, scale=scale) for n in self.inv_names(role)}

    def inv_broadcast_read(self, reg, count=1, role: Optional[str] = None):
        return {n: self.inverters[n].driver.read(reg, count) for n in self.inv_names(role)}

    def close_all(self):
        for s in self.dc.values():
            try: s.close()
            except: pass
        if self.ac:
            try: self.ac.close()
            except: pass
        for node in self.inverters.values():
            try: node.driver.close()
            except: pass

    # ---- DC: configurazione ITECH (SOL/DEF_C) ----
    def dc_config_itech(self, name: str, curve_mode: str = "DEF_C"):
        """
        Imposta l'alimentatore DC 'name' in modalità solare ITECH (FUNC:MODE SOL, SOL:MODE <curve_mode>).
        """
        return self.dc[name].configure_itech_solar(curve_mode=curve_mode)
