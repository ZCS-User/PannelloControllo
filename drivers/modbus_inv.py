from __future__ import annotations
from pymodbus.client import ModbusTcpClient, ModbusSerialClient
import serial
from serial.tools import list_ports
from pymodbus.exceptions import ModbusIOException
import time
from typing import Optional, List, Union
import threading

class Inverter:
    """Driver Modbus per inverter (TCP / RTU / AzzurroHUB)."""

    def __init__(self, proto: str, ip: Optional[str] = None, port: Optional[int] = None,
                 com: Optional[str] = None, slave: int = 1, timeout: float = 1.0):
        self.proto = proto
        self.slave = int(slave)
        self.timeout = timeout
        self.client = None

        if proto == "TCP":
            self.client = ModbusTcpClient(ip, port=port or 8899, timeout=timeout)
        elif proto == "AzzurroHUB":
            self.client = ModbusTcpClient(ip, port=port or 55400, timeout=timeout)
        elif proto == "RTU":
            # Preflight: la porta esiste ed è apribile?
            self._preflight_serial(com)
            self.client = ModbusSerialClient(
            port = com, baudrate = 9600, timeout = timeout, stopbits = 1, bytesize = 8, parity = 'N')
        else:
            raise ValueError("Protocollo non supportato")

        # Primo tentativo di connessione (non bloccante)
        self._lock = threading.RLock()
        try:
            self.client.connect()
        except Exception as e:
            print(f"[WARN] Connessione Modbus iniziale fallita: {e}")

    def _ensure_connected(self) -> bool:
        """Assicura connessione aperta; ritorna True se ok, False se non riuscito."""
        try:
            # il client pymodbus non espone un flag affidabile, richiamiamo connect()
            return bool(self.client and self.client.connect())
        except Exception as e:
            print(f'[WARN] Connessione Modbus non disponibile: {e}')
            return False

    def _preflight_serial(self, port: str):
        # 1) la porta esiste?
        ports = {p.device for p in list_ports.comports()}
        if port not in ports:
            print(f'[WARN] Porta seriale non trovata: {port}')
            return False
        # 2) è apribile (non lockata)? — in caso di accesso negato, non alzare: logga e prosegui
        try:
            s = serial.Serial(port=port, timeout=0.1)
            s.close()
            return True
        except Exception as e:
            print(f'[WARN] Porta seriale occupata o non apribile: {port} ({e}) — riprovo al primo utilizzo.')
            return False

    def _retry(self, fn, attempts: int = 3, delay: float = 0.3):
        last = None
        for _ in range(attempts):
            try:
                res = fn()
                if res and not isinstance(res, ModbusIOException):
                    return res
            except Exception as e:
                last = e
            time.sleep(delay)
        if last:
            raise last
        raise IOError("Operazione Modbus fallita")

    def read(self, reg: Union[int, str], count: int = 1) -> Optional[List[int]]:
        r = int(reg, 16) if isinstance(reg, str) else int(reg)
        if not self._ensure_connected():
            return None
        with self._lock:
            rr = self._retry(lambda: self.client.read_holding_registers(address=r, count=count, slave=self.slave))
        return rr.registers if hasattr(rr, 'registers') else None

    def write(self, reg: Union[int, str], values: Union[int, List[int]], scale: int = 1) -> bool:
        r = int(reg, 16) if isinstance(reg, str) else int(reg)
        if not self._ensure_connected():
            return False
        with self._lock:
            if isinstance(values, list):
                scaled = [int(v / scale) for v in values]
                self._retry(lambda: self.client.write_registers(r, scaled, slave=self.slave))
            else:
                self._retry(lambda: self.client.write_registers(r, [int(values / scale)], slave=self.slave))
        return True

    def close(self):
        try:
            if self.client:
                self.client.close()
        except Exception:
            pass
