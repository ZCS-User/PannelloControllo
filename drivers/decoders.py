# drivers/decoders.py

from typing import Any, Optional

def _to_float(x: Any, default: float = 1.0) -> float:
    """
    Converte x in float in modo robusto (accetta '0,1' e '0.1').
    """
    try:
        s = str(x).strip().replace(",", ".")
        return float(s)
    except Exception:
        return default

def decode_u16_auto(raw: Optional[int], scale: Any = 1.0, signed_hint_thresh: int = 0xF000) -> Optional[float]:
    """
    Decodifica un singolo registro 16-bit:
      - None → None
      - applica two's complement (signed 16) se raw >= signed_hint_thresh (es. 0xF000)
      - applica lo scaling (moltiplicativo)

    Ritorna float (o None).
    """
    if raw is None:
        return None
    v = int(raw) & 0xFFFF
    if v >= int(signed_hint_thresh) & 0xFFFF:
        v = v - 0x10000  # two's complement su 16 bit
    sc = _to_float(scale, 1.0)
    return v * sc


# ---- Stubs futuri (quando/Se aggiungerai tipi in template) ----
# def decode_s16(raw, scale=1.0): ...
# def decode_u32(raw_hi, raw_lo, scale=1.0, word_order="msw_first"): ...
# def decode_f32(raw_hi, raw_lo, word_order="msw_first"): ...
# def decode_ascii(raw_regs, bytes_per_reg=2, strip_null=True): ...
