from tkinter import messagebox
import re
import pandas as pd  # per convertire stringa in lista, se serve
import os
import threading
import time
from .instruments import *
import ast

logging_thread = None
test_thread = None
logging_running = False
logging_paused = False


def _as_py(x):
    # Se in Excel hai messo "[0x1000, 0x1001]" o "[[1,2],[3,4]]", lo converto da stringa a lista
    if isinstance(x, str):
        s = x.strip()
        if s.startswith('[') or s.startswith('(') or s.startswith('{'):
            try: return ast.literal_eval(s)
            except: return x
    return x


def apply_template_writes(ins, role, regs, vals, scale=1):
    regs = _as_py(regs)
    vals = _as_py(vals)

    # Caso A: reg singolo
    if not isinstance(regs, (list, tuple)):
        ins.inv_broadcast_write(regs, vals, scale=scale, role=role)
        return

    # Caso B: reg[] + valore scalare -> ripeti
    if not isinstance(vals, (list, tuple)):
        for r in regs:
            ins.inv_broadcast_write(r, vals, scale=scale, role=role)
        return

    # Caso C: reg[] + value[] (scalari)
    if all(not isinstance(v, (list, tuple)) for v in vals):
        if len(vals) != len(regs):
            # fallback: usa il primo valore per tutti
            for r in regs:
                ins.inv_broadcast_write(r, vals[0], scale=scale, role=role)
        else:
            for r, v in zip(regs, vals):
                ins.inv_broadcast_write(r, v, scale=scale, role=role)
        return

    # Caso D: reg[] + value[][] (blocchi)
    if len(vals) != len(regs):
        raise ValueError("value[][] deve avere la stessa lunghezza di reg[]")
    for r, block in zip(regs, vals):
        ins.inv_broadcast_write(r, block, scale=scale, role=role)


def build_inv_cfgs_from_ui(protocol: str, inverter_data: List[dict]) -> List[dict]:
    """Converte la lista raccolta dal pannello Log in una lista di cfg per Instruments.
    Atteso "inverter_data" nel formato già usato dal tuo codice:
    {"ip"|"address": str, "modbus": int, "slave": bool, "alimentatore": str}
    """
    inv_cfgs = []
    for idx, inv in enumerate(inverter_data, start=1):
        inv_cfgs.append({
            "name": f"INV{idx}",
            "is_slave": bool(inv.get("slave")),
            "alimentatore": inv.get("alimentatore", "Nessuno"),
            "modbus": int(inv["modbus"]),
            "address": inv.get("ip") or inv.get("address"),  # IP per TCP/HUB, COM per RTU
            # Nota: il protocollo è globale nel tuo UI; se in futuro sarà per-inverter,
            # aggiungi qui "proto": inv.get("proto")
        })
    return inv_cfgs


def parse_sn(sn: str):
    """Estrae informazioni strutturate dallo SN.
    Ritorna un dict con: length, raw, family_prefix, family, model_code.
    Se lo SN non è lungo 14/20, ritorna None.
    """
    if not sn:
        return None
    s = str(sn).strip()
    # Accetta alfanumerico; se serve solo numerico, cambia il pattern
    if not re.fullmatch(r"[A-Za-z0-9]{14}|[A-Za-z0-9]{20}", s):
        return None

    if len(s) == 14:
        family_prefix = s[:3]
        family = s[:3]
        model_code = s[3:8]
    elif len(s) == 20:
        family_prefix = s[:6]
        family = s[:6]
        model_code = s[6:9]
    else:
        return None

    return {
        "length": len(s),
        "raw": s,
        "family_prefix": family_prefix,
        "family": family,
        "model_code": model_code,
        }


# Funzione Test
def run_test_from_template(template_file_path, sn, protocol, inverter_data, shared_ins=None):
    info = parse_sn(sn)
    if not info:
        messagebox.showerror("SN non valido", f"Seriale '{sn}' non riconosciuto.")
        return

    family = info["family"]          # es. "ZH1050"
    model_code = info["model_code"]  # es. "050"

    # File DB in base alla family
    db_path = os.path.join("./database", f"{family}.xlsx")
    df_inverter = None
    if os.path.exists(db_path):
        df_inverter = pd.read_excel(db_path, dtype={"Unnamed: 0": str}, keep_default_na=False)
        key_col = "Unnamed: 0" if "Unnamed: 0" in df_inverter.columns else df_inverter.columns[0]
        df_inverter = df_inverter.loc[df_inverter[key_col].astype(str).str.strip() == model_code]
        if df_inverter.empty:
            messagebox.showwarning("Modello non trovato",
                                   f"Nessuna riga in {os.path.basename(db_path)} con {key_col}='{model_code}'")
    else:
        messagebox.showwarning("DB non trovato", f"File database assente: {db_path}")

    def test_logic():
        print(f"[TEST TEMPLATE] Avvio test da template: {template_file_path}")
        if df_inverter is not None and not df_inverter.empty:
            params = df_inverter.iloc[0].to_dict()
            print("[DB] Parametri modello:", params)

        # mapping reale o letto da config
        inv_cfgs = build_inv_cfgs_from_ui(protocol, inverter_data)

        dc_map = {"DC1": "ASRL20::INSTR", "DC2": "ASRL21::INSTR", "DC3": "ASRL22::INSTR"}
        ac_addr = "ASRL5::INSTR"
        ins = shared_ins or Instruments(dc_map=dc_map, ac_addr=ac_addr, inv_cfgs=inv_cfgs, protocol=protocol)
        df_template = pd.read_excel(template_file_path.split('.xlsx')[0]+'.xlsx')
        list_dc = list()
        dc1_yes = ''
        dc2_yes = ''
        dc3_yes = ''
        for _i in df_template['on/off DC1']:
            if _i != 'no':
                dc1_yes = 'yes'
                break
        for _i in df_template['on/off DC2']:
            if _i != 'no':
                dc2_yes = 'yes'
                break
        for _i in df_template['on/off DC3']:
            if _i != 'no':
                dc3_yes = 'yes'
                break
        if 'DC1' in template_file_path or dc1_yes == 'yes':
            list_dc.append('DC1')
        if 'DC2' in template_file_path or dc2_yes == 'yes':
            list_dc.append('DC2')
        if 'DC3' in template_file_path or dc3_yes == 'yes':
            list_dc.append('DC3')
        for ch in list_dc:
            try:
                shared_ins.dc_config_itech(ch, curve_mode="DEF_C")
            except:
                pass
        try:
            if 'custom.xlsx' in template_file_path:
                for i in range(len(df_template[df_template.columns[0]])):
                    if df_template['on/off DC1'][i] != 'no':
                        vmp1 = df_template['tensione DC1'][i]
                        pmp1 = df_template['potenza DC1'][i]
                        pf1 = df_template['pf1'][i]
                        turn_on1 = df_template['on/off DC1'][i]
                        ins.dc_set_iv("DC1", voc=vmp1/pf1, isc=pmp1/(vmp1*pf1), ff=pf1)
                        if turn_on1 == 1:
                            ins.dc_on("DC1")
                        elif turn_on1 == 0:
                            ins.dc_off("DC1")
                    if df_template['on/off DC2'][i] != 'no':
                        vmp2 = df_template['tensione DC2'][i]
                        pmp2 = df_template['potenza DC2'][i]
                        pf2 = df_template['pf2'][i]
                        turn_on2 = df_template['on/off DC2'][i]
                        ins.dc_set_iv("DC2", voc=vmp2 / pf2, isc=pmp2 / (vmp2 * pf2), ff=pf2)
                        if turn_on2 == 1:
                            ins.dc_on("DC2")
                        elif turn_on2 == 0:
                            ins.dc_off("DC2")
                    if df_template['on/off DC3'][i] != 'no':
                        vmp3 = df_template['tensione DC3'][i]
                        pmp3 = df_template['potenza DC3'][i]
                        pf3 = df_template['pf3'][i]
                        turn_on3 = df_template['on/off DC3'][i]
                        ins.dc_set_iv("DC3", voc=vmp3 / pf3, isc=pmp3 / (vmp3 * pf3), ff=pf3)
                        if turn_on3 == 1:
                            ins.dc_on("DC3")
                        elif turn_on3 == 0:
                            ins.dc_off("DC3")
                    if df_template['on/off AC'][i] != 'no':
                        vac = df_template['tensione AC'][i]
                        fac = df_template['frequenza AC'][i]
                        turn_onAC = df_template['on/off AC'][i]
                        fase = df_template['fase'][i]
                        ins.ac_set(vac, fac, phases=fase)
                        if turn_onAC == 1:
                            ins.ac_on()
                        elif turn_onAC == 0:
                            ins.ac_off()
                    if df_template['potenza batteria'][i] != 'no':
                        pbatt = int(df_template['potenza batteria'][i])
                        pbatt = max(0, min(65535, pbatt))

                        pbatt_ctrl = df_template['scarica/carica'][i]
                        ins.inv_broadcast_write("0x1110", 3, scale=1, role=None)
                        if pbatt_ctrl == 'scarica':
                            regs = [65535, 65535-pbatt, 65535, 65535-pbatt]
                        else:
                            regs = [0, pbatt, 0, pbatt]
                        ins.inv_broadcast_write("0x1189", regs, 1, role=None)
                    if df_template['registri master'][i] != 'no':
                        apply_template_writes(
                            ins, "master",
                            df_template['registri master'][i],
                            df_template['value master'][i],
                            scale=1
                        )
                    if df_template['registri slave'][i] != 'no':
                        apply_template_writes(
                            ins, "slave",
                            df_template['registri slave'][i],
                            df_template['value slave'][i],
                            scale=1
                        )
                    time.sleep(df_template['tempo'][i])
            elif 'curva MPPT' in template_file_path:
                row = df_inverter.iloc[0]
                shared_ins.inv_broadcast_write("0x1110", [3], scale=1, role=None)
                shared_ins.inv_broadcast_write("0x1189", [0, 0, 0, 0], scale=1, role=None)
                vmin = max(10, int(row['MIN MPPT']) - 10)
                vmax = int(row['MAX V'])
                #pout = int(row['MAX POUT'])
                imax = int(row['MAX I'])
                pf1 = float(df_template['pf1'][0]) if 'pf1' in df_template.columns else 1.0
                pf2 = float(df_template['pf2'][0]) if 'pf2' in df_template.columns else pf1
                pf3 = float(df_template['pf3'][0]) if 'pf3' in df_template.columns else pf1
                vac = int(df_template['tensione AC'][0])
                fac = float(df_template['frequenza AC'][0])
                fase = str(df_template['fase'][0])
                shared_ins.ac_set(vac, fac, fase)
                shared_ins.ac_on()
                for i in range(vmin, vmax+1, 5):
                    if 'DC1' in list_dc:
                        shared_ins.dc_set_iv("DC1", voc=min(i/pf1, vmax),
                                      isc=imax/pf1, ff=pf1)
                        shared_ins.dc_on('DC1')
                    if 'DC2' in list_dc:
                        shared_ins.dc_set_iv("DC2", voc=min(i/pf2, vmax),
                                      isc=imax/pf2, ff=pf2)
                        shared_ins.dc_on('DC2')
                    if 'DC3' in list_dc:
                        shared_ins.dc_set_iv("DC3", voc=min(i/pf3, vmax),
                                      isc=imax/pf3, ff=pf3)
                        shared_ins.dc_on('DC3')

                    if i == vmin:
                        time.sleep(60)
                    else:
                        time.sleep(int(df_template['tempo'][0]))
            elif 'ciclo batteria' in template_file_path:
                row = df_inverter.iloc[0]
                vnom = int(row['VNOM'])
                #pout = int(row['MAX POUT'])
                #imax = int(row['MAX I'])
                pbat_db = int(row['P BAT'])
                pf1 = float(df_template['pf1'][0]) if 'pf1' in df_template.columns else 1.0
                pf2 = float(df_template['pf2'][0]) if 'pf2' in df_template.columns else pf1
                pf3 = float(df_template['pf3'][0]) if 'pf3' in df_template.columns else pf1
                vac = int(df_template['tensione AC'][0])
                fac = float(df_template['frequenza AC'][0])
                fase = str(df_template['fase'][0])
                shared_ins.ac_set(vac, fac, fase)
                shared_ins.ac_on()
                if dc1_yes != '':
                    shared_ins.dc_set_iv("DC1", voc=vnom,
                                        isc=pbat_db/vnom, ff=pf1)
                    shared_ins.dc_off('DC1')
                if dc2_yes != '':
                    shared_ins.dc_set_iv("DC2", voc=vnom,
                                         isc=pbat_db/vnom, ff=pf2)
                    shared_ins.dc_off('DC2')
                if dc3_yes != '':
                    shared_ins.dc_set_iv("DC3", voc=vnom,
                                         isc=pbat_db/vnom, ff=pf3)
                    shared_ins.dc_off('DC3')
                time.sleep(5)
                shared_ins.inv_broadcast_write("0x1110", [3], scale=1, role=None)
                for i in range(0, len(df_template['pf1'])):
                    if df_template['potenza batteria'][i] != 'P BAT':
                        pbatt = int(df_template['potenza batteria'][i])
                    else:
                        pbatt = pbat_db
                    pbatt_ctrl = df_template['scarica/carica'][i]
                    if pbatt_ctrl == 'scarica':
                        regs = [65535, 65535-pbatt, 65535, 65535-pbatt]
                    else:
                        regs = [0, pbatt, 0, pbatt]
                    shared_ins.inv_broadcast_write("0x1189", regs, 1, role=None)
                    time.sleep(int(df_template['tempo'][i]))
                shared_ins.inv_broadcast_write("0x1189", [0, 0, 0, 0], 1, role=None)
            elif 'MAX SOUT' in template_file_path:
                row = df_inverter.iloc[0]
                shared_ins.inv_broadcast_write("0x1110", [3], scale=1, role=None)
                shared_ins.inv_broadcast_write("0x1189", [0, 0, 0, 0], scale=1, role=None)
                vnom = int(row['VNOM'])
                vmax = int(row['MAX V'])
                #pout = int(row['MAX POUT'])
                imax = int(row['MAX I'])
                pf1 = float(df_template['pf1'][0]) if 'pf1' in df_template.columns else 1.0
                pf2 = float(df_template['pf2'][0]) if 'pf2' in df_template.columns else pf1
                pf3 = float(df_template['pf3'][0]) if 'pf3' in df_template.columns else pf1
                vac = int(df_template['tensione AC'][0])
                fac = float(df_template['frequenza AC'][0])
                fase = str(df_template['fase'][0])
                shared_ins.ac_set(vac, fac, fase)
                shared_ins.ac_on()
                #registri master, value master -> registro e valori da inserire
                if 'DC1' in list_dc:
                    shared_ins.dc_set_iv("DC1", voc=min(vnom/pf1, vmax),
                                  isc=imax/pf1, ff=pf1)
                    shared_ins.dc_on('DC1')
                if 'DC2' in list_dc:
                    shared_ins.dc_set_iv("DC2", voc=min(vnom/pf2, vmax),
                                  isc=imax/pf2, ff=pf2)
                    shared_ins.dc_on('DC2')
                if 'DC3' in list_dc:
                    shared_ins.dc_set_iv("DC3", voc=min(vnom/pf3, vmax),
                                  isc=imax/pf3, ff=pf3)
                    shared_ins.dc_on('DC3')
                for i in range(-900, 900, 50):
                    registro_write = str(df_template['registri master'][0])
                    registri_valori = list()  # df_template['value master'][0]
                    for registro_merretrice in df_template['value master'][0].split('[')[1].split(']')[0].split(','):
                        registri_valori.append(int(registro_merretrice))
                        if i < 0:
                            registri_valori.append(65535+i)
                        else:
                            registri_valori.append(i)
                    shared_ins.inv_broadcast_write(registro_write, registri_valori, scale=1, role=None)
                    time.sleep(int(df_template['tempo'][0]))
                shared_ins.inv_broadcast_write(registro_write, [3, 1000, 1000, 0], scale=1, role=None)
                shared_ins.inv_broadcast_write(registro_write, [0], scale=1, role=None)
            elif '0-INJ' in template_file_path:
                row = df_inverter.iloc[0]
                shared_ins.inv_broadcast_write("0x1110", [3], scale=1, role=None)
                shared_ins.inv_broadcast_write("0x1189", [0, 0, 0, 0], scale=1, role=None)
                vnom = int(row['VNOM'])
                vmax = int(row['MAX V'])
                #pout = int(row['MAX POUT'])
                imax = int(row['MAX I'])
                pf1 = float(df_template['pf1'][0]) if 'pf1' in df_template.columns else 1.0
                pf2 = float(df_template['pf2'][0]) if 'pf2' in df_template.columns else pf1
                pf3 = float(df_template['pf3'][0]) if 'pf3' in df_template.columns else pf1
                vac = int(df_template['tensione AC'][0])
                fac = float(df_template['frequenza AC'][0])
                fase = str(df_template['fase'][0])
                shared_ins.ac_set(vac, fac, fase)
                shared_ins.ac_on()
                #registri master, value master -> registro e valori da inserire
                if 'DC1' in list_dc:
                    shared_ins.dc_set_iv("DC1", voc=min(vnom/pf1, vmax),
                                  isc=imax/pf1, ff=pf1)
                    shared_ins.dc_on('DC1')
                if 'DC2' in list_dc:
                    shared_ins.dc_set_iv("DC2", voc=min(vnom/pf2, vmax),
                                  isc=imax/pf2, ff=pf2)
                    shared_ins.dc_on('DC2')
                if 'DC3' in list_dc:
                    shared_ins.dc_set_iv("DC3", voc=min(vnom/pf3, vmax),
                                  isc=imax/pf3, ff=pf3)
                    shared_ins.dc_on('DC3')
                for i in range(0, 60, 5):
                    registro_write = str(df_template['registri master'][0])
                    registri_valori = list()  # df_template['value master'][0]
                    for registro_merretrice in df_template['value master'][0].split('[')[1].split(']')[0].split(','):
                        registri_valori.append(int(registro_merretrice))
                        registri_valori.append(i)
                    shared_ins.inv_broadcast_write(registro_write, registri_valori, scale=1, role=None)
                    time.sleep(int(df_template['tempo'][0]))
                shared_ins.inv_broadcast_write(registro_write, [0, 6000], scale=1, role=None)


        finally:
            # Sempre: metti DC in stato sicuro e spegni, poi spegni AC
            try:
                for _i in list_dc:
                    (shared_ins or ins).dc_set_iv(_i, 200, 1, 0.9)
                    #(shared_ins or ins).dc_off(_i)
            except Exception as e:
                print(f"[WARN] safe quench DC: {e}")
            # try:
            #     (shared_ins or ins).ac_off()
            # except Exception as e:
            #     print(f"[WARN] ac_off: {e}")
            # Se usiamo shared_ins, NON chiudiamo tutto (lo usa anche il logger),
            # ma rilasciamo sempre i client Modbus per liberare le COM.
            # try:
            #     (shared_ins or ins).inv_disconnect_all()
            # except Exception as e:
            #     print(f"[WARN] inv_disconnect_all: {e}")
            # # Se non è condiviso, chiudiamo completamente le risorse residue.
            if shared_ins is None:
                try:
                    ins.close_all()
                except Exception as e:
                    print(f"[WARN] close_all: {e}")

        time.sleep(2)
        print("[TEST TEMPLATE] Test completato.")


    thread = threading.Thread(target=test_logic, daemon=True)
    thread.start()
    return thread


# ===========================
#  MULTI-TEST: PLAYLIST .TXT
# ===========================
def run_tests_playlist(
    playlist_path: str,
    sn: str,
    protocol: str,
    inverter_data,
    shared_ins=None,
    template_folder: str = "./template",
):
    if not os.path.isfile(playlist_path):
        raise FileNotFoundError(f"Playlist non trovata: {playlist_path}")
    with open(playlist_path, "r", encoding="utf-8") as f:
        names = [ln.strip() for ln in f if ln.strip() and not ln.strip().startswith("#")]
    for name in names:
        tpl = os.path.join(template_folder, f"{name}.xlsx")
        if not os.path.isfile(tpl):
            print(f"[WARN] Template assente: {tpl} — salto.")
            continue
        t = run_test_from_template(tpl, sn, protocol, inverter_data, shared_ins=shared_ins)
        try:
            if hasattr(t, "join"): t.join()
        except Exception as e:
            print(f"[WARN] join test '{name}': {e}")
        time.sleep(30)