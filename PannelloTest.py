import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd
import pyvisa
from pymodbus.client import ModbusTcpClient, ModbusSerialClient
from pymodbus.exceptions import ModbusIOException
import ast  # per convertire stringa in lista, se serve
import os
import openpyxl
import threading
import time
import csv
from datetime import datetime

logging_thread = None
test_thread = None
logging_running = False
logging_paused = False

# --- lettura SN Inverter ---
def read_SN(proto, ip_tcp=None, porta_com=None, ip_hub=None,
            slave_id_rtu=None, slave_id_azzurro=None, max_retries=10):
    registri_read = [
        0x0445, 0x0446, 0x0447, 0x0448, 0x0449,
        0x044a, 0x044b, 0x044c, 0x0470, 0x0471
    ]
    expected_lengths = {14, 20}

    def connect_client():
        if proto == "RTU":
            return ModbusSerialClient(
                port=porta_com,
                baudrate=9600,
                timeout=1,
                stopbits=1,
                bytesize=8,
                parity='N'
            ), int(slave_id_rtu)
        elif proto == "TCP":
            return ModbusTcpClient(ip_tcp, port=8899, timeout=1), 1
        elif proto == "AzzurroHUB":
            return ModbusTcpClient(ip_hub, port=55400, timeout=5), int(slave_id_azzurro)
        else:
            raise ValueError(f"Protocollo non supportato: {proto}")

    for attempt_sn in range(max_retries):  # retry per intero SN
        sn_str = ""
        client, slave_id = connect_client()

        if not client.connect():
            print(f"[ERRORE] Connessione fallita ({proto})")
            return "UNKNOWN"

        try:
            for reg in registri_read:
                val_ok = False
                for _ in range(max_retries):  # retry per singolo registro
                    rr = client.read_holding_registers(address=reg, count=1, slave=slave_id)
                    if not rr.isError():
                        val = rr.registers[0]
                        hex_str = f"{val:04x}"
                        if hex_str == '0000':
                            val_ok = True
                            break
                        sn_str += chr(int(hex_str[0:2], 16))
                        sn_str += chr(int(hex_str[2:4], 16))
                        val_ok = True
                        break
                    # piccolo delay prima del retry
                    if proto == "AzzurroHUB":
                        time.sleep(3)  # attesa obbligatoria per non farsi bloccare dall'HUB
                    else:
                        time.sleep(0.5)  # attesa più breve per TCP/RTU

                if not val_ok:
                    print(f"[WARN] Registro {hex(reg)} non letto dopo {max_retries} tentativi")

        except Exception as e:
            print(f"[ERRORE lettura SN] {e}")
            sn_str = ""
        finally:
            client.close()

        sn_str = sn_str.strip()
        if len(sn_str) in expected_lengths:
            return sn_str  # SN valido trovato

        print(f"[WARN] SN incompleto (len={len(sn_str)}), retry {attempt_sn+1}/{max_retries}")

    return "UNKNOWN"


# --- Lettura DB inverter ---
def load_inverter_db(model, sn):
    db_path = os.path.join("./database", f"{model}.xlsx")
    if not os.path.exists(db_path):
        return None

    wb = openpyxl.load_workbook(db_path)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]

    for row in ws.iter_rows(min_row=2, values_only=True):
        if str(row[0]).strip() == str(sn).strip():
            return headers, row
    return None


# Funzione che crea una nuova finestra con messaggio HelloWorld
def open_subpanel(title):
    sub_win = tk.Toplevel(root)
    sub_win.title(title)
    label = tk.Label(sub_win, text="HelloWorld", font=("Arial", 16))
    label.pack(padx=20, pady=20)


# Funzione Test
def run_test_from_template(template_file_path):
    def test_logic():
        print(f"[TEST TEMPLATE] Avvio test da template: {template_file_path}")
        # Placeholder: fai ciò che serve dopo
        if template_file_path != 'custom':
            df = pd.read_excel(template_file_path)
            for i in range(0, len(df[df.columns[0]])):
                vmp1 = df['tensione DC1'][i]
                pmp1 = df['potenza DC1'][i]
                pf1 = df['pf1'][i]
                turn_on1 = df['on/off DC1'][i]
                vmp2 = df['tensione DC2'][i]
                pmp2 = df['potenza DC2'][i]
                pf2 = df['pf2'][i]
                turn_on2 = df['on/off DC2'][i]
                vmp3 = df['tensione DC3'][i]
                pmp3 = df['potenza DC3'][i]
                pf3 = df['pf3'][i]
                turn_on3 = df['on/off DC3'][i]
                vac = df['tensione AC'][i]
                fac = df['frequenza AC'][i]
                pbatt = df['potenza batteria'][i]
                registri_inv_m = df['registri master'][i]
                registri_value_m = df['value master'][i]
                registri_inv_s = df['registri slave'][i]
                registri_value_s = df['value slave'][i]
                time.sleep(df['tempo'][i])
        time.sleep(2)
        print("[TEST TEMPLATE] Test completato.")
    thread = threading.Thread(target=test_logic, daemon=True)
    thread.start()
    return thread


# apro il thread per il log
def start_logging_routine(
    protocol, inverters, registers, file_path, sampling_time, total_time
):
    global logging_running, logging_paused
    logging_running = True
    logging_paused = False

    def log_loop():
        start_time = time.time()

        # Prepara header CSV
        header = ["timestamp"]
        for inv_index, inv in enumerate(inverters):
            for label, _, _ in registers:
                header.append(f"Inverter{inv_index+1}_{label}")

        try:
            with open(file_path, mode='w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(header)

                while time.time() - start_time < total_time:
                    if not logging_running:
                        break
                    if logging_paused:
                        time.sleep(0.5)
                        continue

                    row = [datetime.now().strftime("%Y-%m-%d %H:%M:%S")]

                    for modbus_id, address in inverters:
                        try:
                            if protocol == "RTU":
                                client = ModbusSerialClient(
                                    port=address,
                                    baudrate=9600,
                                    timeout=1,
                                    stopbits=1,
                                    bytesize=8,
                                    parity='N'
                                )
                            else:
                                port = 8899 if protocol == "TCP" else 55400
                                client = ModbusTcpClient(address, port=port, timeout=1)

                            client.connect()

                            reg_values = []
                            for _, reg, scale in registers:
                                reg = int(reg)
                                scale = int(scale)
                                result = client.read_holding_registers(address=reg, count=1, slave=int(modbus_id))
                                if not result or not hasattr(result, "registers"):
                                    reg_values.append("ERR")
                                else:
                                    reg_values.append(result.registers[0] * scale)

                            client.close()
                            row.extend(reg_values)

                        except Exception as e:
                            row.extend(["ERR"] * len(registers))

                    writer.writerow(row)
                    time.sleep(sampling_time)

            messagebox.showinfo("Logging completato", f"File salvato:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Errore logging", str(e))

    threading.Thread(target=log_loop, daemon=True).start()


# Gestisco il Log del test automatico
def open_log_panel():
    log_win = tk.Toplevel()
    log_win.title("Lancio Log Strumenti")
    log_win.minsize(950, 600)

    protocol_var = tk.StringVar(value="TCP")
    ttk.Label(log_win, text="Protocollo di Comunicazione").pack(anchor="w", padx=10, pady=(10, 0))
    ttk.Combobox(log_win, textvariable=protocol_var, values=["TCP", "RTU", "AzzurroHUB"]).pack(anchor="w", padx=10)

    # Fino a 10 inverter
    inverter_entries = []

    inverter_frame = tk.LabelFrame(log_win, text="Inverter", font=("Arial", 10, "bold"))
    inverter_frame.pack(fill="x", pady=5)

    # Intestazioni colonna (2 blocchi da 5 inverter)
    for block in range(2):  # 2 blocchi di 5 inverter
        col_offset = block * 5
        headers = ["#", "Modbus ID", "IP / COM", "M/S", "Alimentatore"]
        for idx, header in enumerate(headers):
            tk.Label(
                inverter_frame,
                text=header,
                font=("Arial", 9, "bold")
            ).grid(row=0, column=col_offset + idx, padx=5, pady=2)

    # Righe inverter
    for i in range(10):
        block = i // 5
        row = (i % 5) + 1
        col_offset = block * 5

        # Numero inverter
        tk.Label(inverter_frame, text=f"{i + 1}").grid(row=row, column=col_offset, padx=5, pady=2)

        # Modbus ID
        modbus_entry = tk.Entry(inverter_frame, width=5)
        modbus_entry.grid(row=row, column=col_offset + 1, padx=5, pady=2)

        # IP / COM
        ip_entry = tk.Entry(inverter_frame, width=15)
        ip_entry.grid(row=row, column=col_offset + 2, padx=5, pady=2)

        # M/S checkbox
        slave_var = tk.BooleanVar()
        slave_cb = tk.Checkbutton(inverter_frame, variable=slave_var)
        slave_cb.grid(row=row, column=col_offset + 3, padx=5, pady=2)

        # Alimentatore dropdown
        alim_var = tk.StringVar(value="Nessuno")
        alim_menu = ttk.Combobox(
            inverter_frame,
            textvariable=alim_var,
            values=["Nessuno", "DC1", "DC2", "DC3"],
            width=8,
            state="readonly"
        )
        alim_menu.grid(row=row, column=col_offset + 4, padx=5, pady=2)

        inverter_entries.append((modbus_entry, ip_entry, slave_var, alim_var))

    # === Registri di default ===
    default_registers = [
        ("System State", "0x0404", "1"),
        ("Active Output Power [kW]", "0x0485", "100"),
        ("Reactive Output Power [kVAr]", "0x0486", "100"),
        ("Active PCC Power [kW]", "0x0488", "100"),
        ("Reactive PCC Power [kVAr]", "0x0489", "100"),
        ("Voltage DC1 [V]", "0x0584", "10"),
        ("Current DC1 [A]", "0x0585", "100"),
        ("Power DC1 [kW]", "0x0586", "100"),
        ("Voltage DC2 [V]", "0x0587", "10"),
        ("Current DC2 [A]", "0x0588", "100"),
        ("Power DC2 [kW]", "0x0589", "100"),
        ("Charge/Discharge Power [kW]", "0x0667", "10"),
        ("Battery SOC [%]", "0x0668", "1")
    ]

    # Sezione registri (20 registri in 4 colonne da 5 righe)
    registers_frame = tk.LabelFrame(log_win, text="Registri", font=("Arial", 10, "bold"))
    registers_frame.pack(fill="x", pady=5)
    # tk.Label(log_win, text="Registri da leggere", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 0))
    # registers_frame = tk.Frame(log_win)
    # registers_frame.pack(anchor="w", padx=10)

    register_entries = []

    # Intestazioni per ciascuna colonna
    for col in range(4):  # 4 colonne
        tk.Label(registers_frame, text="Label", font=("Arial", 9, "bold")).grid(row=0, column=col * 3, padx=2, pady=2)
        tk.Label(registers_frame, text="Registro", font=("Arial", 9, "bold")).grid(row=0, column=col * 3 + 1, padx=2,
                                                                                   pady=2)
        tk.Label(registers_frame, text="Scaling", font=("Arial", 9, "bold")).grid(row=0, column=col * 3 + 2, padx=2,
                                                                                  pady=2)

    # Registri (20 = 4 colonne × 5 righe)
    for i in range(20):
        col = i // 5
        row = (i % 5) + 1  # +1 per lasciare spazio alle intestazioni

        name_entry = tk.Entry(registers_frame, width=20)
        reg_entry = tk.Entry(registers_frame, width=10)
        scale_entry = tk.Entry(registers_frame, width=6)

        if i < len(default_registers):
            name, reg, scale = default_registers[i]
            name_entry.insert(0, name)
            reg_entry.insert(0, reg)
            scale_entry.insert(0, scale)
        else:
            scale_entry.insert(0, "1")

        name_entry.grid(row=row, column=col * 3, padx=2, pady=2)
        reg_entry.grid(row=row, column=col * 3 + 1, padx=2, pady=2)
        scale_entry.grid(row=row, column=col * 3 + 2, padx=2, pady=2)

        register_entries.append((name_entry, reg_entry, scale_entry))

    # File CSV
    file_frame = tk.Frame(log_win)
    file_frame.pack(anchor="w", padx=10, pady=(10, 0))
    tk.Label(file_frame, text="Percorso file CSV:", font=("Arial", 10, "bold")).pack(side="left")
    file_entry = tk.Entry(file_frame, width=60)
    file_entry.pack(side="left", padx=5)
    def browse_file():
        filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if filepath:
            file_entry.delete(0, tk.END)
            file_entry.insert(0, filepath)
    tk.Button(file_frame, text="Sfoglia", command=browse_file).pack(side="left")

    # Tempi
    time_frame = tk.Frame(log_win)
    time_frame.pack(anchor="w", padx=10, pady=10)
    tk.Label(time_frame, text="Campionamento [s]:", font=("Arial", 10, "bold")).pack(side="left")
    sampling_entry = tk.Entry(time_frame, width=5)
    sampling_entry.insert(0, "5")
    sampling_entry.pack(side="left", padx=5)
    tk.Label(time_frame, text="Durata test [s]:", font=("Arial", 10, "bold")).pack(side="left", padx=(20, 0))
    duration_entry = tk.Entry(time_frame, width=5)
    duration_entry.insert(0, "60")
    duration_entry.pack(side="left", padx=5)

    # Pulsanti
    def on_send_log():
        global logging_thread, logging_running, test_thread
        file_path = file_entry.get().strip()
        if not file_path:
            messagebox.showwarning("Attenzione", "Inserisci un percorso valido per il file CSV.")
            return

        try:
            sampling = float(sampling_entry.get())
            duration = float(duration_entry.get())
        except:
            messagebox.showerror("Errore", "Controlla i valori di tempo")
            return

        inverter_data = []
        for idx, (modbus_entry, ip_entry, slave_var, alim_var) in enumerate(inverter_entries):
            ip_val = ip_entry.get().strip()
            modbus_val = modbus_entry.get().strip()
            if not ip_val or not modbus_val:
                continue
            mode = protocol_var.get().strip()
            if mode == "TCP":
                sn = read_SN(mode, ip_tcp=ip_val)
            elif mode == "RTU":
                sn = read_SN(mode, porta_com=ip_val, slave_id_rtu=int(modbus_val))
            elif mode == "AzzurroHUB":
                sn = read_SN(mode, ip_hub=ip_val, slave_id_azzurro=int(modbus_val))
            else:
                sn = "UNKNOWN"
            inverter_data.append({
                "sn": sn,
                "ip": ip_val,
                "modbus": int(modbus_val),
                "slave": slave_var.get(),
                "alimentatore": alim_var.get()
            })

        print("[INFO] Inverter rilevati:", inverter_data)

        registers = []
        for name_entry, reg_entry, scale_entry in register_entries:
            name = name_entry.get().strip()
            reg = reg_entry.get().strip()
            scale = scale_entry.get().strip()
            if name and reg and scale:
                registers.append((name, reg, scale))

        if not registers:
            messagebox.showerror("Errore", "Nessun registro valido selezionato.")
            return

        # Thread 1: logging
        logging_thread = threading.Thread(target=start_logging_routine, args=(
            protocol_var.get(), inverter_data, registers, file_path, sampling, duration
        ))
        logging_thread.start()

        # Thread 2: test da template (se ≠ custom)
        selected_template = template_var.get()
        if selected_template != "custom":
            template_path = os.path.join(template_folder, selected_template)
            test_thread = run_test_from_template(template_path)

    def pause_logging():
        global logging_paused
        logging_paused = True

    def resume_logging():
        global logging_paused
        logging_paused = False

    button_frame = tk.Frame(log_win)
    button_frame.pack(fill="x", pady=5, padx=5)

    # --- Template Test ---
    tk.Label(button_frame, text="Seleziona template test:", font=("Arial", 10, "bold")).pack(side="left", padx=(0,5))

    # Caricamento file disponibili in ./template/
    template_folder = "./template/"
    template_files = []
    if os.path.exists(template_folder):
        template_files = [f for f in os.listdir(template_folder) if f.lower().endswith((".xlsx", ".xls"))]
    template_options = ["custom"] + template_files

    template_var = tk.StringVar(value="custom")
    template_menu = ttk.OptionMenu(button_frame, template_var, template_options[0], *template_options)

    template_menu.pack(side="left", padx=(0, 20))

    #frame dei bottoni
    tk.Button(button_frame, text="Send", bg="lightgreen", command=on_send_log).pack(side="right", padx=2)
    tk.Button(button_frame, text="Pause", bg="orange", command=pause_logging).pack(side="right", padx=2)
    tk.Button(button_frame, text="Resume", bg="lightblue", command=resume_logging).pack(side="right", padx=2)
    tk.Button(button_frame, text="Exit", bg="red", fg="white", command=log_win.destroy).pack(side="right", padx=2)


    # Frame dinamico per configurazione test
    test_config_frame = tk.Frame(log_win)
    test_config_frame.pack(fill="x", padx=10, pady=10)

    def hide_test_config():
        for widget in test_config_frame.winfo_children():
            widget.destroy()

    def show_test_config():
        hide_test_config()

        # Titolo Configurazione Test
        tk.Label(test_config_frame, text="Configurazione Test", font=("Arial", 10, "bold")).pack(anchor="w", padx=10,
                                                                                                 pady=(10, 5))

        # Frame per i checkbox in linea
        config_frame = tk.Frame(test_config_frame)
        config_frame.pack(fill="x", padx=10, pady=5)

        dc_power_vars = []
        for idx in range(3):
            var = tk.BooleanVar()
            cb = tk.Checkbutton(config_frame, text=f"Alimentatore DC{idx + 1}", variable=var)
            cb.pack(side="left", padx=5)
            dc_power_vars.append(var)

        ac_var = tk.BooleanVar()
        tk.Checkbutton(config_frame, text="Usa simulatore AC", variable=ac_var).pack(side="left", padx=5)

        inv_var = tk.BooleanVar()
        tk.Checkbutton(config_frame, text="Controllo inverter sotto test", variable=inv_var).pack(side="left", padx=5)

    def on_template_change(*args):
        selected = template_var.get()
        if selected == "custom":
            hide_test_config()
        else:
            show_test_config()

    template_var.trace_add("write", on_template_change)


# Gestisco l'inverter
def open_inverter_control_panel():
    inv_win = tk.Toplevel()
    inv_win.title("Controllo Inverter")
    inv_win.minsize("500x600")

    # Variabili
    protocollo = tk.StringVar(value="TCP")
    operazione = tk.StringVar(value="Lettura")
    registro = tk.StringVar()
    num_reg = tk.StringVar()
    scaling = tk.StringVar()
    value = tk.StringVar()
    porta_com = tk.StringVar()
    slave_id_rtu = tk.StringVar()
    ip_tcp = tk.StringVar()
    ip_hub = tk.StringVar()
    slave_id_azzurro = tk.StringVar()

    # Funzione invio comandi (mock)
    def send_command():
        try:
            proto = protocollo.get()
            op = operazione.get()
            reg = int(registro.get(), 16)
            num = int(num_reg.get())
            scale = int(scaling.get())
            result = None

            # === CLIENT CONFIG ===
            if proto == "RTU":
                client = ModbusSerialClient(
                    port=porta_com.get(),
                    baudrate=9600,
                    timeout=1,
                    stopbits=1,
                    bytesize=8,
                    parity='N'
                )
                slave_id = int(slave_id_rtu.get())
            elif proto == "TCP":
                client = ModbusTcpClient(ip_tcp.get(), port=8899, timeout=1)
                slave_id = 1  # Default slave
            elif proto == "AzzurroHUB":
                client = ModbusTcpClient(ip_hub.get(), port=55400, timeout=1)
                slave_id = int(slave_id_azzurro.get())
            else:
                raise ValueError("Protocollo non valido")

            # === CONNESSIONE ===
            if not client.connect():
                raise ConnectionError("Impossibile connettersi al dispositivo")

            if op == "Lettura":
                result = client.read_holding_registers(address=reg, count=num, slave=slave_id)
                if isinstance(result, ModbusIOException) or not result:
                    raise IOError("Errore nella lettura")
                values = [val * scale for val in result.registers]
                messagebox.showinfo("Risultato Lettura", f"Valori: {values}")

            elif op == "Scrittura":
                val_input = value.get()
                try:
                    val_list = ast.literal_eval(val_input)
                    if not isinstance(val_list, list):
                        val_list = [int(val_list)]
                except:
                    raise ValueError("Formato value non valido (usa [1,2,3] o singolo numero)")

                scaled_vals = [int(v / scale) for v in val_list]
                if len(scaled_vals) == 1:
                    client.write_register(reg, scaled_vals[0], slave=slave_id)
                else:
                    client.write_registers(reg, scaled_vals, slave=slave_id)
                messagebox.showinfo("Scrittura", f"Scritti: {val_list} (scalati: {scaled_vals})")

            client.close()

        except Exception as e:
            messagebox.showerror("Errore", str(e))

    def exit_panel():
        inv_win.destroy()

    # Campi dinamici
    def update_fields(*args):
        frame_rtu.forget()
        frame_tcp.forget()
        frame_azzurro.forget()
        frame_value.forget()

        if protocollo.get() == "RTU":
            frame_rtu.pack(pady=5, fill="x")
        elif protocollo.get() == "TCP":
            frame_tcp.pack(pady=5, fill="x")
        elif protocollo.get() == "AzzurroHUB":
            frame_azzurro.pack(pady=5, fill="x")

        if operazione.get() == "Scrittura":
            frame_value.pack(pady=5, fill="x")

    # UI layout base
    ttk.Label(inv_win, text="Protocollo Comunicazione:").pack(pady=(10, 0))
    ttk.OptionMenu(inv_win, protocollo, "TCP", "TCP", "RTU", "AzzurroHUB").pack()

    ttk.Label(inv_win, text="Lettura o Scrittura?").pack(pady=(10, 0))
    ttk.OptionMenu(inv_win, operazione, "Lettura", "Lettura", "Scrittura").pack()

    ttk.Label(inv_win, text="Scrivi il registro in hex (es. 0x0010):").pack(pady=(10, 0))
    tk.Entry(inv_win, textvariable=registro).pack()

    ttk.Label(inv_win, text="# registri da leggere:").pack(pady=(10, 0))
    tk.Entry(inv_win, textvariable=num_reg).pack()

    ttk.Label(inv_win, text="Scaling:").pack(pady=(10, 0))
    tk.Entry(inv_win, textvariable=scaling).pack()

    # Frame value (Scrittura)
    frame_value = tk.Frame(inv_win)
    ttk.Label(frame_value, text="Value (array o stringa):").pack(pady=(0, 5))
    tk.Entry(frame_value, textvariable=value).pack()

    # Frame RTU
    frame_rtu = tk.Frame(inv_win)
    ttk.Label(frame_rtu, text="Porta COM (es. COM3):").pack(pady=(0, 5))
    tk.Entry(frame_rtu, textvariable=porta_com).pack()
    ttk.Label(frame_rtu, text="Indirizzo Modbus Inverter:").pack(pady=(10, 0))
    tk.Entry(frame_rtu, textvariable=slave_id_rtu).pack()

    # Frame TCP
    frame_tcp = tk.Frame(inv_win)
    ttk.Label(frame_tcp, text="IP dell'Inverter:").pack(pady=(0, 5))
    tk.Entry(frame_tcp, textvariable=ip_tcp).pack()

    # Frame Azzurro HUB
    frame_azzurro = tk.Frame(inv_win)
    ttk.Label(frame_azzurro, text="IP dell'HUB:").pack(pady=(0, 5))
    tk.Entry(frame_azzurro, textvariable=ip_hub).pack()
    ttk.Label(frame_azzurro, text="Indirizzo Modbus Inverter:").pack(pady=(10, 0))
    tk.Entry(frame_azzurro, textvariable=slave_id_azzurro).pack()

    # Pulsanti
    button_frame = tk.Frame(inv_win)
    button_frame.pack(pady=20)
    tk.Button(button_frame, text="Send", bg="lightgreen", command=send_command).grid(row=0, column=0, padx=10)
    tk.Button(button_frame, text="Exit", bg="red", fg="white", command=exit_panel).grid(row=0, column=1, padx=10)

    # Tracciamento dinamico
    protocollo.trace_add("write", update_fields)
    operazione.trace_add("write", update_fields)

    update_fields()  # inizializzazione


# Gestisco l'AC Source
def open_ac_control_panel():
    ac_win = tk.Toplevel()
    ac_win.title("Controllo AC")
    ac_win.minsize("400x400")

    # Variabili
    selected_tipo = tk.StringVar(value="Monofase")
    vac = tk.StringVar()
    freq = tk.StringVar()

    inst_ac = None  # Oggetto pyvisa

    def send_command():
        nonlocal inst_ac
        try:
            vac_val = float(vac.get())
            freq_val = float(freq.get())
            tipo = selected_tipo.get()

            rm = pyvisa.ResourceManager()
            inst_ac = rm.open_resource("ASRL5::INSTR")  # <--- Adatta qui la porta!

            inst_ac.write(f'VOLT {vac_val}')
            inst_ac.write(f'FREQ {freq_val}')

            if tipo == "Monofase":
                inst_ac.write("SYST:FUNC ONE")
            else:
                inst_ac.write("SYST:FUNC THREE")

            messagebox.showinfo("Successo", "Comandi AC inviati correttamente.")

        except Exception as e:
            messagebox.showerror("Errore", str(e))

    def turn_on():
        try:
            if inst_ac:
                inst_ac.write('OUTP ON')
        except Exception as e:
            messagebox.showerror("Errore", str(e))

    def turn_off():
        try:
            if inst_ac:
                inst_ac.write('OUTP OFF')
        except Exception as e:
            messagebox.showerror("Errore", str(e))

    def exit_panel():
        ac_win.destroy()

    # UI
    ttk.Label(ac_win, text="Seleziona la tipologia di rete:").pack(pady=5)
    ttk.OptionMenu(ac_win, selected_tipo, "Monofase", "Monofase", "Trifase").pack()

    ttk.Label(ac_win, text="Vrms [V]:").pack(pady=(10, 0))
    tk.Entry(ac_win, textvariable=vac).pack()

    ttk.Label(ac_win, text="Freq. [Hz]:").pack(pady=(10, 0))
    tk.Entry(ac_win, textvariable=freq).pack()

    # Pulsanti ON/OFF
    button_frame1 = tk.Frame(ac_win)
    button_frame1.pack(pady=(20, 5))
    tk.Button(button_frame1, text="Turn On", bg="lightgreen", command=turn_on).grid(row=0, column=0, padx=10)
    tk.Button(button_frame1, text="Turn Off", bg="red", fg="white", command=turn_off).grid(row=0, column=1, padx=10)

    # Pulsanti SEND/EXIT
    button_frame2 = tk.Frame(ac_win)
    button_frame2.pack(pady=5)
    tk.Button(button_frame2, text="Send", bg="lightgreen", command=send_command).grid(row=0, column=0, padx=10)
    tk.Button(button_frame2, text="Exit", bg="red", fg="white", command=exit_panel).grid(row=0, column=1, padx=10)


# Gestisco le DC Sources
def open_dc_control_panel():
    dc_win = tk.Toplevel()
    dc_win.title("Controllo DC")
    dc_win.minsize("400x400")

    # Variabili
    porta_map = {
        "Strumento 1": "ASRL20::INSTR",
        "Strumento 2": "ASRL21::INSTR",
        "Strumento 3": "ASRL22::INSTR"
        # Aggiungi qui se aggiungi altri strumenti DC
    }
    selected_strumento = tk.StringVar(value="Strumento 1")
    voc = tk.StringVar()
    isc = tk.StringVar()
    ff = tk.StringVar(value="1.0")

    inst = None  # Variabile di riferimento per pyvisa

    def send_command():
        nonlocal inst
        try:
            voc_val = float(voc.get())
            isc_val = float(isc.get())
            ff_val = float(ff.get())

            if not (0 <= ff_val <= 1):
                raise ValueError("Fattore di forma deve essere tra 0 e 1")

            porta = porta_map[selected_strumento.get()]
            rm = pyvisa.ResourceManager()
            inst = rm.open_resource(porta)

            inst.write(f'SOL:USER:VOC {voc_val}')
            inst.write(f'SOL:USER:VMP {voc_val * ff_val}')
            inst.write(f'SOL:USER:ISC {isc_val}')
            inst.write(f'SOL:USER:IMP {isc_val * ff_val}')
            messagebox.showinfo("Successo", "Comandi inviati correttamente.")
        except Exception as e:
            messagebox.showerror("Errore", str(e))

    def turn_on():
        try:
            if inst:
                inst.write('OUTP 1')
        except Exception as e:
            messagebox.showerror("Errore", str(e))

    def turn_off():
        try:
            if inst:
                inst.write('OUTP 0')
        except Exception as e:
            messagebox.showerror("Errore", str(e))

    def exit_panel():
        dc_win.destroy()

    # UI Layout
    ttk.Label(dc_win, text="Seleziona lo strumento da controllare:").pack(pady=5)
    ttk.OptionMenu(dc_win, selected_strumento, *porta_map.keys()).pack()

    ttk.Label(dc_win, text="Tensione Voc [V]:").pack(pady=(10, 0))
    tk.Entry(dc_win, textvariable=voc).pack()

    ttk.Label(dc_win, text="Corrente Isc [A]:").pack(pady=(10, 0))
    tk.Entry(dc_win, textvariable=isc).pack()

    ttk.Label(dc_win, text="Fattore di forma (0 - 1):").pack(pady=(10, 0))
    tk.Entry(dc_win, textvariable=ff).pack()

    # Pulsanti Turn On / Turn Off
    button_frame1 = tk.Frame(dc_win)
    button_frame1.pack(pady=(20, 5))
    tk.Button(button_frame1, text="Turn On", bg="lightgreen", command=turn_on).grid(row=0, column=0, padx=10)
    tk.Button(button_frame1, text="Turn Off", bg="red", fg="white", command=turn_off).grid(row=0, column=1, padx=10)

    # Pulsanti Send / Exit (sullo stesso livello)
    button_frame2 = tk.Frame(dc_win)
    button_frame2.pack(pady=5)
    tk.Button(button_frame2, text="Send", bg="lightgreen", command=send_command).grid(row=0, column=0, padx=10)
    tk.Button(button_frame2, text="Exit", bg="red", fg="white", command=exit_panel).grid(row=0, column=1, padx=10)


# Callback per ogni pulsante
def on_dc_control():
    open_dc_control_panel()


def on_ac_control():
    open_ac_control_panel()


def on_inverter_control():
    open_inverter_control_panel()


def on_log_start():
    open_log_panel()


def on_test_start():
    open_subpanel("Lancio Test Automatici")


def on_realtime_monitor():
    open_subpanel("Monitoraggio Realtime")


def on_generate_report():
    open_subpanel("Generazione Report")


def on_kill_exit():
    root.destroy()  # Chiude tutto


# Creazione finestra principale
root = tk.Tk()
root.title("Pannello di Controllo Strumenti")
root.geometry("400x550")

# Lista bottoni con testo, funzione e colore
buttons = [
    ("Controllo DC", on_dc_control, "gold"),
    ("Controllo AC", on_ac_control, "gold"),
    ("Controllo Inverter", on_inverter_control, "gold"),
    ("Lancio Log Strumenti", on_log_start, "lightgreen"),
    ("Lancio Test Automatici", on_test_start, "lightgreen"),
    ("Monitoraggio Realtime", on_realtime_monitor, "lightblue"),
    ("Generazione Report", on_generate_report, "lightblue"),
    ("Exit", on_kill_exit, "red")
]

# Inserimento dei bottoni nel pannello
for text, command, color in buttons:
    fg_color = "white" if color == "red" else "black"
    btn = tk.Button(root, text=text, command=command,
                    font=("Arial", 12), bg=color, fg=fg_color, padx=10, pady=10)
    btn.pack(fill="x", padx=30, pady=5)

root.mainloop()
