"""
Under-Frequency Load Shedding (UFLS) Simulation Module.

This script executes a dynamic UFLS simulation using DIgSILENT PowerFactory.
It iteratively runs the RMS simulation, monitors the system frequency, and
sheds load based on predefined stages and frequency limits.
"""

import os
import re
import sys

import numpy as np
import pandas as pd

# --- Configuration Parameters ---

# --- Generator yang di-trip ---
TARGET_GEN_TRIP = "DG_1"
TIME_GEN_TRIP = 1.0     # Waktu generator trip

# --- Pola Beban (1 = kandidat UFLS, 0 = tidak disentuh) ---
LOAD_SHED_PATTERN = "00011111111111111001100001000000"

# --- Konfigurasi UFLS ---
DELAY_UFLS = 0.46       # Delay sensor & breaker (detik)
STAGES = [
    {"limit": 48.5, "pct": 0.12},  # Block 1: 48.5 Hz, 12% total MW sistem
    {"limit": 48.4, "pct": 0.12},  # Block 2: 48.4 Hz, 12% total MW sistem
    {"limit": 48.3, "pct": 0.12},  # Block 3: 48.3 Hz, 12% total MW sistem
    {"limit": 48.2, "pct": 0.24},  # Block 4: 48.2 Hz, 24% total MW sistem
]

# --- Parameter Settling Time ---
FREQ_TOLERANCE = 0.05
FREQ_SETTLE_WINDOW = 2.0

# --- Waktu Simulasi & Monitoring ---
RMS_TSTOP = 100.0        # Total waktu simulasi (detik)
FREQ_MONITOR_BUS_IDX = 0     
VOLT_MONITOR_BUS_IDX = 17    
EXCEL_OUTPUT = "UFLS_Case 1.xlsx"

# --- PowerFactory ---
VERSION_PYTHON = "3.12"
PATH_APP = r"C:\Program Files\DIgSILENT\PowerFactory 2024"
PATH_API = fr"{PATH_APP}\Python\{VERSION_PYTHON}"
PROJECT_NAME = "Import(1)"

# --- PowerFactory Initialization ---

if PATH_API not in sys.path:
    sys.path.append(PATH_API)
os.environ['PATH'] = PATH_APP + ";" + os.environ['PATH']

try:
    import powerfactory
except ImportError:
    print("Error: Module powerfactory tidak ditemukan.")
    exit()

app = powerfactory.GetApplication()
app.ActivateProject(PROJECT_NAME)
try:
    app.EchoOff()
except:
    pass

inc = app.GetFromStudyCase("ComInc")
sim = app.GetFromStudyCase("ComSim")
elmres = app.GetFromStudyCase("All calculations.ElmRes")

sim.tstop = RMS_TSTOP

# --- Helper Functions ---

def natural_keys(text):
    return [int(c) if c.isdigit() else c for c in re.split(r'(\d+)', text)]

# --- Network Object Retrieval and Stage Clustering ---

raw_loads = app.GetCalcRelevantObjects("*.ElmLod")
all_loads = [obj for obj in raw_loads if obj.GetClassName() == "ElmLod"]
all_loads.sort(key=lambda x: natural_keys(x.GetAttribute("loc_name")))

all_buses = app.GetCalcRelevantObjects("*.ElmTerm")
all_buses.sort(key=lambda x: natural_keys(x.GetAttribute("loc_name")))

all_gens = app.GetCalcRelevantObjects("*.ElmSym") + app.GetCalcRelevantObjects("*.ElmGenstat")
target_gen_obj = next((g for g in all_gens if g.GetAttribute("loc_name") in [TARGET_GEN_TRIP, TARGET_GEN_TRIP.replace("_", " ")]), None)

freq_bus = all_buses[FREQ_MONITOR_BUS_IDX]
volt_bus = all_buses[VOLT_MONITOR_BUS_IDX]

# --- Pengelompokan Beban Berdasarkan Pattern ---
pattern_list = [int(c) for c in LOAD_SHED_PATTERN]

# Hitung total MW SELURUH beban di sistem (basis persentase)
total_system_mw = sum(load.GetAttribute("plini") for load in all_loads)
print(f"\n[UFLS] Total MW seluruh load di sistem: {total_system_mw:.3f} MW")

# Pisahkan load berdasarkan pattern: prioritas (1) dan cadangan (0)
priority_loads = []  # Load ber-pattern 1 (diambil duluan)
backup_loads = []    # Load ber-pattern 0 (diambil jika pattern tidak cukup)

for i, load in enumerate(all_loads):
    info = {"obj": load, "name": load.GetAttribute("loc_name"), "mw": load.GetAttribute("plini"), "idx": i}
    if i < len(pattern_list) and pattern_list[i] == 1:
        priority_loads.append(info)
    else:
        backup_loads.append(info)

total_priority_mw = sum(l["mw"] for l in priority_loads)
print(f"[UFLS] Total beban pattern=1 (prioritas): {total_priority_mw:.3f} MW")
print(f"[UFLS] Total beban pattern=0 (cadangan) : {sum(l['mw'] for l in backup_loads):.3f} MW")

# --- Hitung target MW per stage (dari total MW sistem) ---
stage_targets = [s["pct"] * total_system_mw for s in STAGES]

# --- Pengisian stage: prioritas pattern dulu, lalu non-pattern ---
stage_loads = [[] for _ in range(len(STAGES))]
available_priority = list(priority_loads)  # Copy agar list asli tidak berubah
available_backup = list(backup_loads)

for stage_idx in range(len(STAGES)):
    target_mw = stage_targets[stage_idx]
    acc_mw = 0.0
    
    # Fase 1: Ambil dari load ber-pattern 1 (prioritas)
    remaining_priority = []
    for l in available_priority:
        if acc_mw < target_mw:
            stage_loads[stage_idx].append(l)
            acc_mw += l["mw"]
        else:
            remaining_priority.append(l)
    available_priority = remaining_priority
    
    # Fase 2: Jika masih kurang, ambil dari load ber-pattern 0 (urutan awal)
    if acc_mw < target_mw:
        remaining_backup = []
        for l in available_backup:
            if acc_mw < target_mw:
                stage_loads[stage_idx].append(l)
                acc_mw += l["mw"]
            else:
                remaining_backup.append(l)
        available_backup = remaining_backup

# --- Cetak Rencana UFLS ---
print("\n" + "="*60)
print(f"{'RENCANA UFLS (PENGELOMPOKAN BEBAN)':^60}")
print(f"{'Basis: Total MW Sistem = ' + f'{total_system_mw:.3f} MW':^60}")
print("="*60)
for i, s_loads in enumerate(stage_loads):
    tot_mw = sum(l["mw"] for l in s_loads)
    target_mw = stage_targets[i]
    names = ", ".join([l["name"] for l in s_loads])
    src_info = ", ".join([f"{l['name']}({'P' if l['idx'] < len(pattern_list) and pattern_list[l['idx']] == 1 else 'B'})" for l in s_loads])
    print(f"STAGE {i+1} (Limit: {STAGES[i]['limit']} Hz) | Target: {STAGES[i]['pct']*100:.0f}% = {target_mw:.3f} MW")
    print(f"  -> Terisi : {tot_mw:.3f} MW")
    print(f"  -> Beban  : {src_info}")
    print(f"             (P=Priority/Pattern, B=Backup/Non-Pattern)")
print("="*60 + "\n")

# --- RMS Event Initialization ---

sc = app.GetActiveStudyCase()
evt_folder = app.GetFromStudyCase("IntEvt")
if not evt_folder:
    evt_folder = sc.CreateObject("IntEvt", "Events")
for evt in evt_folder.GetContents():
    evt.Delete()
app.ClearRecycleBin()

# Event Trip Generator
if target_gen_obj:
    gen_evt = evt_folder.CreateObject("EvtSwitch", f"Trip Gen {TARGET_GEN_TRIP}")
    gen_evt.p_target = target_gen_obj
    gen_evt.time = TIME_GEN_TRIP
    try:
        gen_evt.i_all = 0
        gen_evt.i_switch = 0
    except: pass
    gen_evt.outserv = 0

# Event Load Shedding (Dibuat & di-disable di awal)
for stage_idx, s_loads in enumerate(stage_loads):
    for l in s_loads:
        load_evt = evt_folder.CreateObject("EvtSwitch", f"Shed {l['name']}")
        load_evt.p_target = l["obj"]
        try:
            load_evt.i_all = 0
            load_evt.i_switch = 0
        except: pass
        load_evt.outserv = 1  # Disabled by default
        l["evt"] = load_evt

# --- UFLS Iteration Logic ---

def run_simulation(active_stages_times):
    """Menjalankan simulasi RMS dengan daftar waktu trip per stage."""
    # Setup Result Object
    inc.p_resvar = elmres
    elmres.Clear()
    elmres.AddVars(freq_bus, "m:fehz")
    elmres.AddVars(volt_bus, "m:u")

    # Update Load Shedding Events
    for stage_idx, s_loads in enumerate(stage_loads):
        if stage_idx < len(active_stages_times):
            # Stage ini harus aktif
            shed_time = active_stages_times[stage_idx]
            for l in s_loads:
                l["evt"].time = shed_time
                l["evt"].outserv = 0
        else:
            # Stage ini belum/tidak aktif
            for l in s_loads:
                l["evt"].outserv = 1

    # Reset In-Service Status untuk grid
    for l in all_loads: l.SetAttribute("outserv", 0)
    for g in all_gens: g.SetAttribute("outserv", 0)

    # Execute
    err_inc = inc.Execute()
    if err_inc != 0: return None, None
    err_sim = sim.Execute()
    if err_sim != 0: return None, None

    # Ekstrak
    elmres.Load()
    n_rows = elmres.GetNumberOfRows()
    freq_col = elmres.FindColumn(freq_bus, "m:fehz")
    volt_col = elmres.FindColumn(volt_bus, "m:u")

    times, freqs, volts = [], [], []
    for row in range(n_rows):
        t_val = elmres.GetValue(row, -1)
        times.append(t_val[1] if isinstance(t_val, (list, tuple)) else t_val)
        if freq_col >= 0:
            f_val = elmres.GetValue(row, freq_col)
            freqs.append(f_val[1] if isinstance(f_val, (list, tuple)) else f_val)
        if volt_col >= 0:
            v_val = elmres.GetValue(row, volt_col)
            volts.append(v_val[1] if isinstance(v_val, (list, tuple)) else v_val)

    freq_df = pd.DataFrame({"Time_s": times, "Frequency_Hz": freqs}) if freqs else None
    volt_df = pd.DataFrame({"Time_s": times, f"Voltage_{volt_bus.GetAttribute('loc_name')}_pu": volts}) if volts else None
    elmres.Release()

    return freq_df, volt_df

# --- Looping UFLS ---
active_stages_times = []  # Menyimpan waktu pelepasan beban tiap stage
final_freq_df = None
final_volt_df = None
logs = []

print("[UFLS] Memulai Simulasi Iteratif...\n")

for current_stage in range(len(STAGES) + 1):
    if current_stage == 0:
        print("--- ITERASI 0: Base Case (Tanpa UFLS) ---")
    else:
        print(f"--- ITERASI {current_stage}: Telah mengeksekusi Stage 1 hingga Stage {current_stage} ---")
        
    freq_df, volt_df = run_simulation(active_stages_times)
    
    if freq_df is None:
        print("[!] Simulasi Error.")
        break
        
    final_freq_df, final_volt_df = freq_df, volt_df  # Simpan sebagai hasil terbaru
    
    if current_stage < len(STAGES):
        # Cek apakah frekuensi turun menyentuh limit stage berikutnya
        next_limit = STAGES[current_stage]["limit"]
        times = freq_df["Time_s"].values
        freqs = freq_df["Frequency_Hz"].values
        
        hit_time = None
        for t, f in zip(times, freqs):
            # Hanya periksa waktu setelah event sebelumnya (atau setelah 1.0s)
            start_check_time = active_stages_times[-1] if active_stages_times else TIME_GEN_TRIP
            if t > start_check_time and f <= next_limit:
                hit_time = t
                break
                
        if hit_time is not None:
            shed_time = hit_time + DELAY_UFLS
            print(f"  -> [HIT] Frekuensi menyentuh {next_limit} Hz pada t = {hit_time:.4f} s")
            print(f"  -> [ACTION] Menjadwalkan Pelepasan Stage {current_stage+1} pada t = {shed_time:.4f} s\n")
            active_stages_times.append(shed_time)
            logs.append({"Stage": current_stage + 1, "Hit_Freq_Hz": next_limit, "Hit_Time_s": hit_time, "Shed_Time_s": shed_time})
        else:
            print(f"  -> [AMAN] Frekuensi tidak menyentuh {next_limit} Hz. Pemulihan berhasil!")
            break
    else:
        print("  -> [INFO] Semua Stage (1-4) telah dieksekusi.")

# --- Hitung Settling Time ---
settling_time = RMS_TSTOP
if final_freq_df is not None:
    times = final_freq_df["Time_s"].values
    freqs = final_freq_df["Frequency_Hz"].values
    f_final = freqs[-1]
    f_min = f_final - FREQ_TOLERANCE
    f_max = f_final + FREQ_TOLERANCE

    for i in range(len(times)):
        if f_min <= freqs[i] <= f_max:
            settle_ok = True
            window_covered = False
            for j in range(i, len(times)):
                if times[j] - times[i] >= FREQ_SETTLE_WINDOW:
                    window_covered = True
                    break
                if not (f_min <= freqs[j] <= f_max):
                    settle_ok = False
                    break
            if settle_ok and window_covered:
                settling_time = times[i]
                break

logs.append({
    "Stage": "FINAL", 
    "Hit_Freq_Hz": "N/A", 
    "Hit_Time_s": "N/A", 
    "Shed_Time_s": "N/A", 
    "Message": f"Pemulihan Berhasil. Settling Time: {settling_time:.3f} s"
})

# --- Result Extraction & Export ---

print(f"\n[EXPORT] Menyimpan hasil akhir ke {EXCEL_OUTPUT}...")
try:
    with pd.ExcelWriter(EXCEL_OUTPUT, engine='openpyxl') as writer:
        # Sheet 1: Logs
        if logs:
            pd.DataFrame(logs).to_excel(writer, sheet_name="UFLS_Log", index=False)
        else:
            pd.DataFrame([{"Message": "Tidak ada Stage UFLS yang terpicu (Frekuensi Aman)"}]).to_excel(writer, sheet_name="UFLS_Log", index=False)

        # Sheet 2 & 3: Time Series
        if final_freq_df is not None:
            final_freq_df.to_excel(writer, sheet_name="Frequency", index=False)
        if final_volt_df is not None:
            final_volt_df.to_excel(writer, sheet_name="Voltage", index=False)

        # Sheet 4: Load Detail
        load_rows = []
        for i, load in enumerate(all_loads):
            status = "NO SHED"
            for stage_idx, s_loads in enumerate(stage_loads):
                if stage_idx < len(active_stages_times): # Jika stage ini dieksekusi
                    if any(l["idx"] == i for l in s_loads):
                        status = f"SHED (Stage {stage_idx+1} @ {active_stages_times[stage_idx]:.2f}s)"
            
            load_rows.append({
                "No": i + 1,
                "Load_Name": load.GetAttribute("loc_name"),
                "MW": load.GetAttribute("plini"),
                "Pattern": 1 if i < len(pattern_list) and pattern_list[i] == 1 else 0,
                "Status_Akhir": status
            })
        pd.DataFrame(load_rows).to_excel(writer, sheet_name="Load_Detail", index=False)

    print(f"[EXPORT] Berhasil! File tersimpan: {EXCEL_OUTPUT}")
except Exception as e:
    print(f"[ERROR] Gagal export: {e}")

print("\nSimulasi UFLS Bertingkat Selesai.")
