import streamlit as st
import pandas as pd
import math
import io
import copy
import random
from collections import defaultdict
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from openpyxl.drawing.image import Image as xlImage
from openpyxl.utils import get_column_letter

# --- NASTAVENÍ STRÁNKY ---
st.set_page_config(page_title="Konfigurátor Stavinvest", page_icon="✂️", layout="wide")

# ==========================================
# SPOLEČNÝ ALGORITMUS PRO OBĚ VERZE
# ==========================================
def pack_module_strips(items, coil_w, max_l, allow_rotation=True):
    best_modules = None
    best_len = float('inf')
    for iteration in range(100):
        test_items = copy.deepcopy(items)
        random.shuffle(test_items)
        for it in test_items:
            can_std = (it['L'] <= max_l and it['rš'] <= coil_w)
            can_rot = (allow_rotation and it['rš'] <= max_l and it['L'] <= coil_w)
            if can_std and can_rot:
                if random.random() > 0.5: it['dx'], it['dy'], it['rotated'] = it['rš'], it['L'], True
                else: it['dx'], it['dy'], it['rotated'] = it['L'], it['rš'], False
            elif can_rot: it['dx'], it['dy'], it['rotated'] = it['rš'], it['L'], True
            else: it['dx'], it['dy'], it['rotated'] = it['L'], it['rš'], False 
        groups = defaultdict(list)
        for it in test_items: groups[it['dy']].append(it)
        strips = []
        for dy, group_items in groups.items():
            current_strips = []
            for it in group_items:
                placed = False
                for s in current_strips:
                    if s['l'] + it['dx'] <= max_l:
                        it['x'] = s['l']; s['items'].append(it); s['l'] += it['dx']
                        placed = True; break
                if not placed:
                    it['x'] = 0; current_strips.append({'w': dy, 'l': it['dx'], 'items': [it]})
            strips.extend(current_strips)
        strips.sort(key=lambda s: s['l'], reverse=True)
        modules = []
        for s in strips:
            placed = False
            for m in modules:
                if m['used_w'] + s['w'] <= coil_w:
                    s['y'] = m['used_w']
                    for it in s['items']: it['y'] = s['y']
                    m['strips'].append(s); m['used_w'] += s['w']
                    m['l'] = max(m['l'], s['l']); placed = True; break
            if not placed:
                s['y'] = 0
                for it in s['items']: it['y'] = 0
                modules.append({'used_w': s['w'], 'l': s['l'], 'strips': [s]})
        tot_len = sum(m['l'] for m in modules)
        if tot_len < best_len:
            best_len = tot_len; best_modules = modules
    formatted_bins = []
    if best_modules:
        for m in best_modules:
            placed = []
            for s in m['strips']:
                for it in s['items']:
                    it['draw_w'] = it['dx']; it['draw_h'] = it['dy']
                    placed.append(it)
            formatted_bins.append({'w_coil': coil_w, 'odvinuto_mm': m['l'], 'placed': placed})
    return formatted_bins

# ==========================================
# TAJNÝ PŘEPÍNAČ PRO BETA VERZI
# ==========================================
st.sidebar.markdown("---")
secret_mode = st.sidebar.text_input("Tajný mód (Vývojář):", type="password")

if secret_mode == "BETA2026":
    # ---------------------------------------------------------
    # 🧪 BETA VERZE (S VOLNÝM ZADÁVÁNÍM RŠ - BEZ HESEL)
    # ---------------------------------------------------------
    st.title("🧪 Konfigurátor - BETA VERZE (Volné RŠ)")
    st.info("Toto je testovací rozhraní. Zadáváte RŠ ručně.")

    if 'b_config' not in st.session_state: st.session_state.b_config = {"cena_ohyb": 10.0, "max_delka": 4000, "presah": 40, "povolit_rotaci": True}
    if 'b_materialy' not in st.session_state:
        st.session_state.b_materialy = pd.DataFrame([
            {"Materiál": "svitek POZINK 0,55x1000mm", "Šířka (mm)": 1000, "Cena/m2": 200.0, "Max délka tabule (mm)": 50000},
            {"Materiál": "svitek POZINK 0,55x670mm", "Šířka (mm)": 670, "Cena/m2": 218.0, "Max délka tabule (mm)": 50000},
            {"Materiál": "tabule AL 0,6x1000x2000 PES", "Šířka (mm)": 1000, "Cena/m2": 421.0, "Max délka tabule (mm)": 2000}
        ])
    if 'b_prvky' not in st.session_state:
        st.session_state.b_prvky = pd.DataFrame([
            {"Typ prvku": "Závětrná lišta spodní", "Ohyby": 6},
            {"Typ prvku": "Závětrná lišta pultová", "Ohyby": 6},
            {"Typ prvku": "Okapnice", "Ohyby": 2},
            {"Typ prvku": "Lemování ke zdi", "Ohyby": 3},
            {"Typ prvku": "Úžlabí", "Ohyby": 3},
            {"Typ prvku": "Parapet", "Ohyby": 3},
            {"Typ prvku": "Atypický výrobek", "Ohyby": 9}
        ])
    if 'b_zakazka' not in st.session_state: st.session_state.b_zakazka = []

    b_mat_dict = {r["Materiál"]: r for _, r in st.session_state.b_materialy.iterrows()}
    b_prv_dict = {r["Typ prvku"]: r for _, r in st.session_state.b_prvky.iterrows()}

    tab_b_kalk, tab_b_nakres, tab_b_data = st.tabs(["🧮 Kalkulátor (BETA)", "📐 Nákres", "⚙️ Data (BETA)"])

    with tab_b_data:
        st.session_state.b_materialy = st.data_editor(st.session_state.b_materialy, num_rows="dynamic", use_container_width=True)
        st.session_state.b_prvky = st.data_editor(st.session_state.b_prvky, num_rows="dynamic", use_container_width=True)

    with tab_b_kalk:
        col_in, col_res = st.columns([1, 2])
        with col_in:
            v_mat = st.selectbox("Materiál", list(b_mat_dict.keys()))
            v_prvek = st.selectbox("Prvek", list(b_prv_dict.keys()))
            v_rs = st.number_input("Rozvinutá šíře - RŠ (mm)", value=250, min_value=10, step=1)
            v_ohyby = st.number_input("Ohyby", value=int(b_prv_dict[v_prvek]["Ohyby"]))
            v_m = st.number_input("Délka (m)", value=2.0, step=0.1)
            v_ks = st.number_input("Kusů", min_value=1, value=1)
            
            if st.button("➕ Přidat (BETA)", use_container_width=True):
                st.session_state.b_zakazka.append({"Prvek": v_prvek, "RŠ (mm)": v_rs, "Ohyby": v_ohyby, "Metrů": v_m, "Kusů": v_ks})
                st.rerun()
            if st.button("🗑️ Smazat (BETA)"): st.session_state.b_zakazka = []; st.rerun()

        with col_res:
            if st.session_state.b_zakazka:
                df_zak = pd.DataFrame(st.session_state.b_zakazka)
                st.data_editor(df_zak, use_container_width=True)
                if st.button("🚀 SPOČÍTAT (BETA)", type="primary"):
                    items = []
                    for idx, p in enumerate(st.session_state.b_zakazka):
                        for _ in range(int(p["Kusů"])): items.append({"id": idx+1, "Prvek": p['Prvek'], "L": p['Metrů']*1000, "rš": p['RŠ (mm)']})
                    st.session_state.b_vysledky = pack_module_strips(items, b_mat_dict[v_mat]["Šířka (mm)"], st.session_state.b_config["max_delka"])
                    st.session_state.b_mat_zvoleny = b_mat_dict[v_mat]
                    st.rerun()
                    
    with tab_b_nakres:
        if 'b_vysledky' in st.session_state:
            for i, b in enumerate(st.session_state.b_vysledky):
                fig, ax = plt.subplots(figsize=(10, 2.5))
                ax.add_patch(patches.Rectangle((0, 0), b['odvinuto_mm'], b['w_coil'], fill=False, edgecolor='black', lw=2))
                for p in b['placed']:
                    ax.add_patch(patches.Rectangle((p['x'], p['y']), p['draw_w'], p['draw_h'], facecolor='skyblue', edgecolor='black'))
                    ax.text(p['x']+p['draw_w']/2, p['y']+p['draw_h']/2, f"Ř.{p['id']}\n{p['rš']}mm", ha='center', va='center', fontsize=8)
                ax.set_xlim(0, st.session_state.b_config["max_delka"] * 1.05)
                ax.set_title(f"Modul {i+1}: Odvinout {b['odvinuto_mm']/1000:.2f} m")
                st.pyplot(fig)


else:
    # ---------------------------------------------------------
    # 🏢 PRODUKČNÍ VERZE (PŮVODNÍ STAV PRO KOLEGY)
    # ---------------------------------------------------------
    st.title("✂️ Konfigurátor Stavinvest")
    
    UZIVATELE = {"admin@stavinvest.cz": "HlavniKlempir!", "test1@stavinvest.cz": "PlechovaStrecha1"}
    if 'logged_in' not in st.session_state: st.session_state.logged_in = False

    if not st.session_state.logged_in:
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            st.markdown("<h3>🔒 Přihlášení</h3>", unsafe_allow_html=True)
            with st.form("login"):
                email = st.text_input("E-mail")
                heslo = st.text_input("Heslo", type="password")
                if st.form_submit_button("Přihlásit"):
                    if email in UZIVATELE and UZIVATELE[email] == heslo:
                        st.session_state.logged_in = True
                        st.session_state.current_user = email
                        st.rerun()
                    else: st.error("Chyba!")
        st.stop()

    st.sidebar.write(f"👤 Uživatel: **{st.session_state.current_user}**")
    if st.sidebar.button("Odhlásit"): st.session_state.logged_in = False; st.rerun()

    if 'config' not in st.session_state: st.session_state.config = {"cena_ohyb": 10.0, "max_delka": 4000, "presah": 40, "povolit_rotaci": True}
    if 'materialy_df' not in st.session_state:
        st.session_state.materialy_df = pd.DataFrame([
            {"Materiál": "svitek POZINK 0,55x1000mm", "Šířka (mm)": 1000, "Cena/m2": 200.0, "Max délka tabule (mm)": 50000},
            {"Materiál": "tabule AL 0,6x1000x2000", "Šířka (mm)": 1000, "Cena/m2": 421.0, "Max délka tabule (mm)": 2000}
        ])
    if 'prvky_df' not in st.session_state:
        st.session_state.prvky_df = pd.DataFrame([
            {"Typ prvku": "okapnice do r.š. 200", "RŠ (mm)": 200, "Ohyby": 2},
            {"Typ prvku": "okapnice r.š.201-250", "RŠ (mm)": 250, "Ohyby": 2}
        ])
    if 'zakazka' not in st.session_state: st.session_state.zakazka = []

    mat_dict = {r["Materiál"]: r for _, r in st.session_state.materialy_df.iterrows()}
    prv_dict = {r["Typ prvku"]: r for _, r in st.session_state.prvky_df.iterrows()}

    tab_kalk, tab_nakres, tab_data = st.tabs(["🧮 Kalkulátor", "📐 Nákres", "⚙️ Data"])

    with tab_data:
        if st.session_state.current_user == "admin@stavinvest.cz":
            st.session_state.materialy_df = st.data_editor(st.session_state.materialy_df, num_rows="dynamic")
            st.session_state.prvky_df = st.data_editor(st.session_state.prvky_df, num_rows="dynamic")
        else:
            st.dataframe(st.session_state.materialy_df)
            st.dataframe(st.session_state.prvky_df)

    with tab_kalk:
        col_in, col_res = st.columns([1, 2])
        with col_in:
            v_mat = st.selectbox("Materiál", list(mat_dict.keys()))
            v_prvek = st.selectbox("Prvek", list(prv_dict.keys()))
            v_ohyby = st.number_input("Ohyby", value=int(prv_dict[v_prvek]["Ohyby"]) if v_prvek in prv_dict else 0)
            v_m = st.number_input("Délka (m)", value=2.5)
            v_ks = st.number_input("Kusů", value=1)
            
            if st.button("➕ Přidat položku", use_container_width=True):
                st.session_state.zakazka.append({"Prvek": v_prvek, "Ohyby": v_ohyby, "Metrů": v_m, "Kusů": v_ks})
                st.rerun()

        with col_res:
            if st.session_state.zakazka:
                st.dataframe(pd.DataFrame(st.session_state.zakazka))
                if st.button("🚀 SPOČÍTAT ZAKÁZKU", type="primary"):
                    items = []
                    for idx, p in enumerate(st.session_state.zakazka):
                        rs = prv_dict[p["Prvek"]]["RŠ (mm)"]
                        for _ in range(int(p["Kusů"])): items.append({"id": idx+1, "Prvek": p['Prvek'], "L": p['Metrů']*1000, "rš": rs})
                    st.session_state.vysledky = pack_module_strips(items, mat_dict[v_mat]["Šířka (mm)"], st.session_state.config["max_delka"])
                    st.rerun()

    with tab_nakres:
        if 'vysledky' in st.session_state:
            for i, b in enumerate(st.session_state.vysledky):
                fig, ax = plt.subplots(figsize=(10, 2.5))
                ax.add_patch(patches.Rectangle((0, 0), b['odvinuto_mm'], b['w_coil'], fill=False, edgecolor='black', lw=2))
                for p in b['placed']:
                    ax.add_patch(patches.Rectangle((p['x'], p['y']), p['draw_w'], p['draw_h'], facecolor='green', edgecolor='black'))
                    ax.text(p['x']+p['draw_w']/2, p['y']+p['draw_h']/2, f"{p['Prvek']}", ha='center', va='center', fontsize=8)
                ax.set_xlim(0, st.session_state.config["max_delka"] * 1.05)
                st.pyplot(fig)
