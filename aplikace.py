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
st.title("✂️ Konfigurátor Stavinvest")
st.info("💡 **Nová funkce:** Rozvinutou šíři (RŠ) nyní zadáváte ručně v milimetrech pro každý prvek zvlášť.")

# ==========================================
# MODULOVÝ PRUHOVÝ ALGORITMUS (PRO PRŮBĚŽNÉ ŘEZY)
# ==========================================
def pack_module_strips(items, coil_w, max_l, allow_rotation=True):
    best_modules = None
    best_len = float('inf')
    
    for iteration in range(200):
        test_items = copy.deepcopy(items)
        
        if iteration == 0:
            test_items.sort(key=lambda x: x['L'], reverse=True)
        elif iteration == 1:
            test_items.sort(key=lambda x: x['rš'], reverse=True)
        else:
            random.shuffle(test_items)
            
        for it in test_items:
            can_std = (it['L'] <= max_l and it['rš'] <= coil_w)
            can_rot = (allow_rotation and it['rš'] <= max_l and it['L'] <= coil_w)
            
            if iteration < 2:
                if can_std: it['dx'], it['dy'], it['rotated'] = it['L'], it['rš'], False
                elif can_rot: it['dx'], it['dy'], it['rotated'] = it['rš'], it['L'], True
                else: it['dx'], it['dy'], it['rotated'] = it['L'], it['rš'], False 
            else:
                if can_std and can_rot:
                    if random.random() > 0.5: it['dx'], it['dy'], it['rotated'] = it['rš'], it['L'], True
                    else: it['dx'], it['dy'], it['rotated'] = it['L'], it['rš'], False
                elif can_rot: it['dx'], it['dy'], it['rotated'] = it['rš'], it['L'], True
                else: it['dx'], it['dy'], it['rotated'] = it['L'], it['rš'], False

        groups = defaultdict(list)
        for it in test_items:
            groups[it['dy']].append(it)
            
        strips = []
        for dy, group_items in groups.items():
            if iteration % 2 == 0:
                group_items.sort(key=lambda x: x['dx'], reverse=True)
            else:
                random.shuffle(group_items)
                
            current_strips = []
            for it in group_items:
                placed = False
                for s in current_strips:
                    if s['l'] + it['dx'] <= max_l:
                        it['x'] = s['l']
                        s['items'].append(it)
                        s['l'] += it['dx']
                        placed = True
                        break
                if not placed:
                    it['x'] = 0
                    current_strips.append({'w': dy, 'l': it['dx'], 'items': [it]})
            strips.extend(current_strips)
            
        strips.sort(key=lambda s: s['l'], reverse=True)
        modules = []
        
        for s in strips:
            placed = False
            for m in modules:
                if m['used_w'] + s['w'] <= coil_w:
                    s['y'] = m['used_w']
                    for it in s['items']:
                        it['y'] = s['y']
                    m['strips'].append(s)
                    m['used_w'] += s['w']
                    m['l'] = max(m['l'], s['l'])
                    placed = True
                    break
            if not placed:
                s['y'] = 0
                for it in s['items']:
                    it['y'] = 0
                modules.append({'used_w': s['w'], 'l': s['l'], 'strips': [s]})
                
        tot_len = sum(m['l'] for m in modules)
        if tot_len < best_len:
            best_len = tot_len
            best_modules = modules

    formatted_bins = []
    if best_modules:
        for m in best_modules:
            placed = []
            for s in m['strips']:
                for it in s['items']:
                    it['draw_w'] = it['dx']
                    it['draw_h'] = it['dy']
                    placed.append(it)
            formatted_bins.append({
                'w_coil': coil_w,
                'odvinuto_mm': m['l'],
                'placed': placed
            })
            
    return formatted_bins

# --- INICIALIZACE NASTAVENÍ A DAT ---
if 'config' not in st.session_state:
    st.session_state.config = {"cena_ohyb": 10.0, "max_delka": 4000, "presah": 40, "povolit_rotaci": True}

if 'materialy_df' not in st.session_state:
    st.session_state.materialy_df = pd.DataFrame([
        {"Materiál": "svitek POZINK 0,55x1000mm", "Interní kód SI": "0160P003", "Šířka (mm)": 1000, "Cena/m2": 200.0, "Max délka tabule (mm)": 50000},
        {"Materiál": "svitek POZINK 0,55x670mm", "Interní kód SI": "0160P002", "Šířka (mm)": 670, "Cena/m2": 218.0, "Max délka tabule (mm)": 50000},
        {"Materiál": "svitek POZINK 0,5x1250mm PES STANDARD BARVY O+SF", "Interní kód SI": "0160LP0107016O+SF", "Šířka (mm)": 1250, "Cena/m2": 282.0, "Max délka tabule (mm)": 50000},
        {"Materiál": "svitek POZINK 0,5x1250mm PES NESTANDARD O+SF", "Interní kód SI": "0160LP0109010O+SF", "Šířka (mm)": 1250, "Cena/m2": 301.0, "Max délka tabule (mm)": 50000},
        {"Materiál": "Comax FALC POZINK 0,5x620mm PES šedá J+SF", "Interní kód SI": "0160LP0017016J+SF", "Šířka (mm)": 620, "Cena/m2": 456.0, "Max délka tabule (mm)": 50000},
        {"Materiál": "svitek TITANZINEK 0,6x1000mm", "Interní kód SI": "0160T003", "Šířka (mm)": 1000, "Cena/m2": 611.0, "Max délka tabule (mm)": 50000},
        {"Materiál": "svitek TITANZINEK 0,6x670mm", "Interní kód SI": "0160T002", "Šířka (mm)": 670, "Cena/m2": 611.0, "Max délka tabule (mm)": 50000},
        {"Materiál": "svitek MĚĎ 0,55x1000mm", "Interní kód SI": "0160M011000", "Šířka (mm)": 1000, "Cena/m2": 2120.0, "Max délka tabule (mm)": 50000},
        {"Materiál": "svitek MĚĎ 0,55x670mm", "Interní kód SI": "0160M010670", "Šířka (mm)": 670, "Cena/m2": 2120.0, "Max délka tabule (mm)": 50000},
        {"Materiál": "PREFA svitek CLR 0,7x1000 PE", "Interní kód SI": "65P31105", "Šířka (mm)": 1000, "Cena/m2": 457.0, "Max délka tabule (mm)": 30000},
        {"Materiál": "PREFA svitek Prefalz 0,7x1000 hladký", "Interní kód SI": "65P40100", "Šířka (mm)": 1000, "Cena/m2": 578.0, "Max délka tabule (mm)": 30000},
        {"Materiál": "PREFA svitek Prefalz 0,7x650 hladký", "Interní kód SI": "65P40200", "Šířka (mm)": 650, "Cena/m2": 578.0, "Max délka tabule (mm)": 50000},
        {"Materiál": "Comax FALC AL 0,7x600mm", "Interní kód SI": "0160ALCO0706007016", "Šířka (mm)": 600, "Cena/m2": 622.0, "Max délka tabule (mm)": 50000},
        {"Materiál": "tabule AL 0,6x1000x2000 PES jednostranná s folií", "Interní kód SI": "0150AL06100020007016J+SF", "Šířka (mm)": 1000, "Cena/m2": 421.0, "Max délka tabule (mm)": 2000},
        {"Materiál": "tabule PVC 0,6x1000x2000 ROOFPLAN 7035", "Interní kód SI": "0150PVC0037035", "Šířka (mm)": 1000, "Cena/m2": 591.0, "Max délka tabule (mm)": 2000}
    ])

if 'prvky_df' not in st.session_state:
    st.session_state.prvky_df = pd.DataFrame([
        {"Typ prvku": "Závětrná lišta spodní", "Ohyby": 6},
        {"Typ prvku": "Závětrná lišta pultová", "Ohyby": 6},
        {"Typ prvku": "Okapnice", "Ohyby": 2},
        {"Typ prvku": "Lemování ke zdi", "Ohyby": 3},
        {"Typ prvku": "Úžlabí", "Ohyby": 3},
        {"Typ prvku": "Úžlabí s drážkou", "Ohyby": 5},
        {"Typ prvku": "Atikový plech", "Ohyby": 4},
        {"Typ prvku": "L lišta", "Ohyby": 2},
        {"Typ prvku": "Stěnová lišta", "Ohyby": 2},
        {"Typ prvku": "Parapet", "Ohyby": 3},
        {"Typ prvku": "Parapet včetně boků", "Ohyby": 3},
        {"Typ prvku": "Atypický výrobek", "Ohyby": 9}
    ])

if 'zakazka' not in st.session_state:
    st.session_state.zakazka = []

mat_dict = {r["Materiál"]: r for _, r in st.session_state.materialy_df.iterrows()}
prv_dict = {r["Typ prvku"]: r for _, r in st.session_state.prvky_df.iterrows()}

# --- ZÁLOŽKY ---
tab_kalk, tab_nakres, tab_data, tab_nastaveni = st.tabs(["🧮 Kalkulátor", "📐 Nákres 2D Řezů", "⚙️ Data (Ceník)", "🔧 Nastavení"])

# ==========================================
# ZÁLOŽKA: NASTAVENÍ A DATA
# ==========================================
with tab_nastaveni:
    st.header("🔧 Nastavení výroby")
    st.session_state.config["cena_ohyb"] = st.number_input("Cena za ohyb (Kč)", value=float(st.session_state.config["cena_ohyb"]))

with tab_data:
    st.header("⚙️ Správa dat (Ceník a materiály)")
    st.write("Aplikace je plně odemčena pro úpravy ceníku i prvků.")
    st.session_state.materialy_df = st.data_editor(st.session_state.materialy_df, num_rows="dynamic", key="em", use_container_width=True)
    st.session_state.prvky_df = st.data_editor(st.session_state.prvky_df, num_rows="dynamic", key="ep", use_container_width=True)

# ==========================================
# ZÁLOŽKA: KALKULÁTOR
# ==========================================
with tab_kalk:
    
    # --- 1. OBECNÉ ÚDAJE (PŘESUNUTO NAHORU PRO ZAROVNÁNÍ SLOUPCŮ NÍŽE) ---
    st.header("1. Obecné údaje")
    col_top1, col_top2 = st.columns(2)
    with col_top1:
        v_odberatel = st.text_input("Odběratel / Název zakázky", st.session_state.get('odberatel', ''))
        st.session_state.odberatel = v_odberatel
        v_mat = st.selectbox("Materiál (pro celou zakázku)", list(mat_dict.keys()))
    with col_top2:
        st.write("**Parametry výroby:**")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.session_state.config["max_delka"] = st.number_input("Délka ohýbačky (mm)", value=int(st.session_state.config["max_delka"]))
        with col_p2:
            st.session_state.config["presah"] = st.number_input("Přesah spojů (mm)", value=int(st.session_state.config["presah"]))
        st.session_state.config["povolit_rotaci"] = st.checkbox("🔄 Povolit otáčení dílů o 90°", value=st.session_state.config["povolit_rotaci"])
        
    st.markdown("---")

    # --- 2. PŘIDAT POLOŽKU A VÝPOČET (VEDLE SEBE) ---
    col_in, col_res = st.columns([1, 2])
    
    with col_in:
        st.header("2. Přidat položku")
        v_prvek = st.selectbox("Prvek", list(prv_dict.keys()))
        
        v_rs = st.number_input("Rozvinutá šíře - RŠ (mm)", value=250, min_value=10, step=1)
        
        default_ohyby = int(prv_dict[v_prvek]["Ohyby"]) if v_prvek in prv_dict else 0
        v_ohyby = st.number_input("Počet ohybů (lze upravit)", value=default_ohyby, min_value=0)
        v_m = st.number_input("Délka (m)", value=2.5, step=0.1)
        v_ks = st.number_input("Kusů", min_value=1, value=1)
        v_priplatek = st.number_input("Atyp. příplatek/ks (Kč)", value=0.0, step=50.0)
        
        if st.button("➕ Přidat do zakázky", type="primary", use_container_width=True):
            st.session_state.zakazka.append({
                "Prvek": v_prvek,
                "RŠ (mm)": v_rs,
                "Ohyby": v_ohyby,
                "Metrů": v_m, 
                "Kusů": v_ks,
                "Atyp příplatek/ks (Kč)": v_priplatek
            })
            st.rerun()
            
        if st.button("🗑️ Smazat vše", use_container_width=True):
            st.session_state.zakazka = []
            st.session_state.generated_figs = []
            st.session_state.calc_done = False
            st.rerun()

    with col_res:
        st.header("Výpočet a Optimalizace")
        if st.session_state.zakazka:
            df_zakazka = pd.DataFrame(st.session_state.zakazka)
            df_zakazka.insert(0, 'Řádek', range(1, len(df_zakazka) + 1))
            
            edited_zakazka_df = st.data_editor(
                df_zakazka,
                column_config={
                    "Řádek": st.column_config.Column("Řádek", disabled=True),
                    "Prvek": st.column_config.SelectboxColumn("Prvek", options=list(prv_dict.keys()), required=True),
                    "RŠ (mm)": st.column_config.NumberColumn("RŠ (mm)", min_value=10, step=1, required=True),
                    "Ohyby": st.column_config.NumberColumn("Ohyby", min_value=0, step=1, required=True),
                    "Metrů": st.column_config.NumberColumn("Metrů", min_value=0.1, step=0.1, required=True),
                    "Kusů": st.column_config.NumberColumn("Kusů", min_value=1, step=1, required=True),
                    "Atyp příplatek/ks (Kč)": st.column_config.NumberColumn("Atyp příplatek/ks (Kč)", min_value=0.0, step=10.0, required=True)
                },
                hide_index=True,
                num_rows="dynamic",
                use_container_width=True,
                key="editor_zakazka"
            )
            
            updated_zakazka = edited_zakazka_df.drop(columns=['Řádek']).to_dict('records')
            st.session_state.zakazka = updated_zakazka
            
            if st.button("🚀 SPOČÍTAT ZAKÁZKU", type="primary", use_container_width=True):
                with st.spinner("🧠 Vytvářím výrobní moduly pro stroje a kreslím plány..."):
                    items = []
                    cena_prace = 0
