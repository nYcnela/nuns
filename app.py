import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import matplotlib.patches as mpatches
import textwrap
import io
import os
import zipfile
import re
from docx import Document
import numpy as np

# ==========================================
# 1. KONFIGURACJA I STAŁE
# ==========================================
st.set_page_config(page_title="Filigran", layout="wide")

APP_PASSWORD = os.environ.get("APP_PASSWORD", "b12345")

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        st.title("🔐 Logowanie do Narzędzi")
        password = st.text_input("Podaj hasło:", type="password")
        if st.button("Zaloguj"):
            if password == APP_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Błędne hasło!")
        st.stop()

check_password()

def make_lighter(hex_color, alpha=0.3):
    return mcolors.to_rgba(hex_color, alpha=alpha)

COLORS_NUNS = {
    'Gravelines': '#FFD700', 'London': '#B0C4DE', 'Gosfield': '#2ca02c',
    'Rouen': '#1f77b4', 'Haggerston': '#ff7f0e', 'Scorton': '#d62728',
    'Aire': '#e377c2', 'Britwell': '#eaffea', 'Plymouth': '#9467bd',
    'Dunkirk': '#9ACD32', 'Worcester': '#17becf', 'Deceased': '#808080'
}

# ==========================================
# 2. FUNKCJE POMOCNICZE
# ==========================================

def natural_sort_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"([0-9]+)", str(s))]

def extract_ids_and_text(raw):
    raw = (raw or "").strip().replace("\xa0", " ")
    raw = re.sub(r"[ \t]+\n", "\n", raw)
    
    m = re.match(r"^\[([A-Za-z0-9,\s]+)\]\s*\.?\s*(.*)$", raw, flags=re.DOTALL)
    if m:
        ids = [i.strip() for i in m.group(1).split(",") if i.strip()]
        txt = (m.group(2) or "").strip().lstrip(".").strip()
        return ids, txt
    
    m = re.match(r"^([A-Z])\.\s*(.*)$", raw, flags=re.DOTALL)
    if m:
        return [m.group(1)], (m.group(2) or "").strip()
        
    return [], raw

def split_marker(text):
    t = (text or "").strip().lstrip(".").strip()
    m = re.match(r"^(\d+)\.?\s+(.*)$", t, flags=re.DOTALL)
    if m: return m.group(1), (m.group(2) or "").strip()
    m = re.match(r"^([IVXLCDM]+)\.\s+(.*)$", t, flags=re.DOTALL)
    if m: return m.group(1), (m.group(2) or "").strip()
    return "", t

@st.cache_data(show_spinner=False)
def parse_docx_psalms_v2(file_bytes):
    document = Document(io.BytesIO(file_bytes))
    psalms_data = {}
    
    def psalm_from_header(text):
        m = re.search(r"PSALM\s+(\d+)", text, flags=re.I)
        return f"PSALM {m.group(1)}" if m else None

    for table in document.tables:
        try:
            first_row_text = " ".join([c.text.strip() for c in table.rows[0].cells if c.text])
            ps = psalm_from_header(first_row_text)
        except:
            continue
        if not ps: continue

        rows_data = []
        for r_i, row in enumerate(table.rows[1:], start=1):
            cells = row.cells
            if len(cells) < 3: continue
            
            header_like = ("OFFICIUM" in (cells[0].text or "").upper() or "VULGATA" in (cells[1].text or "").upper())
            if header_like: continue

            row_items = {}
            for idx, col_key in enumerate(["A", "B", "C"]):
                raw = (cells[idx].text or "").strip()
                ids, txt = extract_ids_and_text(raw)
                marker, body = split_marker(txt)
                row_items[col_key] = {
                    "ids": ids,
                    "marker": marker,
                    "text": body,
                    "raw": raw
                }
            rows_data.append(row_items)

        if rows_data:
            psalms_data[ps] = rows_data

    return psalms_data

def build_blocks(rows):
    # Zbieramy ID w kolejności ich pierwszego wystąpienia w tabeli
    ordered_ids = []
    seen_ids = set()
    for row in rows:
        for col in ["A", "B", "C"]:
            for uid in row[col]["ids"]:
                if uid not in seen_ids:
                    ordered_ids.append(uid)
                    seen_ids.add(uid)

    if not ordered_ids:
        sorted_ids = [str(i) for i in range(1, len(rows) + 1)]
        id_to_index = {uid: i for i, uid in enumerate(sorted_ids)}
        blocks = {c: [] for c in ["A", "B", "C"]}
        for i, row in enumerate(rows, start=1):
            uid = str(i)
            for c in ["A", "B", "C"]:
                blocks[c].append({"ids": [uid], "marker": row[c]["marker"], "text": row[c]["text"]})
        return sorted_ids, blocks, id_to_index

    # Zachowujemy kolejność z tabeli, nie sortujemy alfabetycznie
    sorted_ids = ordered_ids
    id_to_index = {uid: i for i, uid in enumerate(sorted_ids)}

    blocks = {c: [] for c in ["A", "B", "C"]}
    seen = set()

    for row in rows:
        for c in ["A", "B", "C"]:
            ids = row[c]["ids"]
            if not ids: continue
            txt = (row[c]["text"] or "").strip()
            marker = (row[c]["marker"] or "").strip()
            
            sig = (c, tuple(ids), marker, txt)
            if sig in seen: continue
            seen.add(sig)
            
            blocks[c].append({"ids": ids, "marker": marker, "text": txt})

    return sorted_ids, blocks, id_to_index

def expand_ids_by_merges(target_ids, blocks):
    target_ids = set(target_ids)
    changed = True
    while changed:
        changed = False
        for c in ["A", "B", "C"]:
            for b in blocks[c]:
                b_ids = set(b["ids"])
                if b_ids & target_ids:
                    new_set = target_ids | b_ids
                    if len(new_set) != len(target_ids):
                        target_ids = new_set
                        changed = True
    return target_ids

# ==========================================
# 3. SILNIK GRAFICZNY (FINALNY)
# ==========================================
def draw_pretty_sankey_final(
    title,
    sorted_ids,
    blocks,
    id_to_index,
    colors,
    labels,
    show_links=True,
    link_color="#BFC5D2",
    link_alpha=0.3,
    font_size=10,
    wrap_chars=40,
    compact=False,
    show_stripe=True,
    ribbon_width_scale=0.4,
    show_verse_nums=True,
    show_ids=True,
    show_row_ids_left=True,    
    show_zebra=True,
    badge_text_colors=("#FFFFFF", "#FFFFFF", "#FFFFFF"),
    show_header=True
):
    def wrap_text_content(text):
        lines = []
        for para in (text or "").split("\n"):
            para = para.strip()
            if not para: continue
            lines.extend(textwrap.wrap(para, width=wrap_chars) or [""])
        return "\n".join(lines) if lines else "—"
    
    # Parametry układu
    LINE_HEIGHT = 0.16  
    MIN_ROW_H = 1.2     
    PADDING = 0.8       
    GAP = 0.15 if compact else 0.22
    
    # 1. Obliczanie wysokości
    slot_heights = {}
    wrapped_cache = {}
    
    for uid in sorted_ids:
        max_lines = 1
        for c in ["A", "B", "C"]:
            for b in blocks[c]:
                if uid in b["ids"]:
                    wrapped = wrap_text_content(b.get("text", ""))
                    key = (c, tuple(b["ids"]))
                    wrapped_cache[key] = wrapped
                    num_lines = len(wrapped.split("\n"))
                    lines_per_id = num_lines / max(1, len(b["ids"]))
                    if lines_per_id > max_lines:
                        max_lines = lines_per_id
        
        slot_heights[uid] = max(MIN_ROW_H, (max_lines * LINE_HEIGHT * font_size / 10) + PADDING)
    
    # 2. Pozycje Y
    y_positions = {} 
    current_y = 0
    for uid in sorted_ids:
        h = slot_heights[uid]
        y_top = current_y
        y_bottom = current_y - h
        y_positions[uid] = (y_top, y_bottom)
        current_y = y_bottom - GAP
    
    total_height = abs(current_y)
    
    id_usage_map = {c: {} for c in ["A", "B", "C"]}
    for c in ["A", "B", "C"]:
        for idx, b in enumerate(blocks[c]):
            for uid in b["ids"]:
                if uid not in id_usage_map[c]:
                    id_usage_map[c][uid] = []
                id_usage_map[c][uid].append(idx)
    
    # 3. Rysowanie
    fig_h = max(6, total_height * 1.1)
    fig, ax = plt.subplots(figsize=(18, fig_h))
    ax.set_facecolor("white")

    col_x = {"A": 0.0, "B": 1.55, "C": 3.10}
    col_w = 1.30
    stripe_w = 0.14 
    
    # Ustalanie marginesów na podstawie show_row_ids_left
    x_min = -0.45 if show_row_ids_left else -0.10
    x_max = 4.50

    # Tło i ID wierszy (lewa strona)
    for i, uid in enumerate(sorted_ids):
        y_top, y_bottom = y_positions[uid]
        h = y_top - y_bottom
        y_center = (y_top + y_bottom) / 2
        
        if show_zebra and i % 2 == 0:
            ax.add_patch(mpatches.Rectangle(
                (x_min, y_bottom - 0.02), x_max - x_min, h + 0.04,
                facecolor=mcolors.to_rgba("#111827", 0.03), edgecolor=None, zorder=0
            ))
        
        # Oznaczenia wierszy (duże litery K, M...) tylko jeśli włączone
        if show_row_ids_left:
            ax.text(-0.15, y_center, uid,
                    ha="right", va="center", fontsize=12, fontweight="bold", color="#111827")

    # Nagłówki (opcjonalne)
    if show_header:
        center_x = col_x["B"] + col_w / 2
        ax.text(center_x, 0.8, title, ha="center", va="center", fontsize=18, fontweight="bold", color="#111827")
        for c, lab in zip(["A", "B", "C"], labels):
            ax.text(col_x[c] + col_w / 2, 0.35, lab, ha="center", va="center", fontsize=12, fontweight="bold", color="#111827")

    anchors = {c: {} for c in ["A", "B", "C"]}

    def draw_card(c, block_idx, ids, marker, text):
        indices = [id_to_index[i] for i in ids if i in id_to_index]
        if not indices: return

        valid_ids = [i for i in ids if i in y_positions]
        if not valid_ids: return
        
        # Slicing Logic
        first_id = valid_ids[0]
        usage_list_top = id_usage_map[c].get(first_id, [])
        y_global_top, y_global_bottom_top = y_positions[first_id]
        
        if len(usage_list_top) > 1:
            usage_list_top.sort()
            my_rank = usage_list_top.index(block_idx)
            count = len(usage_list_top)
            seg_h = (y_global_top - y_global_bottom_top) / count
            card_y_top = y_global_top - (my_rank * seg_h)
        else:
            card_y_top = y_global_top

        last_id = valid_ids[-1]
        usage_list_bot = id_usage_map[c].get(last_id, [])
        y_global_top_bot, y_global_bottom_bot = y_positions[last_id]
        
        if len(usage_list_bot) > 1:
            usage_list_bot.sort()
            my_rank = usage_list_bot.index(block_idx)
            count = len(usage_list_bot)
            seg_h = (y_global_top_bot - y_global_bottom_bot) / count
            card_y_bottom = y_global_top_bot - ((my_rank + 1) * seg_h)
        else:
            card_y_bottom = y_global_bottom_bot

        VISUAL_GAP = 0.06
        draw_y_top = card_y_top - VISUAL_GAP
        draw_y_bottom = card_y_bottom + VISUAL_GAP
        h = draw_y_top - draw_y_bottom
        
        if h < 0.2:
            mid = (draw_y_top + draw_y_bottom) / 2
            h = 0.2
            draw_y_top = mid + 0.1
            draw_y_bottom = mid - 0.1

        base = colors[["A", "B", "C"].index(c)]
        badge_txt_color = badge_text_colors[["A", "B", "C"].index(c)]
        x = col_x[c]

        # Kształty
        ax.add_patch(mpatches.FancyBboxPatch(
            (x + 0.02, draw_y_bottom - 0.02), col_w, h,
            boxstyle="round,pad=0.03,rounding_size=0.12",
            linewidth=0, facecolor=(0, 0, 0, 0.08), zorder=2
        ))
        card_shape = mpatches.FancyBboxPatch(
            (x, draw_y_bottom), col_w, h,
            boxstyle="round,pad=0.03,rounding_size=0.12",
            linewidth=0, facecolor="white", zorder=3
        )
        ax.add_patch(card_shape)

        # Pasek
        text_margin_left = 0.08
        if show_stripe:
            stripe = mpatches.Rectangle(
                (x - 0.05, draw_y_bottom - 0.05), stripe_w + 0.05, h + 0.1,
                facecolor=base, zorder=4
            )
            stripe.set_clip_path(card_shape)
            ax.add_patch(stripe)
            content_start_x = x + stripe_w + text_margin_left
        else:
            content_start_x = x + text_margin_left

        id_label = ", ".join(ids)
        
        # Oblicz środek karty w pionie
        card_center_y = (draw_y_top + draw_y_bottom) / 2
        
        # --- MARKER NA KOLOROWYM PASKU (wyśrodkowany) ---
        if show_stripe and show_verse_nums and marker:
            # Środek paska - lekko przesunięty w lewo od środka geometrycznego
            stripe_center_x = x + stripe_w * 0.4
            ax.text(
                stripe_center_x, card_center_y, marker,
                ha="center", va="center", fontsize=10, fontweight="bold", color=badge_txt_color,
                zorder=5, rotation=0
            )

        # Ramka
        ax.add_patch(mpatches.FancyBboxPatch(
            (x, draw_y_bottom), col_w, h,
            boxstyle="round,pad=0.03,rounding_size=0.12",
            linewidth=1, edgecolor="#E5E7EB", facecolor="none", zorder=6
        ))

        # --- ID I TEKST ---
        key = (c, tuple(ids))
        body = wrapped_cache.get(key, wrap_text_content(text))
        
        if show_ids:
            # ID wyświetlane nad wyśrodkowanym tekstem
            # ID w lewym górnym rogu
            ax.text(
                content_start_x, draw_y_top - 0.08, id_label,
                ha='left', va='top', 
                fontsize=9, fontweight='bold', color=base,
                zorder=5, fontfamily="DejaVu Serif"
            )
            # Tekst wyśrodkowany w pionie (przesunięty lekko w dół przez ID)
            ax.text(
                content_start_x, card_center_y, body,
                ha='left', va='center', fontsize=font_size, color="#0B1220", 
                zorder=5, fontfamily="DejaVu Serif", linespacing=1.35, clip_on=True
            )
        else:
            # ID niewyświetlane - tekst wyśrodkowany w pionie
            ax.text(
                content_start_x, card_center_y, body,
                ha='left', va='center', fontsize=font_size, color="#0B1220", 
                zorder=5, fontfamily="DejaVu Serif", linespacing=1.35, clip_on=True
            )

        # Kotwice
        block_ids_sorted = sorted([i for i in ids if i in id_to_index], key=lambda x: id_to_index[x])
        if block_ids_sorted:
            segment_height = h / len(block_ids_sorted)
            for idx, uid in enumerate(block_ids_sorted):
                seg_y_center = draw_y_top - (idx * segment_height) - (segment_height / 2)
                
                if uid not in anchors[c]: anchors[c][uid] = []
                anchors[c][uid].append({
                    "left": (x, seg_y_center), 
                    "right": (x + col_w, seg_y_center), 
                    "height": segment_height
                })

    for c in ["A", "B", "C"]:
        for idx, b in enumerate(blocks[c]):
            draw_card(c, idx, b["ids"], b.get("marker", ""), b.get("text", ""))

    def draw_ribbon_sigmoid(n_src, n_dst, color, alpha):
        x_src, y_src = n_src['right']
        x_dst, y_dst = n_dst['left']
        h_src = n_src.get('height', 0.5)
        h_dst = n_dst.get('height', 0.5)
        ribbon_h = min(h_src, h_dst) * ribbon_width_scale
        x = np.linspace(x_src, x_dst, 150)
        sigmoid = 1 / (1 + np.exp(-12 * (x - (x_src + x_dst) / 2) / (x_dst - x_src)))
        y = y_src + (y_dst - y_src) * sigmoid
        ax.fill_between(x, y - ribbon_h/2, y + ribbon_h/2, color=color, alpha=alpha, zorder=1, edgecolor=None)

    if show_links:
        for uid in sorted_ids:
            if uid in anchors["A"] and uid in anchors["B"]:
                for start in anchors["A"][uid]:
                    for end in anchors["B"][uid]: draw_ribbon_sigmoid(start, end, link_color, link_alpha)
            if uid in anchors["B"] and uid in anchors["C"]:
                for start in anchors["B"][uid]:
                    for end in anchors["C"][uid]: draw_ribbon_sigmoid(start, end, link_color, link_alpha)

    ax.set_xlim(x_min, x_max)
    ax.set_ylim(current_y - 0.5, 1.1)
    ax.axis("off")
    fig.tight_layout()
    return fig

# ==========================================
# ZAKŁADKI GŁÓWNE
# ==========================================
st.title("🗃️ Przybornik Badacza Źródeł")
tab1, tab2 = st.tabs(["📊 Losy Zakonnic", "📜 Porównywarka Psalmów i inne figle"])

# ==========================================
# ZMIENNE GLOBALNE DLA SIDEBARA
# ==========================================
# Wartości domyślne dla Zakonnic (zainicjalizowane tutaj, wypełnione w sidebar)
mappings = {}
defaults = {
    'Gravelines': 'yes, y, yesg, g, yellow, yesy', 'London': 'yesn, london',
    'Gosfield': 'yesz, gosfield', 'Scorton': 'yesc, s, scorton', 'Rouen': 'yesr',
    'Haggerston': 'yesh', 'Aire': 'yesa', 'Britwell': 'yesb', 'Plymouth': 'yesp',
    'Dunkirk': 'yesd', 'Worcester': 'yesw'
}
active_colors_selection = {}

# Wartości domyślne dla Psalmów
filter_input = ""
col_src1 = "#a6cee3"
col_txt1 = "#FFFFFF"
col_src2 = "#6BB72B"
col_txt2 = "#FFFFFF"
col_src3 = "#1f78b4"
col_txt3 = "#FFFFFF"
show_links = True
show_stripe = True
show_markers = True
show_ids = True
show_row_ids_left = True
show_zebra = True
link_color = "#2253BD"
link_opacity = 0.18
ribbon_scale = 0.88
font_size = 10
chars_per_line = 46
compact = False
EXPORT_DPI = 450

# ==========================================
# SIDEBAR - USTAWIENIA ZAKONNIC
# ==========================================
with st.sidebar:
    st.header("⚙️ Ustawienia: Zakonnice")
    
    st.subheader("Konfiguracja skrótów")
    st.info("Wpisz wartości z Excela przypisane do danego koloru (oddzielone przecinkami).")
    
    st.markdown("##### Konfiguracja kolorów i mapowania")
    for location, color_hex in COLORS_NUNS.items():
        if location == 'Deceased': continue
        is_active = st.checkbox(f"{location}", value=True, key=f"active_{location}")
        active_colors_selection[location] = is_active
        if is_active:
            st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;<span style='color:{color_hex}'>■</span> Wartości:", unsafe_allow_html=True)
            val = st.text_input(f"Wartości dla {location}", value=defaults.get(location,''), key=f"inp_{location}", label_visibility="collapsed")
            mappings[location] = [v.strip().lower() for v in val.split(',') if v.strip()]
        else:
            mappings[location] = []
    
    st.markdown("---")
    st.markdown("**Automatyczne:**")
    st.markdown("- `x` → Niepewne (Uncertain)")
    st.markdown("- `z` → Zmarłe (Deceased)")

# ==========================================
# SIDEBAR - USTAWIENIA PSALMÓW
# ==========================================
with st.sidebar:
    st.markdown("---")
    st.header("⚙️ Ustawienia: Psalmy")
    
    st.subheader("🔍 Filtr")
    filter_input = st.text_input("Wpisz ID wiersza (np. E, K):", value="", help="Filtrowanie wierszy zawierających dane ID.")
    
    st.markdown("---")
    st.subheader("Wygląd Wykresu")
    col_src1 = st.color_picker("Kolumna 1 (Officium)", "#a6cee3")
    col_txt1 = st.color_picker("Kolor tekstu zakładki - Kol. 1", "#FFFFFF", key="txt_col1")
    col_src2 = st.color_picker("Kolumna 2 (Vulgata)", "#6BB72B")
    col_txt2 = st.color_picker("Kolor tekstu zakładki - Kol. 2", "#FFFFFF", key="txt_col2")
    col_src3 = st.color_picker("Kolumna 3 (Bellarmine)", "#1f78b4")
    col_txt3 = st.color_picker("Kolor tekstu zakładki - Kol. 3", "#FFFFFF", key="txt_col3")
    
    st.subheader("Widoczność Elementów")
    show_links = st.checkbox("Pokaż Wstęgi (Połączenia)", value=True)
    show_stripe = st.checkbox("Pokaż Kolorowy Pasek", value=True, help="Wyświetla kolorowy pasek po lewej stronie karty.")
    show_markers = st.checkbox("Pokaż Markery (1, 2, V...)", value=True, help="Wyświetla cyfry arabskie lub rzymskie na kolorowym pasku.")
    show_ids = st.checkbox("Pokaż ID wierszy (A, B...)", value=True, help="Wyświetla litery A, B w lewym górnym rogu kafelka.")
    show_row_ids_left = st.checkbox("Pokaż ID Wierszy (poza wykresem - Lewa)", value=True, help="Wyświetla duże litery identyfikacyjne (np. K, M) po lewej stronie całego wykresu.")
    show_zebra = st.checkbox("Pokaż Tło Wierszy (Zebra)", value=True, help="Wyświetla naprzemienne szare tło dla wierszy.")
    
    st.subheader("Szczegóły Techniczne")
    link_color = st.color_picker("Kolor Wstęg", "#2253BD") 
    link_opacity = st.slider("Przezroczystość Wstęg", 0.1, 1.0, 0.18)
    ribbon_scale = st.slider("Szerokość Wstęgi (Skala)", 0.1, 1.0, 0.88, help="Skala szerokości wstęgi względem wysokości kafelka (1.0 = pełna wysokość).")
    font_size = st.slider("Rozmiar Czcionki", 6, 16, 10)
    chars_per_line = st.slider("Znaków w linii (Szerokość)", 20, 100, 46)
    compact = st.checkbox("Tryb Kompaktowy (Mniejsze odstępy)", value=False)

# ==========================================
# TAB 1: ZAKONNICE (ROZBUDOWANA WERSJA)
# ==========================================
with tab1:
    st.header("Generator Wykresów Losów Zakonnic")
    st.markdown("""
    Aplikacja pozwala wgrać plik Excel, wybrać arkusz, skonfigurować wygląd i wygenerować wykres.
    """)

    # --- Główna część ---
    uploaded_file_nuns = st.file_uploader("Wybierz plik Excel (.xlsx)", type=['xlsx'], key="upl_nuns", label_visibility="visible")

    if uploaded_file_nuns:
        try:
            xls = pd.ExcelFile(uploaded_file_nuns)
            sheet_names = xls.sheet_names
            
            st.markdown("---")
            col1, col2 = st.columns([1, 2])
            
            with col1:
                st.subheader("2. Wybór danych")
                selected_sheet = st.selectbox("Wybierz arkusz z danymi:", sheet_names, key="sheet_nuns")
            
            df = pd.read_excel(uploaded_file_nuns, sheet_name=selected_sheet)
            
            # Inteligentne wykrywanie kolumn danych
            data_columns = df.columns.tolist()
            start_idx = next((i for i, c in enumerate(data_columns) if "IMPRISONMENT" in str(c).upper()), 0)
            data_columns = data_columns[start_idx:]
            
            with st.expander(f"Podgląd danych: {selected_sheet}"):
                st.dataframe(df.head(10))
            
            st.markdown("---")
            st.subheader("3. Personalizacja wykresu")
            
            with st.expander("Opcje wyglądu i legendy", expanded=True):
                c1, c2 = st.columns(2)
                with c1:
                    chart_title = st.text_input("Tytuł wykresu", value=f"Population Status: {selected_sheet} Timeline")
                    show_values = st.checkbox("Pokaż liczby na paskach", value=True)
                    show_total = st.checkbox("Pokaż sumę (Total)", value=True)
                with c2:
                    show_legend = st.checkbox("Pokaż legendę", value=True)
                    legend_loc_pl = st.selectbox("Pozycja legendy", ["Prawy górny", "Prawy dolny", "Lewy górny", "Lewy dolny", "Prawy środek"], index=0)
                    # Mapowanie polskich nazw na angielskie dla matplotlib
                    legend_loc_map = {
                        "Prawy górny": "upper right",
                        "Prawy dolny": "lower right", 
                        "Lewy górny": "upper left",
                        "Lewy dolny": "lower left",
                        "Prawy środek": "center right"
                    }
                    legend_loc = legend_loc_map.get(legend_loc_pl, "upper right")
                
                st.markdown("##### Wybór etykiet (okresów)")
                st.info("Wybierz kolumny z Excela, które mają być wyświetlone jako etapy na wykresie.")
                selected_columns = st.multiselect(
                    "Wybierz etykiety (kolumny):",
                    options=data_columns,
                    default=data_columns
                )
                
                st.markdown("##### Własne nazwy w legendzie")
                st.caption("Zostaw puste, aby użyć domyślnych.")
                
                custom_labels = {}
                cols = st.columns(3)
                idx = 0
                priority_order = ['Gravelines', 'London', 'Gosfield', 'Scorton', 'Rouen', 'Haggerston', 'Aire', 'Britwell', 'Plymouth', 'Dunkirk', 'Worcester']
                visible_keys = [k for k in priority_order if active_colors_selection.get(k, False)] + ['Uncertain', 'Deceased']
                
                for key in visible_keys:
                    with cols[idx % 3]:
                        if key == 'Uncertain':
                            default_lbl = "Uncertain (x)"
                        elif key == 'Deceased':
                            default_lbl = "Deceased (z)"
                        else:
                            default_lbl = f"Alive ({key})"
                        
                        custom_labels[key] = st.text_input(f"Etykieta: {key}", placeholder=default_lbl, key=f"lbl_{key}")
                    idx += 1

            # Przycisk generowania
            if st.button("Generuj wykres", type="primary"):
                
                # Przetwarzanie danych
                segments_data = []
                
                for col_name in selected_columns:
                    series = df[col_name].astype(str).str.strip().str.lower()
                    uncertain = series[series.isin(['x'])].count()
                    deceased = series[series.isin(['z'])].count()
                    
                    # Określenie domyślnej kategorii dla kolumny
                    default_cat = 'Gravelines'
                    if "LONDON" in str(col_name).upper(): default_cat = 'London'
                    elif "GOSFIELD" in str(col_name).upper(): default_cat = 'Gosfield'
                    
                    counts = {}
                    for loc, vals in mappings.items():
                        spec_vals = [v for v in vals if v not in ['yes', 'y']]
                        counts[loc] = series[series.isin(spec_vals)].count()
                    
                    gen_count = series[series.isin(['yes', 'y'])].count()
                    counts[default_cat] = counts.get(default_cat, 0) + gen_count
                    
                    bar_segments = []
                    for loc in priority_order:
                        if active_colors_selection.get(loc) and counts.get(loc, 0) > 0:
                            bar_segments.append((counts[loc], COLORS_NUNS[loc], loc))
                    
                    if uncertain > 0:
                        base = COLORS_NUNS.get(default_cat, COLORS_NUNS['Gravelines'])
                        bar_segments.append((uncertain, make_lighter(base), "Uncertain"))
                    if deceased > 0:
                        bar_segments.append((deceased, COLORS_NUNS['Deceased'], "Deceased"))
                        
                    segments_data.append((str(col_name), bar_segments))
                
                # Rysowanie wykresu
                if segments_data:
                    fig, ax = plt.subplots(figsize=(16, 9))
                    y_labels = []
                    
                    for i, (lbl, segs) in enumerate(segments_data):
                        y_labels.append(lbl)
                        left = 0
                        for val, col, key in segs:
                            if val > 0:
                                hatch = '////' if key == "Uncertain" else None
                                ax.barh(i, val, left=left, color=col, edgecolor='black', height=0.6, hatch=hatch)
                                
                                # Kolor tekstu
                                txt_col = 'white' if col in [COLORS_NUNS['Deceased'], COLORS_NUNS['Scorton'], COLORS_NUNS['Rouen'], COLORS_NUNS['Plymouth'], COLORS_NUNS['Worcester']] else 'black'
                                
                                if show_values and val >= 0.8:
                                    center_x = left + val / 2
                                    ax.text(center_x, i, str(int(val)), ha='center', va='center', fontsize=10, fontweight='bold', color=txt_col)
                                
                                left += val
                        
                        if show_total:
                            total = sum([v for v,_,_ in segs])
                            ax.text(left + 0.5, i, f"Total: {total}", ha='left', va='center', fontsize=11, fontweight='bold')
                    
                    ax.set_yticks(range(len(y_labels)))
                    ax.set_yticklabels(y_labels, fontsize=10)
                    ax.invert_yaxis()
                    ax.set_xlabel("Liczba zakonnic", fontsize=12)
                    ax.set_title(chart_title, fontsize=16, pad=20)
                    
                    # Legenda
                    if show_legend:
                        def get_label(key):
                            user_lbl = custom_labels.get(key, "")
                            if user_lbl.strip(): return user_lbl
                            if key == 'Uncertain': return 'Uncertain (x)'
                            if key == 'Deceased': return 'Deceased (z)'
                            return f'Alive ({key})'
                        
                        patches_list = []
                        active_locs = [loc for loc in priority_order if mappings.get(loc) and active_colors_selection.get(loc)]
                        for loc in active_locs:
                            patches_list.append(mpatches.Patch(facecolor=COLORS_NUNS[loc], edgecolor='black', label=get_label(loc)))
                        
                        patches_list.append(mpatches.Patch(facecolor='lightgray', hatch='////', edgecolor='black', label=get_label('Uncertain')))
                        patches_list.append(mpatches.Patch(facecolor=COLORS_NUNS['Deceased'], edgecolor='black', label=get_label('Deceased')))
                        
                        ax.legend(handles=patches_list, loc=legend_loc, title="Legenda")
                    
                    ax.spines['right'].set_visible(False)
                    ax.spines['top'].set_visible(False)
                    ax.grid(axis='x', linestyle='--', alpha=0.5)
                    plt.tight_layout()
                    
                    # Wyświetlenie
                    st.pyplot(fig)
                    
                    # Pobieranie
                    img_buf = io.BytesIO()
                    fig.savefig(img_buf, format='png', dpi=300, bbox_inches='tight')
                    img_buf.seek(0)
                    
                    st.download_button(
                        label="💾 Pobierz wykres (PNG)",
                        data=img_buf,
                        file_name=f"wykres_{selected_sheet}.png",
                        mime="image/png"
                    )
                    
                    plt.close(fig)
                    
        except Exception as e:
            st.error(f"Wystąpił błąd podczas przetwarzania pliku: {e}")
    else:
        st.info("Proszę wgrać plik Excel, aby rozpocząć.")


# ==========================================
# TAB 2: PSALMY
# ==========================================
with tab2:
    st.header("Porównywarka Źródeł Łacińskich")
    st.markdown("""
    **Instrukcja:** Wgraj plik Word. Użyj ID `[M]` lub scalenia `[M,O]` do sterowania układem.
    """)

    # --- Upload i Przetwarzanie ---
    uploaded_docx = st.file_uploader("Wgraj plik Word (.docx)", type=['docx'], key="upl_docx_psalms_new")

    if uploaded_docx:
        st.info("Przetwarzam plik...")
        try:
            psalms_dict = parse_docx_psalms_v2(uploaded_docx.getvalue())
            if not psalms_dict:
                st.warning("Nie znaleziono danych w pliku.")
            else:
                st.success(f"Znaleziono psalmów: {len(psalms_dict)}")
                
                # --- HELPER PRZYGOTOWANIA DANYCH ---
                def prepare_view(selected_psalm, selected_ids=None, filter_id=None):
                    rows = psalms_dict[selected_psalm]
                    sorted_ids, blocks, id_to_index = build_blocks(rows)

                    if filter_id:
                        target = filter_id.strip()
                        if target:
                            expanded = expand_ids_by_merges({target}, blocks)
                            view_ids = [i for i in sorted_ids if i in expanded]
                            blocks_view = {c: [b for b in blocks[c] if set(b["ids"]) & expanded] for c in ["A", "B", "C"]}
                        else:
                            view_ids, blocks_view = sorted_ids, blocks
                    elif selected_ids:
                        expanded = expand_ids_by_merges(set(selected_ids), blocks)
                        view_ids = [i for i in sorted_ids if i in expanded]
                        blocks_view = {c: [b for b in blocks[c] if set(b["ids"]) & expanded] for c in ["A", "B", "C"]}
                    else:
                        view_ids, blocks_view = sorted_ids, blocks

                    id_to_index_view = {uid: i for i, uid in enumerate(view_ids)}
                    return view_ids, blocks_view, id_to_index_view

                # --- 3 TRYBY GENEROWANIA ---
                if filter_input:
                    st.info(f"Tryb filtrowania aktywny dla ID: '{filter_input}'")
                    # Defaults for filter mode
                    col_label_1 = "Officium 1571"
                    col_label_2 = "Vulgata 1592"
                    col_label_3 = "Bellarmine 1611"
                    custom_title_override = ""

                    # Szukamy we wszystkich psalmach
                    for p_name in psalms_dict.keys():
                        view_ids, blocks_view, id_to_index_view = prepare_view(p_name, filter_id=filter_input)
                        if view_ids:
                            st.markdown(f"### {p_name} (Filtr: {filter_input})")
                            final_title = custom_title_override if custom_title_override else f"{p_name} (Filtr: {filter_input})"
                            fig = draw_pretty_sankey_final(
                                title=final_title,
                                sorted_ids=view_ids,
                                blocks=blocks_view,
                                id_to_index=id_to_index_view,
                                colors=[col_src1, col_src2, col_src3],
                                labels=(col_label_1, col_label_2, col_label_3),
                                link_color=link_color,
                                link_alpha=link_opacity,
                                ribbon_width_scale=ribbon_scale,
                                font_size=font_size,
                                wrap_chars=chars_per_line,
                                compact=compact,
                                show_links=show_links,
                                show_stripe=show_stripe,
                                show_verse_nums=show_markers,
                                show_ids=show_ids,
                                show_row_ids_left=show_row_ids_left,
                                show_zebra=show_zebra,
                                badge_text_colors=(col_txt1, col_txt2, col_txt3)
                            )
                            st.pyplot(fig)
                            plt.close(fig)
                else:
                    st.markdown("### Wybierz tryb generowania")
                    mode = st.radio(
                        "Tryb:", 
                        ["Pojedynczy Podgląd", "Wybrane wiersze - Podgląd", "Eksport do ZIP"], 
                        horizontal=True
                    )

                    if mode == "Pojedynczy Podgląd":
                        selected_psalm = st.selectbox("Wybierz Psalm:", list(psalms_dict.keys()))
                        
                        with st.expander("📝 Etykiety i Teksty", expanded=False):
                            custom_title_override = st.text_input("Nadpisz Tytuł", value="", key="t_single")
                            c1, c2, c3 = st.columns(3)
                            col_label_1 = c1.text_input("Kolumna 1", "Officium 1571", key="l1_single")
                            col_label_2 = c2.text_input("Kolumna 2", "Vulgata 1592", key="l2_single")
                            col_label_3 = c3.text_input("Kolumna 3", "Bellarmine 1611", key="l3_single")

                        if selected_psalm:
                            view_ids, blocks_view, id_to_index_view = prepare_view(selected_psalm)
                            final_title = custom_title_override if custom_title_override else selected_psalm
                            
                            fig = draw_pretty_sankey_final(
                                title=final_title,
                                sorted_ids=view_ids,
                                blocks=blocks_view,
                                id_to_index=id_to_index_view,
                                colors=[col_src1, col_src2, col_src3],
                                labels=(col_label_1, col_label_2, col_label_3),
                                link_color=link_color,
                                link_alpha=link_opacity,
                                ribbon_width_scale=ribbon_scale,
                                font_size=font_size,
                                wrap_chars=chars_per_line,
                                compact=compact,
                                show_links=show_links,
                                show_stripe=show_stripe,
                                show_verse_nums=show_markers,
                                show_ids=show_ids,
                                show_row_ids_left=show_row_ids_left,
                                show_zebra=show_zebra,
                                badge_text_colors=(col_txt1, col_txt2, col_txt3)
                            )
                            st.pyplot(fig)
                            img_buf = io.BytesIO()
                            fig.savefig(img_buf, format='png', dpi=EXPORT_DPI, bbox_inches='tight')
                            img_buf.seek(0)
                            st.download_button("Pobierz PNG", data=img_buf, file_name=f"{selected_psalm}.png", mime="image/png")

                    elif mode == "Wybrane wiersze - Podgląd":
                        selected_psalm_view = st.selectbox("Wybierz psalm:", list(psalms_dict.keys()))
                        rows = psalms_dict[selected_psalm_view]
                        sorted_ids_all, _, _ = build_blocks(rows)
                        selected_ids = st.multiselect("Wybierz wiersze do wyświetlenia:", sorted_ids_all, default=[])
                        
                        with st.expander("📝 Etykiety i Teksty", expanded=False):
                            custom_title_override = st.text_input("Nadpisz Tytuł", value="", key="t_custom")
                            c1, c2, c3 = st.columns(3)
                            col_label_1 = c1.text_input("Kolumna 1", "Officium 1571", key="l1_custom")
                            col_label_2 = c2.text_input("Kolumna 2", "Vulgata 1592", key="l2_custom")
                            col_label_3 = c3.text_input("Kolumna 3", "Bellarmine 1611", key="l3_custom")

                        if selected_psalm_view:
                            view_ids, blocks_view, id_to_index_view = prepare_view(selected_psalm_view, selected_ids=selected_ids)
                            suffix = f" (ID: {', '.join(selected_ids)})" if selected_ids else ""
                            final_title = custom_title_override if custom_title_override else f"{selected_psalm_view}{suffix}"
                            
                            fig = draw_pretty_sankey_final(
                                title=final_title,
                                sorted_ids=view_ids,
                                blocks=blocks_view,
                                id_to_index=id_to_index_view,
                                colors=[col_src1, col_src2, col_src3],
                                labels=(col_label_1, col_label_2, col_label_3),
                                link_color=link_color,
                                link_alpha=link_opacity,
                                ribbon_width_scale=ribbon_scale,
                                font_size=font_size,
                                wrap_chars=chars_per_line,
                                compact=compact,
                                show_links=show_links,
                                show_stripe=show_stripe,
                                show_verse_nums=show_markers,
                                show_ids=show_ids,
                                show_row_ids_left=show_row_ids_left,
                                show_zebra=show_zebra,
                                badge_text_colors=(col_txt1, col_txt2, col_txt3)
                            )
                            st.pyplot(fig)
                            
                            img_buf = io.BytesIO()
                            fig.savefig(img_buf, format='png', dpi=EXPORT_DPI, bbox_inches='tight')
                            img_buf.seek(0)
                            st.download_button("Pobierz PNG", data=img_buf, file_name=f"{selected_psalm_view}_custom.png", mime="image/png")
                            
                            plt.close(fig)

                    else:
                        st.markdown("### Eksport wykresów do archiwum ZIP")
                        selected_psalms_zip = st.multiselect("Wybierz psalmy do eksportu:", list(psalms_dict.keys()), default=list(psalms_dict.keys()))
                        
                        with st.expander("📝 Etykiety i tytuły", expanded=False):
                            custom_title_override = st.text_input("Nadpisz tytuł (zostaw puste dla domyślnego)", value="", key="t_zip")
                            c1, c2, c3 = st.columns(3)
                            col_label_1 = c1.text_input("Kolumna 1", "Officium 1571", key="l1_zip")
                            col_label_2 = c2.text_input("Kolumna 2", "Vulgata 1592", key="l2_zip")
                            col_label_3 = c3.text_input("Kolumna 3", "Bellarmine 1611", key="l3_zip")

                        # Sekcja wyboru legendy (tytuł + nagłówki kolumn)
                        with st.expander("🏷️ Legenda (tytuł i nagłówki kolumn)", expanded=True):
                            st.markdown("**Zaznacz, które wykresy mają mieć tytuł i nagłówki kolumn.**")
                            st.markdown("_Przydatne gdy wklejasz do Worda - np. tylko pierwszy wykres na stronie ma legendę._")
                            
                            legend_mode = st.radio(
                                "Tryb legendy:",
                                ["Wszystkie z legendą", "Tylko pierwszy wykres każdego psalmu", "Wybierz ręcznie"],
                                key="legend_mode_zip"
                            )
                            
                            # Wygeneruj listę wszystkich wykresów do wyboru
                            all_charts_info = []
                            for p_name in selected_psalms_zip:
                                rows = psalms_dict[p_name]
                                sorted_ids_temp, blocks_temp, _ = build_blocks(rows)
                                processed_temp = set()
                                chart_idx = 0
                                for uid in sorted_ids_temp:
                                    if uid not in processed_temp:
                                        target_set = expand_ids_by_merges([uid], blocks_temp)
                                        for v_id in target_set:
                                            processed_temp.add(v_id)
                                        ids_str = "-".join(sorted(target_set, key=lambda x: sorted_ids_temp.index(x) if x in sorted_ids_temp else 999))
                                        all_charts_info.append({
                                            "psalm": p_name,
                                            "ids": ids_str,
                                            "label": f"{p_name} ({ids_str})",
                                            "is_first": chart_idx == 0
                                        })
                                        chart_idx += 1
                            
                            # Określ które wykresy mają legendę
                            charts_with_legend = set()
                            if legend_mode == "Wszystkie z legendą":
                                charts_with_legend = {c["label"] for c in all_charts_info}
                            elif legend_mode == "Tylko pierwszy wykres każdego psalmu":
                                charts_with_legend = {c["label"] for c in all_charts_info if c["is_first"]}
                            else:
                                # Ręczny wybór
                                if all_charts_info:
                                    default_selection = [c["label"] for c in all_charts_info if c["is_first"]]
                                    selected_legend_charts = st.multiselect(
                                        "Wybierz wykresy z legendą:",
                                        [c["label"] for c in all_charts_info],
                                        default=default_selection,
                                        key="manual_legend_selection"
                                    )
                                    charts_with_legend = set(selected_legend_charts)

                        if st.button("Generuj archiwum ZIP"):
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            zip_buffer = io.BytesIO()
                            
                            # Najpierw policz wszystkie pliki do wygenerowania
                            total_files = len(all_charts_info)
                            
                            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zf:
                                current_file = 0
                                
                                for p_name in selected_psalms_zip:
                                    rows = psalms_dict[p_name]
                                    sorted_ids, blocks, _ = build_blocks(rows)
                                    processed_ids = set()
                                    
                                    for uid in sorted_ids:
                                        if uid in processed_ids:
                                            continue
                                            
                                        view_ids, blocks_view, id_to_index_view = prepare_view(p_name, selected_ids=[uid])
                                        for v_id in view_ids:
                                            processed_ids.add(v_id)
                                            
                                        ids_str = "-".join(view_ids)
                                        file_name = f"{p_name}_{ids_str}.png"
                                        chart_title = custom_title_override if custom_title_override else f"{p_name} (ID: {', '.join(view_ids)})"
                                        
                                        # Sprawdź czy ten wykres ma mieć legendę
                                        chart_label = f"{p_name} ({ids_str})"
                                        has_legend = chart_label in charts_with_legend
                                        
                                        current_file += 1
                                        progress_bar.progress(current_file / total_files)
                                        legend_info = "z legendą" if has_legend else "bez legendy"
                                        status_text.text(f"Generuję {current_file}/{total_files}: {file_name} ({legend_info})")
                                        
                                        fig = draw_pretty_sankey_final(
                                            title=chart_title,
                                            sorted_ids=view_ids,
                                            blocks=blocks_view,
                                            id_to_index=id_to_index_view,
                                            colors=[col_src1, col_src2, col_src3],
                                            labels=(col_label_1, col_label_2, col_label_3),
                                            link_color=link_color,
                                            link_alpha=link_opacity,
                                            ribbon_width_scale=ribbon_scale,
                                            font_size=font_size,
                                            wrap_chars=chars_per_line,
                                            compact=compact,
                                            show_links=show_links,
                                            show_stripe=show_stripe,
                                            show_verse_nums=show_markers,
                                            show_ids=show_ids,
                                            show_row_ids_left=show_row_ids_left,
                                            show_zebra=show_zebra,
                                            badge_text_colors=(col_txt1, col_txt2, col_txt3),
                                            show_header=has_legend
                                        )
                                        
                                        img_buffer = io.BytesIO()
                                        fig.savefig(img_buffer, format="png", dpi=EXPORT_DPI, bbox_inches='tight')
                                        plt.close(fig)
                                        zf.writestr(file_name, img_buffer.getvalue())
                            
                            status_text.text("")
                            charts_with_legend_count = sum(1 for c in all_charts_info if c["label"] in charts_with_legend)
                            st.success(f"Gotowe! Wygenerowano {total_files} plików ({charts_with_legend_count} z legendą).")
                            st.download_button("📦 Pobierz archiwum ZIP", data=zip_buffer.getvalue(), file_name="psalmy_wykresy.zip", mime="application/zip")

        except Exception as e:
            st.error(f"Błąd: {e}")