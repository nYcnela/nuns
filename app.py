import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import matplotlib.patches as mpatches
import io
import os

# ==========================================
# 1. KONFIGURACJA I STAŁE
# ==========================================
st.set_page_config(page_title="Generator Wykresów Zakonnic", layout="wide")

# Hasło pobierane ze zmiennej środowiskowej (domyślnie "b12345" jeśli nie ustawione)
APP_PASSWORD = os.environ.get("APP_PASSWORD", "b12345")

# ==========================================
# AUTORYZACJA HASŁEM
# ==========================================
def check_password():
    """Sprawdza czy użytkownik podał poprawne hasło."""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        st.title("🔐 Logowanie")
        password = st.text_input("Podaj hasło:", type="password")
        if st.button("Zaloguj"):
            if password == APP_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Błędne hasło!")
        st.stop()

check_password()

COLORS = {
    'Gravelines': '#FFD700',  # Żółty
    'London': '#B0C4DE',      # Jasnoniebieski
    'Gosfield': '#2ca02c',    # Zielony
    'Rouen': '#1f77b4',       # Ciemnoniebieski
    'Haggerston': '#ff7f0e',  # Pomarańczowy
    'Scorton': '#d62728',     # Czerwony
    'Aire': '#e377c2',        # Różowy
    'Britwell': '#eaffea',    # Jasnozielony
    'Plymouth': '#9467bd',    # Fioletowy
    'Dunkirk': '#9ACD32',     # Oliwkowy
    'Worcester': '#17becf',   # Cyjan
    'Deceased': '#808080'     # Szary (Stały)
}

LABELS = [
    "IMPRISONMENT\n1793–1795",
    "LONDON\n1795–1796",
    "GOSFIELD\n1796–1813",
    "MOVE TO GRAVELINES\n1814",
    "STAY AT GRAVELINES\n1814–1825",
    "STAY AT GRAVELINES\n1826–1832",
    "GRAVELINES 1833",
    "GRAVELINES 1834–1837",
    "TRANSFER OF THE GRAVELINES\nHOUSE 1838"
]

def make_lighter(hex_color, alpha=0.3):
    return mcolors.to_rgba(hex_color, alpha=alpha)

def add_value_label(ax, value, left, width, y_pos, text_color='black'):
    if value > 0 and width >= 0.8:
        center_x = left + width / 2
        ax.text(center_x, y_pos, str(int(value)), ha='center', va='center',
                fontsize=10, fontweight='bold', color=text_color)

# ==========================================
# 2. INTERFEJS UŻYTKOWNIKA
# ==========================================

st.title("📊 Generator Wykresów Losów Zakonnic")
st.markdown("""
Aplikacja pozwala wgrać plik Excel, wybrać arkusz, skonfigurować wygląd i wygenerować wykres.
""")

# --- Sidebar: Konfiguracja Mapowania ---
st.sidebar.header("1. Konfiguracja Skrótów")
st.sidebar.info("Wpisz wartości z Excela, które mają być przypisane do danego koloru (oddzielone przecinkami).")

mappings = {}
defaults = {
    'Gravelines': 'yes, y, yesg, g, yellow, yesy',
    'London': 'yesn, london',
    'Gosfield': 'yesz, gosfield',
    'Scorton': 'yesc, s, scorton',
    'Rouen': 'yesr',
    'Haggerston': 'yesh',
    'Aire': 'yesa',
    'Britwell': 'yesb',
    'Plymouth': 'yesp',
    'Dunkirk': 'yesd',
    'Worcester': 'yesw'
}

# Selection for active colors
st.sidebar.subheader("Wybierz aktywne kolory")
active_colors_selection = {}
for location, color_hex in COLORS.items():
    if location == 'Deceased':
        continue
    
    # Checkbox to enable/disable color
    is_active = st.sidebar.checkbox(f"{location}", value=True, key=f"active_{location}")
    active_colors_selection[location] = is_active
    
    if is_active:
        default_val = defaults.get(location, '')
        # Color preview block
        st.sidebar.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;<span style='color:{color_hex}'>■</span> Wartości:", unsafe_allow_html=True)
        user_input = st.sidebar.text_input(f"Wartości dla {location}", value=default_val, key=f"input_{location}", label_visibility="collapsed")
        
        values = [v.strip().lower() for v in user_input.split(',') if v.strip()]
        mappings[location] = values
    else:
        mappings[location] = [] # Empty mapping if disabled

st.sidebar.markdown("---")
st.sidebar.markdown("**Automatyczne:**")
st.sidebar.markdown("- `x` -> Uncertain (Niepewne)")
st.sidebar.markdown("- `z` -> Deceased (Zmarłe)")

# --- Główna część ---

uploaded_file = st.file_uploader("Wybierz plik Excel (.xlsx)", type=['xlsx'])

if uploaded_file:
    try:
        # Load ExcelFile to get sheet names
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        st.write("---")
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.subheader("2. Wybór Danych")
            selected_sheet = st.selectbox("Wybierz arkusz z danymi:", sheet_names)
        
        # Load data from selected sheet
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        
        # Smart detection of data columns
        data_columns = []
        start_col_index = 0
        found_start = False
        
        # Try to find the start column
        for idx, col in enumerate(df.columns):
            if "IMPRISONMENT" in str(col).upper():
                start_col_index = idx
                found_start = True
                break
        
        if found_start:
            # Suggest columns from start_col_index onwards
            data_columns = df.columns[start_col_index:].tolist()
        else:
            # Fallback: use all columns
            data_columns = df.columns.tolist()
        
        with st.expander(f"Podgląd danych: {selected_sheet}"):
            st.dataframe(df.head(10))

        st.write("---")
        st.subheader("3. Personalizacja Wykresu")
        
        with st.expander("Opcje Wyglądu i Legendy", expanded=True):
            c1, c2 = st.columns(2)
            with c1:
                chart_title = st.text_input("Tytuł wykresu", value=f"Population Status: {selected_sheet} Timeline")
                show_values = st.checkbox("Pokaż liczby na paskach", value=True)
                show_total = st.checkbox("Pokaż sumę (Total)", value=True)
            with c2:
                show_legend = st.checkbox("Pokaż legendę", value=True)
                legend_loc = st.selectbox("Pozycja legendy", ["upper right", "lower right", "upper left", "lower left", "center right"], index=0)
            
            st.markdown("##### Wybór Etykiet (Okresów)")
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
            # Create inputs for all locations + Uncertain + Deceased
            all_keys = list(COLORS.keys()) + ['Uncertain']
            
            # Filter keys based on active selection
            visible_keys = [k for k in all_keys if k in ['Uncertain', 'Deceased'] or active_colors_selection.get(k, False)]
            
            for key in visible_keys:
                with cols[idx % 3]:
                    default_lbl = f"Alive ({key})" if key not in ['Uncertain', 'Deceased'] else key
                    if key == 'Uncertain': default_lbl = "Uncertain (x)"
                    if key == 'Deceased': default_lbl = "Deceased (z)"
                    
                    custom_labels[key] = st.text_input(f"Etykieta: {key}", placeholder=default_lbl)
                idx += 1

        # Generate Button
        if st.button("Generuj Wykres", type="primary"):
            
            # ==========================================
            # 3. PRZETWARZANIE DANYCH
            # ==========================================
            
            segments_data = []
            
            # Iteracja po wybranych kolumnach (etykietach)
            for col_name in selected_columns:
                
                # Używamy nazwy kolumny jako etykiety
                label = str(col_name)
                
                # Pobieramy dane z kolumny
                series = df[col_name].astype(str).str.strip().str.lower()

                # Zliczanie Uncertain i Deceased (stałe)
                uncertain = series[series.isin(['x'])].count()
                deceased = series[series.isin(['z'])].count()
                
                bar_segments = []
                
                # Określenie domyślnej kategorii dla danej kolumny
                default_category = 'Gravelines' # Domyślnie
                if "LONDON" in label.upper():
                    default_category = 'London'
                elif "GOSFIELD" in label.upper():
                    default_category = 'Gosfield'
                # Dla TRANSFER 1838 zostawiamy Gravelines jako domyślny
                
                # Zliczanie dla każdej lokalizacji na podstawie mapowania
                counts = {}
                
                # 1. Zliczanie jawnych wystąpień z mapowania
                for loc, values_list in mappings.items():
                    specific_values = [v for v in values_list if v not in ['yes', 'y']]
                    count = series[series.isin(specific_values)].count()
                    counts[loc] = count

                # 2. Obsługa generycznego 'yes'/'y'
                generic_count = series[series.isin(['yes', 'y'])].count()
                
                # Dodajemy generyczne do domyślnej kategorii tego etapu
                if default_category in counts:
                    counts[default_category] += generic_count
                else:
                    counts[default_category] = generic_count

                # Budowanie segmentów paska
                priority_order = ['Gravelines', 'London', 'Gosfield', 'Scorton', 'Rouen', 'Haggerston', 'Aire', 'Britwell', 'Plymouth', 'Dunkirk', 'Worcester']
                
                for loc in priority_order:
                    # Skip if color is disabled
                    if not active_colors_selection.get(loc, False):
                        continue
                        
                    if loc in counts and counts[loc] > 0:
                        bar_segments.append((counts[loc], COLORS[loc], loc)) # Pass loc key for label lookup

                # Dodajemy Uncertain
                if uncertain > 0:
                    base_color = COLORS.get(default_category, COLORS['Gravelines'])
                    bar_segments.append((uncertain, make_lighter(base_color), "Uncertain"))
                
                # Dodajemy Deceased
                if deceased > 0:
                    bar_segments.append((deceased, COLORS['Deceased'], "Deceased"))

                segments_data.append((label, bar_segments))

            # ==========================================
            # 4. RYSOWANIE WYKRESU
            # ==========================================
            if segments_data:
                fig, ax = plt.subplots(figsize=(16, 9))
                y_labels = []

                for i, (label, segments) in enumerate(segments_data):
                    y_labels.append(label)
                    left = 0

                    for value, color, key_name in segments:
                        if value > 0:
                            hatch = '////' if key_name == 'Uncertain' else None
                            edge_color = 'black'

                            ax.barh(i, value, left=left, color=color, edgecolor=edge_color, height=0.6, hatch=hatch)

                            # Kolor tekstu
                            txt_col = 'white' if color in [COLORS['Deceased'], COLORS['Scorton'], COLORS['Rouen'], COLORS['Plymouth'], COLORS['Worcester']] else 'black'
                            
                            if show_values:
                                add_value_label(ax, value, left, value, i, txt_col)
                            
                            left += value

                    # Suma
                    if show_total:
                        total = sum([v for v, _, _ in segments])
                        ax.text(left + 0.5, i, f"Total: {total}", ha='left', va='center', fontsize=11, fontweight='bold')

                ax.set_yticks(range(len(y_labels)))
                ax.set_yticklabels(y_labels, fontsize=10)
                ax.invert_yaxis()
                ax.set_xlabel("Number of Nuns", fontsize=12)
                ax.set_title(chart_title, fontsize=16, pad=20)

                # Legenda
                if show_legend:
                    legend_patches = []
                    
                    # Helper to get label
                    def get_label(key):
                        user_lbl = custom_labels.get(key, "")
                        if user_lbl.strip(): return user_lbl
                        # Defaults
                        if key == 'Uncertain': return 'Uncertain (x)'
                        if key == 'Deceased': return 'Deceased (z)'
                        return f'Alive ({key})'

                    # Add active locations
                    active_locs = [loc for loc in priority_order if mappings[loc]]
                    for loc in active_locs:
                         legend_patches.append(mpatches.Patch(facecolor=COLORS[loc], edgecolor='black', label=get_label(loc)))
                    
                    legend_patches.append(mpatches.Patch(facecolor='lightgray', hatch='////', edgecolor='black', label=get_label('Uncertain')))
                    legend_patches.append(mpatches.Patch(facecolor=COLORS['Deceased'], edgecolor='black', label=get_label('Deceased')))

                    ax.legend(handles=legend_patches, loc=legend_loc, title="Legend")
                
                ax.spines['right'].set_visible(False)
                ax.spines['top'].set_visible(False)
                ax.grid(axis='x', linestyle='--', alpha=0.5)
                plt.tight_layout()

                # Wyświetlenie w Streamlit
                st.pyplot(fig)

                # Pobieranie
                img = io.BytesIO()
                plt.savefig(img, format='png', dpi=300)
                img.seek(0)
                
                st.download_button(
                    label="💾 Pobierz Wykres (PNG)",
                    data=img,
                    file_name=f"wykres_{selected_sheet}.png",
                    mime="image/png"
                )

    except Exception as e:
        st.error(f"Wystąpił błąd podczas przetwarzania pliku: {e}")
else:
    st.info("Proszę wgrać plik Excel, aby rozpocząć.")
