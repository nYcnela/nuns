import matplotlib.pyplot as plt
import pandas as pd
import matplotlib.colors as mcolors
import matplotlib.patches as mpatches

# ==========================================
# 1. KONFIGURACJA KOLORÓW
# ==========================================
FIG_SIZE = (16, 9)

COLORS = {
    'Gravelines': '#FFD700',  # Żółty
    'London': '#B0C4DE',  # Jasnoniebieski
    'Gosfield': '#2ca02c',  # Zielony
    'Rouen': '#1f77b4',  # Ciemnoniebieski
    'Scorton': '#d62728',  # Czerwony
    'Deceased': '#808080'  # Szary
}


def make_lighter(hex_color, alpha=0.3):
    return mcolors.to_rgba(hex_color, alpha=alpha)


def add_value_label(ax, value, left, width, y_pos, text_color='black'):
    """Dodaje liczbę na środku paska, jeśli jest miejsce"""
    if value > 0:
        center_x = left + width / 2
        # Piszemy tylko jeśli pasek jest wystarczająco szeroki
        if width >= 0.8:
            ax.text(center_x, y_pos, str(int(value)), ha='center', va='center',
                    fontsize=10, fontweight='bold', color=text_color)


# ==========================================
# 2. PRZETWARZANIE DANYCH Z PLIKU CSV
# ==========================================

# Definicja etykiet czasowych (muszą odpowiadać kolejności kolumn w Excelu)
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


def load_and_process_data(filename, sheet_name=0):
    try:
        df = pd.read_excel(filename, sheet_name=sheet_name)
        
        # Znajdź indeks kolumny startowej (IMPRISONMENT)
        start_col_index = 0
        for idx, col in enumerate(df.columns):
            if "IMPRISONMENT" in str(col).upper():
                start_col_index = idx
                break
        
        print(f"Dane zaczynają się od kolumny indeks: {start_col_index} ('{df.columns[start_col_index]}')")

        segments_data = []

        for i, label in enumerate(LABELS):
            # Obliczamy właściwy indeks kolumny
            current_col_idx = start_col_index + i
            
            if current_col_idx >= len(df.columns):
                print(f"Brak kolumny dla etykiety: {label}")
                break

            col_name = df.columns[current_col_idx]
            # print(f"Przetwarzanie: {label} -> Kolumna: {col_name}")

            # Pobieramy dane, czyścimy spacje i zamieniamy na małe litery
            series = df[col_name].astype(str).str.strip().str.lower()

            # ZLICZANIE WARTOŚCI WSPÓLNYCH
            uncertain = series[series.isin(['x'])].count()
            deceased = series[series.isin(['z'])].count()
            
            # 1. Determine default category for this column
            default_category = 'Gravelines' # Default
            if "LONDON" in label.upper():
                default_category = 'London'
            elif "GOSFIELD" in label.upper():
                default_category = 'Gosfield'
            # elif "TRANSFER" in label.upper():
            #     default_category = 'Scorton'  <-- ZMIANA: Domyślnie Gravelines (Żółty), żeby YES było żółte, a YESc czerwone
            
            # 2. Count specific values (Explicit overrides)
            # Normalize: lower case
            
            # Counts
            count_gravelines = series[series.isin(['yesg', 'g', 'yellow'])].count() # Explicit Yellow
            count_london = series[series.isin(['yesn'])].count()     # Explicit Blue
            count_gosfield = series[series.isin(['yesz'])].count()   # Explicit Green
            count_scorton = series[series.isin(['yesc', 's'])].count()    # Explicit Red
            
            # Generic YES/Y
            count_generic = series[series.isin(['yes', 'y'])].count()
            
            # Add generic to default
            if default_category == 'Gravelines':
                count_gravelines += count_generic
            elif default_category == 'London':
                count_london += count_generic
            elif default_category == 'Gosfield':
                count_gosfield += count_generic
            elif default_category == 'Scorton':
                count_scorton += count_generic
            
            # Build segments
            # Order: Gravelines, London, Gosfield, Scorton, Uncertain, Deceased
            bar_segments = []
            
            if count_gravelines > 0:
                bar_segments.append((count_gravelines, COLORS['Gravelines'], "Alive Gravelines"))
            if count_london > 0:
                bar_segments.append((count_london, COLORS['London'], "Alive London"))
            if count_gosfield > 0:
                bar_segments.append((count_gosfield, COLORS['Gosfield'], "Alive Gosfield"))
            if count_scorton > 0:
                bar_segments.append((count_scorton, COLORS['Scorton'], "Alive Scorton"))

            # Dodajemy wspólne segmenty (Uncertain, Deceased)
            if uncertain > 0:
                # Kolor kreskowania zależy od głównego koloru etapu
                base_color = COLORS[default_category]
                bar_segments.append((uncertain, make_lighter(base_color), "Uncertain"))
            
            if deceased > 0:
                bar_segments.append((deceased, COLORS['Deceased'], "Deceased"))

            segments_data.append((label, bar_segments))

        return segments_data

    except Exception as e:
        print(f"Błąd przetwarzania pliku: {e}")
        return []


# Wczytanie danych (Podmień nazwę pliku jeśli uruchamiasz lokalnie i jest inna)
# Tutaj używam nazwy pliku, który przesłałeś
file_path = 'ZESTAWIENIE ZAKONNIC_NIEW WYKRES.xlsx'
# Podaj nazwę arkusza (sheet_name) - może być indeks (0, 1, 2...) lub nazwa np. "GRAVELINES"
timeline_data = load_and_process_data(file_path, sheet_name="GRAVELINES")


# ==========================================
# 3. RYSOWANIE WYKRESU
# ==========================================

def create_chart(data):
    if not data:
        print("Brak danych do wyświetlenia.")
        return

    fig, ax = plt.subplots(figsize=FIG_SIZE)
    y_labels = []

    for i, (label, segments) in enumerate(data):
        y_labels.append(label)
        left = 0

        for value, color, desc in segments:
            if value > 0:
                # Kreskowanie tylko dla "Uncertain"
                hatch = '////' if 'Uncertain' in desc else None
                edge_color = 'black'

                # Rysujemy pasek
                ax.barh(i, value, left=left, color=color, edgecolor=edge_color, height=0.6, hatch=hatch)

                # Ustalanie koloru tekstu (biały dla ciemnych teł)
                txt_col = 'white' if color in [COLORS['Deceased'], COLORS['Scorton'], COLORS['Rouen']] else 'black'

                add_value_label(ax, value, left, value, i, txt_col)
                left += value

        # Suma całkowita na końcu
        total = sum([v for v, _, _ in segments])
        ax.text(left + 0.5, i, f"Total: {total}", ha='left', va='center', fontsize=11, fontweight='bold')

    # Formatowanie osi
    ax.set_yticks(range(len(y_labels)))
    ax.set_yticklabels(y_labels, fontsize=10)
    ax.invert_yaxis()  # Najstarsze na górze
    ax.set_xlabel("Number of Nuns", fontsize=12)
    ax.set_title("Population Status: GRAVELINES Timeline (Based on CSV Data)", fontsize=16, pad=20)

    # LEGENDA
    legend_patches = [
        mpatches.Patch(facecolor=COLORS['Gravelines'], edgecolor='black', label='Alive (Gravelines)'),
        mpatches.Patch(facecolor=COLORS['London'], edgecolor='black', label='Alive (London)'),
        mpatches.Patch(facecolor=COLORS['Gosfield'], edgecolor='black', label='Alive (Gosfield)'),
        mpatches.Patch(facecolor=COLORS['Scorton'], edgecolor='black', label='Alive (Scorton)'),
        mpatches.Patch(facecolor='lightgray', hatch='////', edgecolor='black', label='Uncertain (x)'),
        mpatches.Patch(facecolor=COLORS['Deceased'], edgecolor='black', label='Deceased (z)')
    ]
    ax.legend(handles=legend_patches, loc='upper right', title="Legend")

    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.grid(axis='x', linestyle='--', alpha=0.5)

    plt.tight_layout()
    plt.show()


# Uruchomienie
create_chart(timeline_data)