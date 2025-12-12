import matplotlib.pyplot as plt
import matplotlib.patches as patches

# Konfiguracja ogólna
COMMON_X_LIMITS = (1790, 1860)
FIG_SIZE = (15, 3)

# Definicja kolorów
COLORS = {
    'Gravelines': '#FFD700',
    'London': '#B0C4DE',
    'Gosfield': '#2ca02c',
    'Rouen': '#1f77b4',
    'Haggerston': '#ff7f0e',
    'Scorton': '#d62728',
    'Aire': '#e377c2',
    'Britwell': '#eaffea',
    'Plymouth': '#9467bd',
    'Dunkirk': '#9ACD32',
    'Worcester': '#17becf'
}

# Dane
timelines_data = [
    {
        "title": "GRAVELINES → LONDON → GOSFIELD → GRAVELINES (1793–1838)",
        "events": [
            ("Gravelines", 1793, 1795, "Gravelines"),
            ("London", 1795, 1796, "London"),
            ("Gosfield", 1796, 1813, "Gosfield"),
            ("Gravelines", 1813, 1838, "Gravelines")
        ]
    },
    {
        "title": "ROUEN → LONDON → HAGGERSTON → SCORTON (1793–1857)",
        "events": [
            ("Rouen", 1793, 1795, "Rouen"),
            ("London", 1795, 1796, "London"),
            ("Haggerston", 1796, 1807, "Haggerston"),
            ("Scorton", 1807, 1857, "Scorton")
        ]
    },
    {
        "title": "AIRE → BRITWELL → PLYMOUTH → SCORTON (1798–1857)",
        "events": [
            ("Aire", 1798, 1799, "Aire"),
            ("London", 1799, 1799, "London"),
            ("Britwell", 1799, 1813, "Britwell"),
            ("Plymouth", 1813, 1834, "Plymouth"),
            ("Scorton", 1834, 1857, "Scorton")
        ]
    },
    {
        "title": "DUNKIRK → GRAVELINES → WORCESTER → SCORTON (1793–1857)",
        "events": [
            ("Dunkirk", 1793, 1793, "Dunkirk"),
            ("Gravelines", 1793, 1795, "Gravelines"),
            ("Worcester", 1795, '1807/1808', "Worcester"),
            ("Scorton", '1807/1808', 1857, "Scorton")
        ]
    }
]


def parse_year(val):
    """Zamienia daty '1807/1808' na float, resztę zostawia."""
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str) and '/' in val:
        parts = val.split('/')
        return (float(parts[0]) + float(parts[1])) / 2
    return float(val)


def create_timeline(data, filename_suffix):
    fig, ax = plt.subplots(figsize=FIG_SIZE)

    ax.set_xlim(COMMON_X_LIMITS)
    ax.set_ylim(0, 1)

    # Zmienna do śledzenia, gdzie skończył się poprzedni pasek (wizualnie)
    current_visual_cursor = COMMON_X_LIMITS[0]

    for i, (place, start_raw, end_raw, color_key) in enumerate(data["events"]):
        num_start = parse_year(start_raw)
        num_end = parse_year(end_raw)

        # Obliczanie szerokości matematycznej
        math_width = num_end - num_start

        # Ustalanie wizualnego początku
        # Pasek zaczyna się tam, gdzie matematycznie powinien, CHYBA ŻE
        # poprzedni pasek skończył się "dalej" (bo był sztucznie poszerzony).
        visual_start = max(num_start, current_visual_cursor)

        # Ustalanie wizualnego końca
        # Minimalna szerokość paska to 0.6 roku, żeby był widoczny
        visual_width = max(math_width, 0.6)
        visual_end = visual_start + visual_width

        # Aktualizacja kursora dla następnego elementu
        current_visual_cursor = visual_end

        # Rysowanie prostokąta
        rect = patches.Rectangle((visual_start, 0), visual_width, 1,
                                 linewidth=1, edgecolor='black',
                                 facecolor=COLORS.get(color_key, '#cccccc'))
        ax.add_patch(rect)

        # Rysowanie tekstu
        # Środek tekstu obliczamy na podstawie WIZUALNEJ szerokości,
        # dzięki temu będzie idealnie na środku widocznego paska.
        center_x = visual_start + (visual_width / 2)
        center_y = 0.5

        if start_raw == end_raw:
            label_text = f"{place}\n{start_raw}"
        else:
            label_text = f"{place}\n{start_raw}–{end_raw}"

        # Decyzja o obrocie tekstu
        if visual_width < 4:
            ax.text(center_x, center_y, place, rotation=90, ha='center', va='center',
                    fontsize=10, color='black')
        else:
            ax.text(center_x, center_y, label_text, ha='center', va='center',
                    fontsize=11, color='black')

    # Formatowanie osi (bez zmian)
    ax.set_yticks([])
    ax.set_xlabel("Years", fontsize=12)
    ax.set_title(data["title"], fontsize=16, pad=15)

    ax.spines['left'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)

    plt.xticks(range(COMMON_X_LIMITS[0], COMMON_X_LIMITS[1] + 1, 10), fontsize=11)
    plt.tight_layout()
    plt.show()


# Generowanie wykresów
for i, timeline in enumerate(timelines_data):
    create_timeline(timeline, i + 1)