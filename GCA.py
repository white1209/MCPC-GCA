import pandas as pd
from datetime import datetime, timedelta

# Google Sheet CSV export link
sheet_url = "https://docs.google.com/spreadsheets/d/1k9o5KmPbBjcZ9byfvG-ki5Z7qWLVnqsPwqbt5oad9dg/export?format=csv&gid=24781654"

# Read sheet (header from row 7, data starts at row 8)
df = pd.read_csv(sheet_url, header=6, skiprows=[7])
df = df.dropna(how="all")  # Drop completely empty rows

# --- Normalize Place names using substring match ---
def normalize_place(place):
    if not isinstance(place, str):
        return place
    p = place.strip().lower()

    mapping = {
        "arc": "The Arc",
        "mv": "Mutiara Ville",
        "mutiara": "Mutiara Ville",
        "shaft": "Shaftsbury Cyberjaya",
        "edu": "Edusphere",
        "lake": "Lakepoint Residence",
        "mmu": "MMU Bus Stop",
        "hyve": "Hyve"
    }

    for key, value in mapping.items():
        if key in p:  # substring-based matching
            return value

    return place.strip()

if "Place" in df.columns:
    df["Place"] = df["Place"].astype(str).apply(normalize_place)

# --- Function to generate transport arrangement brief ---
def generate_transport_brief(df):
    today = datetime.today()
    days_until_sunday = (6 - today.weekday()) % 7  # Sunday = 6
    coming_sunday = today + timedelta(days=days_until_sunday)
    sunday_date = coming_sunday.strftime("%d %B %Y")

    message = f"Good evening all, this is the transport arrangement brief for Sunday Service for {sunday_date}.\n\n"

    # Count valid names only
    pax_count = df["Name"].dropna().astype(str).str.strip().ne("").sum()
    message += f"Pax: {pax_count}\n\n"

    # Define sections (must match column names in the sheet)
    sections = {
        "Departure Trip": "💒 Departure Trip",
        "After Service": "🏠 After Service",
        "After Youth Fellowship": "🏠 After Youth Fellowship",
        "Worship Enablers": "💒 Worship Enablers"
    }

    # Loop through each section
    for key, title in sections.items():
        if key not in df.columns:
            continue

        df_filtered = df[df[key] == 1]
        if not df_filtered.empty:
            message += f"{title}:\n\n"
            grouped = df_filtered.groupby("Place")["Name"].apply(list).to_dict()
            for place, names in grouped.items():
                message += f"{place}\n"
                for i, name in enumerate(names, 1):
                    message += f"{i}. {name}\n"
                message += "\n"

    if message.strip().endswith(f"Pax: {pax_count}"):
        message += "No transport arrangement available."

    return message.strip()

# --- Display the generated message ---
print("\n" + generate_transport_brief(df))