import io
import sys
import urllib.parse
import pandas as pd
import openrouteservice
import streamlit as st # type: ignore
from itertools import permutations
from datetime import datetime, timedelta

# --- ORS API Key ---
ORS_API_KEY = "eyJvcmciOiI1YjNjZTM1OTc4NTExMTAwMDFjZjYyNDgiLCJpZCI6IjViMjBkMzNjZmJlYjQ4OTg4OWE4ZjYzYjQzMTQ4MDMxIiwiaCI6Im11cm11cjY0In0="
client = openrouteservice.Client(key=ORS_API_KEY)

# --- Streamlit App ---
st.set_page_config(page_title="MCPC GCA", layout="centered")
st.title("🚐 MCPC Gospel Car Arrangement")

if st.button("🚀 Generate Arrangement"):
    buffer = io.StringIO()
    sys.stdout = buffer

    # --- Google Sheet ---
    sheet_url = "https://docs.google.com/spreadsheets/d/1k9o5KmPbBjcZ9byfvG-ki5Z7qWLVnqsPwqbt5oad9dg/export?format=csv&gid=24781654"
    df = pd.read_csv(sheet_url, header=6, skiprows=[7])
    df = df.dropna(how="all")

    # --- Normalize Place Names ---
    def normalize_place(place):
        if not isinstance(place, str):
            return place
        p = place.strip().lower()
        mapping = {
            "arc": "The Arc",
            "mv": "Mutiara Ville",
            "mutiara": "Mutiara Ville",
            "shaftsbury putrajaya": "Shaftsbury Putrajaya",
            "shaft": "Shaftsbury Cyberjaya",
            "edu": "Edusphere",
            "lake": "Lakepoint Residence",
            "mmu": "MMU Bus Stop",
            "serin": "Serin Residency"
        }
        for key, value in mapping.items():
            if key in p:
                return value
        return place.strip().title()

    df["Place"] = df["Place"].apply(normalize_place)

    # --- Coordinates ---
    coords = {
        "MCPC": (2.9170225107127488, 101.6498633796812),
        "The Arc": (2.9257629643936376, 101.63683861036628),
        "Mutiara Ville": (2.922350609640224, 101.6350686085168),
        "Edusphere": (2.9321189224611715, 101.6376606680376),
        "Hyve": (2.92084108875226, 101.6610653950237),
        "Lakepoint Residence": (2.9289648854261663, 101.63512724947454),
        "Shaftsbury Cyberjaya": (2.9244692894170193, 101.65755849840291),
        "MMU Bus Stop": (2.924853141325742, 101.6409283450342),
        "Serin Residency": (2.916432495889349, 101.6457637950237)
    }

    locs = [[v[1], v[0]] for v in coords.values()]
    names = list(coords.keys())

    # --- Distance Matrix ---
    matrix = client.distance_matrix(
        locations=locs,
        profile='driving-car',
        metrics=['distance'],
        units='km'
    )
    distances = matrix['distances']

    def total_distance(route):
        return sum(distances[names.index(route[i])][names.index(route[i+1])] for i in range(len(route)-1))

    # --- Find Best Route ---
    start = "MCPC"
    places = [p for p in names if p != start]
    best_route, best_dist = None, float("inf")
    for perm in permutations(places):
        route = [start] + list(perm) + [start]
        dist = total_distance(route)
        if dist < best_dist:
            best_dist = dist
            best_route = route

    # --- ETA ---
    def estimate_travel_time(distance_km):
        if distance_km < 2:
            return 5
        elif distance_km < 4:
            return 10
        elif distance_km < 6:
            return 15
        elif distance_km < 8:
            return 20
        else:
            return 25

    # --- Venue Grouping ---
    df["Place_Normalize"] = df["Place"].str.lower()
    grouped = df.groupby("Place_Normalize")["Name"].apply(list).to_dict()
    pickup_venues = [v for v in best_route if v.lower() in grouped]
    venue_counts = [(v, len(grouped[v.lower()])) for v in pickup_venues]

    # --- Car Capacity Grouping ---
    CAR_CAPACITY = 6
    trips, current_trip, current_total = [], [], 0
    for venue, count in venue_counts:
        if current_total + count > CAR_CAPACITY and current_trip:
            trips.append(current_trip)
            current_trip = []
            current_total = 0
        current_trip.append((venue, count))
        current_total += count
    if current_trip:
        trips.append(current_trip)

    # --- Arrangement Function ---
    def generate_transport_brief(df):
        today = datetime.today()
        days_until_sunday = (6 - today.weekday()) % 7
        coming_sunday = today + timedelta(days=days_until_sunday)
        sunday_date = coming_sunday.strftime("%d %B %Y")
        pax_count = df["Name"].dropna().astype(str).str.strip().ne("").sum()
        print(f"Hi everyone, this is the transport arrangement brief for Sunday Service on {sunday_date}.\n")
        print(f"Vehicle: Alza VJY3510 \nPax: {pax_count}\n")

    def generate_worship_enablers_trip(df):
        worship_df = df[df["Worship Enablers"].notna()]
        worship_group = worship_df.groupby("Place_Normalize")["Name"].apply(list).to_dict() if not worship_df.empty else {}
        if worship_group:
            print(f"💒 Worship Enablers Trip\n")
        total_travel_minutes = 0
        worship_order = list(worship_group.keys())
        for venue in worship_order:
            if venue in coords:
                d = distances[names.index("MCPC")][names.index(normalize_place(venue))]
                total_travel_minutes += estimate_travel_time(d)
        total_travel_minutes += 10
        arrival_time = datetime.strptime("09:00", "%H:%M")
        depart_time = arrival_time - timedelta(minutes=total_travel_minutes)
        current_time = depart_time
        for venue in worship_order:
            if venue in coords:
                d = distances[names.index("MCPC")][names.index(normalize_place(venue))]
                travel_minutes = estimate_travel_time(d)
                current_time += timedelta(minutes=travel_minutes)
            print(f"{normalize_place(venue)} — ETA: {current_time.strftime('%I:%M %p')}")
            for n in worship_group[venue]:
                print(f"   - {n}")
            print()

    def generate_departure_trip():
        start_time = datetime.strptime("09:15", "%H:%M")
        departure_df = df[df["Departure Trip"] == 1]
        departure_grouped = departure_df.groupby("Place_Normalize")["Name"].apply(list).to_dict()
        for trip_num, trip_venues in enumerate(trips, start=1):
            print(f"🚐 Departure Trip {trip_num}\n")
            current_time = start_time
            previous_venue = "MCPC"
            for venue, _ in trip_venues:
                if venue.lower() not in departure_grouped:
                    continue
                d = distances[names.index(previous_venue)][names.index(venue)]
                travel_minutes = estimate_travel_time(d)
                current_time += timedelta(minutes=travel_minutes)
                print(f"{venue} — ETA: {current_time.strftime('%I:%M %p')}")
                for name in departure_grouped[venue.lower()]:
                    print(f"   - {name}")
                print()
                previous_venue = venue
            d_back = distances[names.index(previous_venue)][names.index("MCPC")]
            travel_back = estimate_travel_time(d_back)
            current_time += timedelta(minutes=travel_back)
            print(f"(Depart to MCPC)\n")
            start_time = current_time

    def generate_carpool_trip(df):
        carpool_group = []
        for _, row in df.iterrows():
            place = row["Place"]
            if pd.notna(place) and place not in coords:
                carpool_group.append(row["Name"])
        if carpool_group:
            print("🚗 Car Pool — ETA: 10:00 AM")
            for name in carpool_group:
                print(f"   - {name}")
            print()

    def generate_after_service_trip(df):
        after_service_df = df[df["After Service"].notna()]
        after_service_group = after_service_df.groupby("Place_Normalize")["Name"].apply(list).to_dict() if not after_service_df.empty else {}
        if after_service_group:
            print("\n🏠 After Service")
            print("Carpool")
            counter = 1
            for venue, names_list in after_service_group.items():
                for name in names_list:
                    print(f"{counter}. {name}")
                    counter += 1
            print()

    def generate_after_youth_trip(df):
        after_youth_df = df[df["After Youth Fellowship"].notna()]
        after_youth_group = after_youth_df.groupby("Place_Normalize")["Name"].apply(list).to_dict() if not after_youth_df.empty else {}
        if after_youth_group:
            print("🏠 After Youth Fellowship")
            print("Gospel Van")
            counter = 1
            for venue, names_list in after_youth_group.items():
                for name in names_list:
                    print(f"{counter}. {name}")
                    counter += 1
            print()

    # --- Output Section ---
    generate_transport_brief(df)
    generate_worship_enablers_trip(df)
    generate_departure_trip()
    generate_carpool_trip(df)
    generate_after_service_trip(df)
    generate_after_youth_trip(df)

    # Restore stdout
    sys.stdout = sys.__stdout__

    # Display result in Streamlit
    output_text = buffer.getvalue()
    edited_text = st.text_area("Transport Arrangement Message", output_text, height=600)

    encoded_msg = urllib.parse.quote(edited_text, safe='')
    encoded_msg = encoded_msg.replace("%0A", "%0A")

    whatsapp_url = f"https://api.whatsapp.com/send?text={encoded_msg}"

    st.markdown(
        f"""
        <a href="{whatsapp_url}" target="_blank" style="text-decoration:none;">
            <button style="
                background-color:#25D366;
                color:white;
                border:none;
                border-radius:8px;
                padding:6px 10px;
                font-size:16px;
                cursor:pointer;">
                📤 Share via WhatsApp
            </button>
        </a>
        """,
        unsafe_allow_html=True
    )
