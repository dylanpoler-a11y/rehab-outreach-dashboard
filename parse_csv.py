import csv
import json
import re
import hashlib

STATE_COORDS = {
    'AL': [32.806671, -86.791130], 'AK': [61.370716, -152.404419],
    'AZ': [33.729759, -111.431221], 'AR': [34.969704, -92.373123],
    'CA': [36.116203, -119.681564], 'CO': [39.059811, -105.311104],
    'CT': [41.597782, -72.755371], 'DE': [39.318523, -75.507141],
    'FL': [27.766279, -81.686783], 'GA': [33.040619, -83.643074],
    'HI': [21.094318, -157.498337], 'ID': [44.240459, -114.478828],
    'IL': [40.349457, -88.986137], 'IN': [39.849426, -86.258278],
    'IA': [42.011539, -93.210526], 'KS': [38.526600, -96.726486],
    'KY': [37.668140, -84.670067], 'LA': [31.169546, -91.867805],
    'ME': [44.693947, -69.381927], 'MD': [39.063946, -76.802101],
    'MA': [42.230171, -71.530106], 'MI': [43.326618, -84.536095],
    'MN': [45.694454, -93.900192], 'MS': [32.741646, -89.678696],
    'MO': [38.456085, -92.288368], 'MT': [46.921925, -110.454353],
    'NE': [41.125370, -98.268082], 'NV': [38.313515, -117.055374],
    'NH': [43.452492, -71.563896], 'NJ': [40.298904, -74.521011],
    'NM': [34.840515, -106.248482], 'NY': [42.165726, -74.948051],
    'NC': [35.630066, -79.806419], 'ND': [47.528912, -99.784012],
    'OH': [40.388783, -82.764915], 'OK': [35.565342, -96.928917],
    'OR': [44.572021, -122.070938], 'PA': [40.590752, -77.209755],
    'RI': [41.680893, -71.511780], 'SC': [33.856892, -80.945007],
    'SD': [44.299782, -99.438828], 'TN': [35.747845, -86.692345],
    'TX': [31.054487, -97.563461], 'UT': [40.150032, -111.862434],
    'VT': [44.045876, -72.710686], 'VA': [37.769337, -78.169968],
    'WA': [47.400902, -121.490494], 'WV': [38.491226, -80.954456],
    'WI': [44.268543, -89.616508], 'WY': [42.755966, -107.302490],
    'DC': [38.897438, -77.026817]
}

# Approximate zip code prefix -> lat/lng offsets for FL (heavy concentration)
FL_ZIP_REGIONS = {
    '320': [30.4, -84.3],   # Tallahassee
    '321': [29.2, -81.0],   # Daytona
    '322': [30.3, -81.7],   # Jacksonville
    '323': [30.3, -81.7],   # Jacksonville
    '324': [30.4, -86.6],   # Panama City
    '325': [30.4, -87.2],   # Pensacola
    '326': [29.2, -82.1],   # Gainesville
    '327': [28.5, -81.4],   # Orlando
    '328': [28.5, -81.4],   # Orlando
    '329': [28.1, -80.6],   # Melbourne
    '330': [25.8, -80.2],   # Miami
    '331': [25.8, -80.2],   # Miami
    '332': [25.8, -80.2],   # Miami
    '333': [26.1, -80.1],   # Fort Lauderdale
    '334': [26.7, -80.1],   # West Palm
    '335': [27.8, -82.6],   # Tampa
    '336': [27.8, -82.6],   # Tampa
    '337': [27.3, -82.5],   # St Pete
    '338': [28.0, -82.0],   # Lakeland
    '339': [26.6, -81.9],   # Fort Myers
    '340': [26.6, -81.9],   # Fort Myers (cont.)
    '341': [26.1, -80.4],   # Deerfield Beach area
    '342': [28.0, -82.5],   # Clearwater/Largo area
    '344': [29.0, -82.5],   # Ocala area
    '346': [27.5, -82.5],   # Sarasota/Bradenton
    '347': [28.2, -82.2],   # Plant City/Brandon area
    '349': [26.4, -80.1],   # Boca Raton area
}

def parse_num(val):
    val = str(val).strip()
    if not val or val == 'NaN':
        return 0
    try:
        parts = [float(x.strip()) for x in val.split(',') if x.strip() and x.strip() != 'NaN']
        return sum(parts)
    except:
        return 0

def get_coords(address, state, name):
    """Get lat/lng from address, with zip-based refinement and jitter."""
    lat, lng = None, None

    zip_match = re.search(r'\b(\d{5})\b', address)

    if zip_match:
        zipcode = zip_match.group(1)
        prefix = zipcode[:3]

        if state == 'FL' and prefix in FL_ZIP_REGIONS:
            lat, lng = FL_ZIP_REGIONS[prefix]
        elif state in STATE_COORDS:
            lat, lng = STATE_COORDS[state]
    elif state in STATE_COORDS:
        lat, lng = STATE_COORDS[state]

    if lat is not None:
        # Add deterministic jitter based on company name to spread overlapping markers
        h = int(hashlib.md5(name.encode()).hexdigest()[:8], 16)
        lat += ((h % 1000) / 1000 - 0.5) * 0.8
        lng += (((h >> 10) % 1000) / 1000 - 0.5) * 0.8

    return lat, lng

def parse_csv(filepath):
    with open(filepath, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    companies = []
    for r in rows:
        name = r.get('Name', '').strip()
        if not name:
            continue

        address = r.get('HQ Address', '').strip()
        state = r.get('HQ State', '').strip()
        ownership = r.get('Ownership', '').strip()
        override = r.get('Override', '').strip() == 'checked'
        website = r.get('Website', '').strip()

        msgs_sent = int(parse_num(r.get('# of Messages Sent (buyer + seller)', '')))
        msgs_responded = int(parse_num(r.get('# of Messages Responded (buyer + seller)', '')))
        intros = int(parse_num(r.get('Intro Count (buyer + seller)', '')))
        meetings = int(parse_num(r.get('# of Meetings Scheduled (buyer + seller)', '')))
        contacted = int(parse_num(r.get('# of Contacted People (buyer + seller)', '')))
        msgs_seller = int(parse_num(r.get('# of Messages Sent (seller)', '')))
        msgs_buyer = int(parse_num(r.get('# of Messages Sent (buyer)', '')))

        assisted = r.get('Assisted Intro Meeting (buyer + seller)', '').strip()
        has_assisted = 'TRUE' in assisted.upper() if assisted else False

        revenue = r.get('LinkedIn Revenue Estimate', '').strip()
        all_states = r.get('All State(s) Operating In', '').strip()
        first_msg = r.get('Date First Cold Message Sent', '').strip()
        last_msg = r.get('Date Last Cold Message Sent', '').strip()
        days_since = r.get('Days Since Last Cold Message Sent', '').strip()
        ai_ownership = r.get('AI Ownership Research', '').strip()
        contacts = r.get('Contacts', '').strip()
        profit_type = r.get('Profit Type', '').strip()
        state_tier = r.get('State Tier', '').strip()

        lat, lng = get_coords(address, state, name)

        companies.append({
            'name': name,
            'address': address,
            'state': state,
            'fullState': r.get('Full HQ State Name', '').strip(),
            'ownership': ownership,
            'override': override,
            'website': website,
            'msgsSent': msgs_sent,
            'msgsSeller': msgs_seller,
            'msgsBuyer': msgs_buyer,
            'msgsResponded': msgs_responded,
            'intros': intros,
            'meetings': meetings,
            'contacted': contacted,
            'hasAssisted': has_assisted,
            'revenue': revenue,
            'allStates': all_states,
            'firstMsg': first_msg,
            'lastMsg': last_msg,
            'daysSince': days_since,
            'aiOwnership': ai_ownership,
            'contacts': contacts,
            'profitType': profit_type,
            'stateTier': state_tier,
            'lat': lat,
            'lng': lng,
        })

    return companies


if __name__ == '__main__':
    companies = parse_csv('/Users/dylanpoler/Downloads/Working Sheet (8).csv')
    with open('/Users/dylanpoler/Downloads/rehab_dashboard/data.json', 'w') as f:
        json.dump(companies, f, indent=2)
    print(f'Parsed {len(companies)} companies')
    print(f'With coordinates: {sum(1 for c in companies if c["lat"])}')
