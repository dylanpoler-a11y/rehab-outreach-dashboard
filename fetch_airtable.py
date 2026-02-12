"""
Fetches data from Airtable and produces data.json for the dashboard.
Tables used:
  - Companies: company info, addresses, ownership, states
  - Cold Outreach: messages, responses, meetings
  - Contacts: links outreach to companies
  - Negotiations: pipeline deals

Env var required: AIRTABLE_PAT (Personal Access Token)
"""

import csv
import json
import re
import hashlib
import os
import io
import sys
import time
import urllib.request
import urllib.parse
from collections import defaultdict

BASE_ID = 'appsXvuuRisy7GiSH'
PAT = os.environ.get('AIRTABLE_PAT', '')

# ============================================================
# GEOCODING
# ============================================================
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

STATE_NAME_TO_ABBR = {
    'Alabama': 'AL', 'Alaska': 'AK', 'Arizona': 'AZ', 'Arkansas': 'AR',
    'California': 'CA', 'Colorado': 'CO', 'Connecticut': 'CT', 'Delaware': 'DE',
    'Florida': 'FL', 'Georgia': 'GA', 'Hawaii': 'HI', 'Idaho': 'ID',
    'Illinois': 'IL', 'Indiana': 'IN', 'Iowa': 'IA', 'Kansas': 'KS',
    'Kentucky': 'KY', 'Louisiana': 'LA', 'Maine': 'ME', 'Maryland': 'MD',
    'Massachusetts': 'MA', 'Michigan': 'MI', 'Minnesota': 'MN', 'Mississippi': 'MS',
    'Missouri': 'MO', 'Montana': 'MT', 'Nebraska': 'NE', 'Nevada': 'NV',
    'New Hampshire': 'NH', 'New Jersey': 'NJ', 'New Mexico': 'NM', 'New York': 'NY',
    'North Carolina': 'NC', 'North Dakota': 'ND', 'Ohio': 'OH', 'Oklahoma': 'OK',
    'Oregon': 'OR', 'Pennsylvania': 'PA', 'Rhode Island': 'RI', 'South Carolina': 'SC',
    'South Dakota': 'SD', 'Tennessee': 'TN', 'Texas': 'TX', 'Utah': 'UT',
    'Vermont': 'VT', 'Virginia': 'VA', 'Washington': 'WA', 'West Virginia': 'WV',
    'Wisconsin': 'WI', 'Wyoming': 'WY', 'Washington DC': 'DC'
}

ABBR_TO_STATE_NAME = {v: k for k, v in STATE_NAME_TO_ABBR.items()}

FL_ZIP_REGIONS = {
    '320': [30.4, -84.3], '321': [29.2, -81.0], '322': [30.3, -81.7],
    '323': [30.3, -81.7], '324': [30.4, -86.6], '325': [30.4, -87.2],
    '326': [29.2, -82.1], '327': [28.5, -81.4], '328': [28.5, -81.4],
    '329': [28.1, -80.6], '330': [25.8, -80.2], '331': [25.8, -80.2],
    '332': [25.8, -80.2], '333': [26.1, -80.1], '334': [26.7, -80.1],
    '335': [27.8, -82.6], '336': [27.8, -82.6], '337': [27.3, -82.5],
    '338': [28.0, -82.0], '339': [26.6, -81.9], '340': [26.6, -81.9],
    '341': [26.1, -80.4], '342': [28.0, -82.5], '344': [29.0, -82.5],
    '346': [27.5, -82.5], '347': [28.2, -82.2], '349': [26.4, -80.1],
}


def resolve_state_abbr(s):
    s = s.strip()
    if len(s) == 2 and s.upper() in STATE_COORDS:
        return s.upper()
    return STATE_NAME_TO_ABBR.get(s, '')


def get_coords(address, state_abbr, name):
    lat, lng = None, None
    zip_match = re.search(r'\b(\d{5})\b', address) if address else None
    if zip_match:
        prefix = zip_match.group(1)[:3]
        if state_abbr == 'FL' and prefix in FL_ZIP_REGIONS:
            lat, lng = FL_ZIP_REGIONS[prefix]
        elif state_abbr in STATE_COORDS:
            lat, lng = STATE_COORDS[state_abbr]
    elif state_abbr in STATE_COORDS:
        lat, lng = STATE_COORDS[state_abbr]
    if lat is not None:
        h = int(hashlib.md5(name.encode()).hexdigest()[:8], 16)
        lat += ((h % 1000) / 1000 - 0.5) * 0.8
        lng += (((h >> 10) % 1000) / 1000 - 0.5) * 0.8
    return lat, lng


# ============================================================
# GEOCODING — Census Bureau Batch API
# ============================================================
def batch_geocode_census(addresses):
    results = {}
    batch_size = 1000
    for batch_start in range(0, len(addresses), batch_size):
        batch = addresses[batch_start:batch_start + batch_size]
        print(f"  Geocoding batch {batch_start // batch_size + 1} ({len(batch)} addresses)...")
        csv_content = io.StringIO()
        writer = csv.writer(csv_content)
        for uid, addr in batch:
            parts = [p.strip() for p in addr.split(',')]
            if len(parts) >= 3:
                street = parts[0]
                city = parts[1]
                state_zip = parts[2]
                sz_match = re.match(r'([A-Za-z\s]+?)\s*(\d{5})', state_zip)
                if sz_match:
                    state, zipcode = sz_match.group(1).strip(), sz_match.group(2)
                else:
                    state, zipcode = state_zip.strip(), ''
                writer.writerow([uid, street, city, state, zipcode])
            elif len(parts) >= 2:
                writer.writerow([uid, parts[0], parts[1], '', ''])
            else:
                writer.writerow([uid, addr, '', '', ''])
        csv_data = csv_content.getvalue().encode('utf-8')
        try:
            url = 'https://geocoding.geo.census.gov/geocoder/locations/addressbatch'
            boundary = '----BatchBoundary'
            body = (
                f'--{boundary}\r\n'
                f'Content-Disposition: form-data; name="addressFile"; filename="addresses.csv"\r\n'
                f'Content-Type: text/csv\r\n\r\n'
            ).encode('utf-8') + csv_data + (
                f'\r\n--{boundary}\r\n'
                f'Content-Disposition: form-data; name="benchmark"\r\n\r\n'
                f'Public_AR_Current\r\n'
                f'--{boundary}--\r\n'
            ).encode('utf-8')
            req = urllib.request.Request(url, data=body,
                                        headers={'Content-Type': f'multipart/form-data; boundary={boundary}'},
                                        method='POST')
            with urllib.request.urlopen(req, timeout=120) as resp:
                response_text = resp.read().decode('utf-8')
            for line in response_text.strip().split('\n'):
                if not line.strip():
                    continue
                for row in csv.reader(io.StringIO(line)):
                    if len(row) >= 6 and row[2].strip().lower() == 'match':
                        uid = row[0].strip().strip('"')
                        coords_str = row[5].strip().strip('"')
                        if coords_str:
                            try:
                                lng_s, lat_s = coords_str.split(',')
                                results[uid] = (float(lat_s.strip()), float(lng_s.strip()))
                            except (ValueError, IndexError):
                                pass
            print(f"    Got {sum(1 for b in batch if b[0] in results)} matches")
        except Exception as e:
            print(f"    Geocoding error: {e}")
        if batch_start + batch_size < len(addresses):
            time.sleep(1)
    return results


# ============================================================
# GEOCODE CACHE
# ============================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
GEOCODE_CACHE_FILE = os.path.join(SCRIPT_DIR, 'geocode_cache.json')


def load_geocode_cache():
    if os.path.exists(GEOCODE_CACHE_FILE):
        with open(GEOCODE_CACHE_FILE, 'r') as f:
            return json.load(f)
    return {}


def save_geocode_cache(cache):
    with open(GEOCODE_CACHE_FILE, 'w') as f:
        json.dump(cache, f, indent=2)


# ============================================================
# AIRTABLE HELPERS
# ============================================================
def airtable_fetch_all(table, fields=None, formula=None):
    """Fetch all records from an Airtable table, handling pagination."""
    records = []
    offset = None
    while True:
        params = {}
        if fields:
            params['fields[]'] = fields
        if formula:
            params['filterByFormula'] = formula
        if offset:
            params['offset'] = offset
        url = f'https://api.airtable.com/v0/{BASE_ID}/{urllib.parse.quote(table)}?{urllib.parse.urlencode(params, doseq=True)}'
        req = urllib.request.Request(url, headers={'Authorization': f'Bearer {PAT}'})
        try:
            data = json.load(urllib.request.urlopen(req, timeout=60))
        except Exception as e:
            print(f"  Error fetching {table}: {e}")
            break
        records.extend(data.get('records', []))
        offset = data.get('offset')
        if not offset:
            break
    return records


def at_val(fields, key, default=''):
    """Extract a value from Airtable fields, handling AI-generated fields."""
    v = fields.get(key, default)
    if isinstance(v, dict):
        # AI-generated fields have {state, value, isStale}
        return v.get('value', default) or default
    if isinstance(v, list):
        return v  # Return lists as-is
    return v if v is not None else default


# ============================================================
# MAIN
# ============================================================
def main():
    if not PAT:
        print("ERROR: Set AIRTABLE_PAT environment variable")
        sys.exit(1)

    print("Fetching data from Airtable...")

    # ---- Fetch all tables in parallel-ish ----
    print("  Fetching Companies...")
    companies_raw = airtable_fetch_all('Companies', [
        'Name', 'HQ Address', 'Full HQ State Name', 'HQ State', 'Ownership',
        'All State(s) Operating In', 'Override', 'Website', 'State Tier',
        'State Tier Categories',
    ])
    print(f"    {len(companies_raw)} companies")

    print("  Fetching Contacts...")
    contacts_raw = airtable_fetch_all('Contacts', [
        'Name', 'Companies', 'Cold Outreach',
    ])
    print(f"    {len(contacts_raw)} contacts")

    print("  Fetching Cold Outreach...")
    outreach_raw = airtable_fetch_all('Cold Outreach', [
        'Contacts', 'Date Sent', 'Message Medium', 'Account', 'Message Type',
        'Responded', 'Scheduled Intro Call', 'Assisted Meeting', 'Not Interested',
        'Meeting Date', 'Opened', 'Follow Up Priority (from Follow Ups)',
    ])
    print(f"    {len(outreach_raw)} outreach messages")

    print("  Fetching Negotiations...")
    negotiations_raw = airtable_fetch_all('Negotiations', [
        'Status', 'Name (from Companies)', 'Companies', 'Name (from Clients)',
        'Assignee',
    ])
    print(f"    {len(negotiations_raw)} negotiations")

    # ============================================================
    # BUILD LOOKUP MAPS
    # ============================================================

    # Company record ID -> company data
    company_by_id = {}
    for r in companies_raw:
        f = r['fields']
        name = f.get('Name', '').strip()
        if not name:
            continue
        address = at_val(f, 'HQ Address', '')
        if isinstance(address, str):
            address = address.strip()
        else:
            address = ''
        website = at_val(f, 'Website', '')
        if isinstance(website, str):
            website = website.strip()
        else:
            website = ''

        all_states_list = f.get('All State(s) Operating In', [])
        all_states_str = ', '.join(all_states_list) if isinstance(all_states_list, list) else str(all_states_list)

        full_state = f.get('Full HQ State Name', '').strip() if isinstance(f.get('Full HQ State Name'), str) else ''
        hq_state_raw = at_val(f, 'HQ State', '')
        hq_state = hq_state_raw.strip() if isinstance(hq_state_raw, str) else ''
        state_abbr = resolve_state_abbr(full_state) or resolve_state_abbr(hq_state)

        ownership = f.get('Ownership', '') or ''
        override = bool(f.get('Override'))
        state_tier_raw = at_val(f, 'State Tier', '')
        state_tier = state_tier_raw.strip() if isinstance(state_tier_raw, str) else ''

        company_by_id[r['id']] = {
            'name': name,
            'key': re.sub(r'\s+', ' ', name.lower().strip().replace('\u202f', ' ').replace('\u00a0', ' ')),
            'address': address,
            'fullState': full_state,
            'state': state_abbr,
            'ownership': ownership,
            'allStates': all_states_str,
            'website': website,
            'override': override,
            'stateTier': state_tier,
        }

    # Contact record ID -> list of company record IDs
    contact_to_companies = {}
    for r in contacts_raw:
        f = r['fields']
        comp_ids = f.get('Companies', [])
        if comp_ids:
            contact_to_companies[r['id']] = comp_ids

    # ============================================================
    # PROCESS COLD OUTREACH → per-company aggregation
    # ============================================================
    # Per-company aggregation
    comp_msgs = defaultdict(lambda: {
        'msgsSent': 0, 'byMedium': defaultdict(int), 'byAccount': defaultdict(int),
        'responded': False, 'respondedCount': 0,
        'scheduledIntro': False, 'assistedMeeting': False,
        'notInterested': False, 'followUpLater': False,
        'contacts': set(), 'firstMsgDate': '', 'lastMsgDate': '',
        'meetingDate': '', 'opened': False, 'viewedProfile': False,
    })

    all_messages = []

    for r in outreach_raw:
        f = r['fields']
        contact_ids = f.get('Contacts', [])
        date_sent = f.get('Date Sent', '')
        medium = f.get('Message Medium', '') or ''
        account = f.get('Account', '') or ''
        msg_type = f.get('Message Type', [])
        responded_val = at_val(f, 'Responded', '')
        scheduled = f.get('Scheduled Intro Call', '') or ''
        assisted = f.get('Assisted Meeting', '') or ''
        not_interested = f.get('Not Interested', [])
        meeting_date = f.get('Meeting Date', '') or ''
        opened = f.get('Opened', '') or ''
        follow_up_priority = f.get('Follow Up Priority (from Follow Ups)', [])

        # Format date
        date_str = ''
        if date_sent:
            try:
                from datetime import datetime
                dt = datetime.fromisoformat(date_sent.replace('Z', '+00:00'))
                date_str = dt.strftime('%m/%d/%Y')
            except:
                date_str = date_sent[:10] if len(date_sent) >= 10 else date_sent

        meeting_date_str = ''
        if meeting_date:
            try:
                from datetime import datetime
                dt = datetime.fromisoformat(meeting_date.replace('Z', '+00:00'))
                meeting_date_str = dt.strftime('%m/%d/%Y')
            except:
                meeting_date_str = meeting_date[:10]

        # Determine which companies this outreach belongs to
        company_ids_for_msg = set()
        for cid in contact_ids:
            for comp_id in contact_to_companies.get(cid, []):
                company_ids_for_msg.add(comp_id)

        is_responded = (str(responded_val).strip() not in ('', '0', 'False', 'false', 'None'))
        is_scheduled = scheduled and scheduled.lower() not in ('', 'no', 'n/a')
        is_assisted = assisted and assisted.lower() not in ('', 'no', 'n/a')
        is_not_interested = bool(not_interested)
        is_opened = opened and opened.lower() not in ('', 'no', 'n/a')
        has_follow_up = bool(follow_up_priority)

        for comp_id in company_ids_for_msg:
            if comp_id not in company_by_id:
                continue
            agg = comp_msgs[comp_id]
            agg['msgsSent'] += 1
            if medium:
                agg['byMedium'][medium] += 1
            if account:
                agg['byAccount'][account] += 1
            for cid in contact_ids:
                agg['contacts'].add(cid)
            if is_responded:
                agg['responded'] = True
                agg['respondedCount'] += 1
            if is_scheduled:
                agg['scheduledIntro'] = True
            if is_assisted:
                agg['assistedMeeting'] = True
            if is_not_interested:
                agg['notInterested'] = True
            if has_follow_up:
                agg['followUpLater'] = True
            if is_opened:
                agg['opened'] = True
            if date_str:
                if not agg['firstMsgDate'] or date_str < agg['firstMsgDate']:
                    agg['firstMsgDate'] = date_str
                if not agg['lastMsgDate'] or date_str > agg['lastMsgDate']:
                    agg['lastMsgDate'] = date_str
            if meeting_date_str:
                agg['meetingDate'] = meeting_date_str

        all_messages.append({
            'date': date_str,
            'medium': medium,
            'account': account,
        })

    # ============================================================
    # FETCH GOOGLE SHEET PIPELINE (primary pipeline source)
    # ============================================================
    GSHEET_ID = '19w2nLn7VrNEWQSRaEIgPQuZwPi88ABNKWnl8Bh7pqVk'
    PIPELINE_GID = '1109153656'
    ACTIONS_GID = '2003109190'

    print("  Fetching Google Sheet pipeline...")
    pipeline = []
    actions = []
    gsheet_pipeline_by_name = {}  # normalized name -> deal

    try:
        # Fetch pipeline detail tab
        pipe_url = f'https://docs.google.com/spreadsheets/d/{GSHEET_ID}/export?format=csv&gid={PIPELINE_GID}'
        req = urllib.request.Request(pipe_url)
        with urllib.request.urlopen(req, timeout=30) as resp:
            pipe_csv = resp.read().decode('utf-8')

        # Skip to the header row (look for "Facility Name")
        pipe_lines = pipe_csv.strip().split('\n')
        header_idx = None
        for i, line in enumerate(pipe_lines):
            if 'Facility Name' in line:
                header_idx = i
                break
        if header_idx is None:
            print("    Could not find pipeline header row!")
            raise Exception("Missing header")
        pipe_body = '\n'.join(pipe_lines[header_idx:])
        reader = csv.DictReader(io.StringIO(pipe_body))
        for row in reader:
            name = (row.get('Facility Name') or '').strip()
            if not name:
                continue
            status = (row.get('Status') or '').strip()
            priority = (row.get('Priority') or '').strip()
            deal_type = (row.get('Type') or '').strip()
            states = (row.get('State(s)') or '').strip()
            ebitda = (row.get('EBITDA / Financials') or '').strip()
            asking_price = (row.get('Asking Price') or '').strip()
            nda_status = (row.get('NDA Status') or '').strip()
            data_room = (row.get('Data Room') or '').strip()
            site_visit = (row.get('Site Visit') or '').strip()
            key_contact = (row.get('Key Contact') or '').strip()
            next_action = (row.get('Next Action') or '').strip()
            action_owner = (row.get('Action Owner') or '').strip()
            deadline = (row.get('Deadline') or '').strip()
            last_update = (row.get('Last Update') or '').strip()
            days_since = (row.get('Days Since Update') or '').strip()
            notes = (row.get('Notes') or '').strip()
            deal_num = (row.get('#') or '').strip()

            deal = {
                'name': name,
                'status': status,
                'type': deal_type,
                'priority': priority,
                'states': states,
                'ebitda': ebitda,
                'askingPrice': asking_price,
                'ndaStatus': nda_status,
                'dataRoom': data_room,
                'siteVisit': site_visit,
                'keyContact': key_contact,
                'nextAction': next_action,
                'actionOwner': action_owner,
                'deadline': deadline,
                'lastUpdate': last_update,
                'daysSinceUpdate': days_since,
                'notes': notes,
                'dealNumber': deal_num,
            }
            pipeline.append(deal)
            norm_name = re.sub(r'\s+', ' ', name.lower().strip().replace('\u202f', ' ').replace('\u00a0', ' '))
            gsheet_pipeline_by_name[norm_name] = deal

        print(f"    {len(pipeline)} pipeline deals from Google Sheet")

        # Fetch action items tab
        act_url = f'https://docs.google.com/spreadsheets/d/{GSHEET_ID}/export?format=csv&gid={ACTIONS_GID}'
        req = urllib.request.Request(act_url)
        with urllib.request.urlopen(req, timeout=30) as resp:
            act_csv = resp.read().decode('utf-8')

        # Skip the header rows (first 2 rows are title + blank)
        lines = act_csv.strip().split('\n')
        # Find the header row with "Priority,Action Item,..."
        header_idx = None
        for i, line in enumerate(lines):
            if 'Action Item' in line and 'Priority' in line:
                header_idx = i
                break
        if header_idx is not None:
            csv_body = '\n'.join(lines[header_idx:])
            reader = csv.DictReader(io.StringIO(csv_body))
            for row in reader:
                action_item = (row.get('Action Item') or '').strip()
                if not action_item:
                    continue
                actions.append({
                    'priority': (row.get('Priority') or '').strip(),
                    'action': action_item,
                    'facility': (row.get('Facility') or '').strip(),
                    'owner': (row.get('Owner') or '').strip(),
                    'deadline': (row.get('Deadline') or '').strip(),
                    'status': (row.get('Status') or '').strip(),
                    'notes': (row.get('Notes') or '').strip(),
                    'pipelineStatus': (row.get('Pipeline Status') or '').strip(),
                })
            print(f"    {len(actions)} action items from Google Sheet")

    except Exception as e:
        print(f"    Google Sheet fetch error: {e}")
        print(f"    Falling back to Airtable Negotiations only")

    # ============================================================
    # MATCH PIPELINE TO COMPANIES (Google Sheet name → Airtable company ID)
    # ============================================================
    # Name aliases from pipeline deal names to Airtable company names
    PIPELINE_ALIASES = {
        'nola detox & recovery center': 'nola detox',
        'second chances': 'second chances addiction recovery center',
        'serenity treatment centers': 'serenity treatment center',
        'sanctuary': 'sanctuary louisiana',
        'asheville detox (healthcare alliance)': 'asheville detox center',
        'recovery now / longbranch': 'longbranch healthcare',
        'dreamlife / crestview': 'dreamlife recovery pa',
        'new waters': 'new waters recovery',
        'momentum recovery': 'momentum recovery',
        'the grove recovery': 'the grove recovery centers',
        'sycamour': 'sycamore behavioral health',
        'new vista / ethan crossing': 'ethan crossing addiction treatment',
        'southeast detox / addiction ctr': 'southeast detox',
        'cardinal': 'cardinal recovery',
        'ghr': 'ghr center for addiction recovery and treatment',
        'peachtree detox (evoraa)': 'peachtree detox',
        'revive recover': 'gateway to sobriety (revive recover)',
        'southern sky': 'southern sky recovery',
        'the wave': 'the wave international',
        'woodlake center': 'woodlake addiction recovery',
        'the sylvia brafman mh center': 'the sylvia brafman mental health center',
        'turning leaf behavioral health': 'turning leaf behavioral health services',
        'centric': 'reign residential treatment center',
    }

    # Build reverse lookup: company key -> company ID
    company_key_to_id = {}
    for comp_id, comp in company_by_id.items():
        company_key_to_id[comp['key']] = comp_id

    # Also build from Airtable Negotiations for direct ID matching
    neg_company_ids = {}
    for r in negotiations_raw:
        f = r['fields']
        company_ids = f.get('Companies', [])
        company_names = f.get('Name (from Companies)', [])
        for comp_id in company_ids:
            neg_company_ids[comp_id] = True

    pipeline_by_company_id = {}
    matched = 0
    for deal in pipeline:
        norm_name = re.sub(r'\s+', ' ', deal['name'].lower().strip().replace('\u202f', ' ').replace('\u00a0', ' '))
        # Try direct match
        comp_id = company_key_to_id.get(norm_name)
        # Try alias
        if not comp_id:
            aliased = PIPELINE_ALIASES.get(norm_name)
            if aliased:
                comp_id = company_key_to_id.get(aliased)
        if comp_id:
            deal['airtableId'] = comp_id
            pipeline_by_company_id[comp_id] = deal
            matched += 1

    print(f"  Pipeline matched to companies: {matched}/{len(pipeline)}")

    # ============================================================
    # GEOCODING
    # ============================================================
    geocode_cache = load_geocode_cache()
    print(f"  Geocode cache: {len(geocode_cache)} entries")

    addresses_to_geocode = []
    for comp_id, comp in company_by_id.items():
        key = comp['key']
        if comp['address'] and key not in geocode_cache:
            addresses_to_geocode.append((key, comp['address']))

    if addresses_to_geocode:
        print(f"  Geocoding {len(addresses_to_geocode)} new addresses...")
        geo_results = batch_geocode_census(addresses_to_geocode)
        for uid, coords in geo_results.items():
            geocode_cache[uid] = list(coords)
        for uid, _ in addresses_to_geocode:
            if uid not in geocode_cache:
                geocode_cache[uid] = None
        save_geocode_cache(geocode_cache)
        print(f"  Geocoded {len(geo_results)} new addresses")
    else:
        print("  All addresses already cached")

    # ============================================================
    # BUILD FINAL COMPANIES LIST
    # ============================================================
    companies = []
    for comp_id, comp in company_by_id.items():
        key = comp['key']
        agg = comp_msgs.get(comp_id, {
            'msgsSent': 0, 'byMedium': {}, 'byAccount': {},
            'responded': False, 'respondedCount': 0,
            'scheduledIntro': False, 'assistedMeeting': False,
            'notInterested': False, 'followUpLater': False,
            'contacts': set(), 'firstMsgDate': '', 'lastMsgDate': '',
            'meetingDate': '', 'opened': False, 'viewedProfile': False,
        })

        # Geocode
        cached = geocode_cache.get(key)
        if cached:
            lat, lng = cached
        else:
            lat, lng = get_coords(comp['address'], comp['state'], comp['name'])

        pipe = pipeline_by_company_id.get(comp_id)

        companies.append({
            'airtableId': comp_id,
            'name': comp['name'],
            'address': comp['address'],
            'state': comp['state'],
            'fullState': comp['fullState'],
            'ownership': comp['ownership'],
            'override': comp['override'],
            'website': comp['website'],
            'allStates': comp['allStates'],
            'stateTier': comp['stateTier'],
            'msgsSent': agg['msgsSent'],
            'byMedium': dict(agg['byMedium']) if isinstance(agg['byMedium'], defaultdict) else agg['byMedium'],
            'byAccount': dict(agg['byAccount']) if isinstance(agg['byAccount'], defaultdict) else agg['byAccount'],
            'responded': agg['responded'],
            'respondedCount': agg['respondedCount'],
            'scheduledIntro': agg['scheduledIntro'],
            'assistedMeeting': agg['assistedMeeting'],
            'notInterested': agg['notInterested'],
            'followUpLater': agg['followUpLater'],
            'contactCount': len(agg['contacts']),
            'contacts': list(agg['contacts']) if isinstance(agg['contacts'], set) else agg['contacts'],
            'firstMsg': agg['firstMsgDate'],
            'lastMsg': agg['lastMsgDate'],
            'meetingDate': agg['meetingDate'],
            'opened': agg['opened'],
            'viewedProfile': agg.get('viewedProfile', False),
            'inPipeline': pipe is not None,
            'pipelineStatus': pipe['status'] if pipe else '',
            'pipelinePriority': pipe.get('priority', '') if pipe else '',
            'pipelineType': pipe.get('type', '') if pipe else '',
            'pipelineEbitda': pipe.get('ebitda', '') if pipe else '',
            'pipelineAskingPrice': pipe.get('askingPrice', '') if pipe else '',
            'pipelineNda': pipe.get('ndaStatus', '') if pipe else '',
            'pipelineDataRoom': pipe.get('dataRoom', '') if pipe else '',
            'pipelineSiteVisit': pipe.get('siteVisit', '') if pipe else '',
            'pipelineKeyContact': pipe.get('keyContact', '') if pipe else '',
            'pipelineNextAction': pipe.get('nextAction', '') if pipe else '',
            'pipelineActionOwner': pipe.get('actionOwner', '') if pipe else '',
            'pipelineDeadline': pipe.get('deadline', '') if pipe else '',
            'pipelineLastUpdate': pipe.get('lastUpdate', '') if pipe else '',
            'pipelineDaysSince': pipe.get('daysSinceUpdate', '') if pipe else '',
            'pipelineNotes': pipe.get('notes', '') if pipe else '',
            'lat': lat,
            'lng': lng,
        })

    # ============================================================
    # COMPUTE AGGREGATES
    # ============================================================
    medium_counts = defaultdict(int)
    account_counts = defaultdict(int)
    monthly_counts = defaultdict(int)
    for m in all_messages:
        if m['medium']:
            medium_counts[m['medium']] += 1
        if m['account']:
            account_counts[m['account']] += 1
        if m['date']:
            match = re.match(r'(\d{1,2})/\d{1,2}/(\d{4})', m['date'])
            if match:
                month_key = f"{match.group(2)}-{int(match.group(1)):02d}"
                monthly_counts[month_key] += 1

    output = {
        'companies': companies,
        'pipeline': pipeline,
        'actions': actions,
        'meta': {
            'totalMessages': len(all_messages),
            'mediumCounts': dict(medium_counts),
            'accountCounts': dict(account_counts),
            'monthlyCounts': dict(sorted(monthly_counts.items())),
        },
    }

    out_path = os.path.join(SCRIPT_DIR, 'data.json')
    with open(out_path, 'w') as f:
        json.dump(output, f, indent=2)

    c = companies
    gc = geocode_cache
    geocoded = sum(1 for x in c if x['lat'] and gc.get(x['name'].lower().strip()))
    print(f"\n=== OUTPUT ===")
    print(f"Total companies: {len(c)}")
    print(f"With coordinates: {sum(1 for x in c if x['lat'])}")
    print(f"Geocoded (precise): {geocoded}")
    print(f"With allStates: {sum(1 for x in c if x['allStates'])}")
    print(f"With messages: {sum(1 for x in c if x['msgsSent'] > 0)}")
    print(f"Total messages: {output['meta']['totalMessages']}")
    print(f"Responded: {sum(1 for x in c if x['responded'])}")
    print(f"Scheduled intro: {sum(1 for x in c if x['scheduledIntro'])}")
    print(f"Assisted meeting: {sum(1 for x in c if x['assistedMeeting'])}")
    print(f"Not interested: {sum(1 for x in c if x['notInterested'])}")
    print(f"In pipeline: {sum(1 for x in c if x['inPipeline'])}")
    print(f"Pipeline deals: {len(pipeline)}")
    print(f"Mediums: {output['meta']['mediumCounts']}")
    print(f"Accounts: {output['meta']['accountCounts']}")


if __name__ == '__main__':
    main()
