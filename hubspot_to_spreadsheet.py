#!/usr/bin/env python3
"""
HubSpot to Spreadsheet Integration Script

This script fetches contacts from HubSpot and updates the State DOE Backlink
Outreach Tracker spreadsheet with contact information.

SETUP:
    1. Install dependencies:
       pip install openpyxl requests

    2. Set API key (optional - defaults to the key below):
       export HUBSPOT_API_KEY="your-api-key-here"

    3. Run the script:
       python hubspot_to_spreadsheet.py

The script will:
1. Fetch all contacts from HubSpot with relevant properties
2. Match contacts to states in the spreadsheet based on company/state fields
3. Update the spreadsheet with contact details (name, title, phone, email)
4. Create a backup before making changes
"""

import requests
import openpyxl
from openpyxl.styles import PatternFill, Font
from datetime import datetime
import os
import sys

# HubSpot API configuration
# Set via environment variable: export HUBSPOT_API_KEY="your-api-key"
HUBSPOT_API_KEY = os.environ.get("HUBSPOT_API_KEY")
HUBSPOT_BASE_URL = "https://api.hubapi.com"

# File paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SPREADSHEET_PATH = os.path.join(SCRIPT_DIR, "State_DOE_Backlink_Outreach_Tracker.xlsx")

# State name variations for matching
STATE_VARIATIONS = {
    "california": ["california", "ca", "cde", "california department of education"],
    "ohio": ["ohio", "oh", "ode", "ohio department of education"],
    "pennsylvania": ["pennsylvania", "pa", "pde", "pennsylvania department of education"],
    "wisconsin": ["wisconsin", "wi", "dpi", "wisconsin dpi"],
    "oklahoma": ["oklahoma", "ok", "osde", "oklahoma state department of education"],
    "texas": ["texas", "tx", "tea", "texas education agency"],
    "florida": ["florida", "fl", "fldoe", "florida department of education"],
    "colorado": ["colorado", "co", "cde", "colorado department of education"],
    "alabama": ["alabama", "al", "alsde"],
    "arkansas": ["arkansas", "ar", "ade"],
    "connecticut": ["connecticut", "ct", "csde"],
    "delaware": ["delaware", "de", "ddoe"],
    "georgia": ["georgia", "ga", "gadoe"],
    "idaho": ["idaho", "id", "isde"],
    "illinois": ["illinois", "il", "isbe"],
    "indiana": ["indiana", "in", "idoe"],
    "iowa": ["iowa", "ia", "iowa doe"],
    "kentucky": ["kentucky", "ky", "kde"],
    "louisiana": ["louisiana", "la", "ldoe"],
    "michigan": ["michigan", "mi", "mde"],
    "minnesota": ["minnesota", "mn", "mde"],
    "mississippi": ["mississippi", "ms", "mde"],
    "missouri": ["missouri", "mo", "dese"],
    "montana": ["montana", "mt", "opi"],
    "nebraska": ["nebraska", "ne", "nde"],
    "nevada": ["nevada", "nv", "doe"],
    "new jersey": ["new jersey", "nj", "njdoe"],
    "new york": ["new york", "ny", "nysed"],
    "north carolina": ["north carolina", "nc", "ncdpi"],
    "north dakota": ["north dakota", "nd", "nddpi"],
    "rhode island": ["rhode island", "ri", "ride"],
    "south carolina": ["south carolina", "sc", "scde"],
    "tennessee": ["tennessee", "tn", "tdoe"],
    "utah": ["utah", "ut", "usbe"],
    "virginia": ["virginia", "va", "vdoe"],
    "west virginia": ["west virginia", "wv", "wvde"],
}


def get_hubspot_headers():
    """Get headers for HubSpot API requests."""
    return {
        "Authorization": f"Bearer {HUBSPOT_API_KEY}",
        "Content-Type": "application/json"
    }


def fetch_hubspot_contacts():
    """
    Fetch all contacts from HubSpot with relevant properties.
    Returns a list of contact dictionaries.
    """
    contacts = []
    url = f"{HUBSPOT_BASE_URL}/crm/v3/objects/contacts"

    # Properties to fetch
    properties = [
        "firstname", "lastname", "email", "phone", "jobtitle",
        "company", "state", "city", "address", "hs_lead_status",
        "notes_last_updated", "hs_object_id"
    ]

    params = {
        "limit": 100,
        "properties": ",".join(properties)
    }

    print("Fetching contacts from HubSpot...")

    while url:
        try:
            response = requests.get(url, headers=get_hubspot_headers(), params=params)
            response.raise_for_status()
            data = response.json()

            for result in data.get("results", []):
                props = result.get("properties", {})
                contact = {
                    "id": result.get("id"),
                    "firstname": props.get("firstname", ""),
                    "lastname": props.get("lastname", ""),
                    "email": props.get("email", ""),
                    "phone": props.get("phone", ""),
                    "jobtitle": props.get("jobtitle", ""),
                    "company": props.get("company", ""),
                    "state": props.get("state", ""),
                    "city": props.get("city", ""),
                    "hs_lead_status": props.get("hs_lead_status", "")
                }
                contacts.append(contact)

            # Check for pagination
            paging = data.get("paging", {})
            next_link = paging.get("next", {}).get("link")
            if next_link:
                url = next_link
                params = {}  # params are included in the next link
            else:
                url = None

        except requests.exceptions.RequestException as e:
            print(f"Error fetching contacts: {e}")
            break

    print(f"Fetched {len(contacts)} contacts from HubSpot")
    return contacts


def fetch_hubspot_companies():
    """
    Fetch all companies from HubSpot with relevant properties.
    Returns a list of company dictionaries.
    """
    companies = []
    url = f"{HUBSPOT_BASE_URL}/crm/v3/objects/companies"

    properties = [
        "name", "domain", "state", "city", "phone", "industry",
        "description", "hs_object_id"
    ]

    params = {
        "limit": 100,
        "properties": ",".join(properties)
    }

    print("Fetching companies from HubSpot...")

    while url:
        try:
            response = requests.get(url, headers=get_hubspot_headers(), params=params)
            response.raise_for_status()
            data = response.json()

            for result in data.get("results", []):
                props = result.get("properties", {})
                company = {
                    "id": result.get("id"),
                    "name": props.get("name", ""),
                    "domain": props.get("domain", ""),
                    "state": props.get("state", ""),
                    "city": props.get("city", ""),
                    "phone": props.get("phone", ""),
                    "industry": props.get("industry", ""),
                    "description": props.get("description", "")
                }
                companies.append(company)

            # Check for pagination
            paging = data.get("paging", {})
            next_link = paging.get("next", {}).get("link")
            if next_link:
                url = next_link
                params = {}
            else:
                url = None

        except requests.exceptions.RequestException as e:
            print(f"Error fetching companies: {e}")
            break

    print(f"Fetched {len(companies)} companies from HubSpot")
    return companies


def match_contact_to_state(contact, state_name):
    """
    Check if a contact matches a given state based on various fields.
    Returns True if there's a match.
    """
    state_lower = state_name.lower()
    variations = STATE_VARIATIONS.get(state_lower, [state_lower])

    # Fields to check for state matches
    check_fields = [
        contact.get("state", "").lower(),
        contact.get("company", "").lower(),
        contact.get("city", "").lower()
    ]

    for field in check_fields:
        for variation in variations:
            if variation in field:
                return True

    # Check for DOE/education keywords combined with state
    company = contact.get("company", "").lower()
    if any(kw in company for kw in ["education", "doe", "department", "school"]):
        for variation in variations:
            if variation in company:
                return True

    return False


def is_education_contact(contact):
    """
    Check if a contact appears to be related to education/government.
    """
    keywords = [
        "education", "department", "doe", "superintendent", "curriculum",
        "director", "coordinator", "specialist", "administrator", "state",
        "government", "school", "academic", "financial literacy"
    ]

    fields_to_check = [
        contact.get("company", "").lower(),
        contact.get("jobtitle", "").lower()
    ]

    for field in fields_to_check:
        for keyword in keywords:
            if keyword in field:
                return True

    return False


def update_spreadsheet(contacts, companies):
    """
    Update the spreadsheet with contact information from HubSpot.
    """
    print(f"\nLoading spreadsheet from: {SPREADSHEET_PATH}")

    try:
        wb = openpyxl.load_workbook(SPREADSHEET_PATH)
    except FileNotFoundError:
        print(f"Error: Spreadsheet not found at {SPREADSHEET_PATH}")
        return False

    # Track updates
    updates_made = 0

    # Process each tier sheet
    tier_sheets = [
        "Tier 1 - Imminent",
        "Tier 2 - Relationships",
        "Tier 3 - Established"
    ]

    for sheet_name in tier_sheets:
        if sheet_name not in wb.sheetnames:
            print(f"Warning: Sheet '{sheet_name}' not found")
            continue

        ws = wb[sheet_name]
        print(f"\nProcessing sheet: {sheet_name}")

        # Find header row (row 3 based on structure)
        header_row = 3
        headers = {cell.value: idx for idx, cell in enumerate(ws[header_row], 1) if cell.value}

        # Column indices
        state_col = headers.get("State", 1)
        contact_col = headers.get("Contact Name", 3)
        title_col = headers.get("Title", 4)
        phone_col = headers.get("Phone", 5)
        email_col = headers.get("Email", 6)

        # Process each state row (starting from row 4)
        for row_num in range(4, ws.max_row + 1):
            state_cell = ws.cell(row=row_num, column=state_col)
            state_name = state_cell.value

            if not state_name:
                continue

            # Check if contact info is already filled
            existing_contact = ws.cell(row=row_num, column=contact_col).value
            if existing_contact and existing_contact.strip():
                print(f"  {state_name}: Already has contact info, skipping")
                continue

            # Find matching contacts for this state
            matching_contacts = []
            for contact in contacts:
                if match_contact_to_state(contact, state_name):
                    # Prioritize education-related contacts
                    if is_education_contact(contact):
                        matching_contacts.insert(0, contact)
                    else:
                        matching_contacts.append(contact)

            if matching_contacts:
                # Use the best match (first education-related contact)
                best_match = matching_contacts[0]

                # Build full name
                firstname = best_match.get("firstname", "") or ""
                lastname = best_match.get("lastname", "") or ""
                full_name = f"{firstname} {lastname}".strip()

                # Update cells if we have data
                if full_name:
                    ws.cell(row=row_num, column=contact_col, value=full_name)
                    updates_made += 1

                if best_match.get("jobtitle"):
                    ws.cell(row=row_num, column=title_col, value=best_match["jobtitle"])

                if best_match.get("phone"):
                    ws.cell(row=row_num, column=phone_col, value=best_match["phone"])

                if best_match.get("email"):
                    ws.cell(row=row_num, column=email_col, value=best_match["email"])

                print(f"  {state_name}: Updated with {full_name} ({best_match.get('email', 'no email')})")
            else:
                print(f"  {state_name}: No matching contacts found")

    # Save the workbook
    if updates_made > 0:
        # Create backup
        backup_path = SPREADSHEET_PATH.replace(".xlsx", f"_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        wb_backup = openpyxl.load_workbook(SPREADSHEET_PATH)
        wb_backup.save(backup_path)
        print(f"\nBackup created: {backup_path}")

        # Save updated workbook
        wb.save(SPREADSHEET_PATH)
        print(f"Spreadsheet updated with {updates_made} new contacts")
    else:
        print("\nNo updates were made to the spreadsheet")

    return True


def print_hubspot_summary(contacts, companies):
    """Print a summary of HubSpot data for review."""
    print("\n" + "="*60)
    print("HUBSPOT DATA SUMMARY")
    print("="*60)

    print(f"\nTotal Contacts: {len(contacts)}")
    print(f"Total Companies: {len(companies)}")

    # Show education-related contacts
    edu_contacts = [c for c in contacts if is_education_contact(c)]
    print(f"\nEducation-related contacts: {len(edu_contacts)}")

    if edu_contacts:
        print("\nEducation Contacts:")
        for c in edu_contacts[:20]:  # Show first 20
            name = f"{c.get('firstname', '')} {c.get('lastname', '')}".strip()
            company = c.get('company', 'N/A')
            title = c.get('jobtitle', 'N/A')
            email = c.get('email', 'N/A')
            state = c.get('state', 'N/A')
            print(f"  - {name} | {title} | {company} | {state} | {email}")

    # Show contacts by state
    print("\n\nContacts by State:")
    state_contacts = {}
    for c in contacts:
        state = c.get('state', 'Unknown')
        if state:
            state_contacts.setdefault(state, []).append(c)

    for state, state_list in sorted(state_contacts.items()):
        print(f"  {state}: {len(state_list)} contacts")


def main():
    """Main function to run the HubSpot to spreadsheet integration."""
    print("="*60)
    print("HubSpot to Spreadsheet Integration")
    print("="*60)
    print(f"\nScript started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # Check for API key
    if not HUBSPOT_API_KEY:
        print("\nError: HUBSPOT_API_KEY environment variable is not set.")
        print("Please set it with your HubSpot API key:")
        print("  export HUBSPOT_API_KEY='your-api-key-here'")
        print("\nThen run this script again.")
        sys.exit(1)

    # Fetch data from HubSpot
    contacts = fetch_hubspot_contacts()
    companies = fetch_hubspot_companies()

    if not contacts and not companies:
        print("\nNo data fetched from HubSpot. Please check:")
        print("  1. API key is valid")
        print("  2. Network connectivity")
        print("  3. HubSpot account has contacts/companies")
        return

    # Print summary
    print_hubspot_summary(contacts, companies)

    # Update spreadsheet
    print("\n" + "="*60)
    print("UPDATING SPREADSHEET")
    print("="*60)

    success = update_spreadsheet(contacts, companies)

    if success:
        print("\n" + "="*60)
        print("INTEGRATION COMPLETE")
        print("="*60)
        print(f"\nSpreadsheet location: {SPREADSHEET_PATH}")
        print("Please review the updates and verify accuracy.")
    else:
        print("\nIntegration failed. Please check errors above.")


if __name__ == "__main__":
    main()
