# Plan: Add Deal to Pipeline from Existing Company

## Overview
Add an "+ Add Deal" button to the Kanban board header that opens a modal where the user can search/select an existing company from `DATA.companies` and add it to the Pipeline Dashboard Google Sheet as a new deal.

## Changes

### 1. HTML — Add button to pipeline header (line ~1202)
- Add an "+ Add Deal" button next to the "Acquisition Pipeline" heading
- Add a new modal/popover for the "Add Deal" form (reuse the existing `.deal-edit-popover` styling pattern)

### 2. CSS — Style the add button and company search
- Style the "+ Add Deal" button to match the dashboard aesthetic
- Add a searchable company dropdown with autocomplete (type-ahead that filters `DATA.companies`)
- Selected company auto-fills: name, states, ownership type

### 3. JS — Add Deal form logic
- **Company search**: Filterable list of `DATA.companies` names. Typing narrows results. Click to select.
- **Pre-filled fields from company data**: name, states, type/ownership
- **User-selected fields**: Status (default "New Lead"), Priority (default "2 - Medium") — using the existing chip selector pattern from the edit popover
- **Optional fields**: Key Contact, Notes — simple text inputs
- **Submit**:
  1. POST to Apps Script with `{ newDeal: { name, status, priority, states, type, keyContact, notes } }`
  2. Add to `DATA.pipeline` locally for instant UI update
  3. Re-render pipeline board
  4. Show success toast

### 4. Apps Script — New `handleNewDeal` route (google_apps_script.js)
- Detect `data.newDeal` payload in `doPost()`
- Find the Pipeline Dashboard sheet and header row
- Append a new row with the provided fields mapped to correct columns
- Return `{ success: true, action: 'newDeal' }`

## UI Flow
1. User clicks "+ Add Deal" button on Kanban board header
2. Modal opens with a company search field at the top
3. User types company name → autocomplete filters matches from `DATA.companies`
4. User selects a company → name fills in, states/type auto-populate
5. User picks Status (chip selector, default "New Lead") and Priority (chip selector, default "2 - Medium")
6. Optional: key contact, notes
7. Click "Add to Pipeline" → writes to Google Sheet + updates board instantly
