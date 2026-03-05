# VCO Car Plan System Design

## Date: 2026-03-05

## Overview
Transportation management system for 48th FW TDY to Orland/Bodo, Norway. Manages 13 rental cars across ~50 personnel staying at 3 hotels. Primary focus: morning push car assignments based on daily flying schedule.

## Architecture
- **Backend**: Google Sheets (VCO edits directly)
- **Frontend**: Static HTML/JS app hosted on GitHub Pages
- **Data flow**: Google Sheets API (read/write), 60s auto-refresh

## Cars (13 total, Norse god names)
| # | Plate | Name | Notes |
|---|-------|------|-------|
| 1 | EP69777 | Skadi | Electric |
| 2 | EN71155 | Thor | |
| 3 | EN35700 | Freya | |
| 4 | EN46650 | Tyr | |
| 5 | NH17739 | Loki | |
| 6 | AS68952 | Heimdall | |
| 7 | EN43555 | Baldur | |
| 8 | EN53375 | Vidar | |
| 9 | EN87446 | Bragi | |
| 10 | EN10871 | Fenrir | |
| 11 | OW75VLN | Ragnar | |
| 12 | VJ36980 | Odin | Permanent - Skooby |
| 13 | EP82941 | Mjolnir | Permanent - Tony/Sean |

## Hotels
- **Flexbase** (yellow) - Pilots
- **Orland Kysthotell** (green) - Support (AFE, SARM, Intel, etc.)
- **Fosen Fjord Hotel** (red) - WX, some others

## Personnel Sections
Pilots, AFE, SARM, Intel, SEL, WX, IDMT, ISSO/GSSO, OMS, ALIS, FSE

## Authorized Fuelers (13 - one per car)
Merritt, Williams, Olmschenk, Cunningham-Wray, Hanlin, Moroz, Quimby, Ponder, Kolmer, Maldonado, Ortega, Evans, Moore

## Key Constraints
- 5 passengers per car max
- 2 cars permanently assigned (Odin, Mjolnir) = 11 assignable
- Pilots ride together (stay late for debrief)
- Morning push is primary focus
- EP-prefix plates are electric
- Personnel arrive/depart at different phases throughout the exercise

## Google Sheet Tabs
1. **Config** - Cars, personnel roster with hotels/roles/fueler auth/on-site dates
2. **Daily Schedule** - Date, show time, destinations, car assignment grid
3. **Availability** - Real-time car status, sign-out log
4. **Roster** - Who's on-site today (derived from dates + manual overrides)

## Auto-Assignment Algorithm
1. Remove Odin + Mjolnir from pool (11 cars remain)
2. Get today's on-site personnel
3. Group by hotel
4. For each hotel group:
   - Separate pilots from non-pilots
   - Calculate cars needed: ceil(group_size / 5)
   - Assign authorized fueler as driver per car
   - Fill pilots together, then support
5. Output pickup manifest per hotel

## Web App Features
- **Dashboard view** (everyone): car assignments, car status cards, clock
- **VCO panel** (toggle): set show time, mark attendance, auto-assign, manual tweaks, save to sheet
