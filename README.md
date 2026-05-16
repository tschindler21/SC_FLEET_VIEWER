# SC Fleet Viewer

An org fleet viewer for Star Citizen. Collect fleet exports from your org members and generate a combined fleet overview page.

## How It Works

1. Each org member exports their fleet using the [SC Fleet Planner](https://github.com/tschindler21/SC_FLEET_PLANNER) — click **Export for Org** to get a clean CSV with ship names and insurance (no financial data).

2. Collect the CSV files and drop them into the `hangars/` folder. Each file becomes one member (filename = member name, e.g. `Matt_Graver.csv`).

3. Run the build script:

```bash
pip install openpyxl
python3 build.py
```

4. Open `index.html` — your org's combined fleet, searchable and grouped by member.

## CSV Format

The build script accepts both the full hangar export and the simplified org export from the fleet planner:

**Simplified (from Fleet Planner → Export for Org):**
```csv
"Member Name","Ship Name","Insurance"
"Matt Graver","Perseus","120 Month Insurance"
"Matt Graver","Paladin","Lifetime Insurance"
```

**Full hangar export (from the browser extension):**
```csv
"Pledge Name","Item Name","Type","Melt Value","Store Value","Insurance","Manufacturer","Date"
"Standalone Ship - Perseus","Perseus","Ship","580.00","580.00","120 Month Insurance","RSI","May 08, 2019"
```

Ship store prices are looked up automatically from `ships.json`.

## Fleet Overview

Open `fleet-overview.html` for an interactive fleet visualization with 4 switchable themes (Holotable, CIC, Infographic, Tactical). You can drag & drop a fleet planner save file (.json) to load any fleet.

## Org Configuration

Edit `org.json` to set your org name and motto:

```json
{
  "name": "YOUR_ORG",
  "motto": "YOUR_ORG Motto",
  "description": "YOUR_ORG Description Text",
  "logo": "logo.png"
}
```

Place your org logo in the `images/` folder.

## GitHub Pages (auto-build)

Push to GitHub and enable Pages — the included GitHub Actions workflow (`.github/workflows/build.yml`) automatically runs `build.py` and deploys `index.html`.

## Credits

- Ship prices from [Star Citizen Wiki API](https://api.star-citizen.wiki)
- Ship images from [Star Citizen Wiki](https://starcitizen.tools)
- Built with the [SC Fleet Planner](https://github.com/tschindler21/SC_FLEET_PLANNER)
- Star Citizen® is a trademark of Cloud Imperium Games

## Support

If this tool saved you some time, feel free to:

☕ [Buy me a coffee](https://buymeacoffee.com/schindi21)

**BTC:** `bc1qdyh5g2zska7s9e4vu27hzqyyre60t6khl4srnx`
