# DPEC Scenario Generation

A small Streamlit app that turns DPEC ("Deep Pollution and Emission Control") pathway data into a batch of Breeze-ready emission control scenario spreadsheets, so they can be fed directly into the CNCAP Breeze Model.

Instead of hand-editing a scenario template for every combination of pollutant, province, sector, and weather year, a user picks their inputs in the UI and the app generates and zips every resulting `.xlsx` file automatically.

**Live app:** [dpec-scenario-generation-m2k9tmwjkyorh78p78rcrt.streamlit.app](https://dpec-scenario-generation-m2k9tmwjkyorh78p78rcrt.streamlit.app/)

## What it does

The user selects, through a browser UI:

- **Base year / Target year** — which DPEC source data to compare (e.g. 2017 vs. 2030) to compute percentage emission reductions.
- **Scenarios** — one or more DPEC pathways (`Baseline`, `CleanAir`, `OTPCA`, `OTPNZCA`, `EPNZCA`).
- **Species** — any subset of SO₂, NOₓ, NH₃, VOC, PM2.5.
- **Provinces** — which of China's 31 provinces/municipalities/autonomous regions receive real reduction values.
- **Sectors** — which sectors (Power, Industry, Transportation, Residential, Agriculture) receive real reduction values.
- **Weather years** — one or more meteorological years to pair the scenario with.

Clicking **Get Excel Files** generates every requested output as a Breeze-formatted `.xlsx`, bundles them into a single `.zip`, and offers it for download.

## How it works

1. For each selected species, the app recursively generates every on/off combination (`generate_scenarios` in `app.py`) — e.g. picking 3 species produces up to 2³ = 8 variants (each species alone, every pair, all three together, and all off).
2. For each variant, it opens the base-year workbook (`data/2017_emission_report.xlsx` or `data/2020_emission_report.xlsx`) and the target-year workbook for the selected scenario (`data/{scenario}_All_Years/{target_year}_Emission.xlsx`).
3. For every selected province × sector pair, it computes the percentage reduction as `(base − target) / base`, clamped to a minimum of 0%, and writes it into a copy of the Breeze scenario template (`管控方案模板.xlsx`). Province/sector pairs *not* selected are left at 0% change, even if the species is "on."
4. One workbook is produced per (scenario × species combination), then duplicated once per selected weather year — the underlying data doesn't change across weather years, only the filename does.
5. All generated workbooks are zipped in memory and served as a single download.

### Output filename convention

```
{date}_{scenario}_{target_year}Target_{species_binary}_{base_year}Base_{weather_year}Met-{base_year}-{weather_year}
```

- `date` — the day the batch was generated (`YYYYMMDD`)
- `species_binary` — 5-digit code for which species are on (`1`) or off (`0`), in order SO₂-NOₓ-NH₃-VOC-PM

**Example:** `20260709_EPNZCA_2035Target_11100_2020Base_2019Met-2020-2019.xlsx` → EPNZCA scenario, 2035 target, SO₂/NOₓ/NH₃ on (VOC/PM off), 2020 base year, paired with 2019 weather data.

## Project structure

```
app.py                     # Streamlit app: UI, scenario/species combinatorics, workbook editing, zip export
管控方案模板.xlsx            # Breeze scenario output template (management plan format)
data/
  2017_emission_report.xlsx  # Base-year emissions, by sector x province, per pollutant sheet
  2020_emission_report.xlsx  # Base-year emissions, by sector x province, per pollutant sheet
  {Scenario}_All_Years/       # One folder per DPEC scenario (Baseline, CleanAir, OTPCA, OTPNZCA, EPNZCA)
    {year}_Emission.xlsx      # Target-year emissions for that scenario, per pollutant sheet
requirements.txt
```

## Data notes

- Base-year workbooks contain one sheet per pollutant (`SO2`, `NOx`, `CO`, `VOC`, `NH3`, `PM10`, `PM25`, `BC`, `OC`), each a sector × province grid of annual emissions.
- Target-year workbooks (under `data/{Scenario}_All_Years/`) follow the same per-pollutant sheet layout, with an additional `category`/`month` breakdown (the app uses the `Annual total` rows).
- Output workbooks follow the Breeze management-plan template (`管控方案`) format: region, sector, species, and reduction factor (%), with region/sector names translated to Chinese via lookup maps in `app.py`.

## Notes / limitations

- This is a small internal utility rather than a general-purpose tool — file paths (source data, template) are hardcoded relative to the working directory, and there's minimal input validation beyond requiring at least one selection in each category before enabling generation.
- All generated workbooks are held in memory (`st.session_state`) before zipping, so very large batches (many scenarios × many species × many weather years) will use more memory and take longer to generate.
