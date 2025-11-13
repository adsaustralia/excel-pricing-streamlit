
# Excel SQM & Pricing Calculator (Streamlit)

This app lets you upload Excel campaign sheets, apply pricing rules per material, 
calculate square metres (SQM) and total price in AUD, and download an updated workbook.

## Features

- Upload `.xlsx` file
- Configure:
  - Start/End columns (e.g. AC → IG)
  - Row for Size, Material, Quantity, SQM output, Price output
- Per-material pricing rules:
  - Jellyfish
  - Ferrous
  - Artboard / B-Flute / 400gsm
  - Synthetic Banner (SS)
  - Synthetic Banner (DS)
  - Backlit Shimmer
- Process:
  - A single selected sheet, **or**
  - All sheets in the workbook
- Shows preview table of calculations
- Writes:
  - SQM to configured SQM row
  - Price to configured Price row
  - Total per sheet into the last column's SQM/Price cells
- Combined summary with total per sheet + grand total
- Download:
  - Updated workbook (`Updated_Pricing_All_Sheets.xlsx`)
  - Summary-only workbook (`Pricing_Summary.xlsx`)

## Local Setup

1. Create a virtual environment (optional but recommended)
2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Run the app:

   ```bash
   streamlit run app.py
   ```

4. Open the URL shown in the terminal (usually `http://localhost:8501`).

## Deploy to Streamlit Cloud

1. Push this folder to a **public GitHub repo**.
2. Go to Streamlit Community Cloud and click **New app**.
3. Select your repo and `app.py` as the entrypoint.
4. Click **Deploy** — your app will build and get a public URL.

## Usage Notes

- Sizes in Row "Size" should contain width × height (usually in mm), e.g. `1200 x 1800mm`.
- Quantity row can contain numbers or text like "3 kinds" — the first number is used.
- Material names are matched loosely (e.g. anything containing "400gsm" will be treated with that rate).
