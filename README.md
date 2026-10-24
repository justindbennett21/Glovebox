# Glovebox Car Log: setup guide

A shared Streamlit app for logging gas fill-ups, car service and receipts.
- Entries go to a **Google Sheet you own** (tabs Cars, Fuel, Service, created automatically).
- Receipts go to a **"Glovebox Receipts" folder in your Google Drive**.

Plan on about 45 minutes the first time.

## Files
| File | What it is |
|---|---|
| `app.py` | The app |
| `requirements.txt` | Packages Streamlit installs |
| `receipts_apps_script.gs` | Small helper you paste into Google Apps Script |
| `secrets.toml.example` | Template for passwords and keys |
| `.gitignore` | Keeps secrets and key files out of GitHub |

---

## 1. Create the Google Sheet
1. Make a new blank Google Sheet, e.g. "Glovebox Car Log".
2. Copy its ID from the URL: `docs.google.com/spreadsheets/d/`**`THIS_PART`**`/edit`.

## 2. Give the app a Google "robot" account (service account)
1. Go to https://console.cloud.google.com and create a project (e.g. "glovebox").
2. **APIs & Services > Library**: enable **Google Sheets API** and **Google Drive API**.
3. **APIs & Services > Credentials > Create credentials > Service account**. Name it, skip optional steps.
4. Open the service account > **Keys > Add key > Create new key > JSON**. A file downloads. Keep it private.
5. Copy the service account's email (ends in `iam.gserviceaccount.com`).
6. In your Google Sheet, click **Share**, paste that email, give it **Editor**.

## 3. Set up the receipts helper (Apps Script)
1. Go to https://script.google.com > **New project**. Name it "Glovebox Receipts".
2. Delete the sample code, paste in all of `receipts_apps_script.gs`.
3. Change `TOKEN` to a long random phrase (e.g. `blue-outback-maple-7731-garage`). Save.
4. Pick `setup` in the function dropdown and click **Run**. Approve the permissions.
   - If you see "Google hasn't verified this app": click **Advanced > Go to Glovebox Receipts**. This is normal for your own scripts.
5. **Deploy > New deployment**, gear icon > **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Click **Deploy** and copy the **Web app URL** (ends in `/exec`).

> "Anyone" means anyone with the URL *and* your token can upload. Keep both private.
> If you edit the script later, use **Deploy > Manage deployments > Edit > New version** so the URL stays the same.

## 4. Share with your wife
- Share the **Google Sheet** with her Gmail (Viewer or Editor).
- Share the **Glovebox Receipts** folder in your Drive with her Gmail (Viewer), so receipt links open for her.

## 5. Test on your computer (optional, VS Code)
```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
mkdir .streamlit                 # then copy secrets.toml.example to .streamlit/secrets.toml and fill it in
streamlit run app.py
```

## 6. Put the code on GitHub
Add `app.py`, `requirements.txt`, `.gitignore`, and optionally `README.md` and `receipts_apps_script.gs`.
Never upload the JSON key file or a filled-in `secrets.toml`.

## 7. Deploy on Streamlit Community Cloud
1. https://share.streamlit.io > **Create app** > pick your repo, main file `app.py`.
2. **Advanced settings > Secrets**: paste your filled-in secrets (same as `.streamlit/secrets.toml`).
3. Deploy. Open the URL on both phones and **Add to Home Screen**.

---

## Using it
- **Fill-up:** date, mileage, total cost, gallons. Forgot one? Change the date to the day you filled up.
- **Service:** pick the service, add cost, shop, notes, and snap a photo of the receipt.
- **Timeline:** everything done to the car, newest first, with receipts. Filter by type or year. The top table shows when each service was last done and miles since.
- **Schedule:** each car's maintenance schedule compared to its odometer. Pick a template the first time (Subaru 2015-2016, Toyota RAV4 Hybrid 2019-2025, or a general one), then edit any interval. Items show 🔴 overdue, 🟠 due soon (within 1,000 mi or 30 days), 🟢 OK, or ⚪ no record. To start tracking an item, log the last time it was done in the Service tab.
- **Trends:** MPG, price per gallon, monthly spending.
- **Fix:** correct a typo, add a receipt to an older service, or delete an entry.

## How MPG works (and missed fill-ups)
MPG = miles since your last **full** tank ÷ gallons bought since then.
- **Partial fill?** Uncheck "Filled the tank all the way." Those gallons roll into the next full fill-up.
- **Forgot to log one but have the receipt?** Log it with the real date. It slots into place.
- **Missed one completely?** On the next fill-up, check "I missed logging a fill-up before this one." That gap is left out of MPG so it doesn't look falsely high. (Your total gas spending will be short by that one fill-up.)

## Troubleshooting
- `SpreadsheetNotFound` / `PERMISSION_DENIED`: the sheet isn't shared with the service account email, or `sheet_id` is wrong.
- Receipt upload fails with "Not allowed": `receipts_token` doesn't match `TOKEN` in the script.
- Receipt upload fails with an HTML/JSON error: the web app wasn't deployed with access "Anyone", or the URL isn't the `/exec` one.
- Wife can't open a receipt: share the Glovebox Receipts folder with her.
