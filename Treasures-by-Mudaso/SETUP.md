# Treasures by Mudaso — setup (new Google Sheet)

1. **Sheet → Extensions → Apps Script.** Delete the starter code, paste all of `google-apps-script.gs`, save.
2. **Reload the Sheet.** Menu **Treasures → Set up sheets** → approve access. It creates the tabs, the 3 official packages and your owner key.
3. **Apps Script → Deploy → New deployment → Web app.** Execute as: *Me*. Who has access: *Anyone*. Copy the URL ending in `/exec`.
4. **Edit `config.js`** in this repo: paste the `/exec` URL and the Sheet URL. Commit and push.
5. **Open the app → `#owner`** → set your PIN → **More** → paste the owner key (**Treasures → Show owner key**) → Save.
6. **More → Dispatch & payment:** save fee, bank details and WhatsApp once.
7. **Catalogue:** add products with **cost price**, and add a cost price to each package. Profit reports use these.

After changing the Apps Script later: **Deploy → Manage deployments → Edit → Version: New version** (the URL stays the same).

## Automatic Apps Script updates (GitHub Actions)

Once set up, every push that changes `google-apps-script.gs` updates the Apps Script and the live web app. The `/exec` link stays the same.

- Secret `CLASPRC_JSON` → contents of `%USERPROFILE%\.clasprc.json` after `clasp login`.
- Script ID is in `apps-script/.clasp.json`.
- The live deployment ID is read from the `/exec` link in `config.js`. You can override it with a repository variable `APPS_SCRIPT_DEPLOYMENT_ID`.
- Run it by hand: GitHub → Actions → Deploy Apps Script → Run workflow.
