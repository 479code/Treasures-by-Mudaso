# Automatic deployments

Every push to `master` updates the live services. Check results in the **Actions** tab.

## Google Apps Script (`deploy-apps-script.yml`)
Runs when `google-apps-script.gs` or `config.js` changes.

| Secret | Required | Value |
| --- | --- | --- |
| `CLASPRC_JSON` | Yes | Output of `Get-Content "$HOME\.clasprc.json" -Raw` after `clasp login` |
| `APPS_SCRIPT_ID` | No | Defaults to the Treasures script ID in the workflow |
| `APPS_SCRIPT_DEPLOYMENT_ID` | No | Defaults to the ID inside the `/exec` link in `config.js` |

It pushes the code, creates a new version and moves the existing web app to it. The `/exec` link never changes.

## Cloudflare Pages (`deploy-cloudflare.yml`)
Runs when website files change. It skips until both secrets exist.

| Secret | Value |
| --- | --- |
| `CLOUDFLARE_API_TOKEN` | Cloudflare API token with Pages edit permission for `treasures-by-mudaso` |
| `CLOUDFLARE_ACCOUNT_ID` | Cloudflare account ID |

Never put secret values in code, commits or chat messages.
