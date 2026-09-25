# Automatic deployments

Once the repository secrets below are added, every push to `master` updates the right live service automatically.

## Cloudflare Pages

Add these two GitHub Actions secrets at **GitHub repository → Settings → Secrets and variables → Actions → New repository secret**:

| Secret | What to paste |
| --- | --- |
| `CLOUDFLARE_API_TOKEN` | A Cloudflare API token restricted to the **Treasures by Mudaso** Pages project with Pages edit permission. |
| `CLOUDFLARE_ACCOUNT_ID` | The Cloudflare account ID that owns the Pages project. |

The workflow builds a clean public folder directly from the committed app source, then deploys it to the existing `treasures-by-mudaso` Pages project. It does not create another site or change the public address. You no longer need to update `dist` for GitHub deployments.

## Google Apps Script

First enable the Apps Script API at https://script.google.com/home/usersettings. Then add these GitHub Actions secrets:

| Secret | What to paste |
| --- | --- |
| `CLASPRC_JSON` | The entire contents of `~/.clasprc.json` created by `clasp login`. Keep this private: it contains a Google refresh token. |
| `APPS_SCRIPT_ID` | The Script ID shown in Apps Script **Project Settings**. |
| `APPS_SCRIPT_DEPLOYMENT_ID` | The existing web-app deployment ID. In Apps Script, open **Deploy → Manage deployments**, open the web app, then copy the Deployment ID. |

After those three secrets exist, changes to `google-apps-script.gs` push to the existing Apps Script project, create a version, and redeploy the existing web-app deployment. Its `/exec` URL stays the same.

## Daily workflow

1. Change the app or Apps Script source.
2. Commit and push to `master`.
3. GitHub Actions deploys the change automatically. Check the **Actions** tab for the result.

Do not put any of the secret values in code, a commit, or a chat message.
