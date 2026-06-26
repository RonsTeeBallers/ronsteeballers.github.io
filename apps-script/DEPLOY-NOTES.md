# Deploying the pairings-email change

Two files changed:
- `apps-script/WebApps.gs`  → goes into the Google Apps Script editor (backend)
- `pairing.html`            → goes into the GitHub repo (front end)

---

## 1. Backend — update Apps Script  (the important part)

1. Open the Golf Group Sheet → **Extensions ▸ Apps Script**.
2. Click **WebApps.gs** in the left file list.
3. Select all (Cmd+A), delete, and paste in the full new contents of
   `apps-script/WebApps.gs`.
4. Click **Save** (disk icon).

### Test before going live
5. In the editor's function dropdown (top toolbar), choose **testBrevoSend**
   and click **Run**. The first run asks for permissions — approve them.
6. Check your **ronsteeballers@gmail.com** inbox for a "TEST — Golf Pairings"
   email, and the **Execution log** should show `{"ok":true}`.
   - If it shows an error about `BREVO_API_KEY`, the key isn't set. Add it
     under **Project Settings ▸ Script Properties** (same key the invites use).

### RE-DEPLOY  ← easy to forget
Editing the code is NOT enough. The live web URL only updates when you
publish a new version:

7. **Deploy ▸ Manage deployments**.
8. Click the pencil (**Edit**) on the existing deployment.
9. **Version** dropdown → **New version** → **Deploy**.

The web app URL stays the same, so nothing on the GitHub side needs to change
for the backend.

---

## 2. Front end — update pairing.html on GitHub

The only change is the confirmation message after sending. Update the copy of
`pairing.html` in the GitHub repo with the new local version.

(See chat — Claude can set up direct `git push` so this becomes automatic, or
you can paste the new file into the GitHub web editor.)

---

## What changed and why

- `savePairings()` used to end with `GmailApp.createDraft(...)`, leaving a
  draft you had to send by hand. It now looks up each foursome player's email
  on the **Players** tab and sends the pairings to them directly through Brevo.
- New shared helper `sendBrevoEmail_()` does the actual Brevo send.
- `testBrevoSend()` is a manual test that emails only you.
- The invite send (`sendInviteEmails`) was already using Brevo — unchanged,
  except a misleading "Resend" error message was corrected to "Brevo".
- Legacy v1 menu functions (Thursday/Monday invites & pairings drafts) have
  now been **removed** — v2 is the only path. The following files were deleted
  from the repo and should also be deleted in the Apps Script editor:
  `Menu.gs`, `ThursdayInvite.gs`, `MondayInvite.gs`, `Pairings.gs`,
  `SyncForms.gs`, `SyncContacts.gs`, `ThursdayForm.html`, `MondayForm.html`.
  The `appsscript.json` OAuth scopes were trimmed to just `spreadsheets` and
  `script.external_request` (the v1-only `gmail.compose`, `forms`, `contacts`,
  and `script.container.ui` scopes are gone), which triggers a one-time
  re-authorization on the next run/deploy.
