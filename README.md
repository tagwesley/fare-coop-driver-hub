# Fare Co-op Driver Hub (Google Apps Script)

> **Portfolio note.** A driver-facing web app and automation layer I built in 2026 (with AI-assisted development) for the San Luis Obispo chapter of Fare Co-op, a driver-owned rideshare cooperative. It sits on a Google Sheets operations hub. Spreadsheet IDs, deployment URLs, sender addresses and collaborator names have been replaced with placeholders. Credentials live in Script Properties and were never in source.

**What it shows:** a sheet-backed data model with schema validation, passwordless email-code auth with hashed session tokens, row-level access (each driver sees only their own records), approval-gated email queues, and scheduled triggers.

---


This is the finished driver-facing web app for the existing **Fare Co-op — Google Operations Hub** spreadsheet.

It connects to spreadsheet ID `YOUR_SPREADSHEET_ID` and gives each driver access only after a one-time sign-in code is sent to the email in the `Driver Index` tab.

## Production status

- **Live driver portal:** https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec
- **Apps Script project:** https://script.google.com/home/projects/YOUR_SCRIPT_PROJECT_ID/edit
- **Deployment ID:** `YOUR_DEPLOYMENT_ID`
- **Schema validation:** passed with all connected driver rows and no missing required fields.
- **Automation:** the daily draft-queue trigger and both 15-minute email/Calendar triggers are installed.
- **Access:** deployed as an `Anyone` web app that executes as the owner; driver data remains behind the email-code gate.

The live sign-in page has been verified at the production URL. No test OTP was sent and no driver, dispatch, task, or email-queue records were changed during the production smoke test.

## What drivers can do

- Sign in without a password using an emailed six-digit code.
- See their Fare Co-op onboarding stage, checklist, and next required step.
- Mark themselves available or unavailable in `Driver Actions`.
- View only their own rows from `Dispatch`.
- Accept an assignment, mark en route, complete it, or report that they cannot take it.
- See upcoming events in their market and express interest.
- Submit support requests into the `Tasks` tab.
- Open Fare Co-op membership, tutorials, driver portal, and support resources.

## Included automation

- `generateDriverNextStepQueue()` creates approval-gated driver email drafts and allows recurring reminders after the driver's configured cadence. Pending drafts are de-duplicated; sent history no longer blocks future reminders forever.
- `sendApprovedEmailQueue()` sends only rows marked `Ready` and `Approved? = Yes`. All onboarding reminders are delivered through the Brevo transactional API from `chapter-sender@example.org`, while each row preserves its Tag or Partner referral attribution. The sender fails closed when the Brevo credential is missing or invalid.
- `syncTasksToCalendar()` creates Calendar events once and writes the event ID back to the task.
- `doPost()` logs authorized Fare webhook payloads to `Fare Sync`. It intentionally does not write to protected Fare driver fields.
- `installAutomationTriggers()` creates the daily/15-minute triggers without duplicates.

## Restore or reinstall in the existing Operations Hub

1. Open the **Fare Co-op — Google Operations Hub** spreadsheet.
2. Choose **Extensions → Apps Script**.
3. Rename the Apps Script project to `Fare Co-op Driver Hub`.
4. Create the project files below and paste in the matching files from this package:
   - `Code.gs`
   - `Auth.gs`
   - `Data.gs`
   - `Automation.gs`
   - `Index.html`
   - `Styles.html`
   - `Client.html`
5. Open **Project Settings**, enable **Show appsscript.json manifest file in editor**, and replace it with this package's `appsscript.json`.
6. Save the project.
7. From the function selector, run `validateHubSchema`. The execution result should show `ok: true` and approximately all connected driver rows.
8. Run `installAutomationTriggers` once. Approve Sheets, email, and Calendar access using the Google account that should own the portal automation.

## Recreate the deployment

1. In Apps Script, choose **Deploy → New deployment**.
2. Select **Web app**.
3. Description: `Fare Co-op Driver Hub v1`.
4. **Execute as:** `Me`.
5. **Who has access:** `Anyone` is recommended because the app has its own email-code gate and the driver roster includes non-Gmail addresses. If your Google Workspace administrator blocks this option, use `Anyone with Google account` until public web apps are approved.
6. Click **Deploy**, approve the requested access, and copy the `/exec` URL.

Do not deploy as **User accessing the web app**. That setting would ask each driver to authorize access to the underlying Google data.

## Smoke test before sharing

1. Open the `/exec` URL in a private browser window.
2. Enter an email that exists in `Driver Index`.
3. Confirm the one-time code arrives and expires after 10 minutes.
4. Sign in and check that the displayed name, driver ID, market, stage, and checklist belong to that one driver.
5. Change availability to **Yes**, then verify the corresponding `Driver Actions` row.
6. Create a harmless test assignment for that driver in `Dispatch`, refresh the portal, accept it, and verify `Assignment Status` and `Confirmed At`.
7. Submit a test support request and verify the new `Tasks` row.
8. Delete or clearly label the harmless test rows after testing.

## Safe operating model

- `Driver Index` stays read-only and is still sourced from the existing Fare Co-op trackers.
- Driver login is verified through the email already present in `Driver Index`.
- Login codes are hashed, rate-limited, expire after 10 minutes, and allow no more than five attempts.
- Session tokens expire after six hours and reveal no driver information by themselves.
- Each dispatch update rechecks that the assignment's `Driver ID` matches the signed-in driver.
- Email sending remains approval-gated in review-first mode. Each queue row records `Referral Owner`, `Sender Email`, and `Delivery Provider`; all onboarding messages use Brevo from `chapter-sender@example.org` and never fall back to the Apps Script owner's personal address.
- The Fare webhook remains logging-only until Fare supplies an approved endpoint, signing/authentication method, schema, and permitted write-back fields.
- API credentials and webhook secrets belong in **Apps Script → Project Settings → Script Properties**, never in spreadsheet cells or source files.


## Fare Co-op sender routing

The reminder queue is owner-aware and provider-aware:

- Shared outbound sender: `chapter-sender@example.org` through the Brevo transactional API for both Tag and Partner referrals. The referral owner and referral code remain in each personalized message. The default per-run safety limit is 75 messages; override with `REMINDER_RUN_LIMIT`.
- Use `Fare Co-op Ops → Check reminder senders` before activation. It validates the Brevo API credential.
- Use `configureReviewFirstReminderAutomation()` to install the daily queue generator and 15-minute sender trigger while keeping new reminders at Draft/No until a reviewer changes them to Ready/Yes.
- Use `Fare Co-op Ops → Activate driver reminder automation` only if fully automatic approval is intentionally authorized. Activation marks current driver reminders Ready/Yes and installs the sender trigger.
- Use `Fare Co-op Ops → Pause driver reminder automation` to remove that user's sender trigger and stop new reminders from auto-entering Ready/Yes.
- Brevo sender/domain verification still has to be completed inside Brevo. The API cannot send from an unregistered sender.

## Optional Fare webhook setup

After Fare technical staff approve an integration contract:

1. Add a strong `FARE_WEBHOOK_SECRET` in Script Properties.
2. Redeploy the web app.
3. Give Fare the deployment URL plus the secret through an approved secure channel.
4. Send test payloads to `WEB_APP_URL?token=SECRET`.
5. Verify each test appears in `Fare Sync`.
6. Keep write-back disabled until Fare approves exact field mapping and conflict rules.

## Updating the live app

After changing code, use **Deploy → Manage deployments → Edit**, choose **New version**, and deploy. The `/exec` URL stays the same.
