# Client Follow-Up Reminder System

Never lose track of client relationships again. Automated follow-up reminders built on Google Sheets + Gmail — no CRM subscription required.

**Demo:** <https://youtu.be/placeholder> (coming soon)  
**Try it:** [Make a copy](https://docs.google.com/spreadsheets/d/1placeholder/copy) (coming soon)

## What it does

- **Tracks clients** — Maintain a master list with contact info, priority levels, and follow-up schedules
- **Logs interactions** — Record every email, call, meeting with automatic date stamping
- **Calculates reminders** — Auto-generates daily follow-up list sorted by priority and due date
- **Sends email digests** — Morning summary of overdue and due follow-ups delivered to your inbox
- **Supports snooze** — Postpone reminders without losing track of the client
- **Integrates with Tool #2** — Auto-import emails logged by Email-to-Spreadsheet Logger

## Why use this

- **Free** — No CRM subscription, no limits on contacts
- **Private** — Your client data stays in your Google account
- **Flexible** — Define your own follow-up intervals per client priority
- **Automated** — Set it once, get daily reminders forever
- **Gmail-first** — Integrates with your existing email workflow

## Quick Start (3 steps)

### 1. Copy the template

[Make a copy of this spreadsheet](https://docs.google.com/spreadsheets/d/1placeholder/copy)

### 2. Set up the script

1. In your copied spreadsheet, click **Extensions → Apps Script**
2. Delete the default `Code.gs` file
3. Click **+** next to **Files** → paste the contents of `Code.gs` from this repo
4. Click **Save** (floppy disk icon)
5. Click **Run** → select `setupTemplate`
6. Authorize the script when prompted (click through permissions)

### 3. Configure and run

1. Switch to the **"Settings"** tab and update the Email Recipient field with your email
2. Edit **"Clients"** tab with your client list (ClientID, Name, Email, Priority)
3. Click **📧 Follow-Ups → 🔄 Calculate Reminders** to generate your first reminder list
4. Click **📧 Follow-Ups → 📨 Send Daily Digest Now** to test email delivery
5. Check your inbox — your follow-up digest should appear!

For automatic daily reminders, set up triggers (see below).

## Sheet Structure

| Sheet | Purpose |
|-------|---------|
| **Settings** | Follow-up intervals, snooze options, email recipient |
| **Clients** | Master client list with priority and next follow-up dates |
| **Interactions** | Log of all client touchpoints (emails, calls, meetings) |
| **Reminders** | Auto-generated follow-up list with priority scoring |

## Sample Clients

| ClientID | Name | Email | Priority | DefaultInterval |
|----------|------|-------|----------|-----------------|
| ACME | Acme Corporation | contact@acme.com | VIP | 7 |
| GLOB | Globex Industries | sales@globex.com | Standard | 14 |
| STARK | Stark Labs | hello@stark.io | Low | 30 |

## Sample Interactions

| Date | ClientID | Type | Summary | NextFollowUp |
|------|----------|------|---------|--------------|
| 2026-02-15 | ACME | Email | Sent proposal | 2026-02-22 |
| 2026-02-10 | GLOB | Call | Discussed Q2 needs | 2026-02-24 |
| 2026-02-01 | STARK | Meeting | Initial consult | 2026-03-01 |

## Setting Up Automatic Reminders

For hands-free operation, add time-driven triggers:

1. In Apps Script, click **Triggers** (⏰ icon in left sidebar)
2. Click **+ Add Trigger** (bottom right)
3. Add first trigger:
   - Function: `calculateReminders`
   - Event source: Time-driven
   - Type: Day timer
   - Time: 8:00 AM to 9:00 AM
4. Add second trigger:
   - Function: `sendDailyDigest`
   - Event source: Time-driven
   - Type: Day timer
   - Time: 8:05 AM to 9:00 AM

Now you'll receive a prioritized follow-up email every morning.

## Features

- ✅ Automatic follow-up scheduling based on client priority
- ✅ Daily email digest with overdue and due items
- ✅ Priority scoring: VIP (⚡), Standard, Low
- ✅ One-click snooze (7, 14, 30 days)
- ✅ Color-coded status in Reminders sheet
- ✅ Interaction history per client
- ✅ Integration with Email-to-Spreadsheet Logger (Tool #2)
- ✅ Customizable follow-up intervals
- ✅ One-click template setup

## Integration with Tool #2 (Email Logger)

If you use the [Email-to-Spreadsheet Logger](../email-to-spreadsheet/), you can auto-import emails as interactions:

1. Ensure Tool #2's `Log` sheet exists in the same spreadsheet (copy both templates into one file)
2. Click **📧 Follow-Ups → 🔄 Sync from Email Logger**
3. Emails matching your clients will be added to Interactions automatically

**Pro tip:** Set a periodic trigger for `syncFromTool2()` to keep interactions updated throughout the day.

## Daily Workflow

1. **Morning** — Check your inbox for the daily digest
2. **Review** — Open Reminders sheet for color-coded status view
3. **Act** — Contact VIPs and overdue clients first
4. **Log** — Add interactions via **📧 Follow-Ups → ➕ Log Interaction**
5. **Snooze** — Postpone non-urgent items as needed

## FAQ

**Q: Can I customize the follow-up intervals?**  
A: Yes — edit the "Default Interval (days)" column in the Clients sheet, or change global defaults in Settings.

**Q: How do I mark a client as contacted?**  
A: Click **📧 Follow-Ups → ➕ Log Interaction**, enter the ClientID, type, and notes. The next follow-up date recalculates automatically.

**Q: What happens if I snooze a reminder?**  
A: The client's next follow-up date is pushed forward by the snooze days (default 7). The reminder disappears until that date.

**Q: Can I have different intervals for different client priorities?**  
A: Yes — VIP clients can have 7-day intervals, Standard 14-day, Low 30-day. Set per client or use global defaults.

**Q: Will this send emails to my clients?**  
A: No. It only sends a daily digest to *you*. Client communication is manual — you decide how to reach out.

**Q: Is there a limit on clients?**  
A: Google Sheets supports up to 10 million cells. Practical limit is thousands of clients — more than enough for freelancers and small businesses.

## Troubleshooting

**"Authorization required" keeps appearing**  
- This is normal for the first run. Click through all authorization prompts. The script needs access to spreadsheets and your email.

**No reminders appearing**  
- Check that your Clients have priority levels set
- Run **Calculate Reminders** manually from the menu
- Verify Interactions have valid ClientIDs matching your Clients sheet

**Daily digest not arriving**  
- Check your spam folder
- Verify the Email Recipient in Settings is correct
- Run `testSendDigest()` from Apps Script to test email delivery
- Check **View → Executions** in Apps Script for errors

**"Exceeded maximum execution time"**  
- This is rare for typical client lists. If it happens, reduce the number of interactions processed at once.

## Roadmap

- [ ] One-click email compose (opens Gmail with client email pre-filled)
- [ ] Slack/Discord notification integration
- [ ] Bulk import clients from CSV
- [ ] Follow-up analytics dashboard
- [ ] Recurring reminder templates

## License

MIT — See [LICENSE](./LICENSE)

## Credits

Built by [PressedCoffee](https://github.com/PressedCoffee) as part of the Automation Tool Loop.

---

**Found this useful?** Star ⭐ the repo and [share what you're building](mailto:Shaddock.Mercer@gmail.com).