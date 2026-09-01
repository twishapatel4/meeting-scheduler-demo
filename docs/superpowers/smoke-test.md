# Manual Smoke Test — Teams Scheduler Demo

Prerequisites: backend running on :5000, frontend running on :5173, Azure App
Registration created with redirect URI `http://localhost:5000/auth/staff/callback`
and delegated scopes `offline_access Calendars.ReadWrite OnlineMeetings.ReadWrite User.Read`.
`.env` populated with real `MS_CLIENT_ID`/`MS_CLIENT_SECRET`.

1. **Connect staff**: visit `http://localhost:5173/admin`, click "Connect Staff
   Account", complete Microsoft OAuth consent with a real account. Expect
   redirect back to `/admin?connect=success` and the account listed as
   "Connected".
2. **Visitor request**: visit `http://localhost:5173/`, submit a booking
   request with a start/end time in the future. Expect it to appear on
   `/admin` with status "Requested".
3. **Schedule**: on `/admin`, click "Schedule w/ <email>". Expect status to
   become "Scheduled" and a "Join" link to appear. Open the link, confirm it
   is a valid Teams meeting join URL. Check the staff member's real Outlook
   calendar — the event should appear with the visitor as an attendee.
4. **Reschedule**: click "Reschedule", enter a new ISO 8601 start/end. Expect
   status "Rescheduled" and the Outlook event's time to have moved. Confirm
   the visitor would receive Graph's automatic notification (Graph handles
   this — no separate email step in this app).
5. **Cancel**: on a different booking (repeat steps 2-3 to create one), click
   "Cancel". Expect status "Cancelled" and the event removed from the staff
   member's Outlook calendar.
6. **Swap host**: connect a second staff account in the *same organization*
   (same email domain) via step 1. Create and schedule a third booking, then
   click "Swap to <second email>". Expect status "Swapped", the original
   staff member's event cancelled, a new event on the second staff member's
   calendar, and a new join URL returned.
7. **Cross-org swap rejection**: connect a third staff account with a
   *different* email domain. Attempt a swap to that staff member via a raw
   API call:
   `curl -X POST http://localhost:5000/api/v1/bookings/<id>/swap -H "Content-Type: application/json" -d '{"newStaffId":"<cross-org-staff-id>"}'`
   Expect a 400 response: "Swap-host is only allowed between staff in the
   same organization".
8. **Disconnected staff handling**: in Azure Portal, revoke the app's
   consent for one connected staff account (Enterprise Apps → the app →
   Users and groups → remove user, or have the user revoke via
   myaccount.microsoft.com). Attempt to schedule a booking with that staff
   member. Expect a 401 and the staff's `connected` flag to flip to `false`
   on the next `GET /api/v1/staff` call.

Record pass/fail for each step; any failure blocks calling the demo done.
