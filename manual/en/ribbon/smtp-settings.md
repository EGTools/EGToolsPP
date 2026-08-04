# SMTP Settings

**Location**: ribbon `EGTools++` tab → `EGTools` group → split-button dropdown menu

A dialog to register the SMTP server used by [Mail Merge](mailmerge.md) for
**sending email** (options 4, 5, 6, 7).

## Fields

| Field | Description |
|---|---|
| Sender name | Display name shown to recipients (optional — if empty, only the address is shown) |
| Sender email | The From address, which is also the **SMTP login account** |
| Outgoing mail server | SMTP server address (e.g. `smtp.gmail.com`) |
| Server port | Default **465** |
| SSL | Whether to use SSL (on by default) |

## Rules

- Settings are stored per user account (in the registry) and are
  **never written into workbooks**.
- **The password is never stored.** It is asked every time you send and used in
  memory only.
- If you choose a sending option in Mail Merge while the settings are empty
  (email, server or port missing), this dialog **opens automatically**. If it is
  still incomplete afterwards, sending is aborted.

## Important limitation — implicit SSL only

The sending engine (Windows CDO) supports **implicit SSL only**, which usually
means **port 465**. **STARTTLS on port 587 is not supported** — entering 587
will make sending fail.

- Gmail: server `smtp.gmail.com`, port `465`, SSL on — use an
  **app password** instead of your account password.
- For other providers use their "SSL (465)" style port as well; accounts with
  two-factor authentication may also require an app password.

## Notes

- Non-ASCII subjects/bodies (e.g. Korean) are sent encoded as UTF-8.
- When sending fails, the error description reported by the server is shown
  verbatim in the failure list of the [Mail Merge](mailmerge.md) completion
  summary.
