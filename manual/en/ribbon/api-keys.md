# Manage API Keys

**Location**: ribbon `EGTools` tab → `EGTools` group → split-button dropdown menu

A dialog to register or remove the API keys used by the public-API functions and
ribbon commands.

The list contains **three** services.

| Service | Purpose | Related features |
|---|---|---|
| juso | Korean road-address search | [SEARCHADDRESS](../SEARCHADDRESS.md) |
| data.go.kr | Public holidays and business-registration status | [KOREANHOLIDAYS](../KOREANHOLIDAYS.md) · [Calendar](calendar.md#holidays) · [BRNSTATUS](../BRNSTATUS.md) |
| vworld | Maps / coordinates | [GEOSEARCH](../GEOSEARCH.md) · [GEOCODER](../GEOCODER.md) · [GEOADDRESS](../GEOADDRESS.md) · [ADDRESSMAP](../ADDRESSMAP.md) |

- Selecting a service shows only **whether a key is registered, masked (●●●●●●)**.
  The stored key is never displayed again.
- Keys are **stored in your user account** and are
  **never written into workbooks or formulas**.
- This coexists with passing the key as the **last argument of a function** — a key
  passed that way is saved to the same store automatically, so it can be omitted
  afterwards.
- **No keys ship with EGTools++.** You must obtain your own key from each service
  (juso.go.kr, data.go.kr, vworld.kr).

## One key for all of data.go.kr

The Korean open-data portal (data.go.kr) issues **a single key per account** and
the same key is used for every service you have applied for. EGTools++ therefore
keeps one `data.go.kr` slot instead of one slot per service.

- That single key is **shared** by [KOREANHOLIDAYS](../KOREANHOLIDAYS.md)
  (holiday information), the temporary-holiday merge in
  [Calendar](calendar.md#holidays), and [BRNSTATUS](../BRNSTATUS.md)
  (NTS business-registration status).
- A key registered **separately for business-registration lookups** by an older
  version is **migrated automatically** to the shared slot on first use — nothing
  to re-enter or clean up.
- You may paste either the Encoding or the Decoding form of the key shown on the
  portal; it is percent-encoded internally when required.
- When a lookup fails in a way that identifies the key as rejected, the **stored
  key is deleted automatically** and a reissue guide is returned. Issue a new key
  and register it again.
