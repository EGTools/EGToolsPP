# Manage API Keys

**Location**: ribbon `EGTools++` tab → `EGTools` group → split-button dropdown menu

A dialog to register or remove per-service API keys used by the public-API functions.

| Service | Purpose | Related functions |
|---|---|---|
| juso | Korean road-address search | [SEARCHADDRESS](../SEARCHADDRESS.md) |
| odcloud | Business-registration status | [BRNSTATUS](../BRNSTATUS.md) |
| vworld | Maps / coordinates | [GEOSEARCH](../GEOSEARCH.md) · [GEOCODER](../GEOCODER.md) · [GEOADDRESS](../GEOADDRESS.md) · [ADDRESSMAP](../ADDRESSMAP.md) |
| data.go.kr | Korean public holidays | [KOREANHOLIDAYS](../KOREANHOLIDAYS.md) · [Calendar](calendar.md#holidays) (for temporary-holiday lookups) |

- Selecting a service shows only **whether a key is registered, masked (●●●●●●)**.
  The stored key is never displayed again.
- Keys are stored in the `HKCU\Software\EGTools\ApiKeys` registry key and are
  **never written into workbooks or formulas**.
- This coexists with passing the key as the **last argument of a function** — a key
  passed that way is saved to the same store automatically, so it can be omitted
  afterwards.
- **No keys ship with EGTools++.** You must obtain your own key from each service
  (juso.go.kr, odcloud/Korean open-data portal, vworld.kr, data.go.kr).
