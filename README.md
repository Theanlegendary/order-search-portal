# Order Search Portal

A clean web portal for support teams to search orders, track delayed/missing bills, and export Excel reports.

## Features
- 🔍 Full-text search across every order field (phone, name, order ID, note, post office)
- 📦 Live 14-day data with 5-minute smart caching
- ⏰ Delayed (>1 day), 🚨 Missing (>7 days), ⚠️ Issue status filters
- ⬇️ Export any view to Excel instantly
- 🔐 Optional password protection

## Deploy to Vercel

[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new)

### Environment Variables (set in Vercel Dashboard)

| Variable | Description | Example |
|---|---|---|
| `API_URL` | API endpoint for order export | `https://gw-express.metfone.com.kh/...` |
| `API_BEARER` | Bearer token for API auth | `eyJhbGci...` |
| `API_BRANCH` | Branch codes to include | `MEGA,PRE,PNP,SVA,KAN` |
| `API_CLIENT_ID` | Client ID header | `TMS_ANDROID` |
| `API_REFERER` | Referer header | `https://opsexpress.metfone.com.kh/` |
| `SEARCH_PASSWORD` | Portal access password (optional) | `mypassword123` |
| `CACHE_TTL_SEC` | Cache duration in seconds (default 300) | `300` |

## Local Development

```bash
pip install -r requirements.txt
python api/index.py
# Open http://localhost:5055
```
