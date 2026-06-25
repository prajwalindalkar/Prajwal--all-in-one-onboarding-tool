# PhonePe Merchant Onboarding Tool v2

Generates official PhonePe onboarding documents for all three merchant types:

| Merchant Type | Documents Generated |
|---|---|
| Sole Proprietor | Merchant Declaration Form (MDF) |
| Partnership Firm | Partner Resolution + Partnership BO Declaration |
| Private Limited | Company BO Declaration |

## Deploy on Render.com
1. Push to GitHub
2. New Web Service → connect repo
3. Build command: `npm install`
4. Start command: `node server.js`
5. Free tier (Singapore) — set `PORT` env auto-detected

## Features
- Agent info saved permanently (one-time setup)
- Mobile-first, iOS Safari safe downloads
- Full validation: PAN format, 100% ownership, >50% present share, BO date = Resolution date
- All filled values bold + underlined in generated Word docs
- Optional letterhead in document header
- Google Sheets usage logging
