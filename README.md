# AI Day Planner

Desktop-first local MVP for generating a daily work plan from Outlook email, Microsoft Calendar, and Jira.

## Apps

- `apps/api`: Express + TypeScript + SQLite backend
- `apps/web`: React + Vite + TypeScript frontend

## Setup

1. Install dependencies:

```bash
npm install
```

2. Copy `apps/api/.env.example` to `apps/api/.env` and fill in the backend values.
3. Copy `apps/web/.env.example` to `apps/web/.env` and fill in the MSAL SPA values.
4. Start the API and web app in separate terminals:

```bash
npm run dev:api
npm run dev:web
```

## Notes

- Microsoft Graph uses delegated OAuth.
- The preferred Microsoft path uses `msal-browser` in the SPA and on-behalf-of token exchange in the API.
- Jira uses base URL + email + API token saved in the local database.
- If OpenAI settings are omitted, email classification falls back to deterministic heuristics.
