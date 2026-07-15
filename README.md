# AI Email Assistant — Outlook Add-in

A task-pane add-in that lets you draft new emails or generate replies using
ChatGPT (OpenAI API), directly inside Outlook.

## What it does
The task pane has four tabs:

- **Draft/Reply** — reading an email: describe what you want to say and it
  opens a reply window pre-filled with an AI draft. Composing: it generates
  text and inserts it at your cursor.
- **Edit** — load the current email's text (or paste your own), then either
  click a quick-edit chip (fix grammar, more formal, more casual, shorten,
  expand, summarize) or type a custom instruction (e.g. "remove the second
  paragraph"). Insert the result back into the email when you're happy with it.
- **Search** — find every occurrence of a word or phrase within the email
  you currently have open, with highlighted context snippets and a match
  count. This searches the one open email, not your whole mailbox.
- **Categorize** — while reading an email, click "Suggest category with AI"
  and it picks a fitting Outlook category (reusing one of your existing
  categories if it fits, or proposing a new one), which you can apply with
  one click. Categories only apply to received/saved emails, not new drafts.

**Note on scope:** search and categorize here work on the single email you
currently have open — that's what's possible with the standard Office.js API
without extra setup. Searching or bulk-categorizing your *entire inbox* would
need Microsoft Graph API access (an Azure app registration + sign-in flow).
Happy to build that as a next step if you want it — just ask.

## Files
```
manifest.xml              ← the add-in's manifest (registers it with Outlook)
src/taskpane/taskpane.html
src/taskpane/taskpane.css
src/taskpane/taskpane.js  ← all the logic (Office.js + OpenAI API calls)
assets/icon-16/32/64/80.png  ← placeholder icons (swap for your own if you like)
```

## Important: it needs HTTPS hosting
Outlook add-ins must be served over HTTPS — you can't just open the HTML file
locally. You have two easy options:

### Option A — Quick local testing (recommended to start)
1. Install Node.js if you don't have it.
2. In this folder, run:
   ```
   npx office-addin-dev-certs install
   npx http-server . -p 3000 --ssl -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key
   ```
   This serves the folder at `https://localhost:3000` with a trusted local
   certificate.
3. Open `manifest.xml` and replace every `YOUR-HOSTING-DOMAIN` with
   `localhost:3000`.
4. Sideload the add-in (see below).

### Option B — Free permanent hosting
Push this folder to a static host like **GitHub Pages**, **Netlify**, or
**Vercel** (all give you free HTTPS). Then replace `YOUR-HOSTING-DOMAIN` in
`manifest.xml` with your real domain, e.g. `yourname.github.io/outlook-ai-addin`.

## Sideloading the add-in
**Outlook on the web:**
1. Go to outlook.office.com → Settings (gear icon) → View all Outlook settings
   → General → Manage add-ins → My add-ins → "Add a custom add-in" → "Add from file"
   → select your edited `manifest.xml`.

**Outlook desktop (Windows/Mac):**
1. Home tab → Get Add-ins → My add-ins → "Add a custom add-in" → "Add from file"
   → select `manifest.xml`.

Once added, open or compose an email and look for the **"AI Draft/Reply"**
button in the ribbon.

## Using it
1. Open the task pane and paste in your OpenAI API key (get one at
   platform.openai.com/api-keys). Click "Save key on this device" — it's
   stored only in your browser's local storage and sent straight to OpenAI,
   never to any server of ours.
2. Type instructions (e.g. "Politely decline, propose Tuesday instead") and
   pick a tone.
3. Click Generate, review the draft, and click "Insert into email".

## Security notes (please read)
- The API key lives in the browser's local storage and is sent directly from
  your machine to `api.openai.com`. That's fine for **personal use**, but:
  - Anyone else who uses the same Windows/Mac profile could see the key.
  - Don't share your `manifest.xml`/hosted files with the key baked in — the
    key is only ever typed by the user into the task pane, never stored in
    the source files.
- For a team/production deployment, put a small backend between the add-in
  and OpenAI (so the API key stays server-side) rather than calling OpenAI
  directly from the browser. I can help you build that proxy if you want it.
- OpenAI usage is billed to your account per API call — keep an eye on
  platform.openai.com/usage.

## Customizing
- Change the model list or default tone options in `taskpane.html`.
- Change the system prompts (how formal/casual/creative it writes) in
  `taskpane.js` inside `generateDraft()`.
- Swap the icons in `assets/` for your own branding.
