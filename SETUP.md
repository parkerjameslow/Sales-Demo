# Setup — Demo Request web app

Two links when you're done:

| Link | Who gets it |
|------|-------------|
| `https://parkerjameslow.github.io/Sales-Demo/` | **Sales reps** — the request form |
| `https://parkerjameslow.github.io/Sales-Demo/dashboard.html` | **You / managers** — the Admin Queue dashboard |

One-time setup, about 5 minutes total.

## 1. Create the database (Supabase — free)

1. Go to [supabase.com](https://supabase.com) → sign up (free) → **New project**.
   Pick any name (e.g. `demo-requests`) and a database password (you won't need
   it day-to-day — store it somewhere safe).
2. In the left sidebar open **SQL Editor** → **New query**, paste the entire
   contents of [`supabase/schema.sql`](supabase/schema.sql), and click **Run**.
   You should see "Success. No rows returned".

## 2. Connect the site to your database

1. In Supabase: **Project Settings** (gear icon) → **API**. You need two values:
   - **Project URL** (looks like `https://abcdefgh.supabase.co`)
   - **anon / public** key (the long string under *Project API keys*)
2. Open [`docs/config.js`](docs/config.js) in this repo and paste them in:

   ```js
   window.CONFIG = {
     SUPABASE_URL: 'https://abcdefgh.supabase.co',
     SUPABASE_ANON_KEY: 'eyJhbGciOi...'
   };
   ```

3. Commit the change (edit the file right on GitHub if that's easiest).

> **Is it safe to commit the anon key?** Yes — it's designed to be public.
> What visitors can actually do is controlled by the row-level-security
> policies in `schema.sql` (submit, read, and update triage fields — nothing
> else, no deletes).

## 3. Turn on the website (GitHub Pages — free)

1. Merge the PR so the `docs/` folder is on your default branch.
2. In this GitHub repo: **Settings → Pages**.
3. Under *Build and deployment*: Source = **Deploy from a branch**,
   Branch = your default branch, Folder = **`/docs`** → **Save**.
4. Wait ~1 minute, then visit `https://parkerjameslow.github.io/Sales-Demo/`.

## 4. Share the links

- Send the form link to your reps — it works on phones, no login needed.
- Keep the dashboard link for yourself and anyone who triages. From there you
  can search, filter by status/priority, sort columns, approve requests,
  open **Details** to change status/priority and add Jira/Sprint numbers,
  and **Export CSV** whenever someone needs a spreadsheet.

## Day-2 notes

- **Where's my data?** In your Supabase project → **Table Editor** →
  `requests`. You can also browse/edit it there directly.
- **Backups / Excel copies:** the dashboard's **Export CSV** button downloads
  exactly what's on screen (respects current filters).
- **Want to lock down the dashboard later?** The schema comment marks the
  update policy to tighten — we can add a team passcode or real logins
  without changing the form reps use.
- **Changing dropdown options** (statuses/priorities): they're defined in
  `supabase/schema.sql` (database checks) and at the top of the script in
  `docs/dashboard.html` — keep the two lists in sync.
