# Demo Request Form

Turns the demo-request form into a **Google Form** so the sales team can fill it
in online and submit it — with every response collecting automatically into a
Google Sheet.

## What's in here

| File | What it does |
|------|--------------|
| `demo-request-form.gs` | A Google Apps Script that builds the form + responses Sheet for you |

## Setup (one time, ~2 minutes)

1. Go to **[script.google.com](https://script.google.com)** → **New project**.
2. Delete the empty `myFunction` stub, then paste in the full contents of
   [`demo-request-form.gs`](./demo-request-form.gs).
3. Click **Save** (💾).
4. In the function dropdown at the top, make sure **`buildDemoRequestForm`** is
   selected, then click **Run** (▶️).
5. The first run asks you to **authorize** the script — approve it with your
   Google account (it only touches Forms/Sheets you own).
6. Open **View → Logs** (or the Execution log at the bottom). You'll see three links:
   - **Edit the form** — tweak wording, add fields, change the theme/colors
   - **Send THIS link to reps** — the link you share with the sales team
   - **Responses Sheet** — every submission lands here as a new row

That's it. Share the "Send THIS link to reps" URL and you're live.

## The fields it creates

**The Basics**
- Request Name *(required)*
- Your Name *(required)*
- Date Requested

**Priority & Tracking**
- Application(s) *(required)*
- Jira #
- Sprint #
- Priority *(required)* — Critical / High / Medium / Low

**Tell Us More**
- Full Description *(required)*

## Notes & customizing

- **Look & feel:** Google Forms can't reproduce the custom card layout or the
  colored priority buttons from the original HTML mockup. You *can* set a theme
  color, header image, and font from the form's **🎨 Customize theme** menu.
- **Email collection:** the script turns on "collect respondent email" so you
  know who submitted each request. If your reps are outside your Google
  Workspace org (or you don't want them to sign in), set
  `form.setCollectEmail(true)` → `false` in the script before running.
- **Changing fields:** edit `demo-request-form.gs` and re-run to generate a
  fresh copy, *or* just edit the live form directly via its edit link.
- **Want the exact custom look instead?** That route is a self-hosted HTML form
  with a form backend (e.g. Formspree / Netlify Forms) — happy to build that
  version too; it keeps the original styling but needs hosting.
