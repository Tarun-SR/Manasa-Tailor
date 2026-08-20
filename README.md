# Manasa Tailor — v1

My mom's boutique website — home-based ladies' tailoring in Tumkur, Karnataka.

## Live Site

<https://tarun-sr.github.io/Manasa-Tailor/>

## Pages

- `index.html` — Public marketing site
- `login.html` — Customer register / login
- `dashboard.html` — Customer portal (orders, status tracking)
- `booking.html` — Book appointment + quotation request
- `profile.html` — Customer profile settings
- `admin.html` — Admin dashboard (Shilpa's portal)
- `404.html` — Branded error page

## Tech Stack

- **Frontend:** Pure HTML, CSS, JavaScript — no frameworks, no build step
- **Backend:** Google Apps Script (deployed as a Web App)
- **Database:** Google Sheets
- **Hosting:** GitHub Pages

## Features (v1)

- Customer registration and login
- 3-step appointment booking with service-specific measurements
- Quotation request → admin review → customer approval flow
- Full order lifecycle tracking with status updates
- Admin dashboard — manage orders, customers, quotations
- Admin walk-in profile creation with auto-generated credentials
- Photo upload for completed orders
- WhatsApp notifications at every key step
- In-app notification bell for both customer and admin
- Customer profile settings (edit details, change password)
- Mobile responsive across all pages

## Setup

1. **Google Sheet.** Create a new Google Sheet, then Extensions → Apps
   Script, delete any starter code, and paste in `apps_script.gs`.
2. **Point it at that Sheet.** Set `SHEET_ID` near the top of
   `apps_script.gs` to the Sheet's ID (the long ID in its URL, between `/d/`
   and `/edit`). Every read/write goes through `SpreadsheetApp.openById`,
   so the script always operates on that specific Sheet regardless of
   whether it's bound to it or run as a standalone project.
3. **Create the tabs.** Select `setupSheets` in the toolbar dropdown and
   click Run once. This creates the `Users`, `Orders`, `Quotations`,
   `Measurements`, `Notifications` and `Admin` tabs with their header rows.
   Safe to re-run later — it skips tabs that already exist.
4. **Create the admin account.** Scroll to `ADMIN_BOOTSTRAP_EMAIL` /
   `ADMIN_BOOTSTRAP_PASSWORD` near the bottom of `apps_script.gs`, fill them
   in, select `createAdminAccount` in the toolbar dropdown and click Run
   once — that's the email/password `admin.html` logs in with. Blank
   `ADMIN_BOOTSTRAP_PASSWORD` out again afterwards so it isn't left sitting
   in the script in plain text. There's no admin login without this step.
5. **Deploy.** Deploy → New deployment → type "Web app" (Execute as: Me,
   Who has access: Anyone), then copy the Web App URL it gives you.
6. **Wire up the frontend.** Every page except `404.html` has its own copy
   of the same line near the bottom of the file:

   ```js
   var GOOGLE_SHEET_WEBAPP_URL = "...";
   ```

   Replace it with your Web App URL in `index.html`, `login.html`,
   `dashboard.html`, `booking.html`, `profile.html` and `admin.html`.
7. **Push to GitHub.** GitHub Pages serves the site automatically once the
   repo is pushed — no build step.

Every time `apps_script.gs` is edited, create a **new deployment** (or
"Manage deployments" → edit → new version) for the change to go live — the
Web App URL itself stays the same either way.

## V2 Roadmap

- Customer testimonials section (#42)
- Instagram feed integration (#43)
- Forgot password via WhatsApp OTP (#44)
- WhatsApp Business API for automated messaging (#45)
