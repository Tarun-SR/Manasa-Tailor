# Manasa-Tailor

My Mom's Boutique — home-based ladies' tailoring, Tumkur, Karnataka.

Live site (GitHub Pages): served from `index.html` in this repo.

## What's in this site

- Single-page site: Home, About, Services, Process, Gallery, Pricing, **Order & Appointment**, Contact.
- WhatsApp click-to-chat + tel: link for direct contact (no backend needed).
- **Order & Appointment booking form** — captures garment, visit purpose, preferred
  date and time slot, saves it as a row in a Google Sheet, then opens WhatsApp
  with the same details pre-filled so nothing gets lost. There's no live
  calendar/availability check (that would need a separate scheduling account) —
  your mom confirms or reschedules each booking over WhatsApp, which fits how
  she already runs things manually.
- Subtle 3D accent (floating particles in the boutique's rose/gold colors)
  behind the hero image, built with Three.js. Automatically skipped on small
  screens and when the visitor's browser prefers reduced motion, so it never
  slows down or gets in the way on mobile.

No payment integration — orders are confirmed and paid for manually, as before.

## Connecting the order form to Google Sheets

The form works out of the box by opening WhatsApp directly. To also save
every enquiry to a Sheet for tracking:

1. Open `apps-script.gs` in this repo and follow the setup steps written at
   the top of the file (create a Sheet, paste the script into
   Extensions → Apps Script, deploy as a Web App).
2. Copy the Web App URL Google gives you.
3. In `index.html`, find this line near the bottom:
   ```js
   var GOOGLE_SHEET_WEBAPP_URL = "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE";
   ```
   Replace the placeholder with your Web App URL.
4. Commit and push — GitHub Pages updates automatically.

Until step 3 is done, the form still works fine; it just skips the Sheet and
goes straight to WhatsApp.

## Editing content

Everything lives in `index.html` — services, pricing, gallery images, and
contact details are plain HTML/CSS, no build step required. Product/gallery
images are the `.png` files in the repo root.