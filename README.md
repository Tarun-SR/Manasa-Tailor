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

## Backend: Google Sheets + Apps Script (v2)

`apps_script.gs` is the full backend API — user accounts, orders,
quotations, measurements and notifications, backed by one Google Sheet with
six tabs: `Users`, `Orders`, `Quotations`, `Measurements`, `Notifications`,
`Admin`. Follow the setup steps written at the top of the file:

1. Create a Sheet, paste `apps_script.gs` into Extensions → Apps Script.
2. Run `setupSheets()` once from the Apps Script editor to create the six
   tabs with their header rows.
3. Add at least one row to the `Admin` tab by hand — `getAllOrders` and
   `updateOrderStatus` require a valid `adminId` from that tab.
4. Deploy → New deployment → Web app (execute as Me, access: Anyone), then
   copy the Web App URL.
5. In `index.html`, find this line near the bottom:

   ```js
   var GOOGLE_SHEET_WEBAPP_URL = "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE";
   ```

   Replace the placeholder with your Web App URL.
6. Commit and push — GitHub Pages updates automatically.

The script exposes `createUser`, `loginUser`, `createOrder`,
`getOrdersByUser`, `getAllOrders`, `updateOrderStatus`, `saveQuotation`,
`saveMeasurements` and `getNotifications` as `action` values on `doGet`/
`doPost`; every response is JSON (`{success, data}` or `{success, error}`).
The current booking form still posts the old plain shape (no `action`),
which the script treats as a legacy enquiry — it auto-creates a guest user
and an order so bookings keep landing in the Sheet. Wiring the rest of the
site (login, account area, admin dashboard) to these new endpoints is a
follow-up piece of work; until then the form still works fine even with no
Web App URL configured — it just skips the Sheet and goes straight to
WhatsApp.

## Editing content

Everything lives in `index.html` — services, pricing, gallery images, and
contact details are plain HTML/CSS, no build step required. Product/gallery
images are the `.png` files in the repo root.
