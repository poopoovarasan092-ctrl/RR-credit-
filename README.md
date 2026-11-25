# RR Credit – Offline V1

A new standalone offline-first RR Mobiles customer ledger.

## Features
- IndexedDB storage on the device
- Customer records and balances
- Credit given and payments received
- Savings records
- Excel workbook export (Dashboard, Customers, Transactions, Collections, Savings)
- Manual JSON backup and restore (only when you choose)
- Installable PWA shell

## Run locally
1. Install Node.js 18 or newer.
2. Extract this ZIP.
3. Run `npm install`
4. Run `npm run dev` for development, or `npm run build` for production.
5. Deploy the `dist` folder to an HTTPS static host.

## Android
This V1 is a PWA and can be installed from Chrome using “Add to Home screen”. A signed APK requires Android/Capacitor build tooling and testing on the target phone.

## Important storage notes
- Records are stored locally in this browser profile using IndexedDB.
- Clearing browser/app storage can erase records. Use JSON backup regularly.
- Excel reports are generated on-device.
- No Google Drive, Gmail, cloud sync, or automatic backup. Data remains on this phone. Use Export Backup only when you choose to create a manual backup.
- This is a new app; no existing customer records are imported.
