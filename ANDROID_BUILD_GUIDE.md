# Build RR Credit APK

This project is prepared for Capacitor 6. To produce an APK, use a computer with:
- Node.js 18+
- Android Studio (with Android SDK installed)
- JDK 17

## Commands
```bash
npm install
npm run build
npx cap add android
npx cap sync android
npx cap open android
```

In Android Studio:
1. Wait for Gradle sync to finish.
2. Select Build > Build Bundle(s) / APK(s) > Build APK(s).
3. The debug APK is normally at `android/app/build/outputs/apk/debug/app-debug.apk`.

For a release APK, configure signing in Android Studio and build a signed APK.

## Notes
- This is a new app, not an update to the previous RR Mobiles V10.
- Data is stored locally using IndexedDB in the app's web runtime.
- Google Drive automatic backup is not yet integrated; it requires OAuth consent and Google Drive API configuration.
- Test on the target Android phone before relying on it for business records.

## Backup mode
This build is offline-only. No Gmail sign-in, Google Drive connection, cloud sync, or automatic backup is included. The user manually exports a JSON backup when desired.
