# Putting MyPDFKit in the Microsoft Store (free)

Microsoft made the Store free for individual developers, and it signs the app itself, so no certificate is needed.
Only you can do steps 1–3 (they need your identity). After step 3, send me the three names and I build the package.

## 1. Make the free developer account (about 15 minutes)
1. Go to **https://storedeveloper.microsoft.com** (this is the free sign-up for individuals; don't start from Partner Center).
2. Sign in with your Microsoft account and choose **Individual**.
3. Microsoft checks your ID (photo of an ID and a selfie). It can take from minutes to a few days.

## 2. Reserve the name
1. In Partner Center: **Apps and games → New product → MSIX or PWA app**.
2. Type **MyPDFKit** and click **Reserve product name**. If it's taken, pick another (tell me which).

## 3. Get the three names for the package
In the app's page: **Product management → Product identity**. Copy these three lines and send them to me:
- **Package/Identity/Name** (looks like `12345YourName.MyPDFKit`)
- **Package/Identity/Publisher** (looks like `CN=ABCD1234-...`)
- **Package/Properties/PublisherDisplayName** (your name as shown in the Store)

I put them in `store-identity.json` and run:
```
npx electron-builder --win dir --publish never --config electron-builder.store.js -c.directories.output=release-store
python store/make_msix.py
```
That makes `release-store\MyPDFKit-Store-<version>.msix` (about 1.5 GB, with the built-in AI).

## 4. Fill in the Store listing (Partner Center → your app → Start your submission)
- **Pricing and availability:** Free. Markets: all.
- **Properties:** Category **Productivity**. Privacy policy URL: required (the app can use the internet).
  The text is ready in `store/PRIVACY.md`; it needs to be on a public web page. Tell me if I should put it on the GitHub releases page.
- **Age ratings:** answer the questionnaire (no violence, no user chat): it comes out **Everyone / 3+**.
- **Packages:** upload the `.msix` file.
- **Store listing (English):**
  - Description: *Edit, sign, convert, print and summarize PDFs. Opens Word, Excel, PowerPoint and pictures even without Microsoft Office. Prints labels on 4x6 label printers, fills forms, locks files, and has a built-in AI that summarizes and answers questions about your files right on your computer: no account, no internet, free.*
  - Screenshots: at least one (1366×768 or larger). I can make them.
- **Submission options → Notes for certification:** *Desktop app (Electron). The built-in AI runs locally (llama.cpp). Online AI is optional and needs the user's own key.*
- **Restricted capability "runFullTrust":** explain: *It's a desktop app that opens and saves the user's files and runs its own PDF engine.*

Then **Submit to the Store**. Microsoft's review usually takes 1–3 days. After that, friends search **MyPDFKit** in the Microsoft Store and click **Get**: no "Windows protected your PC" warning.

## After it's in the Store
- The Store copy updates through the Store (its built-in updater is off). For each new version I build a new `.msix` and you upload it in a new submission.
- The GitHub downloads keep working as they do now.
