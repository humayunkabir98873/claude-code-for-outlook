# Claude Mail Assistant (classic Outlook add-in)

Adds a **Claude** button to the ribbon when you read an email. The task pane has three tabs.

**Email**: Summarize, Spam check, Triage, Actions and dates, Draft reply (tone + instructions, opens a reply window),
Translate, and Ask about this email.
**Search**: describe what you want in plain English ("invoices from Acme last month"), get matching emails,
open them, or have Claude answer from the top results.
**Settings**: paste your Claude API key, pick a model, set your sign-off name. No config files.

Spam check combines real signals (SPF/DKIM/DMARC results, Reply-To and display-name mismatches,
links whose text differs from their destination, shorteners, risky attachments) with Claude's reading of the content.

## Requirements
- Classic Outlook for Windows from Microsoft 365, or Outlook 2021 / LTSC 2021 and later (all use the Edge WebView2
  runtime). Outlook 2016/2019 use an Internet Explorer engine and won't run this code.
- For the web add-in: a Microsoft 365, Exchange or Outlook.com mailbox. Web add-ins don't load for IMAP/POP accounts
  (the Windows installer below does work with them).
- Node.js 18+ (not needed if you use one of the [executables](#executables))
- A Claude API key from the Claude Platform (docs.claude.com). This is separate from a claude.ai login.

## Install on Windows (recommended)
Double-click `ClaudeMailAssistantSetup.exe` (SmartScreen: More info > Run anyway, it isn't code-signed), choose Yes,
and close Outlook when asked. Start Outlook: **Claude** is on the Home tab, like any other add-in (it appears under
File > Options > Add-ins > COM Add-ins). Click it, go to **Settings**, paste your key, Save.

- No Node.js, server window, certificate or manifest. It installs for your Windows account only, without admin
  rights, into `%LOCALAPPDATA%\Programs\ClaudeMailAssistant`. Remove it in Settings > Apps > Claude Mail Assistant.
- Built for classic Outlook 2016 and later, 32- or 64-bit. As a COM add-in it also works on IMAP/POP accounts
  (Search still needs a Microsoft 365 or Outlook.com mailbox). Needs .NET Framework 4.8 and the WebView2 runtime,
  both included in current Windows 10/11.
- The Claude pane belongs to the main Outlook window. Messages opened in their own window aren't covered yet.
- Unlike the web add-in, "Apply category" creates the category if it doesn't exist yet.
- Build it with `npm run build:outlook` (Windows with Office installed). The source is in `outlook-com/`: a C# add-in
  that shows the same `public/` pages in WebView2 and swaps office.js for `outlook-com/AddIn/office-shim.js`, which
  implements the Office calls the pages use through Outlook's own object model. Logs go to
  `%LOCALAPPDATA%\ClaudeMailAssistant\log.txt`.

The next sections are for the web add-in, which also runs on Mac and in Outlook on the web.

## Start it
1. Unzip, then double-click `start.bat`, or run `npm install` then `npm start`. The first run creates a localhost
   certificate and Windows asks whether to trust "Developer CA for Microsoft Office Add-ins": choose Yes.
   (No Node.js? Use a ready-made executable instead, see [Executables](#executables).)
2. In Outlook: open an email, then Get Add-ins > My add-ins > Custom add-ins > Add from file > pick `manifest.xml`.
   If that option is missing or greyed out, open https://aka.ms/olksideload in a browser signed in to the same
   mailbox and add the file there; it syncs to desktop Outlook after a restart. If neither place offers custom
   add-ins, your Exchange admin has turned them off for your account.
3. Click **Claude** on the ribbon (use the pin to keep the pane open), go to **Settings**, paste your key, Save.

Keep the `start.bat` window open while you use the add-in. The server only serves the add-in's files; it never sees your key or your mail.

## Executables
For people without Node.js. `npm install` then `npm run build` creates `dist/ClaudeMailAssistant-windows-x64.zip`,
`-macos-arm64.zip` and `-macos-x64.zip`. Each holds the program (with the `public` folder built in), `manifest.xml`
and `HOW-TO-USE.txt`. Run the build again after changing anything in `public`.
- Windows: unzip, double-click `ClaudeMailAssistant.exe` (SmartScreen: More info > Run anyway, it isn't code-signed)
  and choose Yes when Windows asks to install the certificate. Then sideload `manifest.xml` as in step 2 above.
- Mac: unzip and follow `HOW-TO-USE.txt`. In short, run once in Terminal inside the folder:
  `chmod +x ClaudeMailAssistant && xattr -c ClaudeMailAssistant && codesign --force --sign - ClaudeMailAssistant`,
  then double-click `ClaudeMailAssistant` and enter your password when macOS asks to trust the certificate.

## Troubleshooting
- Blank or "add-in error" pane: check the `start.bat` window is still open, then open https://localhost:3000/taskpane.html
  in Edge. A certificate warning there means the localhost certificate isn't trusted: close the window, start it
  again and choose Yes when Windows asks to install the certificate.
- Something fails inside the pane: right-click in the pane > Inspect to open DevTools, then look at the Console tab.
- After editing `manifest.xml`, remove the add-in (Get Add-ins > My add-ins) and add it again, then restart Outlook.

## Mail search setup (one time, optional)
Outlook add-ins can't search the whole mailbox by themselves, so search uses Microsoft Graph with an app
registration you create (free, about 5 minutes). The steps are also in Settings > Mail search setup:
1. entra.microsoft.com > App registrations > New registration.
2. Account types: "any organizational directory and personal Microsoft accounts".
3. Redirect URI: platform "Single-page application", `https://claude-mail-assistant.example/auth.html` for the
   Windows installer, `https://localhost:3000/auth.html` for the web add-in (add both if you use both).
   Settings > Mail search setup shows the right one.
4. API permissions > Add > Microsoft Graph > Delegated > `Mail.Read`.
5. Paste the Application (client) ID into Settings.
On a work account your IT admin may need to approve the Mail.Read permission.

## Privacy and cost
- Your API key is saved in the add-in's browser storage on your PC in plain text. Use a key with a spend limit.
- Email text (and for spam checks, selected header results and link domains) is sent to the Claude API. Don't use this on mail you may not share with a third party.
- The default model is Haiku 5.5 (cheapest). Change it in Settings.
- "Apply category" (Triage) needs categories named `Claude: Needs reply`, `Claude: Action required`, `Claude: FYI`, `Claude: Promotional` to exist in Outlook.

## Hosting without running anything locally
The `public` folder is plain static files. You can host it on any HTTPS static host, then replace
`https://localhost:3000` in `manifest.xml` with your URL (and add it as a redirect URI in the Azure app registration).
