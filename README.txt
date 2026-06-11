LAB WHATSAPP SENDER - HOW TO GET THE EXE (one time only)

1. Go to github.com and sign in (create a free account if needed).
2. Click the "+" (top right) -> "New repository".
   - Name: lab-whatsapp-sender
   - Choose "Private".
   - Click "Create repository".
3. Click "uploading an existing file" link on the new repo page.
4. Drag ALL the extracted contents of this zip into the upload box:
   - main.py
   - requirements.txt
   - the .github folder (drag the whole folder)
5. Click "Commit changes".
6. Click the "Actions" tab at the top. The build "Build Windows EXE"
   starts automatically. Wait ~5 minutes until it shows a green check.
7. Click the finished run -> scroll down to "Artifacts" ->
   download "LabWhatsAppSender" (a zip containing LabWhatsAppSender.exe).
8. Put LabWhatsAppSender.exe in any folder on your work PC and double-click.
   No installation, no admin needed.

NOTES
- The program creates settings.json, sent_log.csv and a wa_profile folder
  next to the EXE automatically. Keep them with the EXE.
- First send: WhatsApp Web opens and shows a QR code -> scan it once with
  the phone whose number will send the messages. It stays logged in after that.
- Use the "Send test" button with your own number before any real sending.
