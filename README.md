# Word Auto-Save Add-in

This is a Microsoft Word Office Add-in that automatically saves your document every 10 seconds if there are unsaved changes. It works gracefully in the background without interrupting the user.

## Requirements
- Node.js installed (v16 or higher recommended).
- Microsoft Word (Desktop) or a Microsoft 365 account to test on Word Online.

## How to Test Locally

### 1. Install Dependencies
Open your terminal in the `autosave-word` folder and run:
```bash
npm install
```

### 2. Start the Local HTTPS Server
Office Add-ins require an HTTPS connection. The `office-addin-dev-server` will automatically create the necessary development certificates and serve the files on `https://localhost:3000`.

Run:
```bash
npm run dev-server
```
*(If prompted to install developer certificates, accept/allow the prompt).*

### 3. Sideload the Add-in in Word Desktop
While the server is running, open a new terminal window in the same folder and run:
```bash
npm start
```
This command will launch Microsoft Word and automatically sideload the add-in based on the `manifest.xml`.

### 4. Using the Add-in
1. Once Word opens, go to the **Home** tab.
2. Click the **Show Auto-Save** button on the ribbon.
3. The task pane will open on the right side.
4. The auto-save background timer starts automatically.
5. Make some changes to your document. After up to 10 seconds, the add-in will detect the unsaved changes, save the document, and update the "Last saved at HH:MM:SS" timestamp in the task pane.
6. You can toggle the feature ON and OFF using the switch in the task pane.

### Troubleshooting
- **Cannot reach localhost:3000:** Ensure that `npm run dev-server` is actively running.
- **Save throws an error:** If the document has never been saved before (e.g., a brand new Document1 without a file path), you may need to save it manually once so Word has a location to save to.
