let autoSaveIntervalId = null;
const AUTO_SAVE_INTERVAL_MS = 10000; // 10 seconds

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("autosave-toggle").addEventListener("change", toggleAutoSave);
        
        // Start auto-save by default when add-in loads
        startAutoSave();
    }
});

function toggleAutoSave(event) {
    const isChecked = event.target.checked;
    const statusText = document.getElementById("toggle-status");
    const statusIndicator = document.getElementById("status-indicator");
    const statusIcon = document.getElementById("status-icon");

    if (isChecked) {
        statusText.innerText = "ON";
        statusText.style.color = "#2b579a";
        statusIndicator.innerText = "Auto-saving is active...";
        statusIndicator.className = "status-message active";
        statusIcon.style.color = "#107c10"; // Green for active
        startAutoSave();
    } else {
        statusText.innerText = "OFF";
        statusText.style.color = "#605e5c";
        statusIndicator.innerText = "Auto-saving paused.";
        statusIndicator.className = "status-message inactive";
        statusIcon.style.color = "#a4262c"; // Red for inactive
        stopAutoSave();
    }
}

function startAutoSave() {
    // Prevent multiple timers
    if (autoSaveIntervalId !== null) {
        clearInterval(autoSaveIntervalId);
    }

    // Run immediately once (optional, but requested behavior is 'every 10 seconds', so we wait first)
    autoSaveIntervalId = setInterval(async () => {
        await checkAndSaveDocument();
    }, AUTO_SAVE_INTERVAL_MS);
}

function stopAutoSave() {
    if (autoSaveIntervalId !== null) {
        clearInterval(autoSaveIntervalId);
        autoSaveIntervalId = null;
    }
}

async function checkAndSaveDocument() {
    try {
        await Word.run(async (context) => {
            // Load the document saved property to check for unsaved changes
            const doc = context.document;
            doc.load("saved");
            await context.sync();

            if (!doc.saved) {
                // There are unsaved changes, automatically save
                doc.save();
                await context.sync();
                
                updateLastSavedTime();
            }
            // If doc.saved is true, we do nothing to optimize performance
        });
    } catch (error) {
        console.error("Error during auto-save: ", error);
        document.getElementById("status-indicator").innerText = "Error: Could not save.";
        document.getElementById("status-indicator").className = "status-message inactive";
        document.getElementById("status-icon").style.color = "#a4262c";
    }
}

function updateLastSavedTime() {
    const now = new Date();
    // Format to HH:MM:SS
    const timeString = now.toLocaleTimeString([], { hour12: false });
    document.getElementById("last-saved-time").innerText = `Last saved at ${timeString}`;
}
