async function prepareDownloadPage() {
    const page = document.querySelector("[data-download-platform]");
    if (!page) {
        return;
    }

    const platform = page.dataset.downloadPlatform;
    const statusNode = document.querySelector("[data-download-status]");
    const filenameNode = document.querySelector("[data-download-filename]");
    const manualLink = document.querySelector("[data-download-link]");
    const waitingNode = document.querySelector("[data-download-waiting]");

    try {
        const response = await fetch("/downloads/manifest.json", { cache: "no-store" });
        if (!response.ok) {
            throw new Error("Manifest indisponible");
        }

        const manifest = await response.json();
        const entry = manifest?.[platform];
        if (!entry) {
            throw new Error("Plateforme inconnue");
        }

        filenameNode.textContent = entry.filename;
        if (entry.available && entry.url) {
            statusNode.textContent = "Le telechargement va commencer automatiquement.";
            manualLink.href = entry.url;
            manualLink.hidden = false;
            waitingNode.hidden = true;
            setTimeout(() => {
                window.location.assign(entry.url);
            }, 450);
            return;
        }

        statusNode.textContent = "Le telechargement n'est pas encore publie pour ce systeme.";
        manualLink.hidden = true;
        waitingNode.hidden = false;
    } catch (_error) {
        statusNode.textContent = "Impossible de verifier le telechargement pour le moment.";
        manualLink.hidden = true;
        waitingNode.hidden = false;
    }
}

prepareDownloadPage();

