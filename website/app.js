const downloadTargets = {
    windows: {
        label: "Telecharger pour Windows",
        url: "/telecharger/windows",
    },
    macos: {
        label: "Telecharger pour macOS",
        url: "/telecharger/macos",
    },
    linux: {
        label: "Telecharger pour Linux",
        url: "/telecharger/linux",
    },
};

function detectOperatingSystem() {
    const userAgent = navigator.userAgent.toLowerCase();
    const platform = navigator.platform.toLowerCase();

    if (platform.includes("win") || userAgent.includes("windows")) {
        return "windows";
    }
    if (platform.includes("mac") || userAgent.includes("mac os")) {
        return "macos";
    }
    if (platform.includes("linux") || userAgent.includes("linux")) {
        return "linux";
    }
    return "windows";
}

function highlightDetectedOs(osKey) {
    document.querySelectorAll("[data-os-card]").forEach((card) => {
        card.classList.toggle("is-active", card.dataset.osCard === osKey);
    });
}

function updateHeroDownload(osKey) {
    const heroButton = document.querySelector("#hero-download");
    const config = downloadTargets[osKey];
    if (!heroButton || !config) {
        return;
    }
    heroButton.textContent = config.label;
    heroButton.href = config.url;
}

const detectedOs = detectOperatingSystem();
highlightDetectedOs(detectedOs);
updateHeroDownload(detectedOs);
