"use strict";

/* =========================================================
   HANDSIGN - HAND GESTURE PRESENTATION CONTROLLER
   PDF viewer + webcam hand tracking + fullscreen controls
   ========================================================= */

const $ = (id) => document.getElementById(id);

/* ---------- DOM elements ---------- */

const fileInput = $("fileInput");
const fileName = $("fileName");

const inputVideo = $("inputVideo");
const cameraCanvas = $("cameraCanvas");
const cameraContext = cameraCanvas.getContext("2d");

const cameraButton = $("cameraButton");
const cameraStatus = $("cameraStatus");
const cameraPlaceholder = $("cameraPlaceholder");
const cameraLiveLabel = $("cameraLiveLabel");

const detectedGesture = $("detectedGesture");
const appStatus = $("appStatus");

const presentationPanel = $("presentationPanel");
const presentationStage = $("presentationStage");

const presentationTitle = $("presentationTitle");
const presentationSubtitle = $("presentationSubtitle");
const pageCounter = $("pageCounter");

const welcomeViewer = $("welcomeViewer");
const pdfContainer = $("pdfContainer");
const pdfCanvas = $("pdfCanvas");
const pdfContext = pdfCanvas.getContext("2d");

const pptContainer = $("pptContainer");

const loadingOverlay = $("loadingOverlay");
const loadingText = $("loadingText");
const errorMessage = $("errorMessage");

const fullscreenButton = $("fullscreenButton");

const previousButton = $("previousButton");
const nextButton = $("nextButton");
const footerPrevious = $("footerPrevious");
const footerNext = $("footerNext");

const viewerMessage = $("viewerMessage");

/* ---------- PDF.js setup ---------- */

if (window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

/* ---------- Application state ---------- */

let currentMode = "none";
let currentPage = 1;
let totalPages = 0;

let pdfDocument = null;
let renderTask = null;
let renderGeneration = 0;

let pptSlides = [];
let pptCurrentSlide = 0;

let cameraStream = null;
let handsModel = null;
let cameraRunning = false;
let cameraStarting = false;
let cameraFrameBusy = false;

let animationFrameId = null;

let lastGesture = "None";
let candidateGesture = "None";
let candidateSince = 0;
let lastActionTime = 0;

let gestureCooldown = 650;
let gestureStableTime = 120;

let currentObjectUrl = null;

/* ---------- Basic UI helpers ---------- */

function setStatus(message) {
    appStatus.textContent = message;
}

function showLoading(message = "Loading presentation...") {
    loadingText.textContent = message;
    loadingOverlay.classList.remove("hidden");
}

function hideLoading() {
    loadingOverlay.classList.add("hidden");
}

function showError(message) {
    errorMessage.textContent = message;
    errorMessage.classList.remove("hidden");
}

function hideError() {
    errorMessage.textContent = "";
    errorMessage.classList.add("hidden");
}

function updatePageCounter() {
    pageCounter.textContent = `${currentPage} / ${totalPages}`;

    previousButton.disabled = currentPage <= 1;
    footerPrevious.disabled = currentPage <= 1;

    nextButton.disabled = currentPage >= totalPages;
    footerNext.disabled = currentPage >= totalPages;
}

function updateCounterForPPT() {
    pageCounter.textContent =
        `${pptCurrentSlide + 1} / ${pptSlides.length}`;

    previousButton.disabled = pptCurrentSlide <= 0;
    footerPrevious.disabled = pptCurrentSlide <= 0;

    nextButton.disabled =
        pptCurrentSlide >= pptSlides.length - 1;

    footerNext.disabled =
        pptCurrentSlide >= pptSlides.length - 1;
}

function showViewer(mode) {
    currentMode = mode;

    welcomeViewer.classList.toggle("hidden", mode !== "none");
    pdfContainer.classList.toggle("hidden", mode !== "pdf");
    pptContainer.classList.toggle("hidden", mode !== "ppt");

    hideError();
}

function updatePresentationTitle(title, subtitle) {
    presentationTitle.textContent = title;
    presentationSubtitle.textContent = subtitle;
}

/* ---------- File upload ---------- */

fileInput.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];

    if (!file) {
        return;
    }

    fileName.textContent = file.name;
    hideError();

    const extension = file.name.split(".").pop().toLowerCase();

    if (extension === "pdf") {
        await loadPDF(file);
    } else if (extension === "pptx") {
        await loadPPTX(file);
    } else if (extension === "ppt") {
        showError(
            "Legacy .ppt files are not supported by this browser renderer. " +
            "Open the file in PowerPoint and save or export it as PDF or PPTX."
        );

        setStatus("Unsupported PowerPoint format");
        viewerMessage.textContent = "Please upload a PDF or PPTX file.";
    } else {
        showError("Please choose a PDF, PPT, or PPTX file.");
    }

    // Allow the same file to be selected again.
    fileInput.value = "";
});

/* ---------- PDF loading and rendering ---------- */

async function loadPDF(file) {
    showLoading("Opening PDF...");
    setStatus("Loading PDF");
    viewerMessage.textContent = "Loading PDF document...";

    try {
        if (!window.pdfjsLib) {
            throw new Error(
                "PDF.js did not load. Check your internet connection and refresh the page."
            );
        }

        stopExistingPresentation();

        const buffer = await file.arrayBuffer();

        pdfDocument = await pdfjsLib.getDocument({
            data: new Uint8Array(buffer)
        }).promise;

        totalPages = pdfDocument.numPages;
        currentPage = 1;

        showViewer("pdf");

        updatePresentationTitle(
            file.name,
            `${totalPages} page${totalPages === 1 ? "" : "s"} · PDF Document`
        );

        await renderPDFPage(currentPage);

        updatePageCounter();

        setStatus("PDF ready");
        viewerMessage.textContent = "PDF loaded successfully.";

    } catch (error) {
        console.error("PDF loading error:", error);
        showError(`Could not open this PDF. ${error.message || ""}`);
        setStatus("PDF loading failed");
    } finally {
        hideLoading();
    }
}

async function renderPDFPage(pageNumber) {
    if (!pdfDocument) {
        return;
    }

    const generation = ++renderGeneration;

    try {
        if (renderTask) {
            try {
                renderTask.cancel();
            } catch (_) {
                // The previous rendering task may already be complete.
            }

            renderTask = null;
        }

        const page = await pdfDocument.getPage(pageNumber);

        if (generation !== renderGeneration) {
            return;
        }

        const stageWidth = Math.max(
            250,
            presentationStage.clientWidth - 36
        );

        const stageHeight = Math.max(
            250,
            presentationStage.clientHeight - 36
        );

        const baseViewport = page.getViewport({ scale: 1 });

        const fitScale = Math.min(
            stageWidth / baseViewport.width,
            stageHeight / baseViewport.height
        );

        const pixelRatio = Math.min(window.devicePixelRatio || 1, 2);

        const viewport = page.getViewport({
            scale: fitScale * pixelRatio
        });

        pdfCanvas.width = Math.floor(viewport.width);
        pdfCanvas.height = Math.floor(viewport.height);

        pdfCanvas.style.width =
            `${Math.floor(viewport.width / pixelRatio)}px`;

        pdfCanvas.style.height =
            `${Math.floor(viewport.height / pixelRatio)}px`;

        renderTask = page.render({
            canvasContext: pdfContext,
            viewport
        });

        await renderTask.promise;

        if (generation !== renderGeneration) {
            return;
        }

        currentPage = pageNumber;
        updatePageCounter();

    } catch (error) {
        if (error?.name !== "RenderingCancelledException") {
            console.error("PDF rendering error:", error);
            showError("Could not render this PDF page.");
        }
    } finally {
        renderTask = null;
    }
}

/* ---------- PPTX loading ---------- */

/*
   This is a basic PPTX XML reader, not a full PowerPoint renderer.
   It extracts slide text and displays it in a simple slide layout.
   Complex layouts, images, charts, and animations may not be preserved.
*/

async function loadPPTX(file) {
    showLoading("Reading PowerPoint slides...");
    setStatus("Loading PPTX");
    viewerMessage.textContent = "Reading PowerPoint slide text...";

    try {
        if (!window.JSZip) {
            throw new Error(
                "The PPTX reader library did not load. Check your internet connection."
            );
        }

        stopExistingPresentation();

        const zip = await JSZip.loadAsync(file);

        const slideFiles = Object.keys(zip.files)
            .filter((path) =>
                /^ppt\/slides\/slide\d+\.xml$/.test(path)
            )
            .sort((a, b) => {
                const numberA = Number(
                    a.match(/slide(\d+)\.xml$/)?.[1] || 0
                );

                const numberB = Number(
                    b.match(/slide(\d+)\.xml$/)?.[1] || 0
                );

                return numberA - numberB;
            });

        if (!slideFiles.length) {
            throw new Error("No slides were found inside this PPTX file.");
        }

        pptSlides = [];

        for (const path of slideFiles) {
            const xmlText = await zip.files[path].async("text");

            const xmlDocument = new DOMParser().parseFromString(
                xmlText,
                "application/xml"
            );

            const textNodes = Array.from(
                xmlDocument.getElementsByTagNameNS("*", "t")
            );

            const lines = textNodes
                .map((node) => node.textContent.trim())
                .filter(Boolean);

            const title = lines[0] || `Slide ${pptSlides.length + 1}`;

            pptSlides.push({
                title,
                lines
            });
        }

        pptCurrentSlide = 0;

        showViewer("ppt");

        updatePresentationTitle(
            file.name,
            `${pptSlides.length} slides · Basic PPTX text preview`
        );

        renderPPTSlide();

        setStatus("PPTX loaded");
        viewerMessage.textContent =
            "PPTX text preview loaded. Original formatting may differ.";

    } catch (error) {
        console.error("PPTX loading error:", error);

        showError(
            `Could not read this PowerPoint file. ${error.message || ""}`
        );

        setStatus("PPTX loading failed");

    } finally {
        hideLoading();
    }
}

function renderPPTSlide() {
    if (!pptSlides.length) {
        return;
    }

    const slide = pptSlides[pptCurrentSlide];

    pptContainer.innerHTML = "";

    const slideElement = document.createElement("article");
    slideElement.className = "ppt-slide";

    const heading = document.createElement("h2");
    heading.textContent = slide.title;

    slideElement.appendChild(heading);

    const bodyLines = slide.lines.slice(1);

    if (bodyLines.length === 0) {
        const paragraph = document.createElement("p");
        paragraph.textContent = "No additional text on this slide.";
        slideElement.appendChild(paragraph);
    } else {
        bodyLines.forEach((line) => {
            const paragraph = document.createElement("p");
            paragraph.textContent = line;
            paragraph.style.marginBottom = "10px";
            slideElement.appendChild(paragraph);
        });
    }

    pptContainer.appendChild(slideElement);

    updateCounterForPPT();
}

/* ---------- Page navigation ---------- */

async function nextPage() {
    if (currentMode === "pdf" && currentPage < totalPages) {
        await renderPDFPage(currentPage + 1);
        viewerMessage.textContent = "Moved to the next PDF page.";
        return;
    }

    if (currentMode === "ppt" && pptCurrentSlide < pptSlides.length - 1) {
        pptCurrentSlide++;
        renderPPTSlide();
        viewerMessage.textContent = "Moved to the next slide.";
    }
}

async function previousPage() {
    if (currentMode === "pdf" && currentPage > 1) {
        await renderPDFPage(currentPage - 1);
        viewerMessage.textContent = "Moved to the previous PDF page.";
        return;
    }

    if (currentMode === "ppt" && pptCurrentSlide > 0) {
        pptCurrentSlide--;
        renderPPTSlide();
        viewerMessage.textContent = "Moved to the previous slide.";
    }
}

nextButton.addEventListener("click", nextPage);
footerNext.addEventListener("click", nextPage);

previousButton.addEventListener("click", previousPage);
footerPrevious.addEventListener("click", previousPage);

/* ---------- Fullscreen ---------- */

async function enterFullscreen() {
    try {
        if (document.fullscreenElement) {
            return;
        }

        if (presentationPanel.requestFullscreen) {
            await presentationPanel.requestFullscreen();
        } else if (presentationPanel.webkitRequestFullscreen) {
            presentationPanel.webkitRequestFullscreen();
        } else {
            showError("Fullscreen is not supported in this browser.");
        }
    } catch (error) {
        console.warn("Fullscreen request was blocked:", error);

        viewerMessage.textContent =
            "Click the Fullscreen button to enable presentation mode.";
    }
}

async function exitFullscreen() {
    try {
        if (document.fullscreenElement && document.exitFullscreen) {
            await document.exitFullscreen();
        } else if (document.webkitFullscreenElement &&
                   document.webkitExitFullscreen) {
            document.webkitExitFullscreen();
        }
    } catch (error) {
        console.warn("Could not exit fullscreen:", error);
    }
}

function toggleFullscreen() {
    if (document.fullscreenElement || document.webkitFullscreenElement) {
        exitFullscreen();
    } else {
        enterFullscreen();
    }
}

fullscreenButton.addEventListener("click", toggleFullscreen);

document.addEventListener("fullscreenchange", () => {
    const isFullscreen = Boolean(document.fullscreenElement);

    fullscreenButton.textContent = isFullscreen
        ? "⛶ Exit Fullscreen"
        : "⛶ Fullscreen";

    setTimeout(() => {
        if (currentMode === "pdf" && pdfDocument) {
            renderPDFPage(currentPage);
        }
    }, 100);
});

/* ---------- Keyboard controls ---------- */

document.addEventListener("keydown", (event) => {
    const target = event.target;

    if (
        target &&
        (target.tagName === "INPUT" ||
         target.tagName === "TEXTAREA" ||
         target.isContentEditable)
    ) {
        return;
    }

    if (event.key === "ArrowRight" || event.key === "PageDown") {
        event.preventDefault();
        nextPage();
    } else if (
        event.key === "ArrowLeft" ||
        event.key === "PageUp"
    ) {
        event.preventDefault();
        previousPage();
    } else if (event.key.toLowerCase() === "f") {
        toggleFullscreen();
    } else if (event.key === "Escape") {
        exitFullscreen();
    }
});

/* ---------- Camera and MediaPipe ---------- */

cameraButton.addEventListener("click", async () => {
    if (cameraRunning || cameraStarting) {
        stopCamera();
    } else {
        await startCamera();
    }
});

async function startCamera() {
    if (cameraStarting || cameraRunning) {
        return;
    }

    cameraStarting = true;
    cameraButton.disabled = true;

    cameraStatus.textContent = "Requesting camera permission...";
    setStatus("Starting camera");

    try {
        if (!navigator.mediaDevices?.getUserMedia) {
            throw new Error(
                "Camera access is unavailable. Use localhost or HTTPS in a supported browser."
            );
        }

        if (typeof Hands === "undefined") {
            throw new Error(
                "MediaPipe failed to load. Check your internet connection and refresh the page."
            );
        }

        cameraStream = await navigator.mediaDevices.getUserMedia({
            audio: false,
            video: {
                facingMode: "user",
                width: { ideal: 640 },
                height: { ideal: 480 },
                frameRate: { ideal: 30, max: 30 }
            }
        });

        inputVideo.srcObject = cameraStream;

        await new Promise((resolve, reject) => {
            if (inputVideo.readyState >= 2) {
                resolve();
                return;
            }

            inputVideo.onloadedmetadata = resolve;

            inputVideo.onerror = () => {
                reject(new Error("Could not initialize the camera video."));
            };
        });

        await inputVideo.play();

        handsModel = new Hands({
            locateFile: (file) =>
                `https://cdn.jsdelivr.net/npm/@mediapipe/hands/${file}`
        });

        handsModel.setOptions({
            maxNumHands: 1,
            modelComplexity: 0,
            minDetectionConfidence: 0.6,
            minTrackingConfidence: 0.6
        });

        handsModel.onResults(handleHandResults);

        cameraRunning = true;

        cameraPlaceholder.classList.add("hidden");
        cameraLiveLabel.textContent = "CAMERA LIVE";

        cameraButton.textContent = "■ Stop Camera";
        cameraStatus.textContent =
            "Camera is running. Show one hand clearly.";

        setStatus("Camera active");

        resizeCameraCanvas();

        cameraLoop();

    } catch (error) {
        console.error("Camera startup error:", error);

        stopCamera();

        cameraStatus.textContent =
            error.message || "Camera could not be started.";

        setStatus("Camera unavailable");

        showError(
            "Camera could not start. Allow camera access in your browser settings, " +
            "check that another app is not using it, and open this project using VS Code Live Server."
        );

    } finally {
        cameraStarting = false;
        cameraButton.disabled = false;
    }
}

function stopCamera() {
    cameraRunning = false;
    cameraFrameBusy = false;

    if (animationFrameId !== null) {
        cancelAnimationFrame(animationFrameId);
        animationFrameId = null;
    }

    if (cameraStream) {
        cameraStream.getTracks().forEach((track) => track.stop());
        cameraStream = null;
    }

    if (inputVideo) {
        inputVideo.pause();
        inputVideo.srcObject = null;
    }

    if (handsModel) {
        try {
            handsModel.close();
        } catch (_) {
            // Ignore cleanup errors.
        }

        handsModel = null;
    }

    cameraContext.clearRect(
        0,
        0,
        cameraCanvas.width,
        cameraCanvas.height
    );

    cameraPlaceholder.classList.remove("hidden");
    cameraLiveLabel.textContent = "CAMERA OFF";

    cameraButton.textContent = "▶ Start Camera";

    if (!cameraStarting) {
        cameraStatus.textContent = "Camera is stopped.";
        setStatus("System Ready");
    }

    detectedGesture.textContent = "None";

    lastGesture = "None";
    candidateGesture = "None";
    candidateSince = 0;
}

function resizeCameraCanvas() {
    const rect = cameraCanvas.getBoundingClientRect();

    const width = Math.max(1, Math.round(rect.width));
    const height = Math.max(1, Math.round(rect.height));

    if (cameraCanvas.width !== width || cameraCanvas.height !== height) {
        cameraCanvas.width = width;
        cameraCanvas.height = height;
    }
}

async function cameraLoop() {
    if (!cameraRunning || !handsModel) {
        return;
    }

    if (
        inputVideo.readyState >= 2 &&
        !cameraFrameBusy
    ) {
        cameraFrameBusy = true;

        try {
            await handsModel.send({
                image: inputVideo
            });
        } catch (error) {
            console.error("Hand tracking error:", error);
        } finally {
            cameraFrameBusy = false;
        }
    }

    if (cameraRunning) {
        animationFrameId = requestAnimationFrame(cameraLoop);
    }
}

/* ---------- Hand landmark helpers ---------- */

function distance2D(pointA, pointB) {
    const dx = pointA.x - pointB.x;
    const dy = pointA.y - pointB.y;

    return Math.sqrt(dx * dx + dy * dy);
}

function isFingerExtended(landmarks, tipIndex, pipIndex, mcpIndex) {
    const tip = landmarks[tipIndex];
    const pip = landmarks[pipIndex];
    const mcp = landmarks[mcpIndex];
    const wrist = landmarks[0];

    const tipDistance = distance2D(tip, wrist);
    const pipDistance = distance2D(pip, wrist);
    const mcpDistance = distance2D(mcp, wrist);

    return (
        tipDistance > pipDistance * 1.08 &&
        tipDistance > mcpDistance * 1.15
    );
}

function classifyHandGesture(landmarks) {
    /*
       MediaPipe landmark indexes:
       Thumb:  4
       Index:  8, 6, 5
       Middle: 12, 10, 9
       Ring:   16, 14, 13
       Pinky:  20, 18, 17
    */

    const indexExtended = isFingerExtended(landmarks, 8, 6, 5);
    const middleExtended = isFingerExtended(landmarks, 12, 10, 9);
    const ringExtended = isFingerExtended(landmarks, 16, 14, 13);
    const pinkyExtended = isFingerExtended(landmarks, 20, 18, 17);

    const extendedCount = [
        indexExtended,
        middleExtended,
        ringExtended,
        pinkyExtended
    ].filter(Boolean).length;

    const thumbTip = landmarks[4];
    const thumbIP = landmarks[3];
    const thumbMCP = landmarks[2];
    const wrist = landmarks[0];

    const thumbExtended =
        distance2D(thumbTip, wrist) >
        distance2D(thumbIP, wrist) * 1.04 &&
        distance2D(thumbTip, wrist) >
        distance2D(thumbMCP, wrist) * 1.1;

    // All four fingers extended = open palm.
    if (extendedCount >= 4) {
        return "Open Palm";
    }

    // All fingers folded, including the thumb = closed fist.
    if (extendedCount === 0 && !thumbExtended) {
        return "Closed Fist";
    }

    /*
       Pointing gesture:
       Index finger extended and the other three fingers folded.
       Direction is inferred from the index fingertip position.
    */

    if (
        indexExtended &&
        !middleExtended &&
        !ringExtended &&
        !pinkyExtended
    ) {
        const indexTip = landmarks[8];
        const indexBase = landmarks[5];

        const horizontalMovement = indexTip.x - indexBase.x;
        const verticalMovement = indexTip.y - indexBase.y;

        // Avoid treating an upward-pointing finger as left/right.
        if (Math.abs(horizontalMovement) > Math.abs(verticalMovement) * 0.7) {
            /*
               MediaPipe uses mirrored camera coordinates in many setups.
               This mapping matches the mirrored preview.
            */

            if (horizontalMovement < -0.12) {
                return "Point Right";
            }

            if (horizontalMovement > 0.12) {
                return "Point Left";
            }
        }
    }

    return "Other";
}

/* ---------- Hand tracking results ---------- */

function handleHandResults(results) {
    resizeCameraCanvas();

    cameraContext.clearRect(
        0,
        0,
        cameraCanvas.width,
        cameraCanvas.height
    );

    const landmarks = results.multiHandLandmarks?.[0];

    if (!landmarks) {
        detectedGesture.textContent = "No hand";
        candidateGesture = "None";
        candidateSince = 0;
        lastGesture = "None";
        return;
    }

    if (typeof drawConnectors === "function" && typeof HAND_CONNECTIONS !== "undefined") {
        drawConnectors(
            cameraContext,
            landmarks,
            HAND_CONNECTIONS,
            {
                color: "#00F5C8",
                lineWidth: 3
            }
        );

        drawLandmarks(
            cameraContext,
            landmarks,
            {
                color: "#FFFFFF",
                lineWidth: 1,
                radius: 2
            }
        );
    }

    const gesture = classifyHandGesture(landmarks);

    detectedGesture.textContent = gesture;

    processStableGesture(gesture);
}

/* ---------- Gesture stabilization ---------- */

function processStableGesture(gesture) {
    const now = performance.now();

    const actionableGestures = [
        "Open Palm",
        "Closed Fist",
        "Point Right",
        "Point Left"
    ];

    if (!actionableGestures.includes(gesture)) {
        candidateGesture = "None";
        candidateSince = 0;
        lastGesture = "None";
        return;
    }

    if (gesture !== candidateGesture) {
        candidateGesture = gesture;
        candidateSince = now;
        return;
    }

    if (now - candidateSince < gestureStableTime) {
        return;
    }

    // Only trigger once until the hand changes to another gesture.
    if (lastGesture === gesture) {
        return;
    }

    if (now - lastActionTime < gestureCooldown) {
        return;
    }

    lastGesture = gesture;
    lastActionTime = now;

    performGestureAction(gesture);
}

function performGestureAction(gesture) {
    switch (gesture) {
        case "Open Palm":
            enterFullscreen();
            viewerMessage.textContent = "Open palm detected.";
            break;

        case "Closed Fist":
            exitFullscreen();
            viewerMessage.textContent = "Closed fist detected.";
            break;

        case "Point Right":
            nextPage();
            break;

        case "Point Left":
            previousPage();
            break;
    }
}

/* ---------- Cleanup ---------- */

function stopExistingPresentation() {
    renderGeneration++;

    if (renderTask) {
        try {
            renderTask.cancel();
        } catch (_) {
            // Ignore cancellation errors.
        }

        renderTask = null;
    }

    pdfDocument = null;
    totalPages = 0;
    currentPage = 1;

    pptSlides = [];
    pptCurrentSlide = 0;

    pdfContext.clearRect(
        0,
        0,
        pdfCanvas.width,
        pdfCanvas.height
    );

    pptContainer.innerHTML = "";

    if (currentObjectUrl) {
        URL.revokeObjectURL(currentObjectUrl);
        currentObjectUrl = null;
    }
}

/* ---------- Responsive PDF rendering ---------- */

let resizeTimer = null;

window.addEventListener("resize", () => {
    clearTimeout(resizeTimer);

    resizeTimer = setTimeout(() => {
        if (currentMode === "pdf" && pdfDocument) {
            renderPDFPage(currentPage);
        }
    }, 180);

    if (cameraRunning) {
        resizeCameraCanvas();
    }
});

/* ---------- Initial state ---------- */

showViewer("none");

updatePresentationTitle(
    "Presentation Viewer",
    "Upload a file to begin"
);

updatePageCounter();

viewerMessage.textContent = "Ready to present";

setStatus("System Ready");