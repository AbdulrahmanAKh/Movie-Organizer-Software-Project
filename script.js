const uploadBtn = document.getElementById('uploadBtn');
const fileInput = document.getElementById('fileInput');
const viewerContent = document.getElementById('viewerContent');
const fullscreenBtn = document.getElementById('fullscreenBtn');
const slideNav = document.getElementById('slideNav');
const slideCounter = document.getElementById('slideCounter');
const prevBtn = document.getElementById('prevBtn');
const nextBtn = document.getElementById('nextBtn');

let currentSlide = 0;
let slideImages = [];
let zip = null;

uploadBtn.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.name.endsWith('.pptx')) {
        showError('Please select a .pptx file');
        return;
    }

    showLoading();
    try {
        const arrayBuffer = await file.arrayBuffer();
        await parsePPTX(arrayBuffer);
    } catch (error) {
        showError(`Error loading file: ${error.message}`);
        console.error(error);
    }
});

async function parsePPTX(arrayBuffer) {
    try {
        zip = await JSZip.loadAsync(arrayBuffer);
        slideImages = [];

        const slidesFolder = zip.folder('ppt/slides');
        if (!slidesFolder) throw new Error('Invalid PPTX: no slides found');

        const slideFiles = [];
        slidesFolder.forEach((path, file) => {
            if (/^slide\d+\.xml$/.test(path)) slideFiles.push({ name: path, file });
        });
        slideFiles.sort((a, b) => parseInt(a.name.match(/\d+/)) - parseInt(b.name.match(/\d+/)));

        if (slideFiles.length === 0) throw new Error('No slides found');

        for (let i = 0; i < slideFiles.length; i++) {
            const slideImage = await extractSlideImage(i + 1);
            slideImages.push(slideImage);
        }

        currentSlide = 0;
        showSlide(currentSlide);
        slideNav.style.display = 'flex';
        fullscreenBtn.style.display = 'inline-flex';
        slideCounter.style.display = 'block';
        updateSlideCounter();
    } catch (error) {
        showError(`Error parsing PPTX: ${error.message}`);
        console.error(error);
    }
}

async function extractSlideImage(num) {
    try {
        const relsPath = `ppt/slides/_rels/slide${num}.xml.rels`;
        const relsFile = zip.file(relsPath);
        if (!relsFile) return null;

        const relsContent = await relsFile.async('string');
        const relsDoc = new DOMParser().parseFromString(relsContent, 'text/xml');
        const rels = relsDoc.querySelectorAll('Relationship');

        for (const rel of rels) {
            const type = rel.getAttribute('Type');
            const target = rel.getAttribute('Target');
            if (type.includes('image') && target) {
                const imagePath = target.replace('../', 'ppt/');
                const imgFile = zip.file(imagePath);
                if (imgFile) {
                    const data = await imgFile.async('base64');
                    const ext = imagePath.split('.').pop();
                    const mime = ext === 'png' ? 'image/png' : 'image/jpeg';
                    return `data:${mime};base64,${data}`;
                }
            }
        }
        return null;
    } catch (err) {
        console.error(`Slide ${num} image error:`, err);
        return null;
    }
}

function showSlide(index) {
    if (!slideImages.length) return;
    const img = slideImages[index];
    viewerContent.innerHTML = img
        ? `<div class="slide-display"><div class="slide-content"><img src="${img}" alt="Slide ${index + 1}"></div></div>`
        : `<p style="color:#666;text-align:center;">Slide ${index + 1}<br>No preview</p>`;
    updateSlideCounter();
}

function updateSlideCounter() {
    slideCounter.textContent = `Slide ${currentSlide + 1} / ${slideImages.length}`;
}

function showLoading() {
    viewerContent.innerHTML = '<div class="loading">Loading presentation...</div>';
}

function showError(msg) {
    viewerContent.innerHTML = `<div class="placeholder"><div class="error-message">${msg}</div></div>`;
}

prevBtn.addEventListener('click', () => {
    if (currentSlide > 0) showSlide(--currentSlide);
});
nextBtn.addEventListener('click', () => {
    if (currentSlide < slideImages.length - 1) showSlide(++currentSlide);
});

document.addEventListener('keydown', (e) => {
    if (!slideImages.length) return;
    if (e.key === 'ArrowLeft' && currentSlide > 0) showSlide(--currentSlide);
    if (e.key === 'ArrowRight' && currentSlide < slideImages.length - 1) showSlide(++currentSlide);
});

fullscreenBtn.addEventListener('click', () => {
    if (!document.fullscreenElement) {
        viewerContent.requestFullscreen();
        fullscreenBtn.innerHTML = `
            <svg class="icon" viewBox="0 0 24 24">
                <path d="M5 16h3v3h2v-5H5v2zm3-8H5v2h5V5H8v3zm6 11h2v-3h3v-2h-5v5zm2-11V5h-2v5h5V8h-3z"/>
            </svg> Exit Fullscreen`;
    } else {
        document.exitFullscreen();
        fullscreenBtn.innerHTML = `
            <svg class="icon" viewBox="0 0 24 24">
                <path d="M7 14H5v5h5v-2H7v-3zm-2-4h2V7h3V5H5v5zm12 7h-3v2h5v-5h-2v3zM14 5v2h3v3h2V5h-5z"/>
            </svg> Fullscreen`;
    }
});

document.addEventListener('fullscreenchange', () => {
    if (!document.fullscreenElement) {
        fullscreenBtn.innerHTML = `
            <svg class="icon" viewBox="0 0 24 24">
                <path d="M7 14H5v5h5v-2H7v-3zm-2-4h2V7h3V5H5v5zm12 7h-3v2h5v-5h-2v3zM14 5v2h3v3h2V5h-5z"/>
            </svg> Fullscreen`;
    }
});

function downloadProject() {
    window.open('https://www.mediafire.com/file/38i1lw6jokgsyum/FinalProject.rar/file', '_blank');
}
