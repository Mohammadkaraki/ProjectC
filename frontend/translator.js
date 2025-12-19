// Slide Translator JavaScript
// Handles file upload, API communication, and progress tracking

// Auto-detect environment: localhost vs production
const API_URL = (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1')
    ? 'http://localhost:5000/api/translate-slide'  // Local: Direct to Flask
    : '/api/translate-slide';                        // Production: Via nginx proxy

// DOM Elements
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const translateBtn = document.getElementById('translateBtn');
const progressSection = document.getElementById('progressSection');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const resultSection = document.getElementById('resultSection');
const downloadLink = document.getElementById('downloadLink');
const translateAnotherBtn = document.getElementById('translateAnotherBtn');
const errorSection = document.getElementById('errorSection');
const errorText = document.getElementById('errorText');
const retryBtn = document.getElementById('retryBtn');

let selectedFile = null;

// File upload handlers
dropzone.addEventListener('click', () => {
    fileInput.click();
});

dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('drag-over');
});

dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('drag-over');
});

dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('drag-over');

    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFileSelect(files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFileSelect(e.target.files[0]);
    }
});

translateBtn.addEventListener('click', () => {
    if (selectedFile) {
        translateSlide(selectedFile);
    }
});

translateAnotherBtn.addEventListener('click', resetUploader);
retryBtn.addEventListener('click', resetUploader);

// Handle file selection
function handleFileSelect(file) {
    // Validate file type
    const validExtensions = ['.pptx'];
    const fileExt = file.name.toLowerCase().slice(file.name.lastIndexOf('.'));

    if (!validExtensions.includes(fileExt)) {
        alert('Please select a .pptx file');
        return;
    }

    // Validate file size (16 MB max)
    const maxSize = 16 * 1024 * 1024;
    if (file.size > maxSize) {
        alert('File size exceeds 16 MB limit');
        return;
    }

    selectedFile = file;

    // Display file info
    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    fileInfo.style.display = 'block';
    translateBtn.style.display = 'block';
    translateBtn.disabled = false;

    // Hide sections
    progressSection.style.display = 'none';
    resultSection.style.display = 'none';
    errorSection.style.display = 'none';
}

// Format file size
function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(2) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
}

// Translate slide
async function translateSlide(file) {
    // Hide upload UI
    translateBtn.style.display = 'none';
    progressSection.style.display = 'block';
    errorSection.style.display = 'none';
    resultSection.style.display = 'none';

    // Calculate estimated time based on file size
    const fileSizeMB = file.size / (1024 * 1024);
    const isLargeFile = fileSizeMB > 5;

    // Adjust timing based on file size (larger files = slower steps)
    const baseStepDuration = isLargeFile ? 2500 : 1800;

    // Enhanced progress steps with more detail
    const steps = [
        { progress: 5, text: 'Initializing translation... Please be patient', duration: 500 },
        { progress: 15, text: 'Uploading file to server...', duration: baseStepDuration * 0.5 },
        { progress: 25, text: 'Parsing slide structure and extracting text...', duration: baseStepDuration * 0.8 },
        { progress: 40, text: 'Building context map and analyzing hierarchy...', duration: baseStepDuration * 0.6 },
        { progress: 55, text: 'Translating with AI (this may take a moment)...', duration: baseStepDuration * 1.5 },
        { progress: 70, text: 'Processing translations and formatting...', duration: baseStepDuration * 0.7 },
        { progress: 80, text: 'Converting layout to RTL (Right-to-Left)...', duration: baseStepDuration * 0.8 },
        { progress: 90, text: 'Translating layout backgrounds...', duration: baseStepDuration * 1.2 },
        { progress: 95, text: 'Finalizing and packaging output file...', duration: baseStepDuration * 0.5 }
    ];

    // Add patience message for large files
    if (isLargeFile) {
        const patientMsg = document.createElement('p');
        patientMsg.className = 'patience-message';
        patientMsg.style.cssText = 'color: #666; font-size: 14px; margin-top: 10px; font-style: italic;';
        patientMsg.textContent = `⏱️ Large file detected (${fileSizeMB.toFixed(1)} MB). Translation may take 30-60 seconds. Thank you for your patience!`;
        progressSection.appendChild(patientMsg);
    }

    let currentStep = 0;

    // Start progress animation with dynamic timing
    function updateProgress() {
        if (currentStep < steps.length) {
            const step = steps[currentStep];
            progressFill.style.width = step.progress + '%';
            progressText.textContent = step.text;
            currentStep++;

            // Schedule next step with its specific duration
            if (currentStep < steps.length) {
                setTimeout(updateProgress, steps[currentStep].duration);
            }
        }
    }

    // Start the progress animation
    updateProgress();

    try {
        // Create form data
        const formData = new FormData();
        formData.append('file', file);

        // Call API
        const response = await fetch(API_URL, {
            method: 'POST',
            body: formData
        });

        // Translation completed successfully
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.message || 'Translation failed');
        }

        // Get translated file
        const blob = await response.blob();
        const downloadUrl = URL.createObjectURL(blob);

        // Complete progress
        progressFill.style.width = '100%';
        progressText.textContent = 'Translation complete! ✓';

        // Remove patience message if exists
        const patientMsg = document.querySelector('.patience-message');
        if (patientMsg) patientMsg.remove();

        // Show result
        setTimeout(() => {
            progressSection.style.display = 'none';
            resultSection.style.display = 'block';
            downloadLink.href = downloadUrl;
        }, 1000);

    } catch (error) {
        console.error('Translation error:', error);

        // Show error
        progressSection.style.display = 'none';
        errorSection.style.display = 'block';
        errorText.textContent = error.message || 'An unexpected error occurred. Please ensure the backend server is running.';
    }
}

// Reset uploader
function resetUploader() {
    selectedFile = null;
    fileInput.value = '';
    fileInfo.style.display = 'none';
    translateBtn.style.display = 'none';
    progressSection.style.display = 'none';
    resultSection.style.display = 'none';
    errorSection.style.display = 'none';
    progressFill.style.width = '0%';

    // Remove patience message if exists
    const patientMsg = document.querySelector('.patience-message');
    if (patientMsg) patientMsg.remove();
}
