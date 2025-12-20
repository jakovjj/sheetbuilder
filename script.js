class SheetBuilder {
    constructor() {
        this.images = [];
        this.layout = [];
        this.selectedImages = new Set();
        this.lastSelectedId = null;
        this.isRatioLocked = true;
        this.bulkAspectRatio = null;
        this.lastPreviewParams = null;
        this._previewRerenderTimer = null;
        this._exportState = {
            inProgress: false,
            abortController: null
        };
        this.initializeEventListeners();
        this.ensureDefaultPaperSettings();
    }

    getExportUi() {
        return {
            overlay: document.getElementById('exportOverlay'),
            status: document.getElementById('exportStatus'),
            bar: document.getElementById('exportProgressBar'),
            text: document.getElementById('exportProgressText'),
            cancelBtn: document.getElementById('cancelExport')
        };
    }

    setExportOverlayVisible(visible) {
        const { overlay } = this.getExportUi();
        if (!overlay) return;
        overlay.classList.toggle('hidden', !visible);
    }

    setExportProgress(current, total, statusText) {
        const { status, bar, text } = this.getExportUi();
        const safeTotal = Math.max(1, Number.isFinite(total) ? total : 1);
        const safeCurrent = Math.min(safeTotal, Math.max(0, Number.isFinite(current) ? current : 0));
        const percent = Math.round((safeCurrent / safeTotal) * 100);
        if (status && typeof statusText === 'string') status.textContent = statusText;
        if (bar) bar.style.width = `${percent}%`;
        if (text) text.textContent = `${percent}%`;
    }

    cancelExport() {
        if (!this._exportState?.inProgress) return;
        const { cancelBtn } = this.getExportUi();
        if (cancelBtn) {
            cancelBtn.disabled = true;
            cancelBtn.textContent = 'Canceling…';
        }
        try {
            this._exportState.abortController?.abort();
        } catch {
            // ignore
        }
    }

    ensureDefaultPaperSettings() {
        const paperSize = document.getElementById('paperSize');
        const paperWidthEl = document.getElementById('paperWidth');
        const paperHeightEl = document.getElementById('paperHeight');
        const outerMarginEl = document.getElementById('outerMargin');
        if (!paperSize || !paperWidthEl || !paperHeightEl || !outerMarginEl) return;

        // Default to A4 unless the user explicitly chose something else.
        if (!paperSize.value) paperSize.value = 'a4';

        const selected = paperSize.value;
        if (selected === 'a4') {
            const currentW = parseFloat(paperWidthEl.value);
            const currentH = parseFloat(paperHeightEl.value);
            if (!Number.isFinite(currentW) || !Number.isFinite(currentH) || Math.abs(currentW - 210) > 0.1 || Math.abs(currentH - 297) > 0.1) {
                paperWidthEl.value = 210;
                paperHeightEl.value = 297;
            }
        }

        const w = parseFloat(paperWidthEl.value);
        const h = parseFloat(paperHeightEl.value);
        const outer = parseFloat(outerMarginEl.value);
        const maxOuter = Number.isFinite(w) && Number.isFinite(h) ? Math.max(0, (Math.min(w, h) / 2) - 0.1) : 0;
        if (!Number.isFinite(outer) || outer < 0 || outer > maxOuter) {
            outerMarginEl.value = 10;
        }

        this.checkForCustomSize();
        this.validateAllImages();
    }

    async createPreviewDataUrl(file, options = {}) {
        const {
            maxDimension = 360,
            quality = 0.55,
            mimeType = 'image/jpeg'
        } = options;

        // If the browser doesn't support createImageBitmap, fall back to original.
        if (!('createImageBitmap' in window)) {
            return null;
        }

        try {
            const bitmap = await createImageBitmap(file);
            const scale = Math.min(1, maxDimension / Math.max(bitmap.width, bitmap.height));
            const width = Math.max(1, Math.round(bitmap.width * scale));
            const height = Math.max(1, Math.round(bitmap.height * scale));

            const canvas = document.createElement('canvas');
            canvas.width = width;
            canvas.height = height;
            const ctx = canvas.getContext('2d', { alpha: false });
            if (!ctx) return null;

            // JPEG doesn't support transparency; draw onto a white background
            // so transparent pixels don't become black in previews.
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, width, height);
            ctx.drawImage(bitmap, 0, 0, width, height);
            return canvas.toDataURL(mimeType, quality);
        } catch {
            return null;
        }
    }

    getPreviewSrcForPlacedImage(placedImage) {
        if (placedImage.previewDataUrl) return placedImage.previewDataUrl;
        if (placedImage.originalId) {
            const original = this.images.find(img => img.id === placedImage.originalId);
            if (original?.previewDataUrl) return original.previewDataUrl;
            if (original?.dataUrl) return original.dataUrl;
        }
        return placedImage.dataUrl;
    }

    scheduleLayoutPreviewRerender() {
        if (!this.lastPreviewParams || this.layout.length === 0) return;
        if (this._previewRerenderTimer) return;

        this._previewRerenderTimer = window.setTimeout(() => {
            this._previewRerenderTimer = null;
            const { paperWidth, paperHeight, outerMargin } = this.lastPreviewParams;
            this.renderLayoutPreview(paperWidth, paperHeight, outerMargin);
        }, 120);
    }

    drawImageContain(ctx, img, dx, dy, dw, dh) {
        const scale = Math.min(dw / img.width, dh / img.height);
        const drawW = img.width * scale;
        const drawH = img.height * scale;
        const x = dx + (dw - drawW) / 2;
        const y = dy + (dh - drawH) / 2;
        ctx.drawImage(img, x, y, drawW, drawH);
    }

    async drawPagePreviewToCanvas(canvas, page, paperWidth, paperHeight, outerMargin, scale) {
        const dpr = window.devicePixelRatio || 1;
        const cssWidth = canvas.clientWidth;
        const cssHeight = canvas.clientHeight;

        canvas.width = Math.max(1, Math.round(cssWidth * dpr));
        canvas.height = Math.max(1, Math.round(cssHeight * dpr));

        const ctx = canvas.getContext('2d');
        if (!ctx) return;
        ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

        // Background
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, cssWidth, cssHeight);

        // Optional outline of full page and printable area
        ctx.strokeStyle = 'rgba(220, 38, 38, 0.35)';
        ctx.lineWidth = 1;
        ctx.strokeRect(0.5, 0.5, cssWidth - 1, cssHeight - 1);

        const printableX = outerMargin * scale;
        const printableY = outerMargin * scale;
        const printableW = (paperWidth - 2 * outerMargin) * scale;
        const printableH = (paperHeight - 2 * outerMargin) * scale;
        ctx.strokeStyle = 'rgba(0, 0, 0, 0.15)';
        ctx.setLineDash([4, 3]);
        ctx.strokeRect(printableX + 0.5, printableY + 0.5, printableW - 1, printableH - 1);
        ctx.setLineDash([]);

        for (const placed of page.images) {
            const x = (outerMargin + placed.x) * scale;
            const y = (outerMargin + placed.y) * scale;
            const w = placed.width * scale;
            const h = placed.height * scale;

            // Box outline
            ctx.strokeStyle = 'rgba(220, 38, 38, 0.7)';
            ctx.lineWidth = 1;
            ctx.strokeRect(x + 0.5, y + 0.5, Math.max(0, w - 1), Math.max(0, h - 1));

            const src = this.getPreviewSrcForPlacedImage(placed);
            if (!src) continue;

            const img = new Image();
            img.decoding = 'async';
            img.src = src;

            try {
                if (img.decode) {
                    await img.decode();
                } else {
                    await new Promise((resolve, reject) => {
                        img.onload = resolve;
                        img.onerror = reject;
                    });
                }
            } catch {
                continue;
            }

            if (placed.rotated) {
                ctx.save();
                ctx.translate(x + w / 2, y + h / 2);
                ctx.rotate(Math.PI / 2);
                // After 90° rotation, the available box dimensions swap.
                this.drawImageContain(ctx, img, -h / 2, -w / 2, h, w);
                ctx.restore();
            } else {
                this.drawImageContain(ctx, img, x, y, w, h);
            }
        }
    }

    initializeEventListeners() {
        const imageInput = document.getElementById('imageInput');
        const generateBtn = document.getElementById('generateLayout');
        const exportBtn = document.getElementById('exportPDF');
        const paperSize = document.getElementById('paperSize');
        const cancelExportBtn = document.getElementById('cancelExport');
        
        // Paper setting inputs for real-time validation
        const paperWidth = document.getElementById('paperWidth');
        const paperHeight = document.getElementById('paperHeight');
        const outerMargin = document.getElementById('outerMargin');
        
        // Bulk edit controls
        const selectAllBtn = document.getElementById('selectAll');
        const selectNoneBtn = document.getElementById('selectNone');
        const removeSelectedBtn = document.getElementById('removeSelected');
        const applyBulkChangesBtn = document.getElementById('applyBulkChanges');
        const lockRatioBtn = document.getElementById('lockRatio');
        
        // Quick selection controls
        const quickSelectAllBtn = document.getElementById('quickSelectAll');
        const quickSelectNoneBtn = document.getElementById('quickSelectNone');
        
        // Bulk edit inputs
        const bulkWidth = document.getElementById('bulkWidth');
        const bulkHeight = document.getElementById('bulkHeight');
        const bulkCopies = document.getElementById('bulkCopies');

        imageInput.addEventListener('change', (e) => this.handleImageUpload(e));
        generateBtn.addEventListener('click', () => this.generateLayout());
        exportBtn.addEventListener('click', () => this.exportToPDF());
        paperSize.addEventListener('change', (e) => this.handlePaperSizeChange(e));

        if (cancelExportBtn) cancelExportBtn.addEventListener('click', () => this.cancelExport());
        
        // Bulk edit event listeners
        if (selectAllBtn) selectAllBtn.addEventListener('click', () => this.selectAllImages());
        if (selectNoneBtn) selectNoneBtn.addEventListener('click', () => this.selectNoneImages());
        if (removeSelectedBtn) removeSelectedBtn.addEventListener('click', () => this.removeSelectedImages());
        if (applyBulkChangesBtn) applyBulkChangesBtn.addEventListener('click', () => this.applyBulkChanges());
        if (lockRatioBtn) lockRatioBtn.addEventListener('click', () => this.toggleRatioLock());
        
        // Quick selection event listeners
        if (quickSelectAllBtn) quickSelectAllBtn.addEventListener('click', () => this.selectAllImages());
        if (quickSelectNoneBtn) quickSelectNoneBtn.addEventListener('click', () => this.selectNoneImages());
        
        // Bulk input listeners for real-time ratio locking
        if (bulkWidth) bulkWidth.addEventListener('input', (e) => this.handleBulkWidthChange(e));
        if (bulkHeight) bulkHeight.addEventListener('input', (e) => this.handleBulkHeightChange(e));
        
        // Drag and drop functionality
        const fileUpload = document.querySelector('.file-upload');
        if (fileUpload) {
            fileUpload.addEventListener('dragover', (e) => this.handleDragOver(e));
            fileUpload.addEventListener('drop', (e) => this.handleDrop(e));
            fileUpload.addEventListener('dragenter', (e) => this.handleDragEnter(e));
            fileUpload.addEventListener('dragleave', (e) => this.handleDragLeave(e));
        }
        
        // Clipboard paste functionality
        document.addEventListener('paste', (e) => this.handlePaste(e));
        
        // Add real-time validation when paper settings change
        paperWidth.addEventListener('input', () => {
            this.validateAllImages();
            this.checkForCustomSize();
        });
        paperHeight.addEventListener('input', () => {
            this.validateAllImages();
            this.checkForCustomSize();
        });
        outerMargin.addEventListener('input', () => this.validateAllImages());
    }

    handlePaperSizeChange(event) {
        const selectedSize = event.target.value;
        const paperSizes = {
            'a4': { width: 210, height: 297 },
            'a3': { width: 297, height: 420 },
            'a5': { width: 148, height: 210 },
            'letter': { width: 216, height: 279 },
            'legal': { width: 216, height: 356 },
            'tabloid': { width: 279, height: 432 },
            'photo4x6': { width: 102, height: 152 },
            'photo5x7': { width: 127, height: 178 },
            'photo8x10': { width: 203, height: 254 }
        };

        if (selectedSize !== 'custom' && paperSizes[selectedSize]) {
            const size = paperSizes[selectedSize];
            document.getElementById('paperWidth').value = size.width;
            document.getElementById('paperHeight').value = size.height;

            // If the current outer margin is invalid for this preset (or was set to something huge),
            // reset it to a sane default so users don't get negative printable areas.
            const outerMarginInput = document.getElementById('outerMargin');
            if (outerMarginInput) {
                const currentOuter = parseFloat(outerMarginInput.value);
                const maxOuter = Math.max(0, (Math.min(size.width, size.height) / 2) - 0.1);
                if (!Number.isFinite(currentOuter) || currentOuter > maxOuter) {
                    outerMarginInput.value = 10;
                }
            }
            
            // Trigger validation after changing paper size
            this.validateAllImages();
        }
    }

    checkForCustomSize() {
        const currentWidth = parseFloat(document.getElementById('paperWidth').value);
        const currentHeight = parseFloat(document.getElementById('paperHeight').value);
        const paperSizeSelect = document.getElementById('paperSize');
        
        const paperSizes = {
            'a4': { width: 210, height: 297 },
            'a3': { width: 297, height: 420 },
            'a5': { width: 148, height: 210 },
            'letter': { width: 216, height: 279 },
            'legal': { width: 216, height: 356 },
            'tabloid': { width: 279, height: 432 },
            'photo4x6': { width: 102, height: 152 },
            'photo5x7': { width: 127, height: 178 },
            'photo8x10': { width: 203, height: 254 }
        };

        // Check if current dimensions match any preset
        let matchingSize = 'custom';
        for (const [sizeName, dimensions] of Object.entries(paperSizes)) {
            if (Math.abs(dimensions.width - currentWidth) < 0.1 && 
                Math.abs(dimensions.height - currentHeight) < 0.1) {
                matchingSize = sizeName;
                break;
            }
        }

        paperSizeSelect.value = matchingSize;
    }

    validateAllImages() {
        this.images.forEach(image => this.validateImageInRealTime(image));
    }

    handleImageUpload(event) {
        const files = Array.from(event.target.files);
        this.processFiles(files);
    }

    renderImageConfig(imageData) {
        const list = document.getElementById('imagesList');
        const config = document.createElement('div');
        config.className = 'image-config';
        config.dataset.id = imageData.id;
        config.innerHTML = `
            <input type="checkbox" class="select-checkbox" id="select-${imageData.id}" 
                   onclick="sheetBuilder.handleCheckboxClick(event, ${imageData.id})">
            <img src="${imageData.dataUrl}" alt="${imageData.name}" 
                 onclick="sheetBuilder.showImageModal('${imageData.dataUrl}', '${imageData.name}', '${imageData.originalWidth}x${imageData.originalHeight}')">
            <div class="config-inputs">
                <div class="input-with-lock">
                    <div style="width: 100%;">
                        <label>Width (mm):</label>
                        <input type="number" value="${imageData.width.toFixed(1)}" min="1" step="0.1" 
                               onchange="sheetBuilder.updateImageSize(${imageData.id}, 'width', this.value)">
                    </div>
                </div>
                <div>
                    <label>Height (mm):</label>
                    <input type="number" value="${imageData.height.toFixed(1)}" min="1" step="0.1" 
                           onchange="sheetBuilder.updateImageSize(${imageData.id}, 'height', this.value)">
                </div>
                <div>
                    <label>Copies:</label>
                    <input type="number" value="${imageData.copies}" min="1" 
                           onchange="sheetBuilder.updateImageSize(${imageData.id}, 'copies', this.value)">
                </div>
            </div>
            <button class="remove-btn" onclick="sheetBuilder.removeImage(${imageData.id})">Remove</button>
        `;
        list.appendChild(config);
        
        // Show quick selection controls and update selection count
        this.updateQuickSelectionControls();
        this.updateSelectionCount();
    }

    removeImage(id) {
        // Remove from images array
        this.images = this.images.filter(img => img.id !== id);
        
        // Remove from selected images
        this.selectedImages.delete(id);
        
        // Remove from DOM
        const config = document.querySelector(`[data-id="${id}"]`);
        if (config) {
            config.remove();
        }
        
        // Update UI
        this.updateQuickSelectionControls();
        this.updateSelectionCount();
    }

    updateImageSize(id, property, value) {
        const image = this.images.find(img => img.id === id);
        if (image) {
            const numValue = parseFloat(value);
            
            if (property === 'width') {
                image.width = numValue;
                // Maintain aspect ratio
                image.height = numValue / image.aspectRatio;
                // Update the height input field
                const config = document.querySelector(`[data-id="${id}"]`);
                const heightInput = config.querySelector('input[onchange*="height"]');
                heightInput.value = image.height.toFixed(1);
            } else if (property === 'height') {
                image.height = numValue;
                // Maintain aspect ratio
                image.width = numValue * image.aspectRatio;
                // Update the width input field
                const config = document.querySelector(`[data-id="${id}"]`);
                const widthInput = config.querySelector('input[onchange*="width"]');
                widthInput.value = image.width.toFixed(1);
            } else {
                image[property] = numValue;
            }
            
            // Real-time validation
            this.validateImageInRealTime(image);
        }
    }

    validateImageInRealTime(image) {
        const paperWidth = parseFloat(document.getElementById('paperWidth').value);
        const paperHeight = parseFloat(document.getElementById('paperHeight').value);
        const outerMargin = parseFloat(document.getElementById('outerMargin').value);

        const paperIssue = this.getPaperSettingsIssue(paperWidth, paperHeight, outerMargin);
        
        const printableWidth = paperWidth - (2 * outerMargin);
        const printableHeight = paperHeight - (2 * outerMargin);
        
        const config = document.querySelector(`[data-id="${image.id}"]`);
        if (!config) return; // Safety check
        
        const existingWarning = config.querySelector('.size-warning');

        if (paperIssue) {
            if (existingWarning) existingWarning.remove();
            config.classList.add('oversized');
            const warning = document.createElement('div');
            warning.className = 'size-warning';
            warning.innerHTML = `<i class="fa-solid fa-triangle-exclamation" aria-hidden="true"></i> ${paperIssue}`;
            config.appendChild(warning);
            return;
        }
        
        const canFitNormal = image.width <= printableWidth && image.height <= printableHeight;
        const canFitRotated = image.height <= printableWidth && image.width <= printableHeight;
        
        if (!canFitNormal && !canFitRotated) {
            // Add warning if not exists
            if (!existingWarning) {
                config.classList.add('oversized');
                const warning = document.createElement('div');
                warning.className = 'size-warning';
                const maxW = Math.max(0, Math.max(printableWidth, printableHeight));
                const maxH = Math.max(0, Math.min(printableWidth, printableHeight));
                warning.innerHTML = `<i class="fa-solid fa-triangle-exclamation" aria-hidden="true"></i> Too large for paper! Max: ${maxW.toFixed(1)}×${maxH.toFixed(1)}mm`;
                config.appendChild(warning);
            }
        } else {
            // Remove warning if exists
            if (existingWarning) {
                existingWarning.remove();
                config.classList.remove('oversized');
            }
        }
    }

    getPaperSettingsIssue(paperWidth, paperHeight, outerMargin) {
        if (!Number.isFinite(paperWidth) || !Number.isFinite(paperHeight) || !Number.isFinite(outerMargin)) {
            return 'Invalid paper settings.';
        }
        if (paperWidth <= 0 || paperHeight <= 0) {
            return 'Paper width/height must be greater than 0.';
        }
        if (outerMargin < 0) {
            return 'Outer margin cannot be negative.';
        }

        const printableWidth = paperWidth - (2 * outerMargin);
        const printableHeight = paperHeight - (2 * outerMargin);
        if (printableWidth <= 0 || printableHeight <= 0) {
            return `Outer margin is too large for the selected paper size (printable area would be ${printableWidth.toFixed(1)}×${printableHeight.toFixed(1)}mm). Reduce the outer margin or increase paper size.`;
        }

        return null;
    }

    // Selection and Bulk Edit Methods
    handleCheckboxClick(event, id) {
        const checkbox = event.target;
        const isSelected = checkbox.checked;
        
        if (event.shiftKey && this.lastSelectedId !== null) {
            // Shift+click: select range
            this.selectRange(this.lastSelectedId, id, isSelected);
        } else if (event.ctrlKey || event.metaKey) {
            // Ctrl+click: toggle individual selection
            this.toggleImageSelection(id, isSelected);
        } else {
            // Normal click: just toggle this item
            this.toggleImageSelection(id, isSelected);
        }
        
        this.lastSelectedId = id;
    }

    selectRange(startId, endId, select = true) {
        const imageIds = this.images.map(img => img.id);
        const startIndex = imageIds.indexOf(startId);
        const endIndex = imageIds.indexOf(endId);
        
        if (startIndex === -1 || endIndex === -1) return;
        
        const minIndex = Math.min(startIndex, endIndex);
        const maxIndex = Math.max(startIndex, endIndex);
        
        for (let i = minIndex; i <= maxIndex; i++) {
            const imageId = imageIds[i];
            const checkbox = document.getElementById(`select-${imageId}`);
            const config = document.querySelector(`[data-id="${imageId}"]`);
            
            if (select) {
                this.selectedImages.add(imageId);
                if (checkbox) checkbox.checked = true;
                if (config) config.classList.add('selected');
            } else {
                this.selectedImages.delete(imageId);
                if (checkbox) checkbox.checked = false;
                if (config) config.classList.remove('selected');
            }
        }
        
        this.updateSelectionCount();
    }

    toggleImageSelection(id, isSelected) {
        const config = document.querySelector(`[data-id="${id}"]`);
        
        if (isSelected) {
            this.selectedImages.add(id);
            config.classList.add('selected');
        } else {
            this.selectedImages.delete(id);
            config.classList.remove('selected');
        }
        
        this.updateSelectionCount();
    }

    selectAllImages() {
        this.selectedImages.clear();
        
        this.images.forEach(img => {
            this.selectedImages.add(img.id);
            const checkbox = document.getElementById(`select-${img.id}`);
            const config = document.querySelector(`[data-id="${img.id}"]`);
            
            if (checkbox) checkbox.checked = true;
            if (config) config.classList.add('selected');
        });
        
        this.updateSelectionCount();
    }

    selectNoneImages() {
        this.selectedImages.clear();
        
        this.images.forEach(img => {
            const checkbox = document.getElementById(`select-${img.id}`);
            const config = document.querySelector(`[data-id="${img.id}"]`);
            
            if (checkbox) checkbox.checked = false;
            if (config) config.classList.remove('selected');
        });
        
        this.updateSelectionCount();
    }

    removeSelectedImages() {
        if (this.selectedImages.size === 0) {
            alert('No images selected for removal.');
            return;
        }

        if (confirm(`Remove ${this.selectedImages.size} selected image(s)?`)) {
            const idsToRemove = Array.from(this.selectedImages);
            idsToRemove.forEach(id => this.removeImage(id));
        }
    }

    updateSelectionCount() {
        const selectionCount = document.getElementById('selectionCount');
        if (selectionCount) {
            selectionCount.textContent = `${this.selectedImages.size} image(s) selected`;
        }
    }

    updateQuickSelectionControls() {
        const quickControls = document.getElementById('quickSelectionControls');
        if (quickControls) {
            if (this.images.length > 0) {
                quickControls.classList.remove('hidden');
            } else {
                quickControls.classList.add('hidden');
            }
        }
    }

    toggleBulkEdit() {
        const bulkControls = document.getElementById('bulkControls');
        const toggleIcon = document.getElementById('bulkToggleIcon');
        
        if (bulkControls.classList.contains('collapsed')) {
            bulkControls.classList.remove('collapsed');
            toggleIcon.classList.remove('collapsed');
        } else {
            bulkControls.classList.add('collapsed');
            toggleIcon.classList.add('collapsed');
        }
    }

    toggleRatioLock() {
        this.isRatioLocked = !this.isRatioLocked;
        const lockBtn = document.getElementById('lockRatio');
        if (!lockBtn) return;
        
        if (this.isRatioLocked) {
            lockBtn.classList.add('active');
            lockBtn.innerHTML = '<i class="fa-solid fa-lock" aria-hidden="true"></i>';
            lockBtn.title = 'Unlock aspect ratio';
        } else {
            lockBtn.classList.remove('active');
            lockBtn.innerHTML = '<i class="fa-solid fa-lock-open" aria-hidden="true"></i>';
            lockBtn.title = 'Lock aspect ratio';
        }
    }

    handleBulkWidthChange(event) {
        if (!this.isRatioLocked) return;
        
        const width = parseFloat(event.target.value);
        if (isNaN(width) || width <= 0) return;
        
        // Use average aspect ratio of selected images or a default ratio
        let aspectRatio = this.getAverageSelectedAspectRatio();
        if (!aspectRatio) aspectRatio = 1.5; // Default aspect ratio
        
        const height = width / aspectRatio;
        document.getElementById('bulkHeight').value = height.toFixed(1);
    }

    handleBulkHeightChange(event) {
        if (!this.isRatioLocked) return;
        
        const height = parseFloat(event.target.value);
        if (isNaN(height) || height <= 0) return;
        
        // Use average aspect ratio of selected images or a default ratio
        let aspectRatio = this.getAverageSelectedAspectRatio();
        if (!aspectRatio) aspectRatio = 1.5; // Default aspect ratio
        
        const width = height * aspectRatio;
        document.getElementById('bulkWidth').value = width.toFixed(1);
    }

    getAverageSelectedAspectRatio() {
        const selectedImages = this.images.filter(img => this.selectedImages.has(img.id));
        if (selectedImages.length === 0) return null;
        
        const totalAspectRatio = selectedImages.reduce((sum, img) => sum + img.aspectRatio, 0);
        return totalAspectRatio / selectedImages.length;
    }

    applyBulkChanges() {
        console.log('Apply bulk changes called');
        console.log('Selected images:', this.selectedImages);
        
        if (this.selectedImages.size === 0) {
            alert('No images selected for bulk changes.');
            return;
        }

        const bulkWidth = document.getElementById('bulkWidth').value;
        const bulkHeight = document.getElementById('bulkHeight').value;
        const bulkCopies = document.getElementById('bulkCopies').value;

        console.log('Bulk values:', { bulkWidth, bulkHeight, bulkCopies });

        const selectedImageData = this.images.filter(img => this.selectedImages.has(img.id));
        console.log('Selected image data:', selectedImageData.length, 'images');
        
        selectedImageData.forEach((img, index) => {
            console.log(`Processing image ${index + 1}:`, img.name);
            
            const oldValues = { width: img.width, height: img.height, copies: img.copies };
            
            if (bulkWidth) {
                img.width = parseFloat(bulkWidth);
                if (this.isRatioLocked) {
                    // Use the original image aspect ratio, not bulk aspect ratio
                    img.height = img.width / img.aspectRatio;
                }
            }
            
            if (bulkHeight) {
                if (this.isRatioLocked) {
                    // If ratio locked and height is provided, adjust width to maintain ratio
                    img.height = parseFloat(bulkHeight);
                    img.width = img.height * img.aspectRatio;
                } else {
                    // If ratio not locked, just set height
                    img.height = parseFloat(bulkHeight);
                }
            }
            
            if (bulkCopies) {
                img.copies = parseInt(bulkCopies);
            }

            console.log(`Image ${img.name} changed from:`, oldValues, 'to:', 
                       { width: img.width, height: img.height, copies: img.copies });

            // Update the individual controls
            this.updateImageConfigDisplay(img);
            this.validateImageInRealTime(img);
        });

        // Clear bulk inputs after successful application
        document.getElementById('bulkWidth').value = '';
        document.getElementById('bulkHeight').value = '';
        document.getElementById('bulkCopies').value = '';
        
        // Show success message
        alert(`Applied bulk changes to ${selectedImageData.length} image(s).`);
    }

    updateImageConfigDisplay(imageData) {
        const config = document.querySelector(`[data-id="${imageData.id}"]`);
        if (!config) return;

        const widthInput = config.querySelector('input[onchange*="width"]');
        const heightInput = config.querySelector('input[onchange*="height"]');
        const copiesInput = config.querySelector('input[onchange*="copies"]');

        if (widthInput) widthInput.value = imageData.width.toFixed(1);
        if (heightInput) heightInput.value = imageData.height.toFixed(1);
        if (copiesInput) copiesInput.value = imageData.copies;
    }

    showImageModal(dataUrl, name, dimensions) {
        const modal = document.getElementById('imageModal');
        const modalImage = document.getElementById('modalImage');
        const modalImageName = document.getElementById('modalImageName');
        const modalImageSize = document.getElementById('modalImageSize');
        
        modalImage.src = dataUrl;
        modalImageName.textContent = name;
        modalImageSize.textContent = dimensions;
        
        modal.classList.remove('hidden');
        
        // Close modal on background click or ESC key
        const closeModal = () => {
            modal.classList.add('hidden');
            document.removeEventListener('keydown', handleKeydown);
        };
        
        const handleKeydown = (e) => {
            if (e.key === 'Escape') closeModal();
        };
        
        modal.querySelector('.modal-backdrop').onclick = closeModal;
        modal.querySelector('.modal-close').onclick = closeModal;
        document.addEventListener('keydown', handleKeydown);
    }

    // Drag and Drop Methods
    handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    handleDragEnter(e) {
        e.preventDefault();
        e.stopPropagation();
        e.currentTarget.classList.add('drag-over');
    }

    handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        if (!e.currentTarget.contains(e.relatedTarget)) {
            e.currentTarget.classList.remove('drag-over');
        }
    }

    handleDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        e.currentTarget.classList.remove('drag-over');
        
        const files = Array.from(e.dataTransfer.files);
        this.processFiles(files);
    }

    // Clipboard Paste Method
    handlePaste(e) {
        const items = Array.from(e.clipboardData.items);
        const imageItems = items.filter(item => item.type.startsWith('image/'));
        
        if (imageItems.length > 0) {
            e.preventDefault();
            imageItems.forEach(item => {
                const file = item.getAsFile();
                if (file) {
                    this.processFiles([file]);
                }
            });
        }
    }

    processFiles(files) {
        const imageFiles = files.filter(file => file.type.startsWith('image/'));
        
        if (imageFiles.length === 0) {
            alert('Please select image files only.');
            return;
        }

        imageFiles.forEach((file, index) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                // Create an image element to get original dimensions
                const img = new Image();
                img.onload = () => {
                    // Calculate aspect ratio and set default size
                    const aspectRatio = img.width / img.height;
                    const defaultWidth = 50; // default width in mm
                    const defaultHeight = defaultWidth / aspectRatio;
                    
                    const imageData = {
                        id: Date.now() + index + Math.random() * 1000,
                        file: file,
                        dataUrl: e.target.result,
                        previewDataUrl: null,
                        width: defaultWidth,
                        height: defaultHeight,
                        copies: 1,
                        name: file.name,
                        originalWidth: img.width,
                        originalHeight: img.height,
                        aspectRatio: aspectRatio
                    };

                    this.images.push(imageData);
                    this.renderImageConfig(imageData);

                    // Generate a lightweight preview in the background for layout preview.
                    this.createPreviewDataUrl(file).then((previewDataUrl) => {
                        if (previewDataUrl) {
                            imageData.previewDataUrl = previewDataUrl;
                            this.scheduleLayoutPreviewRerender();
                        }
                    });
                };
                img.src = e.target.result;
            };
            reader.readAsDataURL(file);
        });
    }

    generateLayout() {
        if (this.images.length === 0) {
            alert('Please upload some images first!');
            return;
        }

        const paperWidth = parseFloat(document.getElementById('paperWidth').value);
        const paperHeight = parseFloat(document.getElementById('paperHeight').value);
        const outerMargin = parseFloat(document.getElementById('outerMargin').value);
        const innerMargin = parseFloat(document.getElementById('innerMargin').value);
        const allowRotation = Boolean(document.getElementById('rotateImages')?.checked);

        const paperIssue = this.getPaperSettingsIssue(paperWidth, paperHeight, outerMargin);
        if (paperIssue) {
            alert(paperIssue);
            return;
        }

        // Calculate printable area
        const printableWidth = paperWidth - (2 * outerMargin);
        const printableHeight = paperHeight - (2 * outerMargin);

        // Validate that all images can fit on the paper
        // (If rotation is disabled, they must fit in their current orientation.)
        const oversizedImages = this.validateImageSizes(printableWidth, printableHeight);
        if (oversizedImages.length > 0) {
            this.showOversizedWarning(oversizedImages, printableWidth, printableHeight);
            return;
        }

        if (!allowRotation) {
            const requiresRotation = this.images.filter(img => img.width > printableWidth || img.height > printableHeight);
            if (requiresRotation.length > 0) {
                let warningMessage = `WARNING: Rotation is disabled, but some images only fit when rotated:\n\n`;
                warningMessage += `Printable area: ${printableWidth.toFixed(1)}mm × ${printableHeight.toFixed(1)}mm\n\n`;
                requiresRotation.forEach(img => {
                    warningMessage += `• ${img.name}: ${img.width.toFixed(1)}mm × ${img.height.toFixed(1)}mm\n`;
                });
                warningMessage += `\nEnable “Allow Rotation” or reduce the size.`;
                alert(warningMessage);
                return;
            }
        }

        // Expand images based on copies
        const allImages = [];
        this.images.forEach(img => {
            for (let i = 0; i < img.copies; i++) {
                allImages.push({
                    ...img,
                    copyIndex: i,
                    originalId: img.id
                });
            }
        });

        // Apply smart packing algorithm
        this.layout = this.packImages(allImages, printableWidth, printableHeight, innerMargin, allowRotation);
        
        this.renderLayoutPreview(paperWidth, paperHeight, outerMargin);
        this.lastPreviewParams = { paperWidth, paperHeight, outerMargin };
        this.updateLayoutStats();
        
        document.getElementById('exportPDF').disabled = false;
    }

    validateImageSizes(printableWidth, printableHeight) {
        const oversizedImages = [];

        // If printable area is invalid, treat all images as oversized.
        if (!Number.isFinite(printableWidth) || !Number.isFinite(printableHeight) || printableWidth <= 0 || printableHeight <= 0) {
            return this.images.map(img => ({
                ...img,
                maxWidth: 0,
                maxHeight: 0
            }));
        }
        
        this.images.forEach(img => {
            const canFitNormal = img.width <= printableWidth && img.height <= printableHeight;
            const canFitRotated = img.height <= printableWidth && img.width <= printableHeight;
            
            if (!canFitNormal && !canFitRotated) {
                oversizedImages.push({
                    ...img,
                    maxWidth: Math.max(0, Math.max(printableWidth, printableHeight)),
                    maxHeight: Math.max(0, Math.min(printableWidth, printableHeight))
                });
            }
        });
        
        return oversizedImages;
    }

    showOversizedWarning(oversizedImages, printableWidth, printableHeight) {
        let warningMessage = `WARNING: The following images are too large to fit on the paper:\n\n`;
        warningMessage += `Printable area: ${Math.max(0, printableWidth).toFixed(1)}mm × ${Math.max(0, printableHeight).toFixed(1)}mm\n\n`;
        
        oversizedImages.forEach(img => {
            warningMessage += `• ${img.name}: ${img.width.toFixed(1)}mm × ${img.height.toFixed(1)}mm\n`;
            const maxW = Math.max(0, Math.max(printableWidth, printableHeight));
            const maxH = Math.max(0, Math.min(printableWidth, printableHeight));
            warningMessage += `  Max size (any orientation): ${maxW.toFixed(1)}mm × ${maxH.toFixed(1)}mm\n\n`;
        });
        
        warningMessage += `Please reduce the size of these images or increase the paper size/reduce margins.`;
        
        alert(warningMessage);
        
        // Highlight oversized images in the UI
        this.highlightOversizedImages(oversizedImages);
    }

    highlightOversizedImages(oversizedImages) {
        // Remove existing warnings
        document.querySelectorAll('.size-warning').forEach(el => el.remove());
        document.querySelectorAll('.image-config').forEach(el => el.classList.remove('oversized'));
        
        oversizedImages.forEach(img => {
            const config = document.querySelector(`[data-id="${img.id}"]`);
            if (config) {
                config.classList.add('oversized');
                
                const warning = document.createElement('div');
                warning.className = 'size-warning';
                warning.innerHTML = `<i class="fa-solid fa-triangle-exclamation" aria-hidden="true"></i> Too large for paper! Max: ${img.maxWidth.toFixed(1)}×${img.maxHeight.toFixed(1)}mm`;
                config.appendChild(warning);
            }
        });
    }

    // Deterministic 2D rectangle packing (Skyline Bottom-Left).
    // Why Skyline:
    // - Much simpler than MaxRects
    // - Produces very "tidy" looking rows/columns
    // - Efficient for gang-sheet style layouts
    // Rotation is used only when it improves fit (and is never random).
    packImages(images, pageWidth, pageHeight, margin, allowRotation = true) {
        const eps = 1e-6;

        // Sort for packing efficiency and stability (deterministic)
        const sorted = [...images].sort((a, b) => {
            const aArea = a.width * a.height;
            const bArea = b.width * b.height;
            if (bArea !== aArea) return bArea - aArea;
            const aMax = Math.max(a.width, a.height);
            const bMax = Math.max(b.width, b.height);
            if (bMax !== aMax) return bMax - aMax;
            return String(a.name).localeCompare(String(b.name));
        });

        const createPage = () => ({
            images: [],
            _skyline: [{ x: 0, y: 0, width: pageWidth }]
        });

        const pages = [];
        const tryPlaceOnPage = (page, img, rotationAllowed) => {
            const placement = this.skylineFindBestPosition(page._skyline, img, pageWidth, pageHeight, margin, rotationAllowed);
            if (!placement) return false;

            this.skylineAddLevel(page._skyline, placement.x, placement.y, placement.reserveWidth, placement.reserveHeight, eps);
            page.images.push({
                ...img,
                x: placement.x,
                y: placement.y,
                width: placement.placedWidth,
                height: placement.placedHeight,
                rotated: placement.rotated
            });
            return true;
        };

        const tryPlaceOnExistingPages = (img, rotationAllowed) => {
            for (const page of pages) {
                if (tryPlaceOnPage(page, img, rotationAllowed)) return true;
            }
            return false;
        };

        const placeOnNewPage = (img, rotationAllowed) => {
            const page = createPage();
            if (!tryPlaceOnPage(page, img, rotationAllowed)) return false;
            pages.push(page);
            return true;
        };

        for (const img of sorted) {
            // 1) Prefer unrotated on existing pages.
            if (tryPlaceOnExistingPages(img, false)) continue;

            // 2) If rotation is enabled, try rotated ONLY to avoid starting a new page.
            if (allowRotation && tryPlaceOnExistingPages(img, true)) continue;

            // 3) Start a new page (unrotated).
            if (placeOnNewPage(img, false)) continue;

            // 4) If unrotated won't even fit on an empty page, allow rotation on a new page.
            if (allowRotation && placeOnNewPage(img, true)) continue;

            console.error('Image too large to fit on page:', img);
        }

        // Strip internal packing state
        for (const page of pages) {
            delete page._skyline;
        }

        return pages;
    }

    skylineFindBestPosition(skyline, image, pageWidth, pageHeight, margin, allowRotation) {
        const eps = 1e-6;
        const orientations = [
            { w: image.width, h: image.height, rotated: false }
        ];
        if (allowRotation) orientations.push({ w: image.height, h: image.width, rotated: true });

        let best = null;
        for (const o of orientations) {
            for (let i = 0; i < skyline.length; i++) {
                const x = skyline[i].x;

                // Default: reserve trailing spacing on right/bottom.
                // If we are up against an edge, we allow dropping the trailing margin.
                let reserveW = o.w + margin;
                if (x + reserveW > pageWidth + eps && x + o.w <= pageWidth + eps) {
                    reserveW = o.w;
                }

                const fit = this.skylineFindMinY(skyline, i, reserveW, pageWidth, eps);
                if (!fit) continue;

                let reserveH = o.h + margin;
                if (fit.y + reserveH > pageHeight + eps && fit.y + o.h <= pageHeight + eps) {
                    reserveH = o.h;
                }

                if (fit.y + reserveH > pageHeight + eps) continue;

                const rotationPenalty = o.rotated ? 1 : 0;
                const bottom = fit.y + reserveH;

                const candidate = {
                    x,
                    y: fit.y,
                    placedWidth: o.w,
                    placedHeight: o.h,
                    reserveWidth: reserveW,
                    reserveHeight: reserveH,
                    rotated: o.rotated,
                    scoreY: fit.y,
                    scoreX: x,
                    scoreRot: rotationPenalty,
                    scoreBottom: bottom
                };

                if (!best) {
                    best = candidate;
                    continue;
                }

                // Bottom-left: minimize y, then x.
                if (candidate.scoreY !== best.scoreY) {
                    if (candidate.scoreY < best.scoreY) best = candidate;
                    continue;
                }
                if (candidate.scoreX !== best.scoreX) {
                    if (candidate.scoreX < best.scoreX) best = candidate;
                    continue;
                }
                if (candidate.scoreRot !== best.scoreRot) {
                    if (candidate.scoreRot < best.scoreRot) best = candidate;
                    continue;
                }
                if (candidate.scoreBottom !== best.scoreBottom) {
                    if (candidate.scoreBottom < best.scoreBottom) best = candidate;
                    continue;
                }
            }
        }

        return best;
    }

    skylineFindMinY(skyline, startIndex, width, pageWidth, eps) {
        const startX = skyline[startIndex].x;
        if (startX + width > pageWidth + eps) return null;

        let y = skyline[startIndex].y;
        let widthLeft = width;
        let i = startIndex;
        while (widthLeft > eps) {
            y = Math.max(y, skyline[i].y);
            widthLeft -= skyline[i].width;
            i += 1;
            if (i >= skyline.length && widthLeft > eps) return null;
        }

        return { x: startX, y };
    }

    skylineAddLevel(skyline, x, y, width, height, eps) {
        // Insert new node
        const newNode = { x, y: y + height, width };

        // Find insertion index
        let index = 0;
        while (index < skyline.length && skyline[index].x < x - eps) index++;
        skyline.splice(index, 0, newNode);

        // Remove/trim any nodes that overlap the new node horizontally
        for (let i = index + 1; i < skyline.length; i++) {
            const node = skyline[i];
            const prev = skyline[i - 1];
            const prevRight = prev.x + prev.width;

            if (node.x >= prevRight - eps) break;

            const shrink = prevRight - node.x;
            node.x += shrink;
            node.width -= shrink;
            if (node.width <= eps) {
                skyline.splice(i, 1);
                i--;
            }
        }

        // Merge adjacent nodes with the same y
        for (let i = 0; i < skyline.length - 1; i++) {
            if (Math.abs(skyline[i].y - skyline[i + 1].y) <= eps) {
                skyline[i].width += skyline[i + 1].width;
                skyline.splice(i + 1, 1);
                i--;
            }
        }
    }

    maxRectsPlace(page, image, margin, allowRotation) {
        if (!page._freeRects) page._freeRects = [{ x: 0, y: 0, width: Infinity, height: Infinity }];

        const binW = Number.isFinite(page._binWidth) ? page._binWidth : Infinity;
        const binH = Number.isFinite(page._binHeight) ? page._binHeight : Infinity;
        const best = this.maxRectsFindPosition(page._freeRects, image, margin, allowRotation, binW, binH);
        if (!best) return false;

        // Reserve margin on right/bottom by packing with inflated dimensions.
        const packedRect = {
            x: best.x,
            y: best.y,
            width: best.packedWidth,
            height: best.packedHeight
        };

        // Update free rects by splitting any nodes that intersect the placed rect
        const newFree = [];
        for (const free of page._freeRects) {
            if (!this.rectanglesOverlap(packedRect, free)) {
                newFree.push(free);
                continue;
            }

            // Split free rect into up to 4 rects around the placed rect
            // Above
            if (packedRect.y > free.y) {
                newFree.push({
                    x: free.x,
                    y: free.y,
                    width: free.width,
                    height: packedRect.y - free.y
                });
            }
            // Below
            const freeBottom = free.y + free.height;
            const placedBottom = packedRect.y + packedRect.height;
            if (placedBottom < freeBottom) {
                newFree.push({
                    x: free.x,
                    y: placedBottom,
                    width: free.width,
                    height: freeBottom - placedBottom
                });
            }
            // Left
            if (packedRect.x > free.x) {
                const leftWidth = packedRect.x - free.x;
                const top = Math.max(free.y, packedRect.y);
                const bottom = Math.min(free.y + free.height, packedRect.y + packedRect.height);
                const h = bottom - top;
                if (h > 0 && leftWidth > 0) {
                    newFree.push({ x: free.x, y: top, width: leftWidth, height: h });
                }
            }
            // Right
            const freeRight = free.x + free.width;
            const placedRight = packedRect.x + packedRect.width;
            if (placedRight < freeRight) {
                const rightWidth = freeRight - placedRight;
                const top = Math.max(free.y, packedRect.y);
                const bottom = Math.min(free.y + free.height, packedRect.y + packedRect.height);
                const h = bottom - top;
                if (h > 0 && rightWidth > 0) {
                    newFree.push({ x: placedRight, y: top, width: rightWidth, height: h });
                }
            }
        }

        page._freeRects = this.maxRectsPruneFreeList(newFree);

        page.images.push({
            ...image,
            x: best.x,
            y: best.y,
            width: best.placedWidth,
            height: best.placedHeight,
            rotated: best.rotated
        });

        return true;
    }

    maxRectsFindPosition(freeRects, image, margin, allowRotation, binWidth, binHeight) {
        const orientations = [];
        orientations.push({
            placedWidth: image.width,
            placedHeight: image.height,
            packedWidth: image.width + margin,
            packedHeight: image.height + margin,
            rotated: false
        });
        if (allowRotation) {
            orientations.push({
                placedWidth: image.height,
                placedHeight: image.width,
                packedWidth: image.height + margin,
                packedHeight: image.width + margin,
                rotated: true
            });
        }

        let best = null;
        for (const free of freeRects) {
            for (const o of orientations) {
                // Allow placements that fit exactly at the edge without forcing a trailing margin.
                // If there's not enough remaining space for the margin, drop the trailing margin.
                let packedW = o.packedWidth;
                let packedH = o.packedHeight;

                // If width fits but width+margin doesn't, allow packing without the trailing margin.
                if (o.placedWidth <= free.width && packedW > free.width) {
                    packedW = o.placedWidth;
                }
                if (o.placedHeight <= free.height && packedH > free.height) {
                    packedH = o.placedHeight;
                }

                if (packedW > free.width || packedH > free.height) continue;

                const leftoverH = free.width - packedW;
                const leftoverV = free.height - packedH;
                const shortSide = Math.min(leftoverH, leftoverV);
                const longSide = Math.max(leftoverH, leftoverV);

                // Rotation penalty keeps things visually consistent when both fit.
                const rotationPenalty = o.rotated ? 1 : 0;

                const candidate = {
                    x: free.x,
                    y: free.y,
                    placedWidth: o.placedWidth,
                    placedHeight: o.placedHeight,
                    packedWidth: packedW,
                    packedHeight: packedH,
                    rotated: o.rotated,
                    scoreShortSide: shortSide,
                    scoreLongSide: longSide,
                    scoreRotation: rotationPenalty
                };

                if (!best) {
                    best = candidate;
                    continue;
                }

                if (candidate.scoreShortSide !== best.scoreShortSide) {
                    if (candidate.scoreShortSide < best.scoreShortSide) best = candidate;
                    continue;
                }
                if (candidate.scoreLongSide !== best.scoreLongSide) {
                    if (candidate.scoreLongSide < best.scoreLongSide) best = candidate;
                    continue;
                }
                if (candidate.scoreRotation !== best.scoreRotation) {
                    if (candidate.scoreRotation < best.scoreRotation) best = candidate;
                    continue;
                }
                // Stable tie-break for an organized look: top-left first
                if (candidate.y !== best.y) {
                    if (candidate.y < best.y) best = candidate;
                    continue;
                }
                if (candidate.x !== best.x) {
                    if (candidate.x < best.x) best = candidate;
                    continue;
                }
            }
        }

        return best;
    }

    maxRectsPruneFreeList(freeRects) {
        // Remove empty/invalid rects
        const list = freeRects.filter(r => r.width > 0 && r.height > 0);

        // Remove rects that are fully contained in another rect
        const pruned = [];
        for (let i = 0; i < list.length; i++) {
            const a = list[i];
            let contained = false;
            for (let j = 0; j < list.length; j++) {
                if (i === j) continue;
                const b = list[j];
                if (this.rectContains(b, a)) {
                    contained = true;
                    break;
                }
            }
            if (!contained) pruned.push(a);
        }

        // Deterministic order to keep behavior stable
        pruned.sort((r1, r2) => {
            if (r1.y !== r2.y) return r1.y - r2.y;
            if (r1.x !== r2.x) return r1.x - r2.x;
            if (r1.width !== r2.width) return r1.width - r2.width;
            return r1.height - r2.height;
        });

        return pruned;
    }

    rectContains(outer, inner) {
        return (
            inner.x >= outer.x &&
            inner.y >= outer.y &&
            inner.x + inner.width <= outer.x + outer.width &&
            inner.y + inner.height <= outer.y + outer.height
        );
    }

    findBestPlacementOnPage(image, page, pageWidth, pageHeight, margin) {
        const placements = this.getAllPossiblePlacements(image, page, pageWidth, pageHeight, margin);
        
        console.log(`Finding placement for image ${image.name} (${image.width}x${image.height}mm)`);
        console.log(`Page has ${page.images.length} existing images`);
        console.log(`Found ${placements.length} possible placements`);
        
        if (placements.length === 0) {
            return null;
        }
        
        // Sort placements by preference:
        // 1. Orientation that leaves more usable space for future items
        // 2. Bottom-left positioning (y first, then x)
        // 3. Least wasted space
        const occupied = this.getOccupiedRectangles(page, margin);
        placements.sort((a, b) => {
            // Calculate remaining space efficiency for each placement
            const aRemainingSpace = this.calculateRemainingUsableSpace(a, pageWidth, pageHeight, occupied);
            const bRemainingSpace = this.calculateRemainingUsableSpace(b, pageWidth, pageHeight, occupied);
            
            // Strongly prefer placements that leave more usable space
            const spaceDiff = bRemainingSpace - aRemainingSpace;
            if (Math.abs(spaceDiff) > pageWidth * pageHeight * 0.05) { // 5% difference threshold
                return spaceDiff;
            }
            
            // If space efficiency is similar, prefer bottom-left positioning
            const aY = a.y;
            const bY = b.y;
            if (Math.abs(aY - bY) > 1) {
                return aY - bY; // Prefer lower y position
            }
            
            const aX = a.x;
            const bX = b.x;
            if (Math.abs(aX - bX) > 1) {
                return aX - bX; // Prefer left x position
            }

            return a.wastedSpace - b.wastedSpace;
        });        
        
        const bestPlacement = placements[0];
        console.log(`Best placement: (${bestPlacement.x}, ${bestPlacement.y}) ${bestPlacement.width}x${bestPlacement.height} rotated: ${bestPlacement.rotated}`);
        
        return bestPlacement;
    }

    getAllPossiblePlacements(image, page, pageWidth, pageHeight, margin) {
        const placements = [];
        const occupied = this.getOccupiedRectangles(page, margin);
        
        // Try both orientations
        const orientations = [
            { width: image.width, height: image.height, rotated: false },
            { width: image.height, height: image.width, rotated: true }
        ];
        
        for (const orientation of orientations) {
            // Generate candidate positions
            const candidates = this.generateCandidatePositions(occupied, pageWidth, pageHeight, orientation.width, orientation.height);
            
            for (const pos of candidates) {
                if (this.canPlaceAt(pos.x, pos.y, orientation.width, orientation.height, occupied, pageWidth, pageHeight)) {
                    placements.push({
                        x: pos.x,
                        y: pos.y,
                        width: orientation.width,
                        height: orientation.height,
                        rotated: orientation.rotated,
                        wastedSpace: this.calculateWastedSpace(pos.x, pos.y, orientation.width, orientation.height, occupied, pageWidth, pageHeight)
                    });
                }
            }
        }
        
        return placements;
    }

    getOccupiedRectangles(page, margin) {
        return page.images.map(img => ({
            x: img.x,
            y: img.y,
            width: img.width + margin, // Only add margin to the right and bottom
            height: img.height + margin
        }));
    }

    generateCandidatePositions(occupied, pageWidth, pageHeight, itemWidth, itemHeight) {
        const candidates = new Set();
        
        // Add corner positions
        candidates.add(`0,0`);
        
        // Add positions adjacent to existing items
        for (const rect of occupied) {
            // Right edge - this is key for horizontal placement
            const rightX = rect.x + rect.width;
            if (rightX + itemWidth <= pageWidth) {
                candidates.add(`${rightX},${rect.y}`);
                // Also try aligning with the bottom of the existing item
                const bottomAlignY = rect.y + rect.height - itemHeight;
                if (bottomAlignY >= 0) {
                    candidates.add(`${rightX},${bottomAlignY}`);
                }
            }
            
            // Bottom edge
            const bottomY = rect.y + rect.height;
            if (bottomY + itemHeight <= pageHeight) {
                candidates.add(`${rect.x},${bottomY}`);
                // Also try aligning with the right edge of the existing item
                const rightAlignX = rect.x + rect.width - itemWidth;
                if (rightAlignX >= 0) {
                    candidates.add(`${rightAlignX},${bottomY}`);
                }
            }
            
            // Top edge (for items that might fit above)
            const topY = rect.y - itemHeight;
            if (topY >= 0) {
                candidates.add(`${rect.x},${topY}`);
            }
            
            // Left edge (for items that might fit to the left)
            const leftX = rect.x - itemWidth;
            if (leftX >= 0) {
                candidates.add(`${leftX},${rect.y}`);
            }
        }
        
        // Convert to array of position objects and filter valid positions
        return Array.from(candidates).map(pos => {
            const [x, y] = pos.split(',').map(Number);
            return { x, y };
        }).filter(pos => 
            pos.x >= 0 && pos.y >= 0 && 
            pos.x + itemWidth <= pageWidth && 
            pos.y + itemHeight <= pageHeight
        );
    }

    canPlaceAt(x, y, width, height, occupied, pageWidth, pageHeight) {
        // Check bounds
        if (x < 0 || y < 0 || x + width > pageWidth || y + height > pageHeight) {
            return false;
        }
        
        // Check overlap with existing items
        const newRect = { x, y, width, height };
        for (const rect of occupied) {
            if (this.rectanglesOverlap(newRect, rect)) {
                return false;
            }
        }
        
        return true;
    }

    rectanglesOverlap(rect1, rect2) {
        return !(rect1.x >= rect2.x + rect2.width ||
                rect2.x >= rect1.x + rect1.width ||
                rect1.y >= rect2.y + rect2.height ||
                rect2.y >= rect1.y + rect1.height);
    }

    calculateWastedSpace(x, y, width, height, occupied, pageWidth, pageHeight) {
        // Simple heuristic: calculate unused space around the placement
        const rightSpace = pageWidth - (x + width);
        const bottomSpace = pageHeight - (y + height);
        return rightSpace * bottomSpace;
    }

    calculateRemainingUsableSpace(placement, pageWidth, pageHeight, occupied) {
        // Create a temporary occupied list with this placement added
        const newOccupied = [...occupied, {
            x: placement.x,
            y: placement.y,
            width: placement.width,
            height: placement.height
        }];
        
        // Calculate the largest rectangular space that remains free
        let maxRemainingRect = 0;
        
        // Check right side
        const rightX = placement.x + placement.width;
        if (rightX < pageWidth) {
            const rightWidth = pageWidth - rightX;
            const rightHeight = Math.min(pageHeight - placement.y, placement.height);
            maxRemainingRect = Math.max(maxRemainingRect, rightWidth * rightHeight);
        }
        
        // Check bottom side
        const bottomY = placement.y + placement.height;
        if (bottomY < pageHeight) {
            const bottomHeight = pageHeight - bottomY;
            const bottomWidth = Math.min(pageWidth - placement.x, placement.width);
            maxRemainingRect = Math.max(maxRemainingRect, bottomWidth * bottomHeight);
        }
        
        // Check bottom-right corner
        if (rightX < pageWidth && bottomY < pageHeight) {
            const cornerWidth = pageWidth - rightX;
            const cornerHeight = pageHeight - bottomY;
            maxRemainingRect = Math.max(maxRemainingRect, cornerWidth * cornerHeight);
        }
        
        return maxRemainingRect;
    }

    renderLayoutPreview(paperWidth, paperHeight, outerMargin) {
        const preview = document.getElementById('layoutPreview');
        preview.innerHTML = '';

        this.layout.forEach((page, index) => {
            const pageDiv = document.createElement('div');
            pageDiv.className = 'page-preview';
            
            const aspectRatio = paperWidth / paperHeight;
            // Dynamic sizing based on container width and aspect ratio
            const containerWidth = Math.min(300, window.innerWidth * 0.25); // Max 300px or 25% of screen width
            const previewWidth = containerWidth;
            const previewHeight = previewWidth / aspectRatio;
            const scale = previewWidth / paperWidth;

            pageDiv.innerHTML = `
                <div class="page-content" style="width: ${previewWidth}px; height: ${previewHeight}px; overflow: hidden; position: relative;">
                </div>
            `;

            const pageContent = pageDiv.querySelector('.page-content');

            const canvas = document.createElement('canvas');
            canvas.style.width = `${previewWidth}px`;
            canvas.style.height = `${previewHeight}px`;
            canvas.style.display = 'block';
            pageContent.appendChild(canvas);

            preview.appendChild(pageDiv);

            // Draw asynchronously so the UI remains responsive.
            this.drawPagePreviewToCanvas(canvas, page, paperWidth, paperHeight, outerMargin, scale);
        });
    }

    updateLayoutStats() {
        const totalImages = this.images.reduce((sum, img) => sum + img.copies, 0);
        const totalPages = this.layout.length;
        const imagesPerPage = totalPages > 0 ? (totalImages / totalPages).toFixed(1) : 0;

        const stats = document.getElementById('layoutStats');
        stats.innerHTML = `
            <div class="stats-grid">
                <div class="stat-item">
                    <div class="stat-value">${totalImages}</div>
                    <div class="stat-label">Total Images</div>
                </div>
                <div class="stat-item">
                    <div class="stat-value">${totalPages}</div>
                    <div class="stat-label">Pages Required</div>
                </div>
            </div>
        `;
    }

    async exportToPDF() {
        if (this._exportState?.inProgress) {
            return;
        }
        if (this.layout.length === 0) {
            alert('Please generate a layout first!');
            return;
        }

        const exportBtn = document.getElementById('exportPDF');
        const { cancelBtn } = this.getExportUi();
        const abortController = new AbortController();
        this._exportState = {
            inProgress: true,
            abortController
        };

        if (exportBtn) exportBtn.disabled = true;
        if (cancelBtn) {
            cancelBtn.disabled = false;
            cancelBtn.textContent = 'Cancel';
        }

        this.setExportOverlayVisible(true);

        const { jsPDF } = window.jspdf;
        const paperWidth = parseFloat(document.getElementById('paperWidth').value);
        const paperHeight = parseFloat(document.getElementById('paperHeight').value);
        const outerMargin = parseFloat(document.getElementById('outerMargin').value);

        // Convert mm to points (1mm = 2.834645669 points)
        const mmToPt = 2.834645669;
        const pdf = new jsPDF({
            orientation: paperWidth > paperHeight ? 'landscape' : 'portrait',
            unit: 'mm',
            format: [paperWidth, paperHeight]
        });

        const totalImages = this.layout.reduce((sum, page) => sum + (page.images?.length || 0), 0);
        let completed = 0;

        const throwIfAborted = () => {
            if (abortController.signal.aborted) {
                throw new DOMException('Export canceled', 'AbortError');
            }
        };

        try {
            this.setExportProgress(0, totalImages, 'Preparing…');
            throwIfAborted();

            for (let pageIndex = 0; pageIndex < this.layout.length; pageIndex++) {
                throwIfAborted();
                if (pageIndex > 0) {
                    pdf.addPage();
                }

                const page = this.layout[pageIndex];
                this.setExportProgress(completed, totalImages, `Rendering page ${pageIndex + 1}/${this.layout.length}…`);

                for (const img of page.images) {
                    throwIfAborted();
                    try {
                        const imageData = await this.loadImageForPDF(img, abortController.signal);
                        throwIfAborted();

                        pdf.addImage(
                            imageData,
                            'PNG',
                            outerMargin + img.x,
                            outerMargin + img.y,
                            img.width,
                            img.height,
                            undefined,
                            'NONE'
                        );
                    } catch (error) {
                        if (error && (error.name === 'AbortError' || error.message === 'Export canceled')) {
                            throw error;
                        }
                        console.error('Error adding image to PDF:', error);
                    } finally {
                        completed += 1;
                        this.setExportProgress(completed, totalImages, `Rendering page ${pageIndex + 1}/${this.layout.length}…`);
                    }
                }
            }

            throwIfAborted();
            this.setExportProgress(totalImages, totalImages, 'Saving PDF…');
            pdf.save('sheetbuilder-layout.pdf');
            this.setExportProgress(totalImages, totalImages, 'Done');
        } catch (error) {
            if (error && error.name === 'AbortError') {
                this.setExportProgress(completed, totalImages, 'Canceled');
                return;
            }
            console.error('Export failed:', error);
            alert('Export failed. Please try again.');
        } finally {
            window.setTimeout(() => {
                this.setExportOverlayVisible(false);
            }, 150);

            if (exportBtn) exportBtn.disabled = false;
            this._exportState = {
                inProgress: false,
                abortController: null
            };
        }
    }

    loadImageForPDF(imgData, abortSignal) {
        return new Promise((resolve, reject) => {
            const img = new Image();

            const abort = () => {
                try {
                    img.onload = null;
                    img.onerror = null;
                    img.src = '';
                } catch {
                    // ignore
                }
                reject(new DOMException('Export canceled', 'AbortError'));
            };

            if (abortSignal?.aborted) {
                abort();
                return;
            }

            const onAbort = () => abort();
            if (abortSignal) {
                abortSignal.addEventListener('abort', onAbort, { once: true });
            }

            img.onload = () => {
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');

                if (abortSignal) {
                    abortSignal.removeEventListener('abort', onAbort);
                }

                if (abortSignal?.aborted) {
                    reject(new DOMException('Export canceled', 'AbortError'));
                    return;
                }
                
                if (imgData.rotated) {
                    canvas.width = img.height;
                    canvas.height = img.width;
                    
                    // Fill with white background before drawing the image
                    ctx.fillStyle = 'white';
                    ctx.fillRect(0, 0, canvas.width, canvas.height);
                    
                    ctx.translate(canvas.width / 2, canvas.height / 2);
                    ctx.rotate(Math.PI / 2);
                    ctx.drawImage(img, -img.width / 2, -img.height / 2);
                } else {
                    canvas.width = img.width;
                    canvas.height = img.height;
                    
                    // Fill with white background before drawing the image
                    ctx.fillStyle = 'white';
                    ctx.fillRect(0, 0, canvas.width, canvas.height);
                    
                    ctx.drawImage(img, 0, 0);
                }

                // Use lossless PNG export for maximum quality (no JPEG compression).
                resolve(canvas.toDataURL('image/png'));
            };
            img.onerror = (e) => {
                if (abortSignal) {
                    abortSignal.removeEventListener('abort', onAbort);
                }
                reject(e);
            };
            img.src = imgData.dataUrl;
        });
    }
}

// Initialize the application
const sheetBuilder = new SheetBuilder();