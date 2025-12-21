const MM_PER_INCH = 25.4;

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
        // UI unit for all length inputs/labels. Internally we store and compute in mm.
        this.unit = 'mm';
        this._exportState = {
            inProgress: false,
            abortController: null
        };
        this._uploadFeedback = {
            token: 0,
            hideTimer: null
        };
        this.initializeEventListeners();
        this.ensureDefaultPaperSettings();
        this.initializeUnits();
    }

    getUploadStatusEl() {
        return document.getElementById('uploadStatus');
    }

    setUploadStatus(text, { visible } = {}) {
        const el = this.getUploadStatusEl();
        if (!el) return;
        if (typeof text === 'string') {
            el.textContent = text;
        }
        if (typeof visible === 'boolean') {
            el.hidden = !visible;
        }
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

    getPaperSizePresetsMm() {
        return {
            a4: { width: 210, height: 297, label: 'A4' },
            a3: { width: 297, height: 420, label: 'A3' },
            a5: { width: 148, height: 210, label: 'A5' },
            letter: { width: 216, height: 279, label: 'Letter' },
            legal: { width: 216, height: 356, label: 'Legal' },
            tabloid: { width: 279, height: 432, label: 'Tabloid' },
            photo4x6: { width: 102, height: 152, label: 'Photo 4×6' },
            photo5x7: { width: 127, height: 178, label: 'Photo 5×7' },
            photo8x10: { width: 203, height: 254, label: 'Photo 8×10' }
        };
    }

    getUnitLabel() {
        return this.unit === 'in' ? 'in' : 'mm';
    }

    mmToDisplay(mm) {
        if (!Number.isFinite(mm)) return NaN;
        return this.unit === 'in' ? (mm / MM_PER_INCH) : mm;
    }

    displayToMm(value) {
        if (!Number.isFinite(value)) return NaN;
        return this.unit === 'in' ? (value * MM_PER_INCH) : value;
    }

    formatLengthValue(mm) {
        const displayValue = this.mmToDisplay(mm);
        if (!Number.isFinite(displayValue)) return '';
        const decimals = this.unit === 'in' ? 2 : 1;
        return displayValue.toFixed(decimals);
    }

    formatSizePair(mmW, mmH) {
        return `${this.formatLengthValue(mmW)}×${this.formatLengthValue(mmH)}${this.getUnitLabel()}`;
    }

    parseLengthFromInputValue(raw) {
        const num = parseFloat(raw);
        if (!Number.isFinite(num)) return NaN;
        return this.displayToMm(num);
    }

    readLengthInputMm(el) {
        if (!el) return NaN;
        return this.parseLengthFromInputValue(el.value);
    }

    writeLengthInputFromMm(el, mm) {
        if (!el) return;
        if (!Number.isFinite(mm)) {
            el.value = '';
            return;
        }
        el.value = this.formatLengthValue(mm);
    }

    syncLengthInputAttributes(el) {
        if (!el) return;
        const baseMinMm = parseFloat(el.dataset.minMm);
        const baseStepMm = parseFloat(el.dataset.stepMm);

        if (Number.isFinite(baseMinMm)) {
            const minDisplay = this.mmToDisplay(baseMinMm);
            if (Number.isFinite(minDisplay)) el.min = String(minDisplay);
        }
        if (Number.isFinite(baseStepMm)) {
            const stepDisplay = this.mmToDisplay(baseStepMm);
            if (Number.isFinite(stepDisplay)) el.step = String(stepDisplay);
        }
    }

    syncAllLengthInputsAttributes() {
        const ids = ['paperWidth', 'paperHeight', 'outerMargin', 'innerMargin', 'bulkWidth', 'bulkHeight'];
        for (const id of ids) {
            this.syncLengthInputAttributes(document.getElementById(id));
        }
        document.querySelectorAll('.image-config input[data-min-mm][data-step-mm]').forEach((el) => this.syncLengthInputAttributes(el));
    }

    updateUnitLabels() {
        const label = this.getUnitLabel();
        document.querySelectorAll('[data-unit-label]').forEach((el) => {
            el.textContent = label;
        });
    }

    updatePaperSizeOptionLabels() {
        const presets = this.getPaperSizePresetsMm();
        const select = document.getElementById('paperSize');
        if (!select) return;
        Array.from(select.options).forEach((opt) => {
            if (!opt || opt.value === 'custom') return;
            const preset = presets[opt.value];
            if (!preset) return;
            const baseLabel = opt.dataset.label || preset.label || opt.textContent;
            opt.textContent = `${baseLabel} (${this.formatSizePair(preset.width, preset.height)})`;
        });
    }

    getPaperSettingsMm() {
        const paperWidthEl = document.getElementById('paperWidth');
        const paperHeightEl = document.getElementById('paperHeight');
        const outerMarginEl = document.getElementById('outerMargin');
        const innerMarginEl = document.getElementById('innerMargin');

        return {
            paperWidth: this.readLengthInputMm(paperWidthEl),
            paperHeight: this.readLengthInputMm(paperHeightEl),
            outerMargin: this.readLengthInputMm(outerMarginEl),
            innerMargin: this.readLengthInputMm(innerMarginEl)
        };
    }

    rerenderAllImageConfigs() {
        const list = document.getElementById('imagesList');
        if (!list) return;

        const selected = new Set(this.selectedImages);
        list.innerHTML = '';
        for (const img of this.images) {
            this.renderImageConfig(img);
        }

        // Restore selection state
        for (const id of selected) {
            const checkbox = document.getElementById(`select-${id}`);
            const config = document.querySelector(`[data-id="${id}"]`);
            if (checkbox) checkbox.checked = true;
            if (config) config.classList.add('selected');
        }
        this.updateQuickSelectionControls();
        this.updateSelectionCount();
    }

    initializeUnits() {
        const unitSelect = document.getElementById('unitSelect');
        if (unitSelect) {
            unitSelect.addEventListener('change', (e) => this.setUnit(e.target.value));
        }

        let preferred = null;
        try {
            preferred = localStorage.getItem('sheetbuilder_unit');
        } catch {
            preferred = null;
        }

        // Inputs are in mm on first load.
        this.setUnit(preferred === 'in' ? 'in' : 'mm');
    }

    setUnit(nextUnit) {
        const normalized = nextUnit === 'in' ? 'in' : 'mm';
        const current = this.unit;

        // Convert existing inputs from the current unit into mm before switching.
        const paperMm = this.getPaperSettingsMm();

        const bulkWidthEl = document.getElementById('bulkWidth');
        const bulkHeightEl = document.getElementById('bulkHeight');
        const bulkWidthMm = bulkWidthEl && bulkWidthEl.value ? this.readLengthInputMm(bulkWidthEl) : null;
        const bulkHeightMm = bulkHeightEl && bulkHeightEl.value ? this.readLengthInputMm(bulkHeightEl) : null;

        this.unit = normalized;

        try {
            localStorage.setItem('sheetbuilder_unit', normalized);
        } catch {
            // ignore
        }

        const unitSelect = document.getElementById('unitSelect');
        if (unitSelect) unitSelect.value = normalized;

        this.updateUnitLabels();
        this.updatePaperSizeOptionLabels();
        this.syncAllLengthInputsAttributes();

        // Write paper/bulk values back in the new unit.
        this.writeLengthInputFromMm(document.getElementById('paperWidth'), paperMm.paperWidth);
        this.writeLengthInputFromMm(document.getElementById('paperHeight'), paperMm.paperHeight);
        this.writeLengthInputFromMm(document.getElementById('outerMargin'), paperMm.outerMargin);
        this.writeLengthInputFromMm(document.getElementById('innerMargin'), paperMm.innerMargin);
        if (bulkWidthMm !== null) this.writeLengthInputFromMm(bulkWidthEl, bulkWidthMm);
        if (bulkHeightMm !== null) this.writeLengthInputFromMm(bulkHeightEl, bulkHeightMm);

        if (normalized !== current) {
            this.rerenderAllImageConfigs();
        } else {
            // Still ensure labels/inputs inside dynamic configs match.
            this.rerenderAllImageConfigs();
        }

        this.validateAllImages();
        this.checkForCustomSize();
    }

    ensureDefaultPaperSettings() {
        const paperSize = document.getElementById('paperSize');
        const paperWidthEl = document.getElementById('paperWidth');
        const paperHeightEl = document.getElementById('paperHeight');
        const outerMarginEl = document.getElementById('outerMargin');
        if (!paperSize || !paperWidthEl || !paperHeightEl || !outerMarginEl) return;

        // Default to A4 unless the user explicitly chose something else.
        if (!paperSize.value) paperSize.value = 'a4';

        const presets = this.getPaperSizePresetsMm();

        const selected = paperSize.value;
        if (selected === 'a4') {
            const { paperWidth: currentW, paperHeight: currentH } = this.getPaperSettingsMm();
            if (!Number.isFinite(currentW) || !Number.isFinite(currentH) || Math.abs(currentW - presets.a4.width) > 0.1 || Math.abs(currentH - presets.a4.height) > 0.1) {
                this.writeLengthInputFromMm(paperWidthEl, presets.a4.width);
                this.writeLengthInputFromMm(paperHeightEl, presets.a4.height);
            }
        }

        const { paperWidth: w, paperHeight: h, outerMargin: outer } = this.getPaperSettingsMm();
        const maxOuter = Number.isFinite(w) && Number.isFinite(h) ? Math.max(0, (Math.min(w, h) / 2) - 0.1) : 0;
        if (!Number.isFinite(outer) || outer < 0 || outer > maxOuter) {
            this.writeLengthInputFromMm(outerMarginEl, 10);
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
        const exportBtnBottom = document.getElementById('exportPDFBottom');
        const quickCropImagesBtn = document.getElementById('quickCropImages');
        const paperSize = document.getElementById('paperSize');
        const cancelExportBtn = document.getElementById('cancelExport');
        const fillUntilPagesEnabled = document.getElementById('fillUntilPagesEnabled');
        const fillUntilPagesInput = document.getElementById('fillUntilPages');
        
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
        if (exportBtnBottom) exportBtnBottom.addEventListener('click', () => this.exportToPDF());
        paperSize.addEventListener('change', (e) => this.handlePaperSizeChange(e));

        if (quickCropImagesBtn) {
            quickCropImagesBtn.addEventListener('click', async () => {
                if (quickCropImagesBtn.disabled) return;
                const previous = quickCropImagesBtn.innerHTML;
                quickCropImagesBtn.disabled = true;
                quickCropImagesBtn.textContent = 'Cropping…';
                try {
                    await this.cropAllImagesTransparentPadding();
                } finally {
                    quickCropImagesBtn.disabled = false;
                    quickCropImagesBtn.innerHTML = previous;
                }
            });
        }

        if (fillUntilPagesInput) {
            fillUntilPagesInput.addEventListener('input', () => {
                // Do not auto-correct while typing.
            });
            fillUntilPagesInput.addEventListener('change', () => {
                // Normalize to an integer and clamp upward to the computed minimum.
                this.setFillUntilPagesInputValue(this.getFillUntilPages());
                this.setFillUntilPagesMin(this.computeMinimumPagesRequired());

                // Do not apply fill until the user regenerates the layout.
                this.markLayoutStale({ resetFillUntilPages: false });
            });
        }

        if (fillUntilPagesEnabled) {
            fillUntilPagesEnabled.addEventListener('change', () => {
                // Refresh minimum; toggling fill on/off affects layout behavior.
                this.setFillUntilPagesMin(this.computeMinimumPagesRequired());

                // Do not apply fill until the user regenerates the layout.
                this.markLayoutStale({ resetFillUntilPages: false });
            });
        }

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
            this.markLayoutStale({ resetFillUntilPages: false });
        });
        paperWidth.addEventListener('change', () => {
            this.markLayoutStale({ resetFillUntilPages: true });
        });
        paperHeight.addEventListener('input', () => {
            this.validateAllImages();
            this.checkForCustomSize();
            this.markLayoutStale({ resetFillUntilPages: false });
        });
        paperHeight.addEventListener('change', () => {
            this.markLayoutStale({ resetFillUntilPages: true });
        });
        outerMargin.addEventListener('input', () => {
            this.validateAllImages();
            this.markLayoutStale({ resetFillUntilPages: false });
        });
        outerMargin.addEventListener('change', () => {
            this.markLayoutStale({ resetFillUntilPages: true });
        });

        const innerMargin = document.getElementById('innerMargin');
        if (innerMargin) {
            innerMargin.addEventListener('input', () => {
                this.markLayoutStale({ resetFillUntilPages: false });
            });
            innerMargin.addEventListener('change', () => {
                this.markLayoutStale({ resetFillUntilPages: true });
            });
        }

        const rotateImages = document.getElementById('rotateImages');
        if (rotateImages) {
            rotateImages.addEventListener('change', () => {
                this.markLayoutStale({ resetFillUntilPages: true });
            });
        }
    }

    handlePaperSizeChange(event) {
        const selectedSize = event.target.value;
        const paperSizes = this.getPaperSizePresetsMm();

        if (selectedSize !== 'custom' && paperSizes[selectedSize]) {
            const size = paperSizes[selectedSize];
            this.writeLengthInputFromMm(document.getElementById('paperWidth'), size.width);
            this.writeLengthInputFromMm(document.getElementById('paperHeight'), size.height);

            // If the current outer margin is invalid for this preset (or was set to something huge),
            // reset it to a sane default so users don't get negative printable areas.
            const outerMarginInput = document.getElementById('outerMargin');
            if (outerMarginInput) {
                const currentOuter = this.readLengthInputMm(outerMarginInput);
                const maxOuter = Math.max(0, (Math.min(size.width, size.height) / 2) - 0.1);
                if (!Number.isFinite(currentOuter) || currentOuter > maxOuter) {
                    this.writeLengthInputFromMm(outerMarginInput, 10);
                }
            }
            
            // Trigger validation after changing paper size
            this.validateAllImages();
            this.markLayoutStale();
        }
    }

    checkForCustomSize() {
        const { paperWidth: currentWidth, paperHeight: currentHeight } = this.getPaperSettingsMm();
        const paperSizeSelect = document.getElementById('paperSize');
        
        const paperSizes = this.getPaperSizePresetsMm();

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
                        <label>Width (${this.getUnitLabel()}):</label>
                        <input type="number" value="${this.formatLengthValue(imageData.width)}" min="1" step="0.1" data-min-mm="1" data-step-mm="0.1"
                               onchange="sheetBuilder.updateImageSize(${imageData.id}, 'width', this.value)">
                    </div>
                </div>
                <div>
                    <label>Height (${this.getUnitLabel()}):</label>
                    <input type="number" value="${this.formatLengthValue(imageData.height)}" min="1" step="0.1" data-min-mm="1" data-step-mm="0.1"
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

        this.syncAllLengthInputsAttributes();
        
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

        this.markLayoutStale();
    }

    updateImageSize(id, property, value) {
        const image = this.images.find(img => img.id === id);
        if (image) {
            const numValue = property === 'copies' ? parseFloat(value) : this.parseLengthFromInputValue(value);
            
            if (property === 'width') {
                image.width = numValue;
                // Maintain aspect ratio
                image.height = numValue / image.aspectRatio;
                // Update the height input field
                const config = document.querySelector(`[data-id="${id}"]`);
                const heightInput = config.querySelector('input[onchange*="height"]');
                heightInput.value = this.formatLengthValue(image.height);
            } else if (property === 'height') {
                image.height = numValue;
                // Maintain aspect ratio
                image.width = numValue * image.aspectRatio;
                // Update the width input field
                const config = document.querySelector(`[data-id="${id}"]`);
                const widthInput = config.querySelector('input[onchange*="width"]');
                widthInput.value = this.formatLengthValue(image.width);
            } else {
                image[property] = numValue;
            }
            
            // Real-time validation
            this.validateImageInRealTime(image);

            this.markLayoutStale();
        }
    }

    validateImageInRealTime(image) {
        const { paperWidth, paperHeight, outerMargin } = this.getPaperSettingsMm();

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
                warning.innerHTML = `<i class="fa-solid fa-triangle-exclamation" aria-hidden="true"></i> Too large for paper! Max: ${this.formatSizePair(maxW, maxH)}`;
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
            return `Outer margin is too large for the selected paper size (printable area would be ${this.formatSizePair(printableWidth, printableHeight)}). Reduce the outer margin or increase paper size.`;
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
        
        const width = this.parseLengthFromInputValue(event.target.value);
        if (isNaN(width) || width <= 0) return;
        
        // Use average aspect ratio of selected images or a default ratio
        let aspectRatio = this.getAverageSelectedAspectRatio();
        if (!aspectRatio) aspectRatio = 1.5; // Default aspect ratio
        
        const height = width / aspectRatio;
        this.writeLengthInputFromMm(document.getElementById('bulkHeight'), height);
    }

    handleBulkHeightChange(event) {
        if (!this.isRatioLocked) return;
        
        const height = this.parseLengthFromInputValue(event.target.value);
        if (isNaN(height) || height <= 0) return;
        
        // Use average aspect ratio of selected images or a default ratio
        let aspectRatio = this.getAverageSelectedAspectRatio();
        if (!aspectRatio) aspectRatio = 1.5; // Default aspect ratio
        
        const width = height * aspectRatio;
        this.writeLengthInputFromMm(document.getElementById('bulkWidth'), width);
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
                img.width = this.parseLengthFromInputValue(bulkWidth);
                if (this.isRatioLocked) {
                    // Use the original image aspect ratio, not bulk aspect ratio
                    img.height = img.width / img.aspectRatio;
                }
            }
            
            if (bulkHeight) {
                if (this.isRatioLocked) {
                    // If ratio locked and height is provided, adjust width to maintain ratio
                    img.height = this.parseLengthFromInputValue(bulkHeight);
                    img.width = img.height * img.aspectRatio;
                } else {
                    // If ratio not locked, just set height
                    img.height = this.parseLengthFromInputValue(bulkHeight);
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

        this.markLayoutStale();
    }

    updateImageConfigDisplay(imageData) {
        const config = document.querySelector(`[data-id="${imageData.id}"]`);
        if (!config) return;

        const widthInput = config.querySelector('input[onchange*="width"]');
        const heightInput = config.querySelector('input[onchange*="height"]');
        const copiesInput = config.querySelector('input[onchange*="copies"]');

        if (widthInput) widthInput.value = this.formatLengthValue(imageData.width);
        if (heightInput) heightInput.value = this.formatLengthValue(imageData.height);
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

        // Upload feedback: show progress so users know the app is working.
        this._uploadFeedback.token += 1;
        const uploadToken = this._uploadFeedback.token;
        if (this._uploadFeedback.hideTimer) {
            window.clearTimeout(this._uploadFeedback.hideTimer);
            this._uploadFeedback.hideTimer = null;
        }
        const total = imageFiles.length;
        let completed = 0;
        let failed = 0;
        const updateStatus = (final = false) => {
            if (this._uploadFeedback.token !== uploadToken) return;
            if (final) {
                if (failed > 0) {
                    this.setUploadStatus(`Added ${total - failed}/${total} images (some failed).`, { visible: true });
                } else {
                    this.setUploadStatus(`Added ${total} image${total === 1 ? '' : 's'}.`, { visible: true });
                }
                this._uploadFeedback.hideTimer = window.setTimeout(() => {
                    if (this._uploadFeedback.token !== uploadToken) return;
                    this.setUploadStatus('', { visible: false });
                }, 1400);
                return;
            }
            this.setUploadStatus(`Loading ${completed}/${total} image${total === 1 ? '' : 's'}…`, { visible: true });
        };

        updateStatus(false);

        this.markLayoutStale({ resetFillUntilPages: false });

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

                    // Keep Fill-until-pages synced to the new minimum requirement.
                    this.setFillUntilPagesMin(this.computeMinimumPagesRequired());

                    // Generate a lightweight preview in the background for layout preview.
                    this.createPreviewDataUrl(file).then((previewDataUrl) => {
                        if (previewDataUrl) {
                            imageData.previewDataUrl = previewDataUrl;
                            this.scheduleLayoutPreviewRerender();
                        }
                    });

                    completed += 1;
                    updateStatus(completed >= total);
                };
                img.onerror = () => {
                    failed += 1;
                    completed += 1;
                    updateStatus(completed >= total);
                };
                img.src = e.target.result;
            };
            reader.onerror = () => {
                failed += 1;
                completed += 1;
                updateStatus(completed >= total);
            };
            reader.readAsDataURL(file);
        });
    }

    async loadImageElement(src) {
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
        } catch (e) {
            throw e;
        }
        return img;
    }

    computeTransparentPaddingCropBounds(imageData, width, height, options = {}) {
        const {
            alphaThreshold = 8,
            maxNonTransparentRatio = 0.01
        } = options;

        if (!imageData || width <= 0 || height <= 0) return null;
        const data = imageData.data;
        if (!data || data.length < width * height * 4) return null;

        const maxNonTransparentPerRow = Math.max(0, Math.floor(width * maxNonTransparentRatio));
        const maxNonTransparentPerCol = Math.max(0, Math.floor(height * maxNonTransparentRatio));

        const rowIsMostlyTransparent = (y) => {
            let nonTransparent = 0;
            let idx = (y * width * 4) + 3;
            for (let x = 0; x < width; x++) {
                if (data[idx] > alphaThreshold) {
                    nonTransparent++;
                    if (nonTransparent > maxNonTransparentPerRow) return false;
                }
                idx += 4;
            }
            return true;
        };

        const colIsMostlyTransparent = (x, top, bottom) => {
            let nonTransparent = 0;
            for (let y = top; y <= bottom; y++) {
                const idx = ((y * width + x) * 4) + 3;
                if (data[idx] > alphaThreshold) {
                    nonTransparent++;
                    if (nonTransparent > maxNonTransparentPerCol) return false;
                }
            }
            return true;
        };

        let top = 0;
        while (top < height && rowIsMostlyTransparent(top)) top++;
        if (top >= height) return null; // fully transparent

        let bottom = height - 1;
        while (bottom >= top && rowIsMostlyTransparent(bottom)) bottom--;

        let left = 0;
        while (left < width && colIsMostlyTransparent(left, top, bottom)) left++;

        let right = width - 1;
        while (right >= left && colIsMostlyTransparent(right, top, bottom)) right--;

        const cropW = right - left + 1;
        const cropH = bottom - top + 1;
        if (cropW <= 0 || cropH <= 0) return null;

        // No-op crop.
        if (left === 0 && top === 0 && cropW === width && cropH === height) return null;

        return { x: left, y: top, width: cropW, height: cropH };
    }

    async cropImageDataUrlTransparentPadding(dataUrl, options = {}) {
        const img = await this.loadImageElement(dataUrl);

        const canvas = document.createElement('canvas');
        canvas.width = Math.max(1, img.naturalWidth || img.width || 1);
        canvas.height = Math.max(1, img.naturalHeight || img.height || 1);

        const ctx = canvas.getContext('2d', { willReadFrequently: true });
        if (!ctx) return null;
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(img, 0, 0);

        const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const bounds = this.computeTransparentPaddingCropBounds(imageData, canvas.width, canvas.height, options);
        if (!bounds) return null;

        const out = document.createElement('canvas');
        out.width = bounds.width;
        out.height = bounds.height;
        const outCtx = out.getContext('2d');
        if (!outCtx) return null;

        outCtx.clearRect(0, 0, out.width, out.height);
        outCtx.drawImage(canvas, bounds.x, bounds.y, bounds.width, bounds.height, 0, 0, bounds.width, bounds.height);

        return {
            dataUrl: out.toDataURL('image/png'),
            width: bounds.width,
            height: bounds.height
        };
    }

    async cropAllImagesTransparentPadding() {
        if (!this.images || this.images.length === 0) {
            alert('Please upload some images first!');
            return;
        }

        const options = {
            // Treat alpha <= 8 as transparent.
            alphaThreshold: 8,
            // Allow a small amount of noise: rows/cols with <= 1% non-transparent pixels are trimmed.
            maxNonTransparentRatio: 0.01
        };

        let changed = 0;

        for (const img of this.images) {
            if (!img?.dataUrl) continue;

            const originalPixelW = Number(img.originalWidth) || 0;
            const originalPixelH = Number(img.originalHeight) || 0;

            let result = null;
            try {
                result = await this.cropImageDataUrlTransparentPadding(img.dataUrl, options);
            } catch {
                result = null;
            }

            if (!result) continue;

            // Preserve the physical scale of the actual artwork: shrink mm size proportionally
            // to the pixel crop so the visible content doesn't get enlarged.
            const widthMmPerPx = (Number.isFinite(img.width) && originalPixelW > 0) ? (img.width / originalPixelW) : null;
            const heightMmPerPx = (Number.isFinite(img.height) && originalPixelH > 0) ? (img.height / originalPixelH) : null;

            img.dataUrl = result.dataUrl;
            img.previewDataUrl = null;
            img.originalWidth = result.width;
            img.originalHeight = result.height;
            img.aspectRatio = result.width / result.height;

            if (widthMmPerPx !== null) {
                img.width = widthMmPerPx * result.width;
            }
            if (heightMmPerPx !== null) {
                img.height = heightMmPerPx * result.height;
            } else if (widthMmPerPx !== null) {
                img.height = img.width / img.aspectRatio;
            }

            changed += 1;
        }

        if (changed > 0) {
            this.rerenderAllImageConfigs();
            this.validateAllImages();
            this.markLayoutStale({ resetFillUntilPages: true });
        }
    }

    generateLayout() {
        if (this.images.length === 0) {
            alert('Please upload some images first!');
            return;
        }

        const { paperWidth, paperHeight, outerMargin, innerMargin } = this.getPaperSettingsMm();
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
                warningMessage += `Printable area: ${this.formatLengthValue(printableWidth)}${this.getUnitLabel()} × ${this.formatLengthValue(printableHeight)}${this.getUnitLabel()}\n\n`;
                requiresRotation.forEach(img => {
                    warningMessage += `• ${img.name}: ${this.formatLengthValue(img.width)}${this.getUnitLabel()} × ${this.formatLengthValue(img.height)}${this.getUnitLabel()}\n`;
                });
                warningMessage += `\nEnable “Allow Rotation” or reduce the size.`;
                alert(warningMessage);
                return;
            }
        }

        // Expand images based on copies (this is the "requested" set)
        const requestedImages = [];
        this.images.forEach(img => {
            const copies = Math.max(1, Math.floor(Number(img.copies) || 1));
            for (let i = 0; i < copies; i++) {
                requestedImages.push({
                    ...img,
                    isRequired: true,
                    copyIndex: i,
                    originalId: img.id
                });
            }
        });

        // 1) Pack the requested set to find the true minimum pages.
        const baseLayout = this.packImages(requestedImages, printableWidth, printableHeight, innerMargin, allowRotation);
        const minPages = Math.max(1, baseLayout.length);

        // Keep the input's minimum synced, without overwriting larger user-entered values.
        this.setFillUntilPagesMin(minPages);

        // If fill is not enabled, just use the minimum required layout.
        if (!this.isFillUntilEnabled()) {
            this.layout = baseLayout;
        } else {
            const targetPages = Math.max(minPages, this.getFillUntilPages());

            // Always attempt at least one duplicate batch when filling, so "fill 1 page" can add more copies.
            let multiplier = Math.max(2, Math.ceil(targetPages / Math.max(1, minPages)));
            const maxAttempts = 30;
            let finalLayout = baseLayout;

            for (let attempt = 0; attempt < maxAttempts; attempt++) {
                const expanded = [];
                // Required set first
                for (const img of requestedImages) expanded.push(img);

                // Optional duplicates
                for (let batch = 0; batch < (multiplier - 1); batch++) {
                    for (const img of requestedImages) {
                        expanded.push({
                            ...img,
                            isRequired: false,
                            duplicateSetIndex: batch
                        });
                    }
                }

                const layout = this.packImages(expanded, printableWidth, printableHeight, innerMargin, allowRotation);
                finalLayout = layout;

                if (layout.length >= targetPages) {
                    // Safe because required items are packed first.
                    finalLayout = layout.slice(0, targetPages);
                    break;
                }

                multiplier += 1;
            }

            this.layout = finalLayout;
        }
        
        this.renderLayoutPreview(paperWidth, paperHeight, outerMargin);
        this.lastPreviewParams = { paperWidth, paperHeight, outerMargin };
        this.updateLayoutStats();

        this.setExportButtonsDisabled(false);
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
        warningMessage += `Printable area: ${this.formatLengthValue(Math.max(0, printableWidth))}${this.getUnitLabel()} × ${this.formatLengthValue(Math.max(0, printableHeight))}${this.getUnitLabel()}\n\n`;
        
        oversizedImages.forEach(img => {
            warningMessage += `• ${img.name}: ${this.formatLengthValue(img.width)}${this.getUnitLabel()} × ${this.formatLengthValue(img.height)}${this.getUnitLabel()}\n`;
            const maxW = Math.max(0, Math.max(printableWidth, printableHeight));
            const maxH = Math.max(0, Math.min(printableWidth, printableHeight));
            warningMessage += `  Max size (any orientation): ${this.formatLengthValue(maxW)}${this.getUnitLabel()} × ${this.formatLengthValue(maxH)}${this.getUnitLabel()}\n\n`;
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
                warning.innerHTML = `<i class="fa-solid fa-triangle-exclamation" aria-hidden="true"></i> Too large for paper! Max: ${this.formatSizePair(img.maxWidth, img.maxHeight)}`;
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
            const aRequired = Boolean(a.isRequired);
            const bRequired = Boolean(b.isRequired);
            if (aRequired !== bRequired) return aRequired ? -1 : 1;
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
        
        console.log(`Finding placement for image ${image.name} (${this.formatLengthValue(image.width)}x${this.formatLengthValue(image.height)}${this.getUnitLabel()})`);
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
        const totalImages = this.layout.reduce((sum, page) => sum + (page.images?.length || 0), 0);
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

    getFillUntilPages() {
        const el = document.getElementById('fillUntilPages');
        const raw = el ? String(el.value).trim() : '';
        const parsed = Number.parseInt(raw, 10);
        if (!Number.isFinite(parsed) || parsed < 1) return 1;
        return parsed;
    }

    isFillUntilEnabled() {
        return Boolean(document.getElementById('fillUntilPagesEnabled')?.checked);
    }

    setFillUntilPagesMin(minPages) {
        const el = document.getElementById('fillUntilPages');
        if (!el) return;

        const safeMin = Math.max(1, Math.floor(Number(minPages) || 1));
        el.min = String(safeMin);

        const current = this.getFillUntilPages();
        if (current < safeMin) {
            this.setFillUntilPagesInputValue(safeMin);
        }
    }

    getPdfCompressionSetting() {
        const el = document.getElementById('pdfCompression');
        const v = el ? String(el.value || '').trim() : '';
        return v || 'perfect';
    }

    getPdfCompressionOptions() {
        const setting = this.getPdfCompressionSetting();

        // jsPDF's addImage supports an optional compression mode for PNG (Flate) and JPEG.
        // For lossless modes we keep PNG; for tiers we convert to JPEG at varying quality.
        switch (setting) {
            case 'high':
                return { outputFormat: 'JPEG', jsPdfCompression: 'MEDIUM', jpegQuality: 0.9 };
            case 'medium':
                return { outputFormat: 'JPEG', jsPdfCompression: 'MEDIUM', jpegQuality: 0.8 };
            case 'low':
                return { outputFormat: 'JPEG', jsPdfCompression: 'FAST', jpegQuality: 0.65 };
            case 'perfect':
            default:
                return { outputFormat: 'PNG', jsPdfCompression: 'NONE', jpegQuality: null };
        }
    }

    setFillUntilPagesInputValue(value) {
        const el = document.getElementById('fillUntilPages');
        if (!el) return;
        const v = Math.max(1, Math.floor(Number(value)));
        el.value = String(Number.isFinite(v) ? v : 1);
    }

    computeMinimumPagesRequired() {
        if (!this.images || this.images.length === 0) return 1;

        const { paperWidth, paperHeight, outerMargin, innerMargin } = this.getPaperSettingsMm();
        const allowRotation = Boolean(document.getElementById('rotateImages')?.checked);
        const paperIssue = this.getPaperSettingsIssue(paperWidth, paperHeight, outerMargin);
        if (paperIssue) return 1;

        const printableWidth = paperWidth - (2 * outerMargin);
        const printableHeight = paperHeight - (2 * outerMargin);
        if (!Number.isFinite(printableWidth) || !Number.isFinite(printableHeight) || printableWidth <= 0 || printableHeight <= 0) return 1;

        const oversizedImages = this.validateImageSizes(printableWidth, printableHeight);
        if (oversizedImages.length > 0) return 1;

        if (!allowRotation) {
            const requiresRotation = this.images.filter(img => img.width > printableWidth || img.height > printableHeight);
            if (requiresRotation.length > 0) return 1;
        }

        const requestedImages = [];
        this.images.forEach(img => {
            const copies = Math.max(1, Math.floor(Number(img.copies) || 1));
            for (let i = 0; i < copies; i++) {
                requestedImages.push({
                    ...img,
                    copyIndex: i,
                    originalId: img.id
                });
            }
        });

        const layout = this.packImages(requestedImages, printableWidth, printableHeight, innerMargin, allowRotation);
        return Math.max(1, layout.length);
    }

    markLayoutStale({ resetFillUntilPages = true } = {}) {
        this.layout = [];
        this.lastPreviewParams = null;

        this.setExportButtonsDisabled(true);

        const preview = document.getElementById('layoutPreview');
        if (preview) preview.innerHTML = '';

        const stats = document.getElementById('layoutStats');
        if (stats) stats.innerHTML = '';

        if (resetFillUntilPages) {
            this.setFillUntilPagesMin(this.computeMinimumPagesRequired());
        }
    }

    setExportButtonsDisabled(disabled) {
        const btns = [
            document.getElementById('exportPDF'),
            document.getElementById('exportPDFBottom')
        ].filter(Boolean);

        for (const btn of btns) {
            btn.disabled = Boolean(disabled);
        }
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

        this.setExportButtonsDisabled(true);
        if (cancelBtn) {
            cancelBtn.disabled = false;
            cancelBtn.textContent = 'Cancel';
        }

        this.setExportOverlayVisible(true);

        const { jsPDF } = window.jspdf;
        const { paperWidth, paperHeight, outerMargin } = this.getPaperSettingsMm();
        const { outputFormat, jsPdfCompression, jpegQuality } = this.getPdfCompressionOptions();

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
                        const imageData = await this.loadImageForPDF(img, abortController.signal, {
                            outputFormat,
                            jpegQuality
                        });
                        throwIfAborted();

                        pdf.addImage(
                            imageData,
                            outputFormat,
                            outerMargin + img.x,
                            outerMargin + img.y,
                            img.width,
                            img.height,
                            undefined,
                            jsPdfCompression
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

            this.setExportButtonsDisabled(false);
            this._exportState = {
                inProgress: false,
                abortController: null
            };
        }
    }

    loadImageForPDF(imgData, abortSignal, options = {}) {
        const {
            outputFormat = 'PNG',
            jpegQuality = 0.85
        } = options;

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

                if (String(outputFormat).toUpperCase() === 'JPEG') {
                    // JPEG is lossy; keep the white background fill above.
                    const q = Number.isFinite(jpegQuality) ? Math.min(1, Math.max(0.1, jpegQuality)) : 0.85;
                    resolve(canvas.toDataURL('image/jpeg', q));
                    return;
                }

                // Default: lossless PNG.
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