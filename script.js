class FitPrint {
    constructor() {
        this.images = [];
        this.layout = [];
        this.selectedImages = new Set();
        this.lastSelectedId = null;
        this.isRatioLocked = true;
        this.bulkAspectRatio = null;
        this.initializeEventListeners();
        this.initializeTheme();
    }

    initializeEventListeners() {
        const imageInput = document.getElementById('imageInput');
        const generateBtn = document.getElementById('generateLayout');
        const exportBtn = document.getElementById('exportPDF');
        const themeToggle = document.getElementById('themeToggle');
        const paperSize = document.getElementById('paperSize');
        
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
        themeToggle.addEventListener('click', () => this.toggleTheme());
        paperSize.addEventListener('change', (e) => this.handlePaperSizeChange(e));
        
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
        
        // Navigation event listeners
        const navToggle = document.getElementById('navToggle');
        const sidebar = document.getElementById('sidebar');
        const sidebarOverlay = document.getElementById('sidebarOverlay');
        const navLinks = document.querySelectorAll('.nav-link');
        
        if (navToggle) navToggle.addEventListener('click', () => this.toggleSidebar());
        if (sidebarOverlay) sidebarOverlay.addEventListener('click', () => this.closeSidebar());
        
        // Navigation link event listeners
        navLinks.forEach(link => {
            link.addEventListener('click', (e) => {
                e.preventDefault();
                const sectionId = link.getAttribute('data-section');
                this.navigateToSection(sectionId);
                this.closeSidebar(); // Close sidebar on mobile after navigation
            });
        });
        
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

    initializeTheme() {
        const savedTheme = localStorage.getItem('fitprint-theme') || 'light';
        this.setTheme(savedTheme);
    }

    toggleTheme() {
        const currentTheme = document.documentElement.getAttribute('data-theme') || 'light';
        const newTheme = currentTheme === 'light' ? 'dark' : 'light';
        this.setTheme(newTheme);
    }

    setTheme(theme) {
        document.documentElement.setAttribute('data-theme', theme);
        localStorage.setItem('fitprint-theme', theme);
        
        const themeIcon = document.querySelector('.theme-icon');
        if (themeIcon) {
            themeIcon.textContent = theme === 'light' ? '🌙' : '☀️';
        }
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
                   onclick="fitPrint.handleCheckboxClick(event, ${imageData.id})">
            <img src="${imageData.dataUrl}" alt="${imageData.name}" 
                 onclick="fitPrint.showImageModal('${imageData.dataUrl}', '${imageData.name}', '${imageData.originalWidth}x${imageData.originalHeight}')">
            <div class="config-inputs">
                <div class="input-with-lock">
                    <div style="width: 100%;">
                        <label>Width (mm):</label>
                        <input type="number" value="${imageData.width.toFixed(1)}" min="1" step="0.1" 
                               onchange="fitPrint.updateImageSize(${imageData.id}, 'width', this.value)">
                    </div>
                </div>
                <div>
                    <label>Height (mm):</label>
                    <input type="number" value="${imageData.height.toFixed(1)}" min="1" step="0.1" 
                           onchange="fitPrint.updateImageSize(${imageData.id}, 'height', this.value)">
                </div>
                <div>
                    <label>Copies:</label>
                    <input type="number" value="${imageData.copies}" min="1" 
                           onchange="fitPrint.updateImageSize(${imageData.id}, 'copies', this.value)">
                </div>
            </div>
            <button class="remove-btn" onclick="fitPrint.removeImage(${imageData.id})">Remove</button>
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
        
        const printableWidth = paperWidth - (2 * outerMargin);
        const printableHeight = paperHeight - (2 * outerMargin);
        
        const config = document.querySelector(`[data-id="${image.id}"]`);
        if (!config) return; // Safety check
        
        const existingWarning = config.querySelector('.size-warning');
        
        const canFitNormal = image.width <= printableWidth && image.height <= printableHeight;
        const canFitRotated = image.height <= printableWidth && image.width <= printableHeight;
        
        if (!canFitNormal && !canFitRotated) {
            // Add warning if not exists
            if (!existingWarning) {
                config.classList.add('oversized');
                const warning = document.createElement('div');
                warning.className = 'size-warning';
                warning.innerHTML = `⚠️ Too large for paper! Max: ${Math.max(printableWidth, printableHeight).toFixed(1)}×${Math.min(printableWidth, printableHeight).toFixed(1)}mm`;
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
        
        if (this.isRatioLocked) {
            lockBtn.classList.add('active');
            lockBtn.textContent = '🔒';
            lockBtn.title = 'Unlock aspect ratio';
        } else {
            lockBtn.classList.remove('active');
            lockBtn.textContent = '🔓';
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

    // Navigation Methods
    toggleSidebar() {
        const sidebar = document.getElementById('sidebar');
        const overlay = document.getElementById('sidebarOverlay');
        
        if (sidebar.classList.contains('open')) {
            this.closeSidebar();
        } else {
            this.openSidebar();
        }
    }

    openSidebar() {
        const sidebar = document.getElementById('sidebar');
        const overlay = document.getElementById('sidebarOverlay');
        
        sidebar.classList.add('open');
        overlay.classList.add('active');
        document.body.style.overflow = 'hidden'; // Prevent background scrolling
    }

    closeSidebar() {
        const sidebar = document.getElementById('sidebar');
        const overlay = document.getElementById('sidebarOverlay');
        
        sidebar.classList.remove('open');
        overlay.classList.remove('active');
        document.body.style.overflow = ''; // Restore scrolling
    }

    navigateToSection(sectionId) {
        // Remove active class from all nav links
        document.querySelectorAll('.nav-link').forEach(link => {
            link.classList.remove('active');
        });
        
        // Add active class to clicked link
        const activeLink = document.querySelector(`[data-section="${sectionId}"]`);
        if (activeLink) {
            activeLink.classList.add('active');
        }
        
        // Scroll to section
        const section = document.getElementById(sectionId.replace('-section', ''));
        if (section) {
            section.scrollIntoView({ 
                behavior: 'smooth',
                block: 'start'
            });
        }
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

        // Calculate printable area
        const printableWidth = paperWidth - (2 * outerMargin);
        const printableHeight = paperHeight - (2 * outerMargin);

        // Validate that all images can fit on the paper
        const oversizedImages = this.validateImageSizes(printableWidth, printableHeight);
        if (oversizedImages.length > 0) {
            this.showOversizedWarning(oversizedImages, printableWidth, printableHeight);
            return;
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
        this.layout = this.packImages(allImages, printableWidth, printableHeight, innerMargin);
        
        this.renderLayoutPreview(paperWidth, paperHeight, outerMargin);
        this.updateLayoutStats();
        
        document.getElementById('exportPDF').disabled = false;
    }

    validateImageSizes(printableWidth, printableHeight) {
        const oversizedImages = [];
        
        this.images.forEach(img => {
            const canFitNormal = img.width <= printableWidth && img.height <= printableHeight;
            const canFitRotated = img.height <= printableWidth && img.width <= printableHeight;
            
            if (!canFitNormal && !canFitRotated) {
                oversizedImages.push({
                    ...img,
                    maxWidth: Math.max(printableWidth, printableHeight),
                    maxHeight: Math.min(printableWidth, printableHeight)
                });
            }
        });
        
        return oversizedImages;
    }

    showOversizedWarning(oversizedImages, printableWidth, printableHeight) {
        let warningMessage = `⚠️ WARNING: The following images are too large to fit on the paper:\n\n`;
        warningMessage += `Printable area: ${printableWidth.toFixed(1)}mm × ${printableHeight.toFixed(1)}mm\n\n`;
        
        oversizedImages.forEach(img => {
            warningMessage += `• ${img.name}: ${img.width.toFixed(1)}mm × ${img.height.toFixed(1)}mm\n`;
            warningMessage += `  Max size (any orientation): ${Math.max(printableWidth, printableHeight).toFixed(1)}mm × ${Math.min(printableWidth, printableHeight).toFixed(1)}mm\n\n`;
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
                warning.innerHTML = `⚠️ Too large for paper! Max: ${img.maxWidth.toFixed(1)}×${img.maxHeight.toFixed(1)}mm`;
                config.appendChild(warning);
            }
        });
    }

    packImages(images, pageWidth, pageHeight, margin) {
        // Sort images by area (largest first) for better packing efficiency
        const sortedImages = [...images].sort((a, b) => (b.width * b.height) - (a.width * a.height));
        
        const pages = [];
        let currentPageIndex = 0;
        
        for (const image of sortedImages) {
            let placed = false;
            
            // Try to place on existing pages first
            for (let pageIndex = 0; pageIndex < pages.length; pageIndex++) {
                const placement = this.findBestPlacementOnPage(image, pages[pageIndex], pageWidth, pageHeight, margin);
                if (placement) {
                    pages[pageIndex].images.push({
                        ...image,
                        x: placement.x,
                        y: placement.y,
                        width: placement.width,
                        height: placement.height,
                        rotated: placement.rotated
                    });
                    placed = true;
                    break;
                }
            }
            
            // If not placed on existing pages, create a new page
            if (!placed) {
                const newPage = { images: [] };
                const placement = this.findBestPlacementOnPage(image, newPage, pageWidth, pageHeight, margin);
                
                if (placement) {
                    newPage.images.push({
                        ...image,
                        x: placement.x,
                        y: placement.y,
                        width: placement.width,
                        height: placement.height,
                        rotated: placement.rotated
                    });
                    pages.push(newPage);
                } else {
                    console.error('Image too large to fit on page:', image);
                }
            }
        }
        
        return pages;
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
                <div class="page-title">Page ${index + 1}</div>
                <div class="page-content" style="width: ${previewWidth}px; height: ${previewHeight}px; overflow: hidden; position: relative;">
                </div>
            `;

            const pageContent = pageDiv.querySelector('.page-content');

            page.images.forEach(img => {
                const imgDiv = document.createElement('div');
                imgDiv.className = 'placed-image';
                imgDiv.style.left = `${(outerMargin + img.x) * scale}px`;
                imgDiv.style.top = `${(outerMargin + img.y) * scale}px`;
                imgDiv.style.width = `${img.width * scale}px`;
                imgDiv.style.height = `${img.height * scale}px`;
                imgDiv.textContent = `${img.name}${img.rotated ? ' (R)' : ''}`;
                pageContent.appendChild(imgDiv);
            });

            preview.appendChild(pageDiv);
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
        if (this.layout.length === 0) {
            alert('Please generate a layout first!');
            return;
        }

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

        for (let pageIndex = 0; pageIndex < this.layout.length; pageIndex++) {
            if (pageIndex > 0) {
                pdf.addPage();
            }

            const page = this.layout[pageIndex];

            for (const img of page.images) {
                try {
                    // Load image and add to PDF
                    const imageData = await this.loadImageForPDF(img);
                    
                    pdf.addImage(
                        imageData,
                        'JPEG',
                        outerMargin + img.x,
                        outerMargin + img.y,
                        img.width,
                        img.height,
                        undefined,
                        'FAST'
                    );
                } catch (error) {
                    console.error('Error adding image to PDF:', error);
                }
            }
        }

        pdf.save('fitprint-layout.pdf');
    }

    loadImageForPDF(imgData) {
        return new Promise((resolve, reject) => {
            const img = new Image();
            img.onload = () => {
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                
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
                
                resolve(canvas.toDataURL('image/jpeg', 0.8));
            };
            img.onerror = reject;
            img.src = imgData.dataUrl;
        });
    }
}

// Initialize the application
const fitPrint = new FitPrint();