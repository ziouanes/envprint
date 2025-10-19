class EnvelopePrinter {
    constructor() {
        this.addresses = [];
        this.currentAddressIndex = 0;
        this.settings = {
            envelopeSize: 'standard',
            fontSize: 14,
            fontFamily: 'Arial',
            customWidth: null,
            customHeight: null
        };

        this.initializeEventListeners();
    }

    initializeEventListeners() {
        // File input handlers
        const fileInput = document.getElementById('fileInput');
        const removeFile = document.getElementById('removeFile');

        fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        removeFile.addEventListener('click', () => this.removeFile());

        // Settings handlers
        const envelopeSize = document.getElementById('envelopeSize');
        const fontSize = document.getElementById('fontSize');
        const fontFamily = document.getElementById('fontFamily');
        const customWidth = document.getElementById('customWidth');
        const customHeight = document.getElementById('customHeight');

        envelopeSize.addEventListener('change', (e) => this.updateEnvelopeSize(e.target.value));
        fontSize.addEventListener('change', (e) => this.updateFontSize(e.target.value));
        fontFamily.addEventListener('change', (e) => this.updateFontFamily(e.target.value));
        customWidth.addEventListener('input', (e) => this.updateCustomSize('width', e.target.value));
        customHeight.addEventListener('input', (e) => this.updateCustomSize('height', e.target.value));

        // Navigation handlers
        const prevBtn = document.getElementById('prevAddress');
        const nextBtn = document.getElementById('nextAddress');
        
        prevBtn.addEventListener('click', () => this.navigateAddress(-1));
        nextBtn.addEventListener('click', () => this.navigateAddress(1));

        // Action button handlers
        const previewBtn = document.getElementById('previewBtn');
        const printAllBtn = document.getElementById('printAllBtn');
        const printCurrentBtn = document.getElementById('printCurrentBtn');

        previewBtn.addEventListener('click', () => this.showPrintPreview());
        printAllBtn.addEventListener('click', () => this.printAllEnvelopes());
        printCurrentBtn.addEventListener('click', () => this.printCurrentEnvelope());

        // Error message handler
        const closeError = document.getElementById('closeError');
        closeError.addEventListener('click', () => this.hideError());

        // Return address editing
        const returnAddress = document.getElementById('returnAddress');
        returnAddress.addEventListener('blur', () => this.saveReturnAddress());
    }

    async handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;

        this.showProgress('Reading file...', 0);

        try {
            const extension = file.name.split('.').pop().toLowerCase();
            let addresses = [];

            if (extension === 'csv') {
                addresses = await this.parseCSV(file);
            } else if (extension === 'xlsx' || extension === 'xls') {
                addresses = await this.parseExcel(file);
            } else {
                throw new Error('Unsupported file format. Please use CSV or Excel files.');
            }

            this.addresses = addresses;
            this.currentAddressIndex = 0;

            this.showFileInfo(file.name);
            this.showSettings();
            this.showPreview();
            this.showActionButtons();
            this.updatePreview();
            this.hideProgress();

        } catch (error) {
            this.hideProgress();
            this.showError(`Error reading file: ${error.message}`);
            this.removeFile();
        }
    }

    async parseCSV(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const csv = e.target.result;
                    const lines = csv.split('\n').filter(line => line.trim());
                    
                    if (lines.length < 2) {
                        reject(new Error('CSV file must contain headers and at least one address'));
                        return;
                    }

                    const headers = this.parseCSVLine(lines[0]).map(h => h.toLowerCase().trim());
                    const addresses = [];

                    for (let i = 1; i < lines.length; i++) {
                        const values = this.parseCSVLine(lines[i]);
                        if (values.length === 0) continue;

                        const address = this.mapAddressFields(headers, values);
                        if (this.validateAddress(address)) {
                            addresses.push(address);
                        }
                    }

                    if (addresses.length === 0) {
                        reject(new Error('No valid addresses found in the file'));
                        return;
                    }

                    resolve(addresses);
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = () => reject(new Error('Error reading CSV file'));
            reader.readAsText(file);
        });
    }

    async parseExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    if (jsonData.length < 2) {
                        reject(new Error('Excel file must contain headers and at least one address'));
                        return;
                    }

                    const headers = jsonData[0].map(h => String(h).toLowerCase().trim());
                    const addresses = [];

                    for (let i = 1; i < jsonData.length; i++) {
                        const values = jsonData[i];
                        if (!values || values.length === 0) continue;

                        const address = this.mapAddressFields(headers, values);
                        if (this.validateAddress(address)) {
                            addresses.push(address);
                        }
                    }

                    if (addresses.length === 0) {
                        reject(new Error('No valid addresses found in the file'));
                        return;
                    }

                    resolve(addresses);
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = () => reject(new Error('Error reading Excel file'));
            reader.readAsArrayBuffer(file);
        });
    }

    parseCSVLine(line) {
        const result = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            
            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === ',' && !inQuotes) {
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        
        result.push(current.trim());
        return result;
    }

    mapAddressFields(headers, values) {
        const fieldMappings = {
            name: ['name', 'full name', 'recipient', 'to'],
            address1: ['address', 'address1', 'street', 'address line 1'],
            address2: ['address2', 'suite', 'apt', 'apartment', 'address line 2'],
            city: ['city', 'town'],
            state: ['state', 'province', 'region'],
            zip: ['zip', 'zipcode', 'postal', 'postal code', 'postcode'],
            country: ['country']
        };

        const address = {};
        
        for (const [field, possibleHeaders] of Object.entries(fieldMappings)) {
            const headerIndex = headers.findIndex(h => 
                possibleHeaders.some(ph => h.includes(ph))
            );
            
            if (headerIndex !== -1 && values[headerIndex]) {
                address[field] = String(values[headerIndex]).trim();
            }
        }

        return address;
    }

    validateAddress(address) {
        return address.name && (address.address1 || address.city);
    }

    formatAddress(address) {
        const lines = [];
        
        if (address.name) lines.push(address.name);
        if (address.address1) lines.push(address.address1);
        if (address.address2) lines.push(address.address2);
        
        const cityStateZip = [address.city, address.state, address.zip]
            .filter(Boolean)
            .join(', ');
        
        if (cityStateZip) lines.push(cityStateZip);
        if (address.country && address.country.toLowerCase() !== 'usa') {
            lines.push(address.country);
        }

        return lines.join('<br>');
    }

    showFileInfo(fileName) {
        const fileInfo = document.getElementById('fileInfo');
        const fileNameSpan = document.getElementById('fileName');
        
        fileNameSpan.textContent = fileName;
        fileInfo.style.display = 'flex';
    }

    removeFile() {
        const fileInput = document.getElementById('fileInput');
        const fileInfo = document.getElementById('fileInfo');
        
        fileInput.value = '';
        fileInfo.style.display = 'none';
        
        this.addresses = [];
        this.hideSettings();
        this.hidePreview();
        this.hideActionButtons();
    }

    showSettings() {
        document.getElementById('settingsSection').style.display = 'block';
    }

    hideSettings() {
        document.getElementById('settingsSection').style.display = 'none';
    }

    showPreview() {
        document.getElementById('previewSection').style.display = 'block';
    }

    hidePreview() {
        document.getElementById('previewSection').style.display = 'none';
    }

    showActionButtons() {
        document.getElementById('actionButtons').style.display = 'flex';
    }

    hideActionButtons() {
        document.getElementById('actionButtons').style.display = 'none';
    }

    updateEnvelopeSize(size) {
        this.settings.envelopeSize = size;
        const customSizeGroup = document.getElementById('customSizeGroup');
        
        if (size === 'custom') {
            customSizeGroup.style.display = 'block';
        } else {
            customSizeGroup.style.display = 'none';
        }
        
        this.updateEnvelopePreviewSize();
    }

    updateFontSize(size) {
        this.settings.fontSize = parseInt(size);
        this.updatePreviewFonts();
    }

    updateFontFamily(family) {
        this.settings.fontFamily = family;
        this.updatePreviewFonts();
    }

    updateCustomSize(dimension, value) {
        if (dimension === 'width') {
            this.settings.customWidth = parseFloat(value) || null;
        } else {
            this.settings.customHeight = parseFloat(value) || null;
        }
        this.updateEnvelopePreviewSize();
    }

    updateEnvelopePreviewSize() {
        const preview = document.getElementById('envelopePreview');
        const content = preview.querySelector('.envelope-content');
        
        let width, height;
        
        switch (this.settings.envelopeSize) {
            case 'standard':
                width = 456; // 4.125" * 96dpi + padding
                height = 228; // 9.5" * 96dpi scaled down
                break;
            case 'a4':
                width = 600;
                height = 400;
                break;
            case 'legal':
                width = 650;
                height = 450;
                break;
            case 'custom':
                if (this.settings.customWidth && this.settings.customHeight) {
                    width = Math.min(this.settings.customWidth * 50, 800);
                    height = Math.min(this.settings.customHeight * 30, 500);
                } else {
                    return;
                }
                break;
            default:
                width = 456;
                height = 228;
        }
        
        content.style.width = `${width}px`;
        content.style.height = `${height}px`;
    }

    updatePreviewFonts() {
        const returnAddress = document.getElementById('returnAddress');
        const recipientAddress = document.getElementById('recipientAddress');
        
        const fontSize = `${this.settings.fontSize}px`;
        const fontFamily = this.settings.fontFamily;
        
        returnAddress.style.fontSize = fontSize;
        returnAddress.style.fontFamily = fontFamily;
        recipientAddress.style.fontSize = fontSize;
        recipientAddress.style.fontFamily = fontFamily;
    }

    updatePreview() {
        if (this.addresses.length === 0) return;

        const currentAddress = this.addresses[this.currentAddressIndex];
        const recipientAddress = document.getElementById('recipientAddress');
        const currentSpan = document.getElementById('currentAddress');
        const totalSpan = document.getElementById('totalAddresses');
        const prevBtn = document.getElementById('prevAddress');
        const nextBtn = document.getElementById('nextAddress');

        recipientAddress.innerHTML = this.formatAddress(currentAddress);
        currentSpan.textContent = this.currentAddressIndex + 1;
        totalSpan.textContent = this.addresses.length;

        prevBtn.disabled = this.currentAddressIndex === 0;
        nextBtn.disabled = this.currentAddressIndex === this.addresses.length - 1;

        this.updatePreviewFonts();
        this.updateEnvelopePreviewSize();
    }

    navigateAddress(direction) {
        const newIndex = this.currentAddressIndex + direction;
        
        if (newIndex >= 0 && newIndex < this.addresses.length) {
            this.currentAddressIndex = newIndex;
            this.updatePreview();
        }
    }

    saveReturnAddress() {
        // Return address is automatically saved as it's contenteditable
        // Could be extended to save to localStorage
    }

    showPrintPreview() {
        // Create a new window for print preview
        const previewWindow = window.open('', '_blank');
        const envelopes = this.generateAllEnvelopes();
        
        previewWindow.document.write(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>Envelope Print Preview</title>
                <style>
                    body { 
                        font-family: ${this.settings.fontFamily}; 
                        margin: 0; 
                        padding: 20px; 
                        background: #f5f5f5; 
                    }
                    .envelope { 
                        width: 100%; 
                        max-width: 600px; 
                        height: 400px; 
                        background: white; 
                        border: 1px solid #ddd; 
                        margin-bottom: 20px; 
                        padding: 1in; 
                        position: relative; 
                        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
                        page-break-after: always;
                    }
                    .return-address { 
                        position: absolute; 
                        top: 1in; 
                        left: 1in; 
                        font-size: ${this.settings.fontSize}px; 
                        line-height: 1.4;
                    }
                    .recipient-address { 
                        position: absolute; 
                        bottom: 2in; 
                        right: 1in; 
                        font-size: ${this.settings.fontSize}px; 
                        font-weight: 500; 
                        line-height: 1.5;
                    }
                    @media print {
                        body { background: white; padding: 0; }
                        .envelope { 
                            box-shadow: none; 
                            border: none; 
                            margin: 0; 
                            width: 100%; 
                            height: 100vh; 
                        }
                    }
                </style>
            </head>
            <body>
                ${envelopes}
                <script>
                    window.addEventListener('load', function() {
                        setTimeout(() => window.print(), 500);
                    });
                </script>
            </body>
            </html>
        `);
        
        previewWindow.document.close();
    }

    printAllEnvelopes() {
        if (this.addresses.length === 0) {
            this.showError('No addresses to print');
            return;
        }

        this.showProgress('Preparing envelopes for printing...', 0);

        setTimeout(() => {
            try {
                const printWindow = window.open('', '_blank');
                const envelopes = this.generateAllEnvelopes();
                
                printWindow.document.write(this.generatePrintHTML(envelopes));
                printWindow.document.close();
                
                printWindow.addEventListener('load', () => {
                    setTimeout(() => {
                        printWindow.print();
                        this.hideProgress();
                    }, 500);
                });

            } catch (error) {
                this.hideProgress();
                this.showError(`Error preparing envelopes: ${error.message}`);
            }
        }, 100);
    }

    printCurrentEnvelope() {
        if (this.addresses.length === 0) return;

        const currentAddress = this.addresses[this.currentAddressIndex];
        const envelope = this.generateSingleEnvelope(currentAddress);
        
        const printWindow = window.open('', '_blank');
        printWindow.document.write(this.generatePrintHTML(envelope));
        printWindow.document.close();
        
        printWindow.addEventListener('load', () => {
            setTimeout(() => printWindow.print(), 500);
        });
    }

    generateAllEnvelopes() {
        const returnAddressText = document.getElementById('returnAddress').innerHTML;
        
        return this.addresses.map((address, index) => {
            this.updateProgress((index / this.addresses.length) * 100);
            return this.generateSingleEnvelope(address, returnAddressText);
        }).join('');
    }

    generateSingleEnvelope(address, returnAddressText = null) {
        if (!returnAddressText) {
            returnAddressText = document.getElementById('returnAddress').innerHTML;
        }
        
        const recipientAddressText = this.formatAddress(address);
        
        return `
            <div class="envelope">
                <div class="return-address">${returnAddressText}</div>
                <div class="recipient-address">${recipientAddressText}</div>
            </div>
        `;
    }

    generatePrintHTML(envelopesHTML) {
        return `
            <!DOCTYPE html>
            <html>
            <head>
                <title>Envelope Print</title>
                <style>
                    body { 
                        font-family: ${this.settings.fontFamily}; 
                        margin: 0; 
                        padding: 0; 
                        background: white; 
                    }
                    .envelope { 
                        width: 100%; 
                        height: 100vh; 
                        padding: 0.5in; 
                        position: relative; 
                        page-break-after: always; 
                    }
                    .envelope:last-child { page-break-after: avoid; }
                    .return-address { 
                        position: absolute; 
                        top: 0.5in; 
                        left: 0.5in; 
                        font-size: ${this.settings.fontSize}pt; 
                        line-height: 1.4;
                    }
                    .recipient-address { 
                        position: absolute; 
                        bottom: 2in; 
                        right: 1in; 
                        font-size: ${this.settings.fontSize}pt; 
                        font-weight: 500; 
                        line-height: 1.5;
                    }
                    @page { 
                        margin: 0; 
                        size: letter; 
                    }
                </style>
            </head>
            <body>
                ${envelopesHTML}
                <script>
                    window.addEventListener('load', function() {
                        setTimeout(() => window.print(), 500);
                    });
                </script>
            </body>
            </html>
        `;
    }

    showProgress(text, percentage) {
        const progressIndicator = document.getElementById('progressIndicator');
        const progressText = document.getElementById('progressText');
        const progressFill = document.getElementById('progressFill');
        
        progressText.textContent = text;
        progressFill.style.width = `${percentage}%`;
        progressIndicator.style.display = 'block';
    }

    updateProgress(percentage) {
        const progressFill = document.getElementById('progressFill');
        progressFill.style.width = `${percentage}%`;
    }

    hideProgress() {
        document.getElementById('progressIndicator').style.display = 'none';
    }

    showError(message) {
        const errorMessage = document.getElementById('errorMessage');
        const errorText = document.getElementById('errorText');
        
        errorText.textContent = message;
        errorMessage.style.display = 'flex';
        
        setTimeout(() => this.hideError(), 5000);
    }

    hideError() {
        document.getElementById('errorMessage').style.display = 'none';
    }
}

// Initialize the application
document.addEventListener('DOMContentLoaded', () => {
    new EnvelopePrinter();
});