// Travel Voucher Generator JavaScript

class TravelVoucherGenerator {
    constructor() {
        this.daysMapping = {
            "пн": "Monday",
            "вт": "Tuesday",
            "ср": "Wednesday",
            "чт": "Thursday",
            "пт": "Friday",
            "сб": "Saturday",
            "вс": "Sunday"
        };

        this.commentTemplates = {
            "transfer": "{time} – Meet your driver in the {start_poi}",
            "city tour": "{time} – Meet your guide in the lobby of {start_poi}",
            "priv-tour": "{time} – Meet your driver-guide in the lobby of {start_poi}",
            "safari": "{time} – Get ready for the safari at: {start_poi}",
            "group/tix": "{time} – Please arrive to the start point ({start_poi}) in advance."
        };

        this.hotelKeywords = ["hotel", "lodge", "resort"];
        this.bookingKeywords = ["BB", "HB", "DBL", "Superior", "Deluxe"];
        this.emergencyNumber = "+358403581870";
        this.companyName = "Scandinavia";

        this.currentExcursionIndex = 0;
        this.currentExcursions = [];
        this.pendingClassifications = [];
        this.voucherData = null;
        
        this.initializeEventListeners();
    }

    initializeEventListeners() {
        const uploadArea = document.getElementById('upload-area');
        const fileInput = document.getElementById('file-input');
        const fileSelectBtn = document.getElementById('file-select-btn');
        const downloadBtn = document.getElementById('download-btn');
        const printBtn = document.getElementById('print-btn');
        const newFileBtn = document.getElementById('new-file-btn');
        const retryBtn = document.getElementById('retry-btn');
        const demoBtn = document.getElementById('demo-btn');

        // File upload events
        uploadArea.addEventListener('click', () => fileInput.click());
        fileSelectBtn.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e.target.files[0]));

        // Drag and drop events
        uploadArea.addEventListener('dragover', (e) => this.handleDragOver(e));
        uploadArea.addEventListener('dragleave', (e) => this.handleDragLeave(e));
        uploadArea.addEventListener('drop', (e) => this.handleDrop(e));

        // Action buttons
        downloadBtn.addEventListener('click', () => this.downloadWordDocument());
        printBtn.addEventListener('click', () => this.printVoucher());
        newFileBtn.addEventListener('click', () => this.resetApp());
        retryBtn.addEventListener('click', () => this.resetApp());
        demoBtn.addEventListener('click', () => this.showDemoPreview());

        // Modal events
        this.initializeModalEvents();
    }

    initializeModalEvents() {
        const modal = document.getElementById('excursion-modal');
        const options = document.querySelectorAll('.excursion-option');
        
        options.forEach(option => {
            option.addEventListener('click', (e) => {
                const type = e.target.getAttribute('data-type');
                this.handleExcursionClassification(type);
            });
        });
    }

    showDemoPreview() {
        // Create sample voucher data for testing
        this.voucherData = {
            companyName: "Scandinavia",
            tripRef: "SCN-2025-001",
            participants: 4,
            tripDates: "15 Mar 2025 – 22 Mar 2025",
            destinations: "Helsinki – Stockholm – Copenhagen",
            emergencyNumber: "+358403581870",
            days: [
                {
                    dayNumber: 1,
                    dayName: "Monday 15 Mar",
                    pageBreak: false,
                    activities: [
                        {
                            time: "9:00 AM",
                            name: "Airport Transfer to Hotel",
                            comment: "9:00 AM – Meet your driver in the Arrival Hall"
                        },
                        {
                            time: "2:00 PM",
                            name: "Helsinki City Tour",
                            comment: "2:00 PM – Meet your guide in the lobby of Hotel Kämp"
                        }
                    ],
                    hotelInfo: {
                        name: "Hotel Kämp",
                        reference: "SCN-2025-001",
                        booking: "DBL Superior"
                    }
                },
                {
                    dayNumber: 2,
                    dayName: "Tuesday 16 Mar",
                    pageBreak: false,
                    activities: [
                        {
                            time: "10:00 AM",
                            name: "Ferry to Stockholm",
                            comment: "10:00 AM – Please arrive to the start point (West Terminal) in advance."
                        }
                    ],
                    hotelInfo: {
                        name: "Grand Hotel Stockholm",
                        reference: "SCN-2025-001",
                        booking: "DBL Deluxe"
                    }
                },
                {
                    dayNumber: 3,
                    dayName: "Wednesday 17 Mar",
                    pageBreak: false,
                    activities: [
                        {
                            time: "11:00 AM",
                            name: "Stockholm Private City Tour",
                            comment: "11:00 AM – Meet your driver-guide in the lobby of Grand Hotel Stockholm"
                        }
                    ],
                    hotelInfo: null
                }
            ]
        };

        const voucherHTML = this.generateVoucherHTML(this.voucherData);
        document.getElementById('voucher-preview').innerHTML = voucherHTML;
        this.showSection('preview-section');
    }

    handleDragOver(e) {
        e.preventDefault();
        document.getElementById('upload-area').classList.add('drag-over');
    }

    handleDragLeave(e) {
        e.preventDefault();
        document.getElementById('upload-area').classList.remove('drag-over');
    }

    handleDrop(e) {
        e.preventDefault();
        document.getElementById('upload-area').classList.remove('drag-over');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            this.handleFileSelect(files[0]);
        }
    }

    handleFileSelect(file) {
        if (!file) return;

        // Validate file type
        if (!file.name.match(/\.(xlsx|xls)$/i)) {
            this.showError('Please select a valid Excel file (.xlsx or .xls)');
            return;
        }

        this.showFileStatus(file.name, 'success');
        setTimeout(() => {
            this.processFile(file);
        }, 500);
    }

    showFileStatus(fileName, type) {
        const fileStatus = document.getElementById('file-status');
        fileStatus.className = `file-status ${type}`;
        fileStatus.textContent = type === 'success' ? 
            `✓ File selected: ${fileName}` : 
            `✗ Error: ${fileName}`;
        fileStatus.classList.remove('hidden');
    }

    async processFile(file) {
        this.showSection('processing-section');

        try {
            const data = await this.readExcelFile(file);
            this.voucherData = await this.processExcelData(data);
            const voucherHTML = this.generateVoucherHTML(this.voucherData);
            
            document.getElementById('voucher-preview').innerHTML = voucherHTML;
            this.showSection('preview-section');
        } catch (error) {
            console.error('Error processing file:', error);
            this.showError(error.message || 'Failed to process the Excel file. Please check the file format and try again.');
        }
    }

    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                    resolve(jsonData);
                } catch (error) {
                    reject(new Error('Failed to read Excel file'));
                }
            };
            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsArrayBuffer(file);
        });
    }

    async processExcelData(rawData) {
        // Convert array data to objects
        if (rawData.length < 2) {
            throw new Error('Excel file appears to be empty or invalid');
        }

        const headers = rawData[0];
        const rows = rawData.slice(1).map(row => {
            const obj = {};
            headers.forEach((header, index) => {
                obj[header] = row[index] || '';
            });
            return obj;
        });

        // Extract basic info
        const tripRef = this.extractValue(rows, 'Trip / Ref') || 'N/A';
        const participants = parseInt(this.extractValue(rows, 'PAX')) || 0;
        const tripDates = this.extractTripDates(rows);
        const destinations = this.extractDestinations(rows);

        // Process activities by day
        this.currentExcursions = [];
        this.pendingClassifications = [];
        
        const days = await this.processActivitiesByDay(rows, tripRef);

        return {
            tripRef,
            participants,
            tripDates,
            destinations,
            days,
            emergencyNumber: this.emergencyNumber,
            companyName: this.companyName
        };
    }

    extractValue(rows, columnName) {
        for (const row of rows) {
            if (row[columnName] && row[columnName].toString().trim()) {
                return row[columnName].toString().trim();
            }
        }
        return null;
    }

    extractTripDates(rows) {
        const dateRows = rows.filter(row => {
            const place = (row['Place'] || '').toString();
            return /\d{2}\.\d{2}\.\d{2}/.test(place);
        });

        if (dateRows.length === 0) return 'N/A';

        const firstDateRaw = dateRows[0]['Place'].toString().split(',').pop().trim();
        const lastDateRaw = dateRows[dateRows.length - 1]['Place'].toString().split(',').pop().trim();

        const parseDate = (dateStr) => {
            try {
                const cleanDate = dateStr.replace('г.', '').trim();
                const [day, month, year] = cleanDate.split('.');
                const fullYear = 2000 + parseInt(year);
                const date = new Date(fullYear, parseInt(month) - 1, parseInt(day));
                
                const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
                return `${day} ${months[date.getMonth()]} ${fullYear}`;
            } catch (error) {
                return dateStr;
            }
        };

        const firstDate = parseDate(firstDateRaw);
        const lastDate = parseDate(lastDateRaw);

        return `${firstDate} – ${lastDate}`;
    }

    extractDestinations(rows) {
        const places = rows
            .filter(row => {
                const place = (row['Place'] || '').toString();
                return place && !/\d{2}\.\d{2}\.\d{2}/.test(place);
            })
            .map(row => (row['Place'] || '').toString().trim())
            .filter((place, index, array) => place && array.indexOf(place) === index);

        return places.join(' – ');
    }

    async processActivitiesByDay(rows, tripRef) {
        const days = [];
        let currentDay = null;
        let dayCount = 0;
        let hotelInfo = null;

        for (const row of rows) {
            const place = this.normalizeDayName((row['Place'] || '').toString().trim());
            const time = this.convertTimeRange((row['Time'] || '').toString());
            const excursion = (row['Excursion'] || '').toString().trim();
            const startPoi = (row['Start POI'] || '').toString().trim();

            // Check if this is a day header
            if (this.isDayHeader(place)) {
                // Save hotel info for previous day
                if (currentDay) {
                    currentDay.hotelInfo = hotelInfo;
                }

                dayCount++;
                currentDay = {
                    dayNumber: dayCount,
                    dayName: place,
                    activities: [],
                    hotelInfo: null,
                    pageBreak: dayCount % 4 === 1 && dayCount > 1
                };
                days.push(currentDay);
                hotelInfo = null;
                continue;
            }

            // Process activity
            if (excursion && currentDay) {
                const excursionType = await this.classifyExcursion(excursion);
                const comment = this.generateComment(excursionType, time, startPoi);
                
                currentDay.activities.push({
                    time,
                    name: excursion,
                    comment
                });

                // Check for hotel info
                const extractedHotel = this.extractHotelName(excursion) || this.extractHotelName(startPoi);
                if (extractedHotel) {
                    const bookingInfo = this.extractBookingInfo(excursion);
                    hotelInfo = {
                        name: extractedHotel,
                        reference: tripRef,
                        booking: bookingInfo
                    };
                }
            }
        }

        // Add hotel info to last day
        if (currentDay) {
            currentDay.hotelInfo = hotelInfo;
        }

        return days;
    }

    normalizeDayName(text) {
        let normalized = text;
        for (const [ru, en] of Object.entries(this.daysMapping)) {
            if (text.startsWith(ru)) {
                normalized = text.replace(ru, en);
                break;
            }
        }
        // Remove Russian characters
        normalized = normalized.replace(/[А-Яа-яЁё]+/g, '').trim();
        return normalized;
    }

    isDayHeader(place) {
        return Object.values(this.daysMapping).some(day => place.includes(day));
    }

    convertTimeRange(timeStr) {
        if (!timeStr || timeStr === 'undefined') return '';
        
        const parts = timeStr.replace(/\n/g, '-').replace(/–/g, '-').split('-');
        const converted = parts.map(t => {
            t = t.trim();
            if (/^\d{1,2}:\d{2}$/.test(t)) {
                try {
                    const [hours, minutes] = t.split(':');
                    const hour24 = parseInt(hours);
                    const hour12 = hour24 === 0 ? 12 : (hour24 > 12 ? hour24 - 12 : hour24);
                    const ampm = hour24 < 12 ? 'AM' : 'PM';
                    return `${hour12}:${minutes} ${ampm}`;
                } catch (error) {
                    return t;
                }
            }
            return t;
        });
        
        return converted.join(' – ');
    }

    async classifyExcursion(excursion) {
        const e = excursion.toLowerCase();
        
        if (e.includes('transfer')) return 'transfer';
        if ((e.includes('city') && e.includes('tour')) || e.includes('city tour')) return 'city tour';
        if (e.includes('private') && e.includes('tour')) return 'priv-tour';
        if (e.includes('safari')) return 'safari';
        if (e.includes('self') || e.includes('tickets')) return 'group/tix';
        
        // Need manual classification
        return await this.requestManualClassification(excursion);
    }

    requestManualClassification(excursion) {
        return new Promise((resolve) => {
            const modal = document.getElementById('excursion-modal');
            const excursionText = document.getElementById('excursion-text');
            
            excursionText.textContent = excursion;
            modal.classList.remove('hidden');
            
            this.pendingClassifications.push({
                excursion,
                resolve
            });
        });
    }

    handleExcursionClassification(type) {
        const modal = document.getElementById('excursion-modal');
        modal.classList.add('hidden');
        
        if (this.pendingClassifications.length > 0) {
            const current = this.pendingClassifications.shift();
            current.resolve(type);
        }
    }

    generateComment(excursionType, time, startPoi) {
        const template = this.commentTemplates[excursionType];
        if (!template) return '';
        
        return template
            .replace('{time}', time || '')
            .replace('{start_poi}', startPoi || '');
    }

    extractHotelName(text) {
        if (!text) return '';
        
        const tokens = text.split(/\s+/);
        for (let i = 0; i < tokens.length; i++) {
            const token = tokens[i].toLowerCase();
            if (this.hotelKeywords.includes(token)) {
                const parts = [];
                if (i >= 2 && /^[A-Z]/.test(tokens[i - 2])) parts.push(tokens[i - 2]);
                if (i >= 1 && /^[A-Z]/.test(tokens[i - 1])) parts.push(tokens[i - 1]);
                parts.push(tokens[i].charAt(0).toUpperCase() + tokens[i].slice(1));
                if (i + 1 < tokens.length && /^[A-Z]/.test(tokens[i + 1])) {
                    parts.push(tokens[i + 1]);
                }
                return parts.join(' ');
            }
        }
        return '';
    }

    extractBookingInfo(excursion) {
        const e = excursion.toLowerCase();
        for (const keyword of this.bookingKeywords) {
            if (e.includes(keyword.toLowerCase())) {
                return keyword;
            }
        }
        return '';
    }

    async downloadWordDocument() {
        if (!this.voucherData) {
            this.showError('No voucher data available. Please process a file first or use the demo.');
            return;
        }

        if (!window.docx || !window.saveAs) {
            this.showError('Required libraries not loaded. Please refresh the page and try again.');
            return;
        }

        const downloadBtn = document.getElementById('download-btn');
        const originalText = downloadBtn.textContent;
        downloadBtn.classList.add('loading');
        downloadBtn.disabled = true;
        downloadBtn.textContent = 'Generating...';

        try {
            const doc = await this.generateWordDocument(this.voucherData);
            const fileName = `Travel_Voucher_${this.voucherData.tripRef.replace(/[^a-zA-Z0-9]/g, '_')}.docx`;
            
            docx.Packer.toBlob(doc).then(blob => {
                saveAs(blob, fileName);
            }).catch(error => {
                console.error('Error creating blob:', error);
                this.showError('Failed to generate Word document. Please try again.');
            }).finally(() => {
                downloadBtn.classList.remove('loading');
                downloadBtn.disabled = false;
                downloadBtn.textContent = originalText;
            });
        } catch (error) {
            console.error('Error generating Word document:', error);
            this.showError('Failed to generate Word document. Please try again.');
            downloadBtn.classList.remove('loading');
            downloadBtn.disabled = false;
            downloadBtn.textContent = originalText;
        }
    }

    async generateWordDocument(data) {
        const { Document, Paragraph, Table, TableRow, TableCell, TextRun, AlignmentType, BorderStyle, WidthType, HeadingLevel } = docx;

        // Create document sections
        const children = [];

        // Add header
        children.push(
            new Paragraph({
                children: [new TextRun({ text: data.companyName, bold: true, size: 36 })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 }
            })
        );

        // Add trip information table
        const infoRows = [
            new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Trip Dates', bold: true })] })] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: data.tripDates })] })] })
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Destinations', bold: true })] })] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: data.destinations })] })] })
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Trip Reference', bold: true })] })] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: data.tripRef })] })] })
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Participants', bold: true })] })] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: data.participants.toString() })] })] })
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Emergency Number', bold: true })] })] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: data.emergencyNumber })] })] })
                ]
            })
        ];

        children.push(
            new Table({
                rows: infoRows,
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: {
                    top: { style: BorderStyle.SINGLE, size: 1 },
                    bottom: { style: BorderStyle.SINGLE, size: 1 },
                    left: { style: BorderStyle.SINGLE, size: 1 },
                    right: { style: BorderStyle.SINGLE, size: 1 },
                    insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
                    insideVertical: { style: BorderStyle.SINGLE, size: 1 }
                }
            })
        );

        children.push(new Paragraph({ children: [], spacing: { after: 400 } }));

        // Add days
        data.days.forEach((day, index) => {
            // Add page break for every 4 days (except the first)
            if (day.pageBreak) {
                children.push(new Paragraph({ children: [], pageBreakBefore: true }));
            }

            // Add day header
            children.push(
                new Paragraph({
                    children: [new TextRun({ text: `Day ${day.dayNumber} — ${day.dayName}`, bold: true, size: 24 })],
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 200, after: 200 }
                })
            );

            // Add activities table if there are activities
            if (day.activities.length > 0) {
                const activityRows = [
                    new TableRow({
                        children: [
                            new TableCell({ 
                                children: [new Paragraph({ children: [new TextRun({ text: 'Start Time', bold: true })] })] 
                            }),
                            new TableCell({ 
                                children: [new Paragraph({ children: [new TextRun({ text: 'Name', bold: true })] })] 
                            }),
                            new TableCell({ 
                                children: [new Paragraph({ children: [new TextRun({ text: 'Comment', bold: true })] })] 
                            })
                        ]
                    }),
                    ...day.activities.map(activity => 
                        new TableRow({
                            children: [
                                new TableCell({ 
                                    children: [new Paragraph({ children: [new TextRun({ text: activity.time })] })] 
                                }),
                                new TableCell({ 
                                    children: [new Paragraph({ children: [new TextRun({ text: activity.name })] })] 
                                }),
                                new TableCell({ 
                                    children: [new Paragraph({ children: [new TextRun({ text: activity.comment })] })] 
                                })
                            ]
                        })
                    )
                ];

                children.push(
                    new Table({
                        rows: activityRows,
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: {
                            top: { style: BorderStyle.SINGLE, size: 1 },
                            bottom: { style: BorderStyle.SINGLE, size: 1 },
                            left: { style: BorderStyle.SINGLE, size: 1 },
                            right: { style: BorderStyle.SINGLE, size: 1 },
                            insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
                            insideVertical: { style: BorderStyle.SINGLE, size: 1 }
                        }
                    })
                );
            }

            // Add hotel information table if available
            if (day.hotelInfo) {
                children.push(new Paragraph({ children: [], spacing: { before: 200 } }));

                const hotelRows = [
                    new TableRow({
                        children: [
                            new TableCell({ 
                                children: [new Paragraph({ children: [new TextRun({ text: 'Hotel Name', bold: true })] })] 
                            }),
                            new TableCell({ 
                                children: [new Paragraph({ children: [new TextRun({ text: 'Reference', bold: true })] })] 
                            }),
                            new TableCell({ 
                                children: [new Paragraph({ children: [new TextRun({ text: 'Room & Booking type', bold: true })] })] 
                            })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ 
                                children: [new Paragraph({ children: [new TextRun({ text: day.hotelInfo.name })] })] 
                            }),
                            new TableCell({ 
                                children: [new Paragraph({ children: [new TextRun({ text: day.hotelInfo.reference })] })] 
                            }),
                            new TableCell({ 
                                children: [new Paragraph({ children: [new TextRun({ text: day.hotelInfo.booking })] })] 
                            })
                        ]
                    })
                ];

                children.push(
                    new Table({
                        rows: hotelRows,
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: {
                            top: { style: BorderStyle.SINGLE, size: 1 },
                            bottom: { style: BorderStyle.SINGLE, size: 1 },
                            left: { style: BorderStyle.SINGLE, size: 1 },
                            right: { style: BorderStyle.SINGLE, size: 1 },
                            insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
                            insideVertical: { style: BorderStyle.SINGLE, size: 1 }
                        }
                    })
                );
            }

            children.push(new Paragraph({ children: [], spacing: { after: 300 } }));
        });

        return new Document({
            sections: [{
                children: children,
                properties: {}
            }]
        });
    }

    generateVoucherHTML(data) {
        let html = `
            <div class="voucher-header">
                <h1>${data.companyName}</h1>
                <div class="voucher-info">
                    <div class="info-item">
                        <strong>Trip Dates</strong>
                        <span>${data.tripDates}</span>
                    </div>
                    <div class="info-item">
                        <strong>Destinations</strong>
                        <span>${data.destinations}</span>
                    </div>
                    <div class="info-item">
                        <strong>Trip Reference</strong>
                        <span>${data.tripRef}</span>
                    </div>
                    <div class="info-item">
                        <strong>Participants</strong>
                        <span>${data.participants}</span>
                    </div>
                    <div class="info-item">
                        <strong>Emergency Number</strong>
                        <span>${data.emergencyNumber}</span>
                    </div>
                </div>
            </div>
        `;

        data.days.forEach(day => {
            html += `
                <div class="day-section ${day.pageBreak ? 'page-break' : ''}">
                    <div class="day-header">
                        <h3>Day ${day.dayNumber} — ${day.dayName}</h3>
                    </div>
            `;

            if (day.activities.length > 0) {
                html += `
                    <table class="activity-table">
                        <thead>
                            <tr>
                                <th>Start Time</th>
                                <th>Name</th>
                                <th>Comment</th>
                            </tr>
                        </thead>
                        <tbody>
                `;

                day.activities.forEach(activity => {
                    html += `
                        <tr>
                            <td>${activity.time}</td>
                            <td>${activity.name}</td>
                            <td>${activity.comment}</td>
                        </tr>
                    `;
                });

                html += '</tbody></table>';
            }

            if (day.hotelInfo) {
                html += `
                    <table class="hotel-table">
                        <thead>
                            <tr>
                                <th>Hotel Name</th>
                                <th>Reference</th>
                                <th>Room & Booking type</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>${day.hotelInfo.name}</td>
                                <td>${day.hotelInfo.reference}</td>
                                <td>${day.hotelInfo.booking}</td>
                            </tr>
                        </tbody>
                    </table>
                `;
            }

            html += '</div>';
        });

        return html;
    }

    printVoucher() {
        window.print();
    }

    showSection(sectionId) {
        // Hide all sections
        const sections = ['upload-section', 'processing-section', 'preview-section', 'error-section'];
        sections.forEach(id => {
            document.getElementById(id).classList.add('hidden');
        });

        // Show target section
        document.getElementById(sectionId).classList.remove('hidden');
    }

    showError(message) {
        document.getElementById('error-message').textContent = message;
        this.showSection('error-section');
    }

    resetApp() {
        this.showSection('upload-section');
        document.getElementById('file-status').classList.add('hidden');
        document.getElementById('file-input').value = '';
        document.getElementById('voucher-preview').innerHTML = '';
        this.pendingClassifications = [];
        this.voucherData = null;
    }
}

// Initialize the application when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    new TravelVoucherGenerator();
});