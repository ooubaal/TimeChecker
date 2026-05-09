// State management
const state = {
    selectedMonth: new Date().getMonth(),
    selectedYear: new Date().getFullYear(),
    holidays: [], // List of dates (1-31)
    employees: [], // { name: '', scans: { 'YYYY-MM-DD': { in: '', out: '' } } }
    summary: [] // { name: '', late: 0, early: 0, leave: 0, absence: 0, ot: 0 }
};

// DOM Elements
const monthSelect = document.getElementById('monthSelect');
const yearSelect = document.getElementById('yearSelect');
const holidayDate = document.getElementById('holidayDate');
const addHolidayBtn = document.getElementById('addHolidayBtn');
const holidayList = document.getElementById('holidayList');
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const tableBody = document.getElementById('tableBody');
const totalStaff = document.getElementById('totalStaff');
const totalLate = document.getElementById('totalLate');
const totalOT = document.getElementById('totalOT');

// Initialization
function init() {
    monthSelect.value = state.selectedMonth;
    yearSelect.value = state.selectedYear;

    // Listeners
    monthSelect.addEventListener('change', (e) => {
        state.selectedMonth = parseInt(e.target.value);
        calculateAll();
    });

    yearSelect.addEventListener('change', (e) => {
        state.selectedYear = parseInt(e.target.value);
        calculateAll();
    });

    addHolidayBtn.addEventListener('click', addHoliday);

    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.style.borderColor = 'var(--primary)';
    });
    dropZone.addEventListener('dragleave', () => {
        dropZone.style.borderColor = 'var(--border)';
    });
    dropZone.addEventListener('drop', handleDrop);

    fileInput.addEventListener('change', handleFileSelect);

    // Sortable
    new Sortable(tableBody, {
        animation: 150,
        handle: '.drag-handle',
        onEnd: () => {
            // Update internal order if needed
        }
    });
}

// Holiday Management
function addHoliday() {
    const day = parseInt(holidayDate.value);
    if (!day || day < 1 || day > 31) return;
    if (state.holidays.includes(day)) return;

    state.holidays.push(day);
    state.holidays.sort((a, b) => a - b);
    renderHolidays();
    calculateAll();
    holidayDate.value = '';
}

function renderHolidays() {
    holidayList.innerHTML = '';
    state.holidays.forEach(day => {
        const item = document.createElement('div');
        item.className = 'holiday-item';
        item.innerHTML = `
            <span>วันที่ ${day}</span>
            <button onclick="removeHoliday(${day})">&times;</button>
        `;
        holidayList.appendChild(item);
    });
}

window.removeHoliday = (day) => {
    state.holidays = state.holidays.filter(h => h !== day);
    renderHolidays();
    calculateAll();
};

// File Handling
async function handleDrop(e) {
    e.preventDefault();
    dropZone.style.borderColor = 'var(--border)';
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
}

function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) processFile(file);
}

async function processFile(file) {
    const extension = file.name.split('.').pop().toLowerCase();
    
    if (extension === 'xlsx' || extension === 'xls') {
        parseExcel(file);
    } else if (extension === 'pdf') {
        parsePDF(file);
    } else {
        alert('กรุณาอัพโหลดไฟล์ .xlsx, .xls หรือ .pdf');
    }
}

// Parsing Logic
function parseExcel(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(firstSheet);
        
        // Transform JSON to state.employees
        // Expected format: { Name, Date, TimeIn, TimeOut } or similar
        processRawData(json);
    };
    reader.readAsArrayBuffer(file);
}

async function parsePDF(file) {
    // Basic PDF text extraction
    const reader = new FileReader();
    reader.onload = async (e) => {
        const data = new Uint8Array(e.target.result);
        const pdf = await pdfjsLib.getDocument(data).promise;
        let fullText = "";
        
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            fullText += content.items.map(item => item.str).join(' ') + "\n";
        }
        
        // Simple heuristic for parsing text (requires specific format)
        // This is a placeholder for actual pattern matching
        console.log("PDF Text Extracted:", fullText);
        alert("ระบบกำลังพัฒนาการอ่าน PDF ที่มีความซับซ้อน แนะนำให้ใช้ Excel เพื่อความแม่นยำ");
    };
    reader.readAsArrayBuffer(file);
}

// Data Processing
function processRawData(data) {
    console.log("Raw Data Received:", data);
    if (!data || data.length === 0) {
        alert("ไม่พบข้อมูลในไฟล์ หรือรูปแบบไฟล์ไม่ถูกต้อง");
        return;
    }

    const employeeMap = {};
    let processedCount = 0;

    data.forEach((row, index) => {
        // More flexible header matching
        const findValue = (keys) => {
            for (let key of keys) {
                const foundKey = Object.keys(row).find(k => k.trim().toLowerCase() === key.toLowerCase());
                if (foundKey) return row[foundKey];
            }
            return null;
        };

        const name = findValue(['Name', 'ชื่อ', 'พนักงาน', 'staff', 'employee']);
        const dateStr = findValue(['Date', 'วันที่', 'day']);
        const timeIn = findValue(['TimeIn', 'เวลาเข้า', 'เข้า', 'checkin', 'in']);
        const timeOut = findValue(['TimeOut', 'เวลาออก', 'ออก', 'checkout', 'out']);

        if (!name || !dateStr) {
            console.warn(`Row ${index} missing Name or Date:`, row);
            return;
        }

        if (!employeeMap[name]) {
            employeeMap[name] = { name, scans: {} };
        }

        const date = normalizeDate(dateStr);
        if (date) {
            employeeMap[name].scans[date] = { in: String(timeIn || ''), out: String(timeOut || '') };
            processedCount++;
        }
    });

    state.employees = Object.values(employeeMap);
    console.log("Processed Employees:", state.employees);

    if (state.employees.length === 0) {
        alert("ไม่สามารถประมวลผลข้อมูลได้ กรุณาตรวจสอบหัวตาราง (Name, Date, TimeIn, TimeOut)");
    } else {
        calculateAll();
        alert(`นำเข้าข้อมูลพนักงาน ${state.employees.length} คน เรียบร้อยแล้ว`);
    }
}

function normalizeDate(str) {
    if (!str) return null;
    
    // If it's already a number (Excel date serial)
    if (typeof str === 'number') {
        const date = new Date((str - 25569) * 86400 * 1000);
        return date.toISOString().split('T')[0];
    }

    const s = String(str).trim();
    
    // Try DD/MM/YYYY or DD-MM-YYYY
    const parts = s.split(/[/-]/);
    if (parts.length === 3) {
        let day, month, year;
        // Assume DD/MM/YYYY or YYYY/MM/DD
        if (parts[0].length === 4) { // YYYY
            [year, month, day] = parts;
        } else {
            [day, month, year] = parts;
        }
        // Handle Buddhist Era
        if (parseInt(year) > 2500) year = parseInt(year) - 543;
        
        const d = new Date(year, month - 1, day);
        if (!isNaN(d.getTime())) {
            return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
        }
    }

    const d = new Date(s);
    if (!isNaN(d.getTime())) {
        return d.toISOString().split('T')[0];
    }
    
    return null;
}

// Attendance Calculation Engine
function calculateAll() {
    if (state.employees.length === 0) return;

    const daysInMonth = new Date(state.selectedYear, state.selectedMonth + 1, 0).getDate();
    const results = [];

    state.employees.forEach(emp => {
        const summary = {
            name: emp.name,
            lateMinutes: 0,
            earlyExitMinutes: 0,
            morningLeave: 0,
            afternoonLeave: 0,
            absence: 0,
            otHours: 0,
            isHolidayOT: false
        };

        for (let day = 1; day <= daysInMonth; day++) {
            const dateStr = `${state.selectedYear}-${String(state.selectedMonth + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
            const scan = emp.scans[dateStr];
            
            const isWeekend = new Date(state.selectedYear, state.selectedMonth, day).getDay() % 6 === 0;
            const isHoliday = state.holidays.includes(day);
            const isWorkDay = !isWeekend && !isHoliday;

            if (isWorkDay) {
                if (!scan || (!scan.in && !scan.out)) {
                    summary.absence++;
                } else {
                    // Rule 4: Late check (08:30)
                    if (scan.in) {
                        const inMinutes = timeToMinutes(scan.in);
                        if (inMinutes > timeToMinutes('08:30')) {
                            summary.lateMinutes += (inMinutes - timeToMinutes('08:30'));
                        }
                    }

                    // Rule 5: Early exit (16:30)
                    if (scan.out) {
                        const outMinutes = timeToMinutes(scan.out);
                        if (outMinutes < timeToMinutes('16:30')) {
                            summary.earlyExitMinutes += (timeToMinutes('16:30') - outMinutes);
                        }
                    }

                    // Rule 6: Morning Leave
                    // Work day, in before 13:00, out after 16:30
                    if (scan.in && scan.out) {
                        const inMin = timeToMinutes(scan.in);
                        const outMin = timeToMinutes(scan.out);
                        if (inMin > timeToMinutes('08:30') && inMin <= timeToMinutes('13:00') && outMin >= timeToMinutes('16:30')) {
                            summary.morningLeave++;
                        }
                        // Rule 7: Afternoon Leave
                        // Work day, in before 08:30, out after 12:00
                        if (inMin <= timeToMinutes('08:30') && outMin >= timeToMinutes('12:00') && outMin < timeToMinutes('16:30')) {
                            summary.afternoonLeave++;
                        }
                    }

                    // Rule 2: OT Work Day (8:30 - 19:30)
                    if (scan.out && timeToMinutes(scan.out) >= timeToMinutes('19:30')) {
                        summary.otHours += 3; // 16:30 to 19:30 is 3 hours
                    }
                }
            } else if (isHoliday || isWeekend) {
                // Rule 3 & 8: Holiday OT
                if (scan && scan.in && scan.out) {
                    const inMin = timeToMinutes(scan.in);
                    const outMin = timeToMinutes(scan.out);

                    // Rule 8: Compensation & Penalty
                    if (inMin > timeToMinutes('09:30')) {
                        // Cannot do OT
                    } else if (inMin > timeToMinutes('08:45')) {
                        // Fine logic would go here (not specified where to display)
                        // Still count as OT if stays late
                        if (outMin >= timeToMinutes('16:30') + (inMin - timeToMinutes('08:30'))) {
                            summary.otHours += 8;
                        }
                    } else {
                        // Standard Holiday OT (8:30 - 16:30)
                        const lateIn = Math.max(0, inMin - timeToMinutes('08:30'));
                        if (outMin >= timeToMinutes('16:30') + lateIn) {
                            summary.otHours += 8;
                        }
                    }
                }
            }
        }
        results.push(summary);
    });

    state.summary = results;
    renderTable();
}

function timeToMinutes(timeStr) {
    if (!timeStr || timeStr === 'undefined' || timeStr === 'null') return 0;
    
    // Clean string and handle both . and :
    const cleanStr = String(timeStr).trim().replace(/[^\d:.]/g, '');
    const parts = cleanStr.split(/[:.]/).map(Number);
    
    if (parts.length >= 2) {
        const [hrs, mins] = parts;
        return (hrs * 60) + (mins || 0);
    }
    
    // Fallback if only hours or something weird
    if (parts.length === 1 && !isNaN(parts[0])) {
        return parts[0] * 60;
    }
    
    return 0;
}

function renderTable() {
    tableBody.innerHTML = '';
    let totalLateVal = 0;
    let totalOTVal = 0;

    state.summary.forEach((row, index) => {
        totalLateVal += row.lateMinutes;
        totalOTVal += row.otHours;

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><span class="drag-handle">☰</span> ${index + 1}</td>
            <td><strong>${row.name}</strong></td>
            <td><span class="badge ${row.lateMinutes > 0 ? 'badge-late' : ''}">${row.lateMinutes}</span></td>
            <td><span class="badge ${row.earlyExitMinutes > 0 ? 'badge-early' : ''}">${row.earlyExitMinutes}</span></td>
            <td>${row.morningLeave}/${row.afternoonLeave}</td>
            <td>${row.absence}</td>
            <td><span class="badge badge-ot">${row.otHours}</span></td>
        `;
        tableBody.appendChild(tr);
    });

    totalStaff.innerText = state.summary.length;
    totalLate.innerText = totalLateVal;
    totalOT.innerText = totalOTVal;
}

// Run init
init();
