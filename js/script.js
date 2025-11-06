// --- APLIKÁCIA ROZPIS POHOTOVOSTI - KANBAN BOARD ---

// Konfigurácia načítaná z JSON
let allEmployees = [];
let employeeGroups = [];
let employeeSignatures = {};

// Centrálny stav (State)
let appState = {
    selectedMonth: new Date().getMonth(),
    selectedYear: new Date().getFullYear(),
    dutyAssignments: {}, // { "2025-W44": [employee1, employee2, ...] }
    reporting: {}, // { "2025-W44-employeeId": true }
    selectedDutyForSwap: null,
    employeeToReplace: null, // <<< NOVÁ PREMENNÁ PRE NAHRADENIE (ČERVENÝ RÁMIK)
    autoRotation: true // Automatická rotácia skupín
};

// --- KONŠTANTY ---
const CONFIG_URL = 'files/config.json';
const FONT_URL = 'fonts/Roboto-Regular.ttf';
const DOCX_TEMPLATE_URL = 'files/vykaz_pohotovosti.docx';

// Typy notifikácií
const TOAST_SUCCESS = 'success';
const TOAST_ERROR = 'error';
const TOAST_INFO = 'info';

// DOM IDčká
const ID_MONTH_SELECT = 'monthSelect';
const ID_YEAR_SELECT = 'yearSelect';
const ID_EXPORT_BUTTON = 'exportButton';
const ID_SAVE_BUTTON = 'saveButton';
const ID_CLEAR_BUTTON = 'clearButton';
const ID_CLOSE_MODAL_BUTTON = 'closeModalButton';
const ID_DOWNLOAD_PDF_BUTTON = 'downloadPdfButton';
const ID_KANBAN_BOARD = 'kanbanBoard';
const ID_GROUPS_LIST = 'groupsList';

// *** PRIDANÉ: Globálne názvy mesiacov pre DOCX (zo script2.js) ***
const monthNames = [
    "január", "február", "marec", "apríl", "máj", "jún", "júl", "august", "september", "október", "november", "december"
];

// --- INICIALIZÁCIA ---
document.addEventListener('DOMContentLoaded', async () => {
    try {
        await loadConfiguration();
        initializeUI();
        loadStateFromStorage();
        renderKanbanBoard();
        renderGroupsSidebar();
        
        zobrazOznamenie('Aplikácia úspešne načítaná', TOAST_SUCCESS);
    } catch (error) {
        console.error('Chyba pri inicializácii:', error);
        zobrazOznamenie('Chyba pri načítaní aplikácie', TOAST_ERROR);
    }
});

// --- NAČÍTANIE KONFIGURÁCIE ---
async function loadConfiguration() {
    try {
        const response = await fetch(CONFIG_URL);
        if (!response.ok) throw new Error('Nepodarilo sa načítať konfiguráciu');
        
        const config = await response.json();
        
        // Spracovanie zamestnancov
        employeeGroups = config.zamestnanci || [];
        allEmployees = employeeGroups.flatMap(group => 
            group.moznosti.map(emp => ({
                ...emp,
                skupina: group.skupina
            }))
        );
        
        // Spracovanie podpisov
        employeeSignatures = config.podpisy || {};
        
        console.log('Konfigurácia načítaná:', { employeeGroups, allEmployees });
    } catch (error) {
        console.error('Chyba pri načítaní konfigurácie:', error);
        throw error;
    }
}

// --- INICIALIZÁCIA UI ---
function initializeUI() {
    // Vyplnenie selectu rokov
    const yearSelect = document.getElementById(ID_YEAR_SELECT);
    const currentYear = new Date().getFullYear();
    for (let year = currentYear - 1; year <= currentYear + 2; year++) {
        const option = document.createElement('option');
        option.value = year;
        option.textContent = year;
        yearSelect.appendChild(option);
    }
    yearSelect.value = appState.selectedYear;
    
    // Nastavenie selectu mesiaca
    const monthSelect = document.getElementById(ID_MONTH_SELECT);
    monthSelect.value = appState.selectedMonth;
    
    // Event listenery
    monthSelect.addEventListener('change', (e) => {
        appState.selectedMonth = parseInt(e.target.value);
        renderKanbanBoard();
        saveStateToStorage();
    });
    
    yearSelect.addEventListener('change', (e) => {
        appState.selectedYear = parseInt(e.target.value);
        renderKanbanBoard();
        saveStateToStorage();
    });
    
    document.getElementById(ID_EXPORT_BUTTON).addEventListener('click', exportToPdf);
    document.getElementById(ID_SAVE_BUTTON).addEventListener('click', generateDocx);
    document.getElementById(ID_CLEAR_BUTTON).addEventListener('click', clearSchedule);
    document.getElementById(ID_CLOSE_MODAL_BUTTON).addEventListener('click', closeModal);
    document.getElementById(ID_DOWNLOAD_PDF_BUTTON).addEventListener('click', downloadPdfFromModal);
}

// --- RENDER KANBAN BOARD ---
function renderKanbanBoard() {
    const kanbanBoard = document.getElementById(ID_KANBAN_BOARD);
    kanbanBoard.innerHTML = '';
    
    const weeks = getWeeksInMonth(appState.selectedYear, appState.selectedMonth);
    
    weeks.forEach(week => {
        const weekColumn = createWeekColumn(week);
        kanbanBoard.appendChild(weekColumn);
    });
}

// --- VYTVORIŤ STĹPEC TÝŽĎŇA ---
function createWeekColumn(week) {
    const column = document.createElement('div');
    column.className = 'week-column';
    column.dataset.weekKey = week.key;
    
    // Header
    const header = document.createElement('div');
    header.className = 'week-column-header';
    
    // --- ZAČIATOK ZMENY (DISPLAY DATES) ---
    // Použijeme week.displayStart a week.displayEnd namiesto week.startDate a week.endDate
    header.innerHTML = `
        <span class="week-number">Týždeň ${week.weekNumber}</span>
        <span class="week-dates">${formatDate(week.displayStart)} - ${formatDate(week.displayEnd)}</span>
    `;
    // --- KONIEC ZMENY ---
    
    // Body
    const body = document.createElement('div');
    body.className = 'week-column-body';
    body.dataset.weekKey = week.key;
    
    // Načítanie priradených zamestnancov
    const assignedEmployees = appState.dutyAssignments[week.key] || [];
    
    if (assignedEmployees.length === 0) {
        body.classList.add('empty');
    } else {
        // FIX: Nevytvárame kartu pre každého zamestnanca, ale jednu kartu pre celú skupinu
        const dutyCard = createDutyCard(assignedEmployees, week.key);
        body.appendChild(dutyCard);
    }
    
    // Sortable.js pre drag & drop
    new Sortable(body, {
        group: 'kanban',
        animation: 200,
        ghostClass: 'sortable-ghost',
        dragClass: 'sortable-drag',
        onAdd: function(evt) {
            handleGroupDrop(evt);
        },
        onEnd: function(evt) {
            saveStateToStorage();
        }
    });
    
    column.appendChild(header);
    column.appendChild(body);
    
    return column;
}

// --- VYTVORIŤ KARTU SLUŽBY ---
// FIX: Funkcia teraz prijíma pole zamestnancov (skupinu) a nie jedného zamestnanca
function createDutyCard(employees, weekKey) {
    const card = document.createElement('div');
    card.className = 'duty-card';
    card.dataset.weekKey = weekKey;

    // Získame názov skupiny z prvého zamestnanca
    const groupName = employees[0]?.skupina || 'Priradená služba';

    const header = document.createElement('div');
    header.className = 'duty-card-header';
    header.textContent = groupName; // Hlavička karty je názov skupiny
    
    const employeesDiv = document.createElement('div');
    employeesDiv.className = 'duty-card-employees';
    
    // FIX: Vytvoríme "badge" (odznak) pre každého zamestnanca v skupine
    employees.forEach(employee => {
        const badge = createEmployeeBadge(employee, weekKey);
        employeesDiv.appendChild(badge);
    });
    
    card.appendChild(header);
    card.appendChild(employeesDiv);
    
    return card;
}

// --- VYTVORIŤ BADGE ZAMESTNANCA ---
function createEmployeeBadge(employee, weekKey) {
    const badge = document.createElement('div');
    badge.className = 'employee-badge';
    badge.dataset.employeeId = employee.id;
    badge.dataset.weekKey = weekKey;
    
    const reportingKey = `${weekKey}-${employee.id}`;
    const isReporting = appState.reporting[reportingKey] || false;
    
    if (isReporting) {
        badge.classList.add('reporting');
    }
    
    badge.innerHTML = `
        <span>${employee.meno}</span>
        <span style="font-size: 0.75rem; opacity: 0.7;">${employee.telefon || ''}</span>
    `;
    
    // --- ZAČIATOK ÚPRAV ---
    
    // Vizuálne označenie pre VÝMENU (Swap) - FIALOVÁ
    if (appState.selectedDutyForSwap && 
        appState.selectedDutyForSwap.employeeId === employee.id && 
        appState.selectedDutyForSwap.weekKey === weekKey) {
        badge.classList.add('swap-selected');
    }

    // Vizuálne označenie pre NAHRADENIE (Replacement) - ČERVENÁ
    if (appState.employeeToReplace && 
        appState.employeeToReplace.employeeId === employee.id && 
        appState.employeeToReplace.weekKey === weekKey) {
        badge.classList.add('replace-selected');
    }
    
    // Event listeners
    badge.addEventListener('click', () => {
        // Ak práve vyberáme náhradu (niekto je červený)...
        if (appState.employeeToReplace) {
            // Spustíme logiku nahradenia
            performReplacement(employee.id, weekKey);
        } else {
            // Inak robíme pôvodnú akciu (hlásenie)
            toggleReporting(employee.id, weekKey);
        }
    });

    badge.addEventListener('dblclick', () => {
        // Zrušíme prípadný aktívny swap, aby sa to nebilo
        appState.selectedDutyForSwap = null;
        
        // Označíme tohto zamestnanca na nahradenie
        appState.employeeToReplace = { employeeId: employee.id, weekKey: weekKey };
        
        renderKanbanBoard();
        zobrazOznamenie('DVOJKLIK: Vyberte zamestnanca (klik), ktorý prevezme túto službu.', TOAST_INFO);
    });
    
    badge.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        // Zrušíme prípadné aktívne nahrádzanie
        appState.employeeToReplace = null;
        initiateSwap(employee.id, weekKey);
    });
    
    // --- KONIEC ÚPRAV ---
    
    return badge;
}

// --- RENDER GROUPS SIDEBAR ---
function renderGroupsSidebar() {
    const groupsList = document.getElementById(ID_GROUPS_LIST);
    groupsList.innerHTML = '';
    
    employeeGroups.forEach(group => {
        const groupCard = createGroupCard(group);
        groupsList.appendChild(groupCard);
    });
    
    // Sortable.js pre drag & drop zo sidebaru
    new Sortable(groupsList, {
        group: {
            name: 'kanban',
            pull: 'clone',
            put: false
        },
        animation: 200,
        sort: false
    });
}

// --- VYTVORIŤ KARTU SKUPINY ---
function createGroupCard(group) {
    const card = document.createElement('div');
    card.className = 'group-card';
    card.dataset.groupName = group.skupina;
    
    const header = document.createElement('div');
    header.className = 'group-card-header';
    header.textContent = group.skupina;
    
    const members = document.createElement('div');
    members.className = 'group-card-members';
    members.innerHTML = group.moznosti.map(emp => emp.meno).join('<br>');
    
    card.appendChild(header);
    card.appendChild(members);
    
    return card;
}

// --- SPRACOVANIE DROPU SKUPINY ---
function handleGroupDrop(evt) {
    const weekKey = evt.to.dataset.weekKey;
    const droppedElement = evt.item;
    
    // Zistenie, či ide o skupinu zo sidebaru
    const groupName = droppedElement.dataset.groupName;
    
    if (groupName) {
        // Odstránenie dočasného elementu
        droppedElement.remove();
        
        // Nájdenie skupiny
        const group = employeeGroups.find(g => g.skupina === groupName);
        if (!group) return;
        
        // Priradenie všetkých členov skupiny
        appState.dutyAssignments[weekKey] = group.moznosti.map(emp => ({
            ...emp,
            skupina: group.skupina
        }));
        
        // Automatická rotácia
        if (appState.autoRotation) {
            applyAutoRotation(weekKey, group);
        }
        
        // Re-render
        renderKanbanBoard();
        saveStateToStorage();
        
        zobrazOznamenie(`${group.skupina} priradená do týždňa ${weekKey}`, TOAST_SUCCESS);
    }
}

// --- AUTOMATICKÁ ROTÁCIA ---
function applyAutoRotation(startWeekKey, startGroup) {
    const allWeeks = getWeeksInMonth(appState.selectedYear, appState.selectedMonth);
    const startIndex = allWeeks.findIndex(w => w.key === startWeekKey);
    
    if (startIndex === -1) return;
    
    const groupIndex = employeeGroups.findIndex(g => g.skupina === startGroup.skupina);
    if (groupIndex === -1) return;
    
    // Aplikovať rotáciu na nasledujúce týždne
    for (let i = startIndex + 1; i < allWeeks.length; i++) {
        const week = allWeeks[i];
        const nextGroupIndex = (groupIndex + (i - startIndex)) % employeeGroups.length;
        const nextGroup = employeeGroups[nextGroupIndex];
        
        // Priradiť iba ak týždeň je prázny
        if (!appState.dutyAssignments[week.key] || appState.dutyAssignments[week.key].length === 0) {
            appState.dutyAssignments[week.key] = nextGroup.moznosti.map(emp => ({
                ...emp,
                skupina: nextGroup.skupina
            }));
        }
    }
    
    zobrazOznamenie('Automatická rotácia aplikovaná na nasledujúce týždne', TOAST_INFO);
}

// --- TOGGLE REPORTING ---
function toggleReporting(employeeId, weekKey) {
    const reportingKey = `${weekKey}-${employeeId}`;
    appState.reporting[reportingKey] = !appState.reporting[reportingKey];
    
    renderKanbanBoard();
    saveStateToStorage();
    
    // FIX: Logika toastu je teraz správna
    const status = appState.reporting[reportingKey] ? 'pridané' : 'odobrané';
    zobrazOznamenie(`Hlásenie ${status}`, TOAST_INFO);
}

// --- INITIATE SWAP ---
function initiateSwap(employeeId, weekKey) {
    // --- ZAČIATOK ÚPRAVY ---
    // Zrušíme režim nahradenia, ak bol aktívny
    appState.employeeToReplace = null;
    // --- KONIEC ÚPRAVY ---

    if (appState.selectedDutyForSwap) {
        // Vykonať výmenu
        performSwap(appState.selectedDutyForSwap, { employeeId, weekKey });
        appState.selectedDutyForSwap = null;
    } else {
        // Označiť na výmenu
        appState.selectedDutyForSwap = { employeeId, weekKey };
        zobrazOznamenie('PRAVÝ KLIK: Vyberte druhého zamestnanca pre výmenu (pravý klik)', TOAST_INFO);
    }
    
    renderKanbanBoard();
}

// --- PERFORM SWAP ---
function performSwap(duty1, duty2) {
    const assignments1 = appState.dutyAssignments[duty1.weekKey] || [];
    const assignments2 = appState.dutyAssignments[duty2.weekKey] || [];
    
    const emp1Index = assignments1.findIndex(e => e.id === duty1.employeeId);
    const emp2Index = assignments2.findIndex(e => e.id === duty2.employeeId);
    
    if (emp1Index === -1 || emp2Index === -1) {
        zobrazOznamenie('Chyba pri výmene zamestnancov', TOAST_ERROR);
        return;
    }
    
    // Výmena
    const temp = assignments1[emp1Index];
    assignments1[emp1Index] = assignments2[emp2Index];
    assignments2[emp2Index] = temp;
    
    appState.dutyAssignments[duty1.weekKey] = assignments1;
    appState.dutyAssignments[duty2.weekKey] = assignments2;
    
    renderKanbanBoard();
    saveStateToStorage();
    
    zobrazOznamenie('Výmena úspešná', TOAST_SUCCESS);
}

// --- NOVÁ FUNKCIA: PERFORM REPLACEMENT ---
function performReplacement(replacingEmployeeId, replacingEmployeeWeekKey) {
    if (!appState.employeeToReplace) return;

    const { employeeId: replacedId, weekKey: replacedWeekKey } = appState.employeeToReplace;

    // 1. Nájdeme zamestnanca, ktorý nahrádza (napr. p. Kováč)
    const sourceWeekAssignments = appState.dutyAssignments[replacingEmployeeWeekKey] || [];
    const replacingEmployee = sourceWeekAssignments.find(e => e.id === replacingEmployeeId);

    if (!replacingEmployee) {
        zobrazOznamenie('Chyba: Nahrádzajúci zamestnanec nenájdený.', TOAST_ERROR);
        return;
    }

    // 2. Odstránime "chorého" zamestnanca (napr. p. Novák) z jeho týždňa (W44)
    let targetWeekAssignments = appState.dutyAssignments[replacedWeekKey] || [];
    targetWeekAssignments = targetWeekAssignments.filter(e => e.id !== replacedId);

    // 3. Pridáme nahrádzajúceho (p. Kováč) do týždňa W44
    // (p. Kováč zostáva aj vo svojom pôvodnom týždni W45)
    targetWeekAssignments.push({ ...replacingEmployee }); // Pridáme ako kópiu
    appState.dutyAssignments[replacedWeekKey] = targetWeekAssignments;

    // 4. Resetujeme stav
    appState.employeeToReplace = null;

    // 5. Uložíme a prekreslíme
    renderKanbanBoard();
    saveStateToStorage();
    zobrazOznamenie('Zastupovanie úspešne zaznamenané.', TOAST_SUCCESS);
}


// --- HELPER FUNKCIE ---

// --- ZAČIATOK ZMENY (DISPLAY DATES) ---
function getWeeksInMonth(year, month) {
    const weeks = [];
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    
    // Pridané: Hranice mesiaca
    const firstDayOfMonth = new Date(year, month, 1);
    const lastDayOfMonth = new Date(year, month + 1, 0);

    let currentDate = new Date(firstDay);
    
    // Začiatok prvého týždňa (pondelok)
    const dayOfWeek = currentDate.getDay();
    const diff = (dayOfWeek === 0 ? -6 : 1) - dayOfWeek;
    currentDate.setDate(currentDate.getDate() + diff);
    
    while (currentDate <= lastDay || currentDate.getMonth() === month) {
        const weekStart = new Date(currentDate);
        const weekEnd = new Date(currentDate);
        weekEnd.setDate(weekEnd.getDate() + 6);
        
        const weekInfo = getWeekNumber(weekStart);
        const weekKey = `${weekInfo.year}-W${String(weekInfo.week).padStart(2, '0')}`;
        
        // Pridané: Výpočet orezaných dátumov
        let displayStartDate = new Date(weekStart);
        let displayEndDate = new Date(weekEnd);
        
        if (displayStartDate < firstDayOfMonth) {
            displayStartDate = firstDayOfMonth;
        }
        if (displayEndDate > lastDayOfMonth) {
            displayEndDate = lastDayOfMonth;
        }
        
        weeks.push({
            key: weekKey,
            weekNumber: weekInfo.week,
            startDate: new Date(weekStart), // Pôvodný začiatok pre logiku
            endDate: new Date(weekEnd),     // Pôvodný koniec pre logiku
            displayStart: displayStartDate, // Začiatok pre zobrazenie
            displayEnd: displayEndDate      // Koniec pre zobrazenie
        });
        
        currentDate.setDate(currentDate.getDate() + 7);
        
        if (weeks.length > 6) break;
    }
    
    return weeks;
}
// --- KONIEC ZMENY ---

// *** NAHRADENÉ: getWeekNumber (verzia zo script2.js) ***
/**
 * Získa číslo týždňa pre daný dátum (ISO 8601)
 */
function getWeekNumber(date) {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    
    // Uložíme si správny ISO rok
    const isoYear = d.getUTCFullYear();
    
    const yearStart = new Date(Date.UTC(isoYear, 0, 1));
    const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    
    // Vrátime objekt namiesto jednoduchého čísla
    return { week: weekNo, year: isoYear };
}


// Formátovanie dátumu
function formatDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    return `${day}.${month}.`;
}

// Formátovanie dátumu pre PDF/DOCX
function formatDateFull(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}.${month}.${year}`;
}

// Názov dňa v týždni
function getDayName(date) {
    const days = ['Nedeľa', 'Pondelok', 'Utorok', 'Streda', 'Štvrtok', 'Piatok', 'Sobota'];
    return days[date.getDay()];
}

// --- ZAČIATOK NOVEJ FUNKCIE ---
/**
 * Získa 2-písmenovú skratku dňa.
 * @param {Date} date - Dátum
 * @returns {string} Skratka dňa (napr. 'po', 'ut', 'ne')
 */
function getDayAbbreviation(date) {
    const days = ['ne', 'po', 'ut', 'st', 'št', 'pi', 'so'];
    return days[date.getDay()];
}
// --- KONIEC NOVEJ FUNKCIE ---


// --- ZAČIATOK NOVEJ FUNKCIE ---
/**
 * Spraví, koľko dní z daného týždňa spadá do cieľového mesiaca.
 * @param {Date} startDate - Začiatok týždňa
 * @param {Date} endDate - Koniec týždňa
 * @param {number} targetYear - Cieľový rok (napr. 2025)
 * @param {number} targetMonth - Cieľový mesiac (0-11)
 * @returns {number} Počet dní patriacich do cieľového mesiaca
 */
function countDaysInMonthForWeek(startDate, endDate, targetYear, targetMonth) {
    let count = 0;
    const currentDate = new Date(startDate);
    const end = new Date(endDate);

    // Iterujeme po dňoch v rámci týždňa
    while (currentDate <= end) {
        if (currentDate.getFullYear() === targetYear && currentDate.getMonth() === targetMonth) {
            count++;
        }
        currentDate.setDate(currentDate.getDate() + 1);
    }
    return count;
}
// --- KONIEC NOVEJ FUNKCIE ---


// --- UKLADANIE A NAČÍTANIE STAVU ---
function saveStateToStorage() {
    try {
        localStorage.setItem('kanban_dutyAssignments', JSON.stringify(appState.dutyAssignments));
        localStorage.setItem('kanban_reporting', JSON.stringify(appState.reporting));
        localStorage.setItem('kanban_selectedMonth', appState.selectedMonth.toString());
        localStorage.setItem('kanban_selectedYear', appState.selectedYear.toString());
    } catch (error) {
        console.error('Chyba pri ukladaní stavu:', error);
    }
}

function loadStateFromStorage() {
    try {
        const savedAssignments = localStorage.getItem('kanban_dutyAssignments');
        const savedReporting = localStorage.getItem('kanban_reporting');
        const savedMonth = localStorage.getItem('kanban_selectedMonth');
        const savedYear = localStorage.getItem('kanban_selectedYear');
        
        if (savedAssignments) {
            appState.dutyAssignments = JSON.parse(savedAssignments);
        }
        
        if (savedReporting) {
            appState.reporting = JSON.parse(savedReporting);
        }
        
        if (savedMonth) {
            appState.selectedMonth = parseInt(savedMonth);
        }
        
        if (savedYear) {
            appState.selectedYear = parseInt(savedYear);
        }
    } catch (error) {
        console.error('Chyba pri načítaní stavu:', error);
    }
}

// --- CLEAR SCHEDULE ---
function clearSchedule() {
    if (!confirm('Naozaj chcete vymazať celý rozpis pre tento mesiac?')) {
        return;
    }
    
    const weeks = getWeeksInMonth(appState.selectedYear, appState.selectedMonth);
    weeks.forEach(week => {
        delete appState.dutyAssignments[week.key];
        
        // Vymazať aj reporting pre tento týždeň
        Object.keys(appState.reporting).forEach(key => {
            if (key.startsWith(week.key)) {
                // FIX: Bol tu preklep 'appBuda' namiesto 'appState'
                delete appState.reporting[key];
            }
        });
    });
    
    renderKanbanBoard();
    saveStateToStorage();
    
    zobrazOznamenie('Rozpis bol vymazaný', TOAST_SUCCESS);
}

// --- EXPORT TO PDF (UPRAVENÉ) ---
async function exportToPdf() {
    try {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();
        
        // Načítanie fontu
        let fontBase64;
        try {
            const fontResponse = await fetch(FONT_URL);
            const fontBlob = await fontResponse.blob();
            fontBase64 = await blobToBase64(fontBlob);
            doc.addFileToVFS('Roboto-Regular.ttf', fontBase64.split(',')[1]);
            doc.addFont('Roboto-Regular.ttf', 'Roboto', 'normal');
            doc.setFont('Roboto');
        } catch (err) {
            console.warn('Font sa nepodarilo načítať, použije sa fallback', err);
        }
        
        // Hlavička
        doc.setFontSize(16);
        doc.text('Rozpis pohotovosti - OKR BB', 105, 15, { align: 'center' });
        
        const localMonthNames = ['Január', 'Február', 'Marec', 'Apríl', 'Máj', 'Jún',
                           'Júl', 'August', 'September', 'Október', 'November', 'December'];
        
        doc.setFontSize(12);
        doc.text(`${localMonthNames[appState.selectedMonth]} ${appState.selectedYear}`, 105, 22, { align: 'center' });
        
        // Dáta pre tabuľku
        const weeks = getWeeksInMonth(appState.selectedYear, appState.selectedMonth);
        const tableData = [];
        
        weeks.forEach(week => {
            const assignments = appState.dutyAssignments[week.key] || [];
            
            // --- ZAČIATOK ZMENY (SKRATKA DNÍ) ---
            const dayStartAbbr = getDayAbbreviation(week.displayStart);
            const dayEndAbbr = getDayAbbreviation(week.displayEnd);
            const dayRangeAbbr = `${dayStartAbbr}-${dayEndAbbr}`;
            // --- KONIEC ZMENY ---

            // Použijeme pôvodné startDate/endDate pre výpočet dní
            const daysInMonthForWeek = countDaysInMonthForWeek(
                week.startDate,
                week.endDate,
                appState.selectedYear,
                appState.selectedMonth
            );

            if (assignments.length > 0) {
                assignments.forEach((emp, idx) => {
                    const reportingKey = `${week.key}-${emp.id}`;
                    const isReporting = appState.reporting[reportingKey] ? 'hlásenia' : '';
                    
                    // --- ZAČIATOK ZMENY (PRIDANIE \n) ---
                    tableData.push([
                        idx === 0 ? `Týždeň ${week.weekNumber}` : '',
                        // Pridáme skratku dňa s novým riadkom
                        idx === 0 ? `${formatDate(week.displayStart)} - ${formatDate(week.displayEnd)}\n${dayRangeAbbr}` : '',
                        emp.meno,
                        emp.telefon || '',
                        idx === 0 ? daysInMonthForWeek.toString() : '', // Nový stĺpec Σ
                        isReporting
                    ]);
                    // --- KONIEC ZMENY ---
                });
            } else {
                // --- ZAČIATOK ZMENY (PRIDANIE \n) ---
                tableData.push([
                    `Týždeň ${week.weekNumber}`,
                    // Pridáme skratku dňa s novým riadkom
                    `${formatDate(week.displayStart)} - ${formatDate(week.displayEnd)}\n${dayRangeAbbr}`,
                    'Nepriradené',
                    '',
                    daysInMonthForWeek.toString(), // Nový stĺpec Σ
                    ''
                ]);
                // --- KONIEC ZMENY ---
            }
        });
        
        // Vytvorenie tabuľky
        doc.autoTable({
            startY: 30,
            head: [['Týždeň', 'Dátum', 'Meno', 'Telefón', 'Σ', 'Poznámka']],
            body: tableData,
            theme: 'grid',
            styles: {
                font: 'Roboto',
                fontSize: 10,
                cellPadding: 3,
                overflow: 'linebreak'
            },
            headStyles: {
                fillColor: [0, 51, 102],
                textColor: 255,
                fontStyle: 'bold'
            },
            columnStyles: {
                0: { cellWidth: 25 }, // Týždeň
                1: { cellWidth: 35 }, // Dátum
                // 2: Meno - auto
                3: { cellWidth: 35 }, // Telefón
                4: { cellWidth: 15, halign: 'center' }, // Nový stĺpec Σ
                5: { cellWidth: 30 }  // Poznámka
            },
            
            didParseCell: function(data) {
                if (data.section === 'body') {
                    const firstCellText = data.row.cells[0]?.text[0] || '';
                    if (firstCellText.startsWith('Týždeň')) {
                        data.cell.styles.fillColor = '#F3F4F6'; // --gray-100
                    }
                }
            }
        });
        
        // --- ZAČIATOK NOVEJ ZMENY (PODPISY) ---
        // Získame pozíciu Y, kde skončila tabuľka
        const tableEndY = doc.autoTable.previous.finalY;
        let signatureY = tableEndY + 27; // Pridáme 27px medzeru
        
        const leftMargin = 20;
        const rightMargin = 171;
        
        doc.setFontSize(10);
        doc.setFont('Roboto', 'normal');

        // Podpis vľavo (Zodpovedá)
        if (employeeSignatures.zodpoveda) {
            doc.text('Zodpovedá:', leftMargin, signatureY);
            doc.text(employeeSignatures.zodpoveda.meno, leftMargin, signatureY + 5);
            doc.text(employeeSignatures.zodpoveda.funkcia, leftMargin, signatureY + 10);
        }

        // Podpis vpravo (Schvaľuje)
        if (employeeSignatures.schvaluje) {
            doc.text('Schvaľuje:            ', rightMargin, signatureY, { align: 'right' });
            doc.text(employeeSignatures.schvaluje.meno, rightMargin, signatureY + 5, { align: 'right' });
            doc.text(employeeSignatures.schvaluje.funkcia, rightMargin, signatureY + 10, { align: 'right' });
        }
        // --- KONIEC NOVEJ ZMENY (PODPISY) ---


        // --- ZAČIATOK NOVEJ ZMENY (SKRYTIE UI) ---
        // Zobrazenie v modali
        const pdfBlob = doc.output('blob');
        const pdfUrl = URL.createObjectURL(pdfBlob);
        
        // Pridáme parametre na skrytie UI prehliadača PDF
        const pdfUrlWithParams = `${pdfUrl}#toolbar=0&navpanes=0&scrollbar=0`;
        
        document.getElementById('pdfPreviewFrame').src = pdfUrlWithParams;
        document.getElementById('previewModal').style.display = 'block';
        // --- KONIEC NOVEJ ZMENY (SKRYTIE UI) ---
        
        // Uloženie pre download
        window.currentPdfBlob = pdfBlob;
        
        zobrazOznamenie('PDF náhľad vygenerovaný', TOAST_SUCCESS);
    } catch (error) {
        console.error('Chyba pri generovaní PDF:', error);
        zobrazOznamenie('Chyba pri generovaní PDF', TOAST_ERROR);
    }
}

// --- DOWNLOAD PDF FROM MODAL ---
function downloadPdfFromModal() {
    if (window.currentPdfBlob) {
        const localMonthNames = ['Januar', 'Februar', 'Marec', 'April', 'Maj', 'Jun',
                           'Jul', 'August', 'September', 'Oktober', 'November', 'December'];
        const filename = `Rozpis_${localMonthNames[appState.selectedMonth]}_${appState.selectedYear}.pdf`;
        
        saveAs(window.currentPdfBlob, filename);
        zobrazOznamenie('PDF stiahnuté', TOAST_SUCCESS);
    }
}


// --- GENERATE DOCX ---
// *** OPRAVENÁ LOGIKA: Zamestnanec je v skupine, nie modulo ***
async function generateDocx() {
    try {
        const response = await fetch(DOCX_TEMPLATE_URL);
        const arrayBuffer = await response.arrayBuffer();
        
        const zip = new PizZip(arrayBuffer);
        const doc = new window.docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            nullGetter: () => "", 
            delimiters: { start: '{{', end: '}}' }
        });
        
        // Príprava dát
        const data = {
            mesiac: monthNames[appState.selectedMonth],
            rok: appState.selectedYear
        };
        
        const daysInMonth = new Date(appState.selectedYear, appState.selectedMonth + 1, 0).getDate();

        // Iterujeme cez všetkých zamestnancov (max 9, podľa šablóny)
        for (let i = 0; i < 9; i++) {
            const employee = allEmployees[i];

            if (employee) {
                data[i.toString()] = employee.meno;
                data["oc" + i] = employee.coz || ''; 
                
                const employeeDates = [];
                let sum_popi = 0;
                let sum_sonesv = 0;
                let sum_hodinyP5 = 0; 
                let sum_hodinySn10 = 0;

                // Iterujeme dni od 1 do X
                for (let day = 1; day <= daysInMonth; day++) {
                    const currentDate = new Date(appState.selectedYear, appState.selectedMonth, day);
                    
                    // Zistíme, do ktorého týždňa tento deň patrí
                    const weekInfo = getWeekNumber(currentDate);
                    const weekKey = `${weekInfo.year}-W${String(weekInfo.week).padStart(2, '0')}`;
                    
                    const assignments = appState.dutyAssignments[weekKey] || [];
                    if (assignments.length === 0) continue;

                    // *** TOTO JE OPRAVENÁ LOGIKA ***
                    // Zistíme, či je zamestnanec (z vonkajšej slučky) členom priradenej skupiny
                    const isEmployeeInAssignedGroup = assignments.some(person => person.id === employee.id);
                    
                    // Ak je zamestnanec v skupine, ktorá má v tento deň službu:
                    if (isEmployeeInAssignedGroup) {
                        
                        const dayOfWeek = currentDate.getDay(); // 0=Ne, 1=Po, ..., 6=So
                        const isPoPi = dayOfWeek >= 1 && dayOfWeek <= 5;
                        const isSoNeSv = !isPoPi; // 0 (Nedeľa) alebo 6 (Sobota)

                        const dateEntry = {
                            date: formatDateFull(currentDate),
                            popi: "",
                            sonesv: "",
                            p5: "", 
                            sn10: ""
                        };

                        // Aplikujeme logiku hodín z script2.js
                        if (isPoPi) {
                            dateEntry.popi = "1";
                            sum_popi++;
                            dateEntry.p5 = "16"; 
                            sum_hodinyP5 += 16;
                        } else if (isSoNeSv) {
                            dateEntry.sonesv = "1";
                            sum_sonesv++;
                            dateEntry.sn10 = "24";
                            sum_hodinySn10 += 24;
                        }
                        employeeDates.push(dateEntry);
                    }
                    // *** KONIEC OPRAVENEJ LOGIKY ***
                }

                // Priradíme dáta pre šablónu
                data["dates" + i] = employeeDates;
                data["sum1" + i] = sum_popi || "";
                data["sum2" + i] = sum_sonesv || "";
                data["sum3" + i] = sum_hodinyP5 || "";
                data["sum4" + i] = sum_hodinySn10 || "";

            } else {
                // Vyplníme prázdne dáta pre nevyužité sloty v šablóne
                data[i.toString()] = "";
                data["oc" + i] = "";
                data["dates" + i] = [];
                data["sum1" + i] = "";
                data["sum2" + i] = "";
                data["sum3" + i] = "";
                data["sum4" + i] = "";
            }
        }
        
        doc.render(data);
        
        const blob = doc.getZip().generate({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
        
        const filename = `Vykaz_${monthNames[appState.selectedMonth]}_${appState.selectedYear}.docx`;
        saveAs(blob, filename);
        
        zobrazOznamenie('Výkaz (.docx) stiahnutý', TOAST_SUCCESS);
    } catch (error) {
        console.error('Chyba pri generovaní DOCX:', error);
        zobrazOznamenie('Chyba pri generovaní výkazu', TOAST_ERROR);
    }
}


// --- CLOSE MODAL ---
function closeModal() {
    document.getElementById('previewModal').style.display = 'none';
    document.getElementById('pdfPreviewFrame').src = '';
}

// --- NOTIFIKÁCIE ---
function zobrazOznamenie(message, type = TOAST_INFO) {
    const existingToast = document.querySelector('.oznamenie');
    if (existingToast) {
        existingToast.remove();
    }
    
    const toast = document.createElement('div');
    toast.className = `oznamenie ${type}`;
    toast.textContent = message;
    
    document.body.appendChild(toast);
    
    setTimeout(() => toast.classList.add('show'), 100);
    
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 500);
    }, 3000);
}

// --- HELPER: BLOB TO BASE64 ---
function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

// --- WINDOW CLICK OUTSIDE MODAL ---
window.onclick = function(event) {
    const modal = document.getElementById('previewModal');
    if (event.target === modal) {
        closeModal();
    }
};