// --- APLIKÁCIA ROZPIS POHOTOVOSTI ---

// Konfigurácia načítaná z JSON
let allEmployees = [];
let employeeGroups = []; 
let employeeSignatures = {};

// Centrálny stav (State)
let appState = {
    selectedMonth: new Date().getMonth(),
    selectedYear: new Date().getFullYear(),
    dutyAssignments: {}, 
    reporting: {},
    selectedDutyForSwap: null,
    serviceOverrides: {}
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
// ZMENA: ID pre 3 stĺpce
const ID_AVAILABLE_LIST_1 = 'availableList1';
const ID_AVAILABLE_LIST_2 = 'availableList2';
const ID_AVAILABLE_LIST_3 = 'availableList3';
const ID_HEADER_GROUP_1 = 'headerGroup1';
const ID_HEADER_GROUP_2 = 'headerGroup2';
const ID_HEADER_GROUP_3 = 'headerGroup3';

const ID_CALENDAR_INFO = 'currentMonthInfo';
const ID_PDF_PREVIEW_FRAME = 'pdfPreviewFrame';
const ID_PREVIEW_MODAL = 'previewModal';

// ODSTRÁNENÉ: GROUP_ID_AVAILABLE

// CSS Triedy
const CSS_EMPLOYEE_ITEM = 'employee-item';
const CSS_GROUP_DRAG_ITEM = 'group-drag-item';
const CSS_REPORTING = 'reporting';
const CSS_IS_ASSIGNED = 'is-assigned';
const CSS_SORTABLE_GHOST = 'sortable-ghost'; 
const CSS_CALENDAR_DROP_GHOST = 'calendar-drop-ghost'; 
const CSS_SORTABLE_DRAG = 'sortable-drag';
const CSS_SWAP_SELECTED = 'swap-selected'; 

// Názvy pre PDF font
const PDF_FONT_FILENAME = 'Roboto-Regular.ttf';
const PDF_FONT_INTERNAL_NAME = 'Roboto-Regular';

// Pomocné globálne premenné
// --- OPRAVA: Názvy mesiacov s malým začiatočným písmenom ---
const monthNames = [
    "január", "február", "marec", "apríl", "máj", "jún", "júl", "august", "september", "október", "november", "december"
];
const dayNames = ["Pondelok", "Utorok", "Streda", "Štvrtok", "Piatok", "Sobota", "Nedeľa"];

// Premenné pre PDF náhľad
let generatedPDFDataURI = null;
let generatedPDFFilename = '';

// PREMENNÁ PRE NAČÍTANÝ FONT
let customFontBase64 = null;

// --- DOM ELEMENTY (Cache) ---
let elMonthSelect, elYearSelect, elExportButton, elSaveButton, elClearButton;
let elCloseModalButton, elDownloadPdfButton;
// ZMENA: Cache pre 3 stĺpce
let elAvailableList1, elAvailableList2, elAvailableList3;
let elHeaderGroup1, elHeaderGroup2, elHeaderGroup3;
let elCalendarInfo, elPdfPreviewFrame, elPreviewModal;


// --- INICIALIZÁCIA APLIKÁCIE ---

document.addEventListener('DOMContentLoaded', async () => {
    // 1. Cache-ovanie DOM elementov
    elMonthSelect = document.getElementById(ID_MONTH_SELECT);
    elYearSelect = document.getElementById(ID_YEAR_SELECT);
    elExportButton = document.getElementById(ID_EXPORT_BUTTON);
    elSaveButton = document.getElementById(ID_SAVE_BUTTON);
    elClearButton = document.getElementById(ID_CLEAR_BUTTON);
    elCloseModalButton = document.getElementById(ID_CLOSE_MODAL_BUTTON);
    elDownloadPdfButton = document.getElementById(ID_DOWNLOAD_PDF_BUTTON);
    // ZMENA: Cache pre 3 stĺpce
    elAvailableList1 = document.getElementById(ID_AVAILABLE_LIST_1);
    elAvailableList2 = document.getElementById(ID_AVAILABLE_LIST_2);
    elAvailableList3 = document.getElementById(ID_AVAILABLE_LIST_3);
    elHeaderGroup1 = document.getElementById(ID_HEADER_GROUP_1);
    elHeaderGroup2 = document.getElementById(ID_HEADER_GROUP_2);
    elHeaderGroup3 = document.getElementById(ID_HEADER_GROUP_3);

    elCalendarInfo = document.getElementById(ID_CALENDAR_INFO);
    elPdfPreviewFrame = document.getElementById(ID_PDF_PREVIEW_FRAME);
    elPreviewModal = document.getElementById(ID_PREVIEW_MODAL);
    
    // 2. Načítať konfiguráciu
    await loadConfig();
    
    // 3. Načítať font (asynchrónne na pozadí)
    loadFontData();
    
    // 4. Inicializovať UI komponenty
    populateYearSelect();
    elMonthSelect.value = appState.selectedMonth;
    // ZMENA: Inicializácia 3 D&D zoznamov
    initSortableList(elAvailableList1);
    initSortableList(elAvailableList2);
    initSortableList(elAvailableList3);
    initEventListeners();
    
    // 5. Vykresliť počiatočný stav
    render();
});

/**
 * Načíta konfiguráciu zo súboru config.json
 */
async function loadConfig() {
    try {
        const response = await fetch(CONFIG_URL);
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        
        const config = await response.json();
        
        allEmployees = config.zamestnanci.flatMap(group => group.moznosti);
        employeeGroups = config.zamestnanci; 
        employeeSignatures = config.podpisy;
        
    } catch (error) {
        console.error('Chyba pri načítaní files/config.json:', error);
        zobrazOznamenie('Nepodarilo sa načítať konfiguráciu zamestnancov. Aplikácia je nefunkčná.', TOAST_ERROR);
        
        elExportButton.disabled = true;
        elSaveButton.disabled = true;
        elClearButton.disabled = true;
    }
}

/**
 * Naplní <select> rokmi
 */
function populateYearSelect() {
    const currentYear = new Date().getFullYear();
    for (let year = currentYear - 5; year <= currentYear + 5; year++) {
        const option = document.createElement('option');
        option.value = year;
        option.textContent = year;
        elYearSelect.appendChild(option);
    }
    elYearSelect.value = currentYear;
    appState.selectedYear = currentYear;
}

/**
 * Inicializuje všetky hlavné event listenery
 */
function initEventListeners() {
    elMonthSelect.addEventListener('change', handleDateChange);
    elYearSelect.addEventListener('change', handleDateChange);
    
    elExportButton.addEventListener('click', showSchedulePreview);
    elSaveButton.addEventListener('click', generateDocxReport);
    elClearButton.addEventListener('click', clearSchedule); 

    // Listenery pre modálne okno
    elCloseModalButton.addEventListener('click', closeModal);
    elDownloadPdfButton.addEventListener('click', downloadSchedulePDF);

    // Dvojité kliknutie pre výber výmeny
    elCalendarInfo.addEventListener('dblclick', handleCalendarDutyDblClick);
    
    // ZMENA: Kliknutie na zamestnanca v 3 D&D zoznamoch pre dokončenie výmeny
    elAvailableList1.addEventListener('click', handleEmployeeListClick);
    elAvailableList2.addEventListener('click', handleEmployeeListClick);
    elAvailableList3.addEventListener('click', handleEmployeeListClick);

    // Kliknutie na položku v kalendári (pre reporting)
    elCalendarInfo.addEventListener('click', handleCalendarDutyClick);
}

// --- SPRÁVA STAVU (STATE MANAGEMENT) ---

/**
 * Aktualizuje stav (appState) pri zmene dátumu a prekreslí UI
 */
function handleDateChange(e) {
    const { id, value } = e.target;
    if (id === ID_MONTH_SELECT) {
        appState.selectedMonth = parseInt(value);
    } else if (id === ID_YEAR_SELECT) {
        appState.selectedYear = parseInt(value);
    }
    render(); 
}

/**
 * Spracuje dvojité kliknutie na zamestnanca v kalendári (Vyberá cieľovú službu pre výmenu)
 * (BEZO ZMENY)
 */
function handleCalendarDutyDblClick(e) {
    const employeeItem = e.target.closest('.employee-item-calendar');
    if (!employeeItem) return;

    const { employeeId, weekKey } = employeeItem.dataset;
    
    document.querySelectorAll(`.${CSS_SWAP_SELECTED}`).forEach(el => el.classList.remove(CSS_SWAP_SELECTED));
    
    if (appState.selectedDutyForSwap && appState.selectedDutyForSwap.weekKey === weekKey && appState.selectedDutyForSwap.originalId === employeeId) {
        appState.selectedDutyForSwap = null;
        zobrazOznamenie('Výber služby pre výmenu zrušený.', TOAST_INFO);
    } else {
        appState.selectedDutyForSwap = { weekKey, originalId: employeeId };
        employeeItem.classList.add(CSS_SWAP_SELECTED);
        
        const originalEmployee = findEmployeeById(employeeId);
        const employeeName = originalEmployee ? originalEmployee.meno : 'Neznámy zamestnanec';
        
        zobrazOznamenie(`VYBRANÁ SLUŽBA PRE VÝMENU: ${employeeName} v týždni ${weekKey}. Teraz kliknite na náhradníka v zoznamoch dostupných.`, TOAST_INFO);
    }
}

/**
 * (UPRAVENÉ) Spracuje kliknutie na zamestnanca v D&D zozname (Dostupní). Ak je vybraná služba pre SWAP, vykoná prepis.
 */
function handleEmployeeListClick(e) {
    const employeeItem = e.target.closest('.employee-item');
    if (!employeeItem) return;

    if (!appState.selectedDutyForSwap) return; 

    const newEmployeeId = employeeItem.dataset.id;
    if (!newEmployeeId) return;

    const { weekKey, originalId } = appState.selectedDutyForSwap;

    if (newEmployeeId === originalId) {
        zobrazOznamenie('Náhradník nemôže byť tá istá osoba. Vyberte iného zamestnanca.', TOAST_ERROR);
        return;
    }

    if (!appState.serviceOverrides[weekKey]) {
        appState.serviceOverrides[weekKey] = {};
    }
    appState.serviceOverrides[weekKey][originalId] = newEmployeeId;

    appState.selectedDutyForSwap = null;
    document.querySelectorAll(`.${CSS_SWAP_SELECTED}`).forEach(el => el.classList.remove(CSS_SWAP_SELECTED));
    zobrazOznamenie(`Služba bola úspešne prepísaná v týždni ${weekKey}.`, TOAST_SUCCESS);

    // --- OPRAVA CHYBY ---
    // renderCalendar();  // PÔVODNÝ KÓD (chybný)
    render();            // NOVÝ KÓD (opravený)
}


/**
 * Spracuje kliknutie na položku zamestnanca priamo v kalendári (Reporting)
 * (BEZO ZMENY)
 */
function handleCalendarDutyClick(e) {
    const employeeItem = e.target.closest('.employee-item-calendar');
    if (!employeeItem) return;

    const { employeeId, weekKey } = employeeItem.dataset;
    if (!employeeId || !weekKey) return;

    if (appState.selectedDutyForSwap) { }

    // --- Logika REPORTING ---
    if (!appState.reporting[weekKey]) {
        appState.reporting[weekKey] = [];
    }
    const reportingArray = appState.reporting[weekKey];
    const index = reportingArray.indexOf(employeeId);
    let isNowReporting; 

    if (index > -1) {
        reportingArray.splice(index, 1);
        isNowReporting = false;
    } else {
        reportingArray.push(employeeId);
        isNowReporting = true;
    }
    
    employeeItem.classList.toggle(CSS_REPORTING, isNowReporting);

    const allReportingIds = Object.values(appState.reporting).flat();
    const isReportingInAnyWeek = allReportingIds.includes(employeeId);

    document.querySelectorAll(`.employee-item[data-id="${employeeId}"]`).forEach(el => {
        el.classList.toggle(CSS_REPORTING, isReportingInAnyWeek);
    });
}

/**
 * Vymaže celý rozpis z kalendára a prekreslí UI
 * (BEZO ZMENY)
 */
function clearSchedule() {
    appState.dutyAssignments = {};
    appState.reporting = {};
    appState.serviceOverrides = {};
    appState.selectedDutyForSwap = null;
    
    render(); 
    zobrazOznamenie('Celý rozpis služieb a prepisy boli vymazané.', TOAST_INFO);
}

/**
 * Pomocná funkcia na nájdenie zamestnanca v globálnej konfigurácii
 * (BEZO ZMENY)
 */
function findEmployeeById(id) {
    return allEmployees.find(emp => emp.id === id);
}

// --- VYKRESLOVACIE FUNKCIE (RENDER) ---

/**
 * Hlavná funkcia na prekreslenie celého UI na základe appState
 * (BEZO ZMENY)
 */
function render() {
    renderGroupLists();
    renderCalendar();
    initCalendarSortable();
}

/**
 * (UPRAVENÉ) Vykreslí 3 stĺpce dostupných zamestnancov
 */
function renderGroupLists() {
    const assignedIds = new Set(
        Object.values(appState.dutyAssignments).flat().map(p => p.id)
    );
    const allReportingIds = new Set(
        Object.values(appState.reporting).flat()
    );

    const buildListContent = (group) => {
        if (!group) return ''; 
        let html = '';
        const availableInGroup = group.moznosti;
        if (availableInGroup.length === 0) return '';
        
        const allMembersAvailable = group.moznosti.every(emp => !assignedIds.has(emp.id));

        if (allMembersAvailable) {
            const employeeIds = JSON.stringify(group.moznosti.map(emp => emp.id));
            // --- ZMENA: Text tlačidla skupiny ---
            html += `
                <div class="${CSS_GROUP_DRAG_ITEM}" 
                     data-group-ids='${employeeIds}' 
                     data-group-name="${group.skupina}">
                     ${group.skupina}
                </div>
            `;
        }

        html += availableInGroup.map(employee => {
            const isAssigned = assignedIds.has(employee.id);
            const isReporting = allReportingIds.has(employee.id);
            
            let classes = CSS_EMPLOYEE_ITEM;
            if (isAssigned) classes += ` ${CSS_IS_ASSIGNED}`;
            if (isReporting) classes += ` ${CSS_REPORTING}`;

            return `
                <div class="${classes}" data-id="${employee.id}">
                    ${employee.meno}
                </div>
            `;
        }).join('');
        
        return html;
    };

    const group1 = employeeGroups[0];
    if (group1) {
        elHeaderGroup1.textContent = group1.skupina;
        elAvailableList1.innerHTML = buildListContent(group1);
    } else {
        elHeaderGroup1.textContent = '';
        elAvailableList1.innerHTML = '';
    }

    const group2 = employeeGroups[1];
    if (group2) {
        elHeaderGroup2.textContent = group2.skupina;
        elAvailableList2.innerHTML = buildListContent(group2);
    } else {
        elHeaderGroup2.textContent = '';
        elAvailableList2.innerHTML = '';
    }

    const group3 = employeeGroups[2];
    if (group3) {
        elHeaderGroup3.textContent = group3.skupina;
        elAvailableList3.innerHTML = buildListContent(group3);
    } else {
        elHeaderGroup3.textContent = '';
        elAvailableList3.innerHTML = '';
    }
}


/**
 * Vytvorí HTML reťazec pre jedného zamestnanca (používané v D&D)
 * (BEZO ZMENY)
 */
function createEmployeeItemHTML(employee, allReportingIds) {
    const isReporting = allReportingIds.has(employee.id);
    
    return `
        <div class="${CSS_EMPLOYEE_ITEM} ${isReporting ? CSS_REPORTING : ''}" data-id="${employee.id}">
            ${employee.meno}
        </div>
    `;
}

/**
 * (UPRAVENÉ) Vykreslí "živý" kalendár služieb na základe appState
 * Odstránená pomlčka '—'
 */
function renderCalendar() {
    const currentMonth = appState.selectedMonth;
    const currentYear = appState.selectedYear;

    const firstDay = new Date(currentYear, currentMonth, 1);
    const lastDay = new Date(currentYear, currentMonth + 1, 0);

    let htmlContent = `<h2>${monthNames[currentMonth]} ${currentYear}</h2>
        <table>
            <thead>
                <tr>
                    <th class="week-column">Týždeň</th>
                    <th class="day-column">Deň</th>
                    <th class="date-column">Dátum</th>
                    <th class="duty-column-header">Služba</th>
                </tr>
            </thead>
            <tbody>`;

    let lastRenderedWeekKey = -1; 
    let dutyCycleDayIndex = 0; 
    let previousGroupEmployees = []; 

    const addGhostRows = (startIndex, group, weekKey) => {
        for (let i = startIndex; i < group.length; i++) {
            const employee = group[i];
            
            let employeeToRender = employee;
            if (employee) {
                const overridesForWeek = appState.serviceOverrides[weekKey];
                if (overridesForWeek) {
                    const overrideId = overridesForWeek[employee.id];
                    if (overrideId) {
                        const newEmployee = findEmployeeById(overrideId);
                        if (newEmployee) employeeToRender = newEmployee;
                    }
                }
            }
            if (!employeeToRender) continue;
            
            const reportersForWeek = appState.reporting[weekKey] || [];
            const isReporting = reportersForWeek.includes(employeeToRender.id);
            
            const isSwapSelected = appState.selectedDutyForSwap && 
                                   appState.selectedDutyForSwap.weekKey === weekKey &&
                                   appState.selectedDutyForSwap.originalId === employeeToRender.id;
            const extraClass = isSwapSelected ? CSS_SWAP_SELECTED : '';

            const ghostDutyCellContent = `
                <div class="employee-item-calendar ${isReporting ? CSS_REPORTING : ''} ${extraClass}" 
                     data-employee-id="${employeeToRender.id}" 
                     data-week-key="${weekKey}">
                    ${employeeToRender.meno}
                </div>
            `;
            
            htmlContent += `
                <tr class="ghost-row">
                    <td></td>
                    <td></td>
                    <td style="text-align: right; opacity: 0.6; font-size: 0.8em;">(pokr. t. ${weekKey.split('-')[1]})</td>
                    <td class="duty-column" data-week-key="${weekKey}">
                        ${ghostDutyCellContent}
                    </td>
                </tr>
            `;
        }
    };


    for (let day = 1; day <= lastDay.getDate(); day++) {
        const date = new Date(currentYear, currentMonth, day);
        const dayOfWeek = (date.getDay() + 6) % 7; 
        const weekInfo = getWeekNumber(date);
        const weekNumber = weekInfo.week; // Ponechané pre logiku 'weekCellContent'
        const reportingKey = `${weekInfo.year}-${weekInfo.week}`; // Použije správny ISO rok 
        
        if (reportingKey !== lastRenderedWeekKey) {
            
            if (lastRenderedWeekKey !== -1 && 
                day > 1 && day <= 7 && 
                previousGroupEmployees.length > dutyCycleDayIndex) 
            {
                addGhostRows(dutyCycleDayIndex, previousGroupEmployees, lastRenderedWeekKey);
            }

            dutyCycleDayIndex = 0;
            lastRenderedWeekKey = reportingKey;
        }

        const activeGroupEmployees = appState.dutyAssignments[reportingKey] || [];
        previousGroupEmployees = activeGroupEmployees; 
        
        let dutyCellContent = ''; 

        if (activeGroupEmployees.length > 0) {
            let employee = activeGroupEmployees[dutyCycleDayIndex]; 
            
            let employeeToRender = employee;
            if (employee) {
                const overridesForWeek = appState.serviceOverrides[reportingKey];
                if (overridesForWeek) {
                    const overrideId = overridesForWeek[employee.id];
                    if (overrideId) {
                        const newEmployee = findEmployeeById(overrideId);
                        if (newEmployee) {
                            employeeToRender = newEmployee; 
                        }
                    }
                }
            }
            
            if (employeeToRender) {
                const reportersForWeek = appState.reporting[reportingKey] || [];
                const isReporting = reportersForWeek.includes(employeeToRender.id);
                
                const isSwapSelected = appState.selectedDutyForSwap && 
                                       appState.selectedDutyForSwap.weekKey === reportingKey &&
                                       appState.selectedDutyForSwap.originalId === employeeToRender.id;
                                       
                const extraClass = isSwapSelected ? CSS_SWAP_SELECTED : '';

                dutyCellContent = `
                    <div class="employee-item-calendar ${isReporting ? CSS_REPORTING : ''} ${extraClass}" 
                         data-employee-id="${employeeToRender.id}" 
                         data-week-key="${reportingKey}">
                        ${employeeToRender.meno}
                    </div>
                `;
            } else {
                // --- ZMENA: Odstránená pomlčka ---
                dutyCellContent = '';
            }
            
            dutyCycleDayIndex++;
            
        } else {
            dutyCellContent = ''; 
            dutyCycleDayIndex = 0; 
            previousGroupEmployees = []; 
        }
        
        const weekCellContent = (dayOfWeek === 0 || (day === 1 && date.getDay() !== 1)) ? weekNumber : '';

        htmlContent += `<tr class="${dayOfWeek === 0 ? 'new-week' : ''}">
            <td class="week-column">${weekCellContent}</td>
            <td class="day-column">${dayNames[dayOfWeek]}</td>
            <td class="date-column">${date.getDate()}. ${monthNames[currentMonth]} ${currentYear}</td>
            <td class="duty-column" 
                id="duty-${date.getDate()}-${currentMonth}-${currentYear}" 
                data-week-key="${reportingKey}">
                ${dutyCellContent}
            </td>
        </tr>`;
    }

    if (previousGroupEmployees.length > dutyCycleDayIndex) {
        addGhostRows(dutyCycleDayIndex, previousGroupEmployees, lastRenderedWeekKey);
    }

    htmlContent += `</tbody></table>`;
    elCalendarInfo.innerHTML = htmlContent;
}

// --- LOGIKA A POMOCNÉ FUNKCIE ---

/**
 * (UPRAVENÉ) Inicializuje D&D pre JEDEN zoznam dostupných
 * Ponecháva možnosť presunúť len celé skupiny (.group-drag-item)
 */
function initSortableList(listElement) {
    if (listElement) {
        const options = {
            animation: 150,
            ghostClass: CSS_SORTABLE_GHOST,
            dragClass: CSS_SORTABLE_DRAG,
            group: { name: 'shared', pull: 'clone', put: false },
            filter: '.available-group-header', 
            // --- ZMENA: Umožní ťahať iba celé skupiny ---
            draggable: `.${CSS_GROUP_DRAG_ITEM}`, 
        };
        new Sortable(listElement, options);
    }
}

/**
 * Inicializuje D&D pre bunky kalendára.
 * (BEZO ZMENY)
 */
function initCalendarSortable() {
    const calendarCells = document.querySelectorAll('.duty-column');
    calendarCells.forEach(cell => {
        new Sortable(cell, {
            group: 'shared',
            animation: 150,
            ghostClass: CSS_CALENDAR_DROP_GHOST,
            onAdd: handleDragToCalendar 
        });
    });
}

/**
 * (UPRAVENÉ) Handler pre D&D udalosť 'onAdd' (pustenie položky do bunky kalendára)
 * Logika cyklu bola upravená, aby umožnila priradenie služieb presahujúcich
 * do januára nasledujúceho roka.
 */
function handleDragToCalendar(evt) {
    const { item, to } = evt; 
    const weekKey = to.dataset.weekKey;
    
    if (!weekKey) return;

    // Kontrola, či je to skupina
    const isGroupDrag = item.classList.contains(CSS_GROUP_DRAG_ITEM);
    if (!isGroupDrag) {
        item.remove();
        return; 
    }

    let employeesToAssign = [];

    const employeeIds = JSON.parse(item.dataset.groupIds); 
    employeesToAssign = employeeIds.map(findEmployeeById).filter(Boolean);

    item.remove(); // Odstráni "ghost" element z kalendára

    const [startYear, startWeek] = weekKey.split('-').map(Number);
    const targetMonth = appState.selectedMonth; 
    const targetYear = appState.selectedYear; 

    let currentYear = startYear;
    let currentWeek = startWeek;

    // Pôvodné 'while(true)' nahradené bezpečným 'for' cyklom s poistkou (max 52 iterácií = 3 roky)
    for (let i = 0; i < 52; i++) {
        const mondayOfCurrentWeek = getDateOfISOWeek(currentWeek, currentYear);

        // --- OPRAVENÁ LOGIKA PRE UKONČENIE CYKLU ---
        // Pôvodné prísne 'if... break;' boli odstránené.
        // Nová logika povoľuje "presah" do januára nasledujúceho roka,
        // ale zastaví sa, akonáhle sa dostane do februára (alebo ďalej).

        // 1. Skontrolujeme, či sme neprekročili cieľový rok o VIAC ako 1
        if (mondayOfCurrentWeek.getFullYear() > targetYear + 1) {
            break; // Sme napr. v roku 2027, keď sme začali v 2025
        }

        // 2. Skontrolujeme, či sme presne o 1 rok ďalej
        if (mondayOfCurrentWeek.getFullYear() === targetYear + 1) {
            // A ak áno, či sme už ďalej ako v Januári (index mesiaca 0)
            if (mondayOfCurrentWeek.getMonth() > 0) {
                break; // Sme v Februári (alebo neskôr) 2026, končíme.
            }
        }
        // --- KONIEC OPRAVENEJ LOGIKY ---

        const currentKey = `${currentYear}-${currentWeek}`;
        appState.dutyAssignments[currentKey] = employeesToAssign;

        currentWeek += 3;

        // (Volanie getISOWeeks už je opravené nižšie)
        const weeksInYear = getISOWeeks(currentYear);
        if (currentWeek > weeksInYear) {
            currentWeek = currentWeek - weeksInYear;
            currentYear += 1;
        }
    }

    render(); 
}


/**
 * Získa číslo týždňa pre daný dátum (ISO 8601)
 * (BEZO ZMENY)
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

/**
 * Získa celkový počet ISO týždňov v danom roku.
 * (OPRAVENÉ - číta 'week' property z vráteného objektu)
 */
function getISOWeeks(year) {
    const date = new Date(year, 11, 28); 
    // getWeekNumber teraz vracia { week: ..., year: ... }
    // Musíme vrátiť len počet týždňov.
    const weekInfo = getWeekNumber(date);
    return weekInfo.week;
}

/**
 * Vráti dátum (pondelok) pre zadané ISO týždeň a rok.
 * (BEZO ZMENY)
 */
function getDateOfISOWeek(w, y) {
    const jan4 = new Date(Date.UTC(y, 0, 4)); 
    const jan4Day = (jan4.getUTCDay() + 6) % 7; 
    const mondayOfW1 = new Date(jan4.valueOf() - jan4Day * 86400000);
    return new Date(mondayOfW1.valueOf() + (w - 1) * 7 * 86400000);
}


// --- GENEROVANIE PDF (jsPDF-AutoTable) ---

/**
 * (UPRAVENÉ) Vygeneruje a zobrazí náhľad PDF rozpisu
 * Pridaný stĺpec Kontakt, veľkosť písma a okraje vrátené na pôvodné.
 */
async function showSchedulePreview() {
    if (Object.keys(appState.dutyAssignments).length === 0) {
        zobrazOznamenie("Priraďte aspoň jedného zamestnanca alebo skupinu do kalendára.", TOAST_ERROR);
        return;
    }
    
    zobrazOznamenie('Generujem náhľad PDF...', TOAST_INFO);

    if (!customFontBase64) {
        zobrazOznamenie('Načítavam dáta fontu...', TOAST_INFO);
        await loadFontData(); 
        if (!customFontBase64) {
            zobrazOznamenie('Chyba: Font sa nepodarilo načítať. PDF nemusí mať správnu diakritiku.', TOAST_ERROR);
        }
    }

    try {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('p', 'mm', 'a4');
        
        if (customFontBase64) {
            doc.addFileToVFS(PDF_FONT_FILENAME, customFontBase64);
            doc.addFont(PDF_FONT_FILENAME, PDF_FONT_INTERNAL_NAME, 'normal');
            doc.setFont(PDF_FONT_INTERNAL_NAME, 'normal'); 
        }

        const selectedMonth = appState.selectedMonth;
        const selectedYear = appState.selectedYear;
        const monthName = monthNames[selectedMonth];
        
        doc.setFontSize(14);
        doc.text(`Rozpis pohotovosti zamestnancov odboru krízového riadenia`, doc.internal.pageSize.getWidth() / 2, 15, { align: 'center' });
        
        const pageCenter = doc.internal.pageSize.getWidth() / 2;
        const yPos = 22; 
        doc.setFontSize(14);

        const text1 = `Okresného úradu Banská Bystrica - `;
        const text2 = `${monthName} ${selectedYear}`;
        doc.setTextColor(0, 0, 0); 
        const text1Width = doc.getTextWidth(text1);
        doc.setTextColor(255, 0, 0); 
        const text2Width = doc.getTextWidth(text2);
        const totalWidth = text1Width + text2Width;
        const startX1 = pageCenter - (totalWidth / 2); 
        const startX2 = startX1 + text1Width; 

        doc.setTextColor(0, 0, 0); 
        doc.text(text1, startX1, yPos);
        doc.setTextColor(255, 0, 0); 
        doc.text(text2, startX2, yPos);
        doc.setTextColor(0, 0, 0); 
        
        // --- ÚPRAVA 1: Pridaný stĺpec 'Kontakt' ---
        const tableHead = [['Dátum', 'Deň', 'Meno a priezvisko', 'Kontakt', 'Σ', 'Poznámka']];
        const tableBody = [];
        
        const lastDayOfMonth = new Date(selectedYear, selectedMonth + 1, 0);
        let currentDate = new Date(selectedYear, selectedMonth, 1);

        const lightGray = [230, 230, 230];
        
        while (currentDate.getMonth() === selectedMonth) {
            const weekInfo = getWeekNumber(currentDate);
            const reportingKey = `${weekInfo.year}-${weekInfo.week}`; // Použije správny ISO rok
            
            const activeGroup = appState.dutyAssignments[reportingKey] || [];
            
            const formattedDate = `${currentDate.getDate()}.${currentDate.getMonth() + 1}.${currentDate.getFullYear()}`;
            const dayName = dayNames[(currentDate.getDay() + 6) % 7];
            
            const reportersForWeek = appState.reporting[reportingKey] || [];
            const overridesForWeek = appState.serviceOverrides[reportingKey];

            activeGroup.forEach((person, index) => {
                let employeeToRender = person;
                let originalId = person.id;
                
                if (overridesForWeek && overridesForWeek[originalId]) {
                    const newEmployee = findEmployeeById(overridesForWeek[originalId]);
                    if (newEmployee) {
                        employeeToRender = newEmployee;
                    }
                }
                
                // --- ÚPRAVA 2: Rozdelenie mena a telefónu ---
                const employeeName = employeeToRender.meno;
                const employeePhone = employeeToRender.telefon || '';
                
                const poznamka = (reportersForWeek.includes(employeeToRender.id))
                                 ? 'hlásenia' 
                                 : '';

                tableBody.push([
                    index === 0 ? formattedDate : '',
                    index === 0 ? dayName : '',
                    employeeName, 
                    employeePhone, // Nový stĺpec
                    '',
                    poznamka 
                ]);
            });

            const nextSunday = getNextSunday(currentDate);
            
            if (nextSunday.getMonth() === selectedMonth) {
                const daysInCycle = Math.floor((nextSunday - currentDate) / (24 * 60 * 60 * 1000)) + 1;
                
                tableBody.push([
                    { content: `${nextSunday.getDate()}.${nextSunday.getMonth() + 1}.${nextSunday.getFullYear()}`, styles: { fontStyle: 'bold', fillColor: lightGray } },
                    { content: dayNames[6], styles: { fontStyle: 'bold', fillColor: lightGray } }, 
                    { content: '', styles: { fillColor: lightGray } },
                    { content: '', styles: { fillColor: lightGray } }, // Nová prázdna bunka pre Kontakt
                    { content: daysInCycle.toString(), styles: { fontStyle: 'bold', halign: 'center', fillColor: lightGray } },
                    { content: '', styles: { fillColor: lightGray } }
                ]);
                currentDate = new Date(nextSunday);
                currentDate.setDate(currentDate.getDate() + 1);
            } else {
                const daysInCycle = Math.floor((lastDayOfMonth - currentDate) / (24 * 60 * 60 * 1000)) + 1;
                const endDayOfWeek = (lastDayOfMonth.getDay() + 6) % 7;
                const endDayName = dayNames[endDayOfWeek];
                const isSunday = (endDayOfWeek === 6);
                
                const cellStyles = { fontStyle: 'bold' };
                if (isSunday) {
                    cellStyles.fillColor = lightGray;
                }
                const emptyCellStyles = isSunday ? { fillColor: lightGray } : {};

                tableBody.push([
                    { content: `${lastDayOfMonth.getDate()}.${lastDayOfMonth.getMonth() + 1}.${lastDayOfMonth.getFullYear()}`, styles: { ...cellStyles } },
                    { content: endDayName, styles: { ...cellStyles } },
                    { content: '', styles: emptyCellStyles },
                    { content: '', styles: emptyCellStyles }, // Nová prázdna bunka pre Kontakt
                    { content: daysInCycle.toString(), styles: { ...cellStyles, halign: 'center' } },
                    { content: '', styles: emptyCellStyles }
                ]);
                break; 
            }
        }

        doc.autoTable({
            head: tableHead,
            body: tableBody,
            startY: 30, 
            theme: 'grid',
            // --- ÚPRAVA 3: Vrátenie pôvodnej veľkosti písma hlavičky ---
            headStyles: {
                fillColor: [230, 230, 230],
                textColor: [33, 33, 33],
                fontStyle: 'bold',
                fontSize: 10 // Vrátené na 10
            },
            // --- ÚPRAVA 4: Vrátenie pôvodnej veľkosti písma a odstránenie cellPadding ---
            styles: {
                font: customFontBase64 ? PDF_FONT_INTERNAL_NAME : 'helvetica',
                fontSize: 10, // Vrátené na 10
                // cellPadding: 1 // Odstránené
            },
            // --- ÚPRAVA 5: Posun stĺpca 'Σ' ---
            columnStyles: {
                4: { halign: 'center' } // Posunuté z 3 na 4
            }
        });

        const finalY = doc.lastAutoTable.finalY;
        doc.setFontSize(10);
        
        // --- ÚPRAVA 6: Vrátenie pôvodných medzier pre podpisy ---
        doc.text('Zodpovedá:', 40, finalY + 20); // Vrátené na +20
        doc.text(employeeSignatures.zodpoveda.meno, 40, finalY + 27); // Vrátené na +27
        doc.text(employeeSignatures.zodpoveda.funkcia, 40, finalY + 34); // Vrátené na +34
        
        doc.text('Schvaľuje:', doc.internal.pageSize.getWidth() - 80, finalY + 20); // Vrátené na +20
        doc.text(employeeSignatures.schvaluje.meno, doc.internal.pageSize.getWidth() - 80, finalY + 27); // Vrátené na +27
        doc.text(employeeSignatures.schvaluje.funkcia, doc.internal.pageSize.getWidth() - 80, finalY + 34); // Vrátené na +34

        if (generatedPDFDataURI && generatedPDFDataURI.startsWith('blob:')) {
            URL.revokeObjectURL(generatedPDFDataURI);
        }
        
        generatedPDFDataURI = doc.output('bloburl');
        generatedPDFFilename = `rozpis_pohotovosti_OKR_OUBB_${monthName}_${selectedYear}.pdf`;

        elPdfPreviewFrame.src = generatedPDFDataURI + '#toolbar=0&navpanes=0';
        elPreviewModal.style.display = 'block';

    } catch (err) {
        console.error("Chyba pri generovaní PDF:", err);
        zobrazOznamenie("Nepodarilo sa vygenerovať PDF náhľad.", TOAST_ERROR);
    }
}

/**
 * Pomocná funkcia pre PDF - nájde nasledujúcu nedeľu
 * (BEZO ZMENY)
 */
function getNextSunday(date) {
    const result = new Date(date);
    const dayOfWeek = (date.getDay() + 6) % 7;
    if (dayOfWeek === 6) return result;
    const daysUntilSunday = 6 - dayOfWeek;
    result.setDate(result.getDate() + daysUntilSunday);
    return result;
}

// --- MODÁLNE OKNO A SŤAHOVANIE ---
// (BEZO ZMENY)

function downloadSchedulePDF() {
    if (generatedPDFDataURI) {
        const a = document.createElement('a');
        a.href = generatedPDFDataURI; 
        a.download = generatedPDFFilename; 
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        zobrazOznamenie('Súbor sa sťahuje...', TOAST_SUCCESS);
    } else {
        zobrazOznamenie('Chyba sťahovania. Skúste náhľad vygenerovať znova.', TOAST_ERROR);
    }
}

function closeModal() {
    elPreviewModal.style.display = 'none';
    elPdfPreviewFrame.src = 'about:blank';
    
    if (generatedPDFDataURI && generatedPDFDataURI.startsWith('blob:')) {
        URL.revokeObjectURL(generatedPDFDataURI);
    }
    generatedPDFDataURI = null; 
    generatedPDFFilename = '';
}

// --- GENEROVANIE DOCX (Refaktorované pre appState) ---

/**
 * Vygeneruje DOCX výkaz pohotovosti
 * (BEZO ZMENY)
 */
async function generateDocxReport() {
    
    const allPeopleInGroups = Object.values(appState.dutyAssignments).flat();
    
    const uniquePeopleIds = new Set(allPeopleInGroups.map(p => p.id));
    const allPeople = Array.from(uniquePeopleIds)
                           .map(id => allPeopleInGroups.find(p => p.id === id))
                           .slice(0, 10);
    
    if (allPeople.length === 0) {
        zobrazOznamenie("Priraďte aspoň jedného zamestnanca alebo skupinu do kalendára.", TOAST_ERROR);
        return;
    }
    
    try {
        zobrazOznamenie('Načítavam šablónu...', TOAST_INFO);
        const response = await fetch(DOCX_TEMPLATE_URL);
        if (!response.ok) {
            throw new Error('Nepodarilo sa načítať šablónu files/vykaz_pohotovosti.docx');
        }
        const arrayBuffer = await response.arrayBuffer();
        
        const pizZipInstance = new PizZip(arrayBuffer);
        const doc = new window.docxtemplater(pizZipInstance, {
            paragraphLoop: true,
            linebreaks: true,
            nullGetter: () => "", 
            delimiters: { start: '{{', end: '}}' }
        });

        const selectedMonth = appState.selectedMonth;
        const selectedYear = appState.selectedYear;
        
        const templateData = {};
        templateData['mesiac'] = monthNames[selectedMonth];
        templateData['rok'] = selectedYear;

        const employeeRows = {};
        const daysInMonth = new Date(selectedYear, selectedMonth + 1, 0).getDate();

        allPeople.forEach(person => {
            employeeRows[person.id] = { person: person, dates: [] };
        });

        for (let day = 1; day <= daysInMonth; day++) {
            const currentDate = new Date(selectedYear, selectedMonth, day);
            
            const weekInfo = getWeekNumber(currentDate);
            const reportingKey = `${weekInfo.year}-${weekInfo.week}`; // Použije správny ISO rok
            
            const activeGroup = appState.dutyAssignments[reportingKey] || [];
            
            const formattedDate = `${day}.${selectedMonth + 1}.${selectedYear}`;
            const overridesForWeek = appState.serviceOverrides[reportingKey] || {};

            activeGroup.forEach(person => {
                let currentPersonId = person.id;
                
                if (overridesForWeek[currentPersonId]) {
                    currentPersonId = overridesForWeek[currentPersonId];
                }

                if (employeeRows[currentPersonId]) {
                    employeeRows[currentPersonId].dates.push({
                        date: formattedDate,
                        dayOfWeek: currentDate.getDay() // 0=Ne, 1=Po, ... 6=So
                    });
                }
            });
        }
        
        allPeople.forEach((person, i) => {
            if (i > 8) return; 

            templateData[i.toString()] = person.meno;
            templateData[`oc${i}`] = person.coz || "";
            
            const dates = employeeRows[person.id]?.dates || [];
            
            let sumPracovneDni = 0, sumVikendy = 0, sumHodinyP5 = 0, sumHodinySn10 = 0;
            
            templateData[`dates${i}`] = dates.map(dateObj => {
                const dayOfWeek = dateObj.dayOfWeek;
                const isPracovnyDen = dayOfWeek >= 1 && dayOfWeek <= 5;
                const isVikend = dayOfWeek === 0 || dayOfWeek === 6;
                
                if (isPracovnyDen) { sumPracovneDni++; sumHodinyP5 += 16; }
                if (isVikend) { sumVikendy++; sumHodinySn10 += 24; }
                
                return {
                    date: dateObj.date,
                    popi: isPracovnyDen ? 1 : "",
                    sonesv: isVikend ? 1 : "",
                    p5: isPracovnyDen ? 16 : "",
                    sn10: isVikend ? 24 : "",
                    oc: person.coz || ""
                };
            });
            
            templateData[`sum1${i}`] = sumPracovneDni > 0 ? sumPracovneDni : "";
            templateData[`sum2${i}`] = sumVikendy > 0 ? sumVikendy : "";
            templateData[`sum3${i}`] = sumHodinyP5 > 0 ? sumHodinyP5 : "";
            templateData[`sum4${i}`] = sumHodinySn10 > 0 ? sumHodinySn10 : "";
        });

        doc.render(templateData);
        const docBuffer = doc.getZip().generate({ type: 'blob' });
        saveAs(docBuffer, `vykaz_pohotovosti_${monthNames[selectedMonth]}_${selectedYear}.docx`);
        
        zobrazOznamenie(`Súbor bol úspešne vytvorený.`, TOAST_SUCCESS);
        
    } catch (error) {
        console.error('Chyba pri spracovaní DOCX súboru:', error);
        if (error.properties && error.properties.errors) {
            console.error('Detaily chyby:', error.properties.errors);
        }
        zobrazOznamenie('Nastala chyba pri generovaní dokumentu: ' + error.message, TOAST_ERROR);
    }
}


// --- NOTIFIKÁCIE (Toast) ---
// (BEZO ZMENY)

function zobrazOznamenie(text, typ = TOAST_SUCCESS) {
    const oznamenie = document.createElement('div');
    oznamenie.className = `oznamenie ${typ}`;
    oznamenie.textContent = text;
    document.body.appendChild(oznamenie);
    
    setTimeout(() => {
        oznamenie.classList.add('show');
    }, 10);

    setTimeout(() => {
        oznamenie.classList.remove('show');
        oznamenie.addEventListener('transitionend', () => {
            if (document.body.contains(oznamenie)) {
                document.body.removeChild(oznamenie);
            }
        });
    }, 3000);
}


// --- POMOCNÉ FUNKCIE PRE NAČÍTANIE FONTU ---

/**
 * (OPRAVENÉ) Načíta dáta fontu a skonvertuje ich na Base64
 */
async function loadFontData() {
    if (customFontBase64) return; 

    try {
        const fontUrl = FONT_URL;
        const response = await fetch(fontUrl);
        if (!response.ok) throw new Error(`HTTP chyba! status: ${response.status}`);
        
        const fontBuffer = await response.arrayBuffer();
        customFontBase64 = arrayBufferToBase64(fontBuffer);
        console.log('Vlastný font pre PDF bol úspešne načítaný.');
        
    } catch (error) {
        console.error('Chyba pri načítaní fontu pre PDF:', error);
    }
}

/**
 * (OPRAVENÉ) Konvertuje ArrayBuffer na Base64 string
 */
function arrayBufferToBase64(buffer) {
    let binary = '';
    // --- OPRAVA CHYBY ---
    // const bytes = new UintArray(buffer); // PÔVODNÝ KÓD (chybný)
    const bytes = new Uint8Array(buffer); // NOVÝ KÓD (opravený)
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
}