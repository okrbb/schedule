// Globálna premenná pre konfiguráciu zamestnancov
let employeeConfig = [];

// PRIDANÉ: Globálne premenné pre náhľad a stiahnutie PDF
let generatedPDFDataURI = null; // Toto teraz bude ObjectURL
let generatedPDFFilename = '';

/**
 * Načíta konfiguráciu zamestnancov zo súboru config.json
 */
async function loadConfig() {
    try {
        const response = await fetch('files/config.json');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        employeeConfig = await response.json();
    } catch (error) {
        console.error('Chyba pri načítaní files/config.json:', error);
        alert('Nepodarilo sa načítať konfiguráciu zamestnancov. Skontrolujte konzolu.');
    }
}

/**
 * Naplní <select> zoznamom ľudí z načítanej konfigurácie
 */
function populatePeopleSelect() {
    const listPeople = document.getElementById('listPeople');
    if (!listPeople) return;

    listPeople.innerHTML = '<option value="0"></option>'; // Vyčistiť a ponechať prázdnu možnosť

    employeeConfig.forEach(group => {
        const optgroup = document.createElement('optgroup');
        optgroup.label = group.skupina;

        group.moznosti.forEach(person => {
            const option = document.createElement('option');
            option.value = person.id;
            // Zostaví text presne ako bol v pôvodnom HTML (Meno, Telefón)
            option.textContent = person.telefon ? `${person.meno}, ${person.telefon}` : person.meno;
            optgroup.appendChild(option);
        });
        listPeople.appendChild(optgroup);
    });
}

document.addEventListener('DOMContentLoaded', async function() {
    // Najprv načítať konfiguráciu, potom inicializovať zvyšok
    await loadConfig();
    populatePeopleSelect(); // Naplní zoznam ľudí

    function populateYearSelect() {
        const yearSelect = document.getElementById('yearSelect');
        const currentYear = new Date().getFullYear();
        for (let year = currentYear - 10; year <= currentYear + 10; year++) {
            const option = document.createElement('option');
            option.value = year;
            option.textContent = year;
            yearSelect.appendChild(option);
        }
        yearSelect.value = currentYear; // Set current year as default
    }

    function showSchedule() {
        // ... (kód pre showSchedule zostáva nezmenený) ...
        // Získanie vybraného dátumu
        const selectedMonth = document.getElementById('monthSelect').value;
        const selectedYear = document.getElementById('yearSelect').value;
        const currentDate = new Date(selectedYear, selectedMonth, 1);
        const currentMonth = currentDate.getMonth();
        const currentYear = currentDate.getFullYear();

        // Názvy mesiacov v slovenčine
        const monthNames = [
            "Január", "Február", "Marec", "Apríl", "Máj", "Jún", "Júl", "August", "September", "Október", "November", "December"
        ];

        // Názvy dní v slovenčine (začínajúc pondelkom)
        const dayNames = ["Pondelok", "Utorok", "Streda", "Štvrtok", "Piatok", "Sobota", "Nedeľa"];

        // Funkcia na získanie čísla týždňa
        function getWeekNumber(date) {
            const startOfYear = new Date(date.getFullYear(), 0, 1);
            const pastDaysOfYear = (date - startOfYear) / 86400000;
            return Math.ceil((pastDaysOfYear + startOfYear.getDay() + 1) / 7);
        }

        // Zobrazenie informácie o aktuálnom mesiaci
        const firstDay = new Date(currentYear, currentMonth, 1);
        const firstDayOfWeek = firstDay.getDay();
        const currentWeek = getWeekNumber(currentDate);

        let htmlContent = `<table>
            <thead>
                <tr>
                    <th class="week-column">Týždeň</th>
                    <th class="day-column">Deň</th>
                    <th class="date-column">Dátum</th>
                </tr>
            </thead>
            <tbody>`;

        // Zobrazenie posledného dňa predchádzajúceho mesiaca
        const lastDayPrevMonth = new Date(currentYear, currentMonth, 0);
        const lastDayOfWeekPrevMonth = (lastDayPrevMonth.getDay() + 6) % 7;
        const weekNumberPrevMonth = getWeekNumber(lastDayPrevMonth);

        htmlContent += `<tr class="${lastDayOfWeekPrevMonth === 0 ? 'new-week' : ''}">
            <td class="week-column">${lastDayOfWeekPrevMonth === 0 ? weekNumberPrevMonth : ''}</td>
            <td class="day-column">${dayNames[lastDayOfWeekPrevMonth]}</td>
            <td class="date-column">${lastDayPrevMonth.getDate()}. ${monthNames[lastDayPrevMonth.getMonth()]} ${lastDayPrevMonth.getFullYear()}</td>
            <td class="duty-column" id="duty-${lastDayPrevMonth.getDate()}-${lastDayPrevMonth.getMonth()}-${lastDayPrevMonth.getFullYear()}"></td>
        </tr>`;

        // Zobrazenie dní aktuálneho mesiaca
        for (let day = 1; day <= 31; day++) {
            const date = new Date(currentYear, currentMonth, day);
            if (date.getMonth() !== currentMonth) break;

            const dayOfWeek = (date.getDay() + 6) % 7; // Posunúť dni tak, aby pondelok bol prvý deň
            const weekNumber = getWeekNumber(date);

            htmlContent += `<tr class="${dayOfWeek === 0 ? 'new-week' : ''}">
                <td class="week-column">${dayOfWeek === 0 ? weekNumber : ''}</td>
                <td class="day-column">${dayNames[dayOfWeek]}</td>
                <td class="date-column">${date.getDate()}. ${monthNames[currentMonth]} ${currentYear}</td>
                <td class="duty-column" id="duty-${date.getDate()}-${currentMonth}-${currentYear}"></td>
            </tr>`;
        }

        // Zobrazenie prvého dňa nasledujúceho mesiaca
        const firstDayNextMonth = new Date(currentYear, currentMonth + 1, 1);
        const firstDayOfWeekNextMonth = (firstDayNextMonth.getDay() + 6) % 7;
        const weekNumberNextMonth = getWeekNumber(firstDayNextMonth);

        htmlContent += `<tr class="${firstDayOfWeekNextMonth === 0 ? 'new-week' : ''}">
            <td class="week-column">${firstDayOfWeekNextMonth === 0 ? weekNumberNextMonth : ''}</td>
            <td class="day-column">${dayNames[firstDayOfWeekNextMonth]}</td>
            <td class="date-column">${firstDayNextMonth.getDate()}. ${monthNames[firstDayNextMonth.getMonth()]} ${firstDayNextMonth.getFullYear()}</td>
            <td class="duty-column" id="duty-${firstDayNextMonth.getDate()}-${firstDayNextMonth.getMonth()}-${firstDayNextMonth.getFullYear()}"></td>
        </tr>`;

        document.getElementById('currentMonthInfo').innerHTML = htmlContent;
    }

    function showSelectedPerson() {
        // ... (kód pre showSelectedPerson zostáva nezmenený) ...
        const listPeople = document.getElementById('listPeople');
        const selectedValue = listPeople.value;
        const selectedPerson = listPeople.options[listPeople.selectedIndex].text;

        const peopleList1 = document.getElementById('peopleList1');
        const peopleList2 = document.getElementById('peopleList2');
        const peopleList3 = document.getElementById('peopleList3');

        let table1 = peopleList1.querySelector('.table1');
        let table2 = peopleList2.querySelector('.table2');
        let table3 = peopleList3.querySelector('.table3');

        if (!table1) {
            table1 = document.createElement('table');
            table1.classList.add('table1');
            table1.innerHTML = `<thead>
                <tr>
                    <th></th>
                </tr>
            </thead>
            <tbody></tbody>`;
            peopleList1.appendChild(table1);
        }

        if (!table2) {
            table2 = document.createElement('table');
            table2.classList.add('table2');
            table2.innerHTML = `<thead>
                <tr>
                    <th></th>
                </tr>
            </thead>
            <tbody></tbody>`;
            peopleList2.appendChild(table2);
        }

        if (!table3) {
            table3 = document.createElement('table');
            table3.classList.add('table3');
            table3.innerHTML = `<thead>
                <tr>
                    <th></th>
                </tr>
            </thead>
            <tbody></tbody>`;
            peopleList3.appendChild(table3);
        }

        const tables = [table1, table2, table3];
        const personCount = peopleList1.querySelectorAll('td').length + peopleList2.querySelectorAll('td').length + peopleList3.querySelectorAll('td').length;
        const tableIndex = Math.floor(personCount / 3);

        if (selectedValue.startsWith('group')) {
            // Nájdenie správnej skupiny v JSON dátach
            const groupData = employeeConfig.find(g => g.moznosti.some(p => p.id === selectedValue));
            if (groupData) {
                // Filtrujeme "Celá skupina" a berieme len reálnych ľudí
                const groupOptions = groupData.moznosti.filter(p => !p.id.startsWith('group'));
                
                groupOptions.forEach(option => {
                    if (tableIndex < tables.length) {
                        const tbody = tables[tableIndex].querySelector('tbody');
                        const row = document.createElement('tr');
                        const cell = document.createElement('td');
                        // Zostavíme text (Meno, Telefón)
                        cell.textContent = option.telefon ? `${option.meno}, ${option.telefon}` : option.meno;
                        cell.addEventListener('click', editPerson);
                        row.appendChild(cell);
                        tbody.appendChild(row);
                    }
                });
            }

        } else if (selectedValue !== "0") { // Ignorujeme prázdnu voľbu
            if (tableIndex < tables.length) {
                const tbody = tables[tableIndex].querySelector('tbody');
                const row = document.createElement('tr');
                const cell = document.createElement('td');
                cell.textContent = selectedPerson;
                cell.addEventListener('click', editPerson);
                row.appendChild(cell);
                tbody.appendChild(row);
            }
        }

        // Add event listeners to existing cells for editing
        [peopleList1, peopleList2, peopleList3].forEach(peopleList => {
            peopleList.querySelectorAll('td').forEach(cell => {
                cell.removeEventListener('click', editPerson); // Remove any existing listeners to avoid duplicates
                cell.addEventListener('click', editPerson);
            });
        });

        // Update the schedule table with the selected person's name
        const dateSelect = document.getElementById('dateSelect');
        if (dateSelect) {
            const selectedDate = dateSelect.value;
            const dutyCell = document.getElementById(`duty-${selectedDate}`);
            if (dutyCell) {
                dutyCell.textContent = selectedPerson;
            }
        }

        // Update the schedule table with values from table1, table2, table3
        const scheduleTable = document.getElementById('currentMonthInfo').querySelector('table');
        const scheduleRows = scheduleTable.querySelectorAll('tr');
        scheduleRows.forEach(row => {
            const dateCell = row.querySelector('.date-column');
            if (dateCell) {
                const dateText = dateCell.textContent;
                const dutyCell = row.querySelector('.duty-column');
                if (dutyCell) {
                    tables.forEach(table => {
                        const tableHeader = table.querySelector('thead tr th').textContent;
                        const employeeRows = table.querySelectorAll('tbody tr');
                        employeeRows.forEach(employeeRow => {
                            const employeeCell = employeeRow.querySelector('td');
                            if (employeeCell && employeeCell.textContent === selectedPerson) {
                                dutyCell.textContent = tableHeader;
                            }
                        });
                    });
                }
            }
        });
    }

    function editPerson() {
        const newValue = prompt('Edit person:', this.textContent);
        if (newValue !== null) {
            this.textContent = newValue;
        }
    }

    // Funkcia na generovanie a zobrazenie náhľadu PDF
    async function showSchedulePreview() {
        // 1. Skontrolujeme, či sú vybrané nejaké položky
        const peopleList1Items = document.getElementById('peopleList1').querySelectorAll('td').length;
        const peopleList2Items = document.getElementById('peopleList2').querySelectorAll('td').length;
        const peopleList3Items = document.getElementById('peopleList3').querySelectorAll('td').length;
        
        if (peopleList1Items === 0 && peopleList2Items === 0 && peopleList3Items === 0) {
            alert("Vyberte členov výjazdovej skupiny");
            return; 
        }
        
        // 2. Získanie dát (rovnaké ako v pôvodnej funkcii)
        const selectedMonth = parseInt(document.getElementById('monthSelect').value);
        const selectedYear = parseInt(document.getElementById('yearSelect').value);
        
        const monthNames = [
            "Január", "Február", "Marec", "Apríl", "Máj", "Jún", 
            "Júl", "August", "September", "Október", "November", "December"
        ];
        const dayNames = ["Pondelok", "Utorok", "Streda", "Štvrtok", "Piatok", "Sobota", "Nedeľa"];
        
        // 3. Vytvorenie HTML obsahu (rovnaké ako v pôvodnej funkcii)
        let htmlString = `
            <div style="font-family: Arial, sans-serif; margin: 20px; color: #000; background: #fff; width: 210mm; padding: 10mm;">
                <h2 style="text-align: center;">Rozpis pohotovosti zamestnancov odboru krízového riadenia Okresného úradu Banská Bystrica <span style="color: red;">${monthNames[selectedMonth]}, ${selectedYear}</span></h2>
                <table style="width: 100%; border-collapse: collapse; margin-top: 40px;">
                    <thead>
                        <tr>
                            <th style="border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f2f2f2;">Dátum</th>
                            <th style="border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f2f2f2;">Deň</th>
                            <th style="border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f2f2f2;">Meno a priezvisko</th>
                            <th style="border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f2f2f2;">Σ</th>
                            <th style="border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f2f2f2; width: 150px;">Poznámka</th> 
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        // Získanie zoznamu ľudí
        function getPeopleFromList(listId) {
            const peopleList = document.getElementById(listId);
            const people = [];
            if (peopleList) {
                const cells = peopleList.querySelectorAll('td');
                cells.forEach(cell => {
                    people.push(cell.textContent);
                });
            }
            return people;
        }
        
        const group1People = getPeopleFromList('peopleList1');
        const group2People = getPeopleFromList('peopleList2');
        const group3People = getPeopleFromList('peopleList3');
        
        function formatDate(date) {
            return `${date.getDate()}.${date.getMonth() + 1}.${date.getFullYear()}`;
        }
        
        function getDayName(date) {
            const dayOfWeek = (date.getDay() + 6) % 7; 
            return dayNames[dayOfWeek];
        }
        
        function getNextSunday(date) {
            const result = new Date(date);
            const dayOfWeek = (date.getDay() + 6) % 7;
            if (dayOfWeek === 6) return result;
            const daysUntilSunday = 6 - dayOfWeek;
            result.setDate(result.getDate() + daysUntilSunday);
            return result;
        }
        
        function addPeopleRows(date, people) {
            const dayName = getDayName(date);
            const formattedDate = formatDate(date);
            for (let i = 0; i < people.length; i++) {
                htmlString += `
                    <tr>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;">${i === 0 ? formattedDate : ''}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;">${i === 0 ? dayName : ''}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;">${people[i]}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;"></td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;"></td> 
                    </tr>
                `;
            }
        }
        
        const firstDayOfMonth = new Date(selectedYear, selectedMonth, 1);
        const lastDayOfMonth = new Date(selectedYear, selectedMonth + 1, 0);
        
        const groupSequence = [group1People, group2People, group3People];
        let currentDate = new Date(firstDayOfMonth);
        let groupIndex = 0;
        
        while (currentDate.getMonth() === selectedMonth) {
            const currentGroup = groupSequence[groupIndex % 3];
            addPeopleRows(currentDate, currentGroup);
            
            const nextSunday = getNextSunday(currentDate);
            
            if (nextSunday.getMonth() === selectedMonth) {
                // --- ZAČIATOK ZMENY: Vraciame premennú do 4. stĺpca ---
                const daysUntilSunday = Math.floor((nextSunday - currentDate) / (24 * 60 * 60 * 1000)) + 1;
                htmlString += `
                    <tr style="background-color: #f2f2f2;">
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;">${formatDate(nextSunday)}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;">Nedeľa</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;"></td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;">${daysUntilSunday}</td> 
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;"></td> 
                    </tr>
                `;
                // --- KONIEC ZMENY ---
                currentDate = new Date(nextSunday);
                currentDate.setDate(currentDate.getDate() + 1);
            } else {
                // --- ZAČIATOK ZMENY: Vraciame premennú do 4. stĺpca ---
                const daysUntilEndOfMonth = Math.floor((lastDayOfMonth - currentDate) / (24 * 60 * 60 * 1000)) + 1;
                htmlString += `
                    <tr style="background-color: #f2f2f2;">
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;">${formatDate(lastDayOfMonth)}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;">${getDayName(lastDayOfMonth)}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;"></td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;">${daysUntilEndOfMonth}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: left;"></td> 
                    </tr>
                `;
                // --- KONIEC ZMENY ---
                break;
            }
            groupIndex++;
        }
        
        htmlString += `
                    </tbody>
                </table>
                <div style="display: flex; justify-content: space-between; margin-top: 20px;">
                    <div style="margin-left: 50px;">
                        <p>Zodpovedá:</p>
                        <p>Ing. Vladimír Melikant</p>
                        <p>vedúci oddelenia COaKP</p>
                    </div>
                    <div style="margin-right: 90px;">
                        <p>Schvaľuje:</p>
                        <p>Mgr. Mário Banič</p>
                        <p>vedúci odboru</p>
                    </div>
                </div>
            </div>
        `;

        // 4. Vytvorenie skrytého kontajnera na renderovanie
        const renderContainer = document.createElement('div');
        renderContainer.style.position = 'absolute';
        renderContainer.style.left = '-9999px'; 
        renderContainer.style.top = '0';
        renderContainer.style.background = 'white'; 
        renderContainer.innerHTML = htmlString;
        document.body.appendChild(renderContainer);

        // 5. Použitie html2canvas na vykreslenie HTML a jsPDF na vytvorenie PDF
        try {
            zobrazOznamenie('Generujem náhľad PDF...', 'info');
            const canvas = await html2canvas(renderContainer, {
                scale: 2, 
                useCORS: true,
                width: renderContainer.scrollWidth,
                height: renderContainer.scrollHeight
            });
            
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF('p', 'mm', 'a4'); 
            
            const imgData = canvas.toDataURL('image/png');
            const imgProps = doc.getImageProperties(imgData);
            
            const pdfPageWidth = doc.internal.pageSize.getWidth();
            const pdfPageHeight = doc.internal.pageSize.getHeight();
            
            const scaledImgHeight = (imgProps.height * pdfPageWidth) / imgProps.width;

            if (scaledImgHeight <= pdfPageHeight) {
                doc.addImage(imgData, 'PNG', 0, 0, pdfPageWidth, scaledImgHeight);
            } else {
                const originalRatio = imgProps.width / imgProps.height;
                const finalHeight = pdfPageHeight; 
                const finalWidth = finalHeight * originalRatio; 
                const posX = (pdfPageWidth - finalWidth) / 2;
                doc.addImage(imgData, 'PNG', posX, 0, finalWidth, finalHeight);
            }
            
            // 6. Uloženie dát pre náhľad a stiahnutie
            if (generatedPDFDataURI && generatedPDFDataURI.startsWith('blob:')) {
                URL.revokeObjectURL(generatedPDFDataURI);
            }
            
            const pdfBlob = doc.output('blob');
            generatedPDFDataURI = URL.createObjectURL(pdfBlob); 

            generatedPDFFilename = `rozpis_pohotovosti_OKR_OUBB_${monthNames[selectedMonth]}_${selectedYear}.pdf`;

            // 7. Zobrazenie náhľadu v modálnom okne
            document.getElementById('pdfPreviewFrame').src = generatedPDFDataURI + '#toolbar=0';
            document.getElementById('previewModal').style.display = 'block';

        } catch (err) {
            console.error("Chyba pri generovaní PDF:", err);
            alert("Nepodarilo sa vygenerovať PDF náhľad.");
            zobrazOznamenie('Chyba pri generovaní PDF.', 'error');
        } finally {
            // 8. Odstránenie skrytého kontajnera
            document.body.removeChild(renderContainer);
        }
    }

    // Funkcia na stiahnutie PDF, volaná z modálneho okna
    function downloadSchedulePDF() {
        if (generatedPDFDataURI) {
            const a = document.createElement('a');
            a.href = generatedPDFDataURI.split('#')[0]; // Pre stiahnutie chceme čistý blob URL
            a.download = generatedPDFFilename; 
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            zobrazOznamenie('Súbor sa sťahuje...', 'success');
        } else {
            alert('Chyba: Neboli nájdené dáta na stiahnutie. Skúste náhľad vygenerovať znova.');
            zobrazOznamenie('Chyba sťahovania.', 'error');
        }
    }


    function clearSelectedEmployees() {
        const peopleLists = ['peopleList1', 'peopleList2', 'peopleList3'];
        peopleLists.forEach(listId => {
            const peopleList = document.getElementById(listId);
            if (peopleList) {
                peopleList.innerHTML = ''; 
            }
        });

        const dutyCells = document.querySelectorAll('.duty-column');
        dutyCells.forEach(cell => {
            cell.textContent = '';
        });

        const listPeople = document.getElementById('listPeople');
        if (listPeople) {
            listPeople.selectedIndex = 0; 
        }
    }

    populateYearSelect(); 
    document.getElementById('monthSelect').addEventListener('change', showSchedule);
    document.getElementById('yearSelect').addEventListener('change', showSchedule);
    document.getElementById('listPeople').addEventListener('change', showSelectedPerson);
    
    document.getElementById('exportButton').addEventListener('click', showSchedulePreview); 
    
    document.getElementById('clearButton').addEventListener('click', clearSelectedEmployees); 

    document.getElementById('closeModalButton').addEventListener('click', () => {
        document.getElementById('previewModal').style.display = 'none';
        document.getElementById('pdfPreviewFrame').src = 'about:blank'; 
        
        if (generatedPDFDataURI && generatedPDFDataURI.startsWith('blob:')) {
            // Odstránime #toolbar=0 pred revokeObjectURL, inak to nemusí fungovať
            URL.revokeObjectURL(generatedPDFDataURI.split('#')[0]);
        }

        generatedPDFDataURI = null; 
        generatedPDFFilename = '';
    });

    document.getElementById('downloadPdfButton').addEventListener('click', downloadSchedulePDF);

    showSchedule(); 
});

// --- Výkaz pohotovosti ---
// (Zvyšok súboru script.js zostáva nezmenený)

// Pomocná keš pre rýchle vyhľadávanie dát zamestnancov
let employeeDataCache = new Map();

/**
 * Nájde dáta zamestnanca v globálnej konfigurácii podľa textu zobrazeného v zozname.
 * @param {string} nameWithPhone - Text v tvare "Meno Priezvisko, Telefón"
 * @returns {object | null} - Objekt zamestnanca z config.json alebo null
 */
function findEmployeeData(nameWithPhone) {
    if (employeeDataCache.has(nameWithPhone)) {
        return employeeDataCache.get(nameWithPhone);
    }
    for (const group of employeeConfig) {
        for (const person of group.moznosti) {
            // Zrekonštruujeme text, ako je v <option>
            const personText = person.telefon ? `${person.meno}, ${person.telefon}` : person.meno;
            if (personText === nameWithPhone) {
                employeeDataCache.set(nameWithPhone, person);
                return person;
            }
        }
    }
    // Ak sa nenájde (napr. ručne upravená bunka), skúsime nájsť aspoň podľa mena
    for (const group of employeeConfig) {
        for (const person of group.moznosti) {
             if (person.meno === nameWithPhone) {
                 employeeDataCache.set(nameWithPhone, person);
                return person;
             }
        }
    }
    
    employeeDataCache.set(nameWithPhone, null); // Nenájdené
    return null;
}

/**
 * Získa čisté meno zamestnanca.
 * @param {string} fullName - Text z bunky (napr. "Ing. Vladimír Melikant, 0911580860")
 * @returns {string} - Čisté meno (napr. "Ing. Vladimír Melikant")
 */
function extractName(fullName) {
    const employee = findEmployeeData(fullName);
    if (employee) {
        return employee.meno; // Vráti čisté meno z JSON
    }
    // Fallback, ak bol zamestnanec pridaný/upravený manuálne
    if (fullName.includes(',')) {
        return fullName.split(',')[0].trim();
    }
    return fullName;
}

/**
 * Získa osobné číslo (coz) zamestnanca.
 * @param {string} fullName - Text z bunky (napr. "Ing. Vladimír Melikant, 0911580860")
 * @returns {string} - Osobné číslo (coz) alebo ""
 */
function getPersonalNumber(fullName) {
    const employee = findEmployeeData(fullName);
    return employee ? (employee.coz || "") : "";
}


document.getElementById('saveButton').addEventListener('click', async function() {
    
    // Vyčistiť keš pre prípadné manuálne úpravy
    employeeDataCache.clear();

    try {
        // 1. Načítanie šablóny .docx z adresára files/
        zobrazOznamenie('Načítavam šablónu...', 'info');
        const response = await fetch('files/vykaz_pohotovosti.docx');
        if (!response.ok) {
            throw new Error('Nepodarilo sa načítať šablónu files/vykaz_pohotovosti.docx');
        }
        const arrayBuffer = await response.arrayBuffer();
        
        // Pôvodný kód na spracovanie súboru
        const pizZipInstance = new PizZip(arrayBuffer);
        
        const content = pizZipInstance.files['word/document.xml'].asText();
        console.log('Obsah šablóny:', content.substr(0, 500)); 

        const selectedMonth = parseInt(document.getElementById('monthSelect').value);
        const selectedYear = parseInt(document.getElementById('yearSelect').value);
        const monthNames = ["Január", "Február", "Marec", "Apríl", "Máj", "Jún", "Júl", "August", "September", "Október", "November", "December"];
        
        // Získanie zoznamu zamestnancov z jednotlivých skupín
        const group1People = Array.from(document.getElementById('peopleList1').querySelectorAll('td')).map(cell => cell.textContent);
        const group2People = Array.from(document.getElementById('peopleList2').querySelectorAll('td')).map(cell => cell.textContent);
        const group3People = Array.from(document.getElementById('peopleList3').querySelectorAll('td')).map(cell => cell.textContent);
        
        // Vytvoríme mapu, ktorá bude obsahovať všetkých zamestnancov
        const allPeople = [...new Set([...group1People, ...group2People, ...group3People])].slice(0, 10);
         
        // Ak nie sú vybraní žiadni ľudia, zastavíme
        if (allPeople.length === 0) {
            zobrazOznamenie("Najprv vyberte zamestnancov do výjazdových skupín.", "error");
            return;
        }

        // Funkcia na určenie, ktorá skupina má službu v daný deň
        function getGroupForDate(date) {
            const firstDayOfMonth = new Date(selectedYear, selectedMonth, 1);
            let currentDate = new Date(firstDayOfMonth);
            let groupIndex = 0;
            
            while (currentDate <= date) {
                const nextSunday = new Date(currentDate);
                // Nájde nasledujúcu nedeľu (alebo posledný deň v mesiaci)
                const dayOfWeek = (currentDate.getDay() + 6) % 7; // 0=Po, 6=Ne
                nextSunday.setDate(currentDate.getDate() + (6 - dayOfWeek));

                // Orezanie na koniec mesiaca, ak nedeľa presahuje
                if (nextSunday.getMonth() !== selectedMonth) {
                    nextSunday.setDate(new Date(selectedYear, selectedMonth + 1, 0).getDate());
                }
                
                // Ak dátum spadá do aktuálneho týždňa služieb
                if (date >= currentDate && date <= nextSunday) {
                    return groupIndex % 3;
                }
                
                currentDate = new Date(nextSunday);
                currentDate.setDate(currentDate.getDate() + 1);
                groupIndex++;
            }
            
            return -1; // Nenašla sa zodpovedajúca skupina
        }

        // Instead of creating a row for each day of the month,
        // we'll create rows for each employee with their specific duty dates
        const employeeRows = {};
        const daysInMonth = new Date(selectedYear, selectedMonth + 1, 0).getDate();

        // Initialize rows for each employee
        allPeople.forEach(person => {
            employeeRows[person] = { dates: [] };
        });

        // Populate each employee's dates
        for (let day = 1; day <= daysInMonth; day++) {
            const currentDate = new Date(selectedYear, selectedMonth, day);
            const groupIndex = getGroupForDate(currentDate);
            
            // Vyberáme správnu skupinu zamestnancov pre daný deň
            let activeGroup = [];
            if (groupIndex === 0) activeGroup = group1People;
            else if (groupIndex === 1) activeGroup = group2People;
            else if (groupIndex === 2) activeGroup = group3People;
            
            const formattedDate = `${day}.${selectedMonth + 1}.${selectedYear}`;
            
            // Add this date to each active employee's list
            activeGroup.forEach(person => {
                if (employeeRows[person]) {
                    employeeRows[person].dates.push({
                        date: formattedDate,
                        dayOfWeek: currentDate.getDay() // 0=Ne, 1=Po, ... 6=So
                    });
                }
            });
        }

        // Upravená časť: Vytvorenie dát pre šablónu vo formáte, ktorý očakáva
        // Dôležité: upravené pre správne vloženie dátumov do placeholderov
        const templateData = {};
        
        // Pridáme všeobecné údaje o mesiaci a roku
        templateData['mesiac'] = monthNames[selectedMonth];
        templateData['rok'] = selectedYear;
        
        
        // Pridáme mená zamestnancov do príslušných placeholderov {{0}} až {{8}}
        // a ich osobné čísla
        allPeople.forEach((person, i) => {
            if (i <= 8) {
                // Použijeme nové funkcie na získanie dát z JSON
                templateData[i.toString()] = extractName(person);
                templateData[`oc${i}`] = getPersonalNumber(person);
            }
        });
        
        // Pridáme dátumy pre každého zamestnanca
        // Pre každého zamestnanca vytvoríme pole s jeho dátumami
        allPeople.forEach((person, i) => {
            if (i <= 8) {
                const dates = employeeRows[person]?.dates || [];
                const personCoz = getPersonalNumber(person); // Získame COZ raz
                
                // Vypočítame sumárne hodnoty pre daného zamestnanca
                let sumPracovneDni = 0;  // Počet pracovných dní (počet popi)
                let sumVikendy = 0;      // Počet víkendových dní (počet sonesv)
                let sumHodinyP5 = 0;     // Celkový počet hodín počas pracovných dní
                let sumHodinySn10 = 0;   // Celkový počet hodín počas víkendov
                
                // Pre placeholder {{dates0}}, {{dates1}}, atď. vložíme pole s dátumami
                // Každý objekt v poli bude mať vlastnosť date, popi, sonesv, p5 a sn10
                templateData[`dates${i}`] = dates.map(dateObj => {
                    const dayOfWeek = dateObj.dayOfWeek;
                    
                    // Určíme, či je pracovný deň (1-5 = pondelok až piatok) alebo víkend (0, 6 = nedeľa, sobota)
                    const isPracovnyDen = dayOfWeek >= 1 && dayOfWeek <= 5;
                    const isVikend = dayOfWeek === 0 || dayOfWeek === 6;
                    
                    // Pripočítame hodnoty do súčtov
                    if (isPracovnyDen) {
                        sumPracovneDni++;
                        sumHodinyP5 += 16;  // 16 hodín za každý pracovný deň
                    }
                    if (isVikend) {
                        sumVikendy++;
                        sumHodinySn10 += 24;  // 24 hodín za každý víkendový deň
                    }
                    
                    return {
                        date: dateObj.date,
                        popi: isPracovnyDen ? 1 : "",
                        sonesv: isVikend ? 1 : "",
                        p5: isPracovnyDen ? 16 : "",     // 16 hodín pre pracovné dni, prázdne pre víkendy
                        sn10: isVikend ? 24 : "",        // 24 hodín pre víkendy, prázdne pre pracovné dni
                        oc: personCoz                    // Pridáme osobné číslo k dátumu
                    };
                });
                
                // Pridáme sumárne hodnoty do dátového objektu pre daného zamestnanca
                templateData[`sum1${i}`] = sumPracovneDni > 0 ? sumPracovneDni : "";
                templateData[`sum2${i}`] = sumVikendy > 0 ? sumVikendy : "";
                templateData[`sum3${i}`] = sumHodinyP5 > 0 ? sumHodinyP5 : "";
                templateData[`sum4${i}`] = sumHodinySn10 > 0 ? sumHodinySn10 : "";
            }
        });
        
        // Pre debug účely - vypisujeme hodnoty, ktoré sa budú vkladať do šablóny
        console.log("Dáta na vloženie do šablóny:", JSON.stringify(templateData, null, 2));
        
        // Vytvorenie nových inštancií
        try {
            const doc = new window.docxtemplater(pizZipInstance, {
                paragraphLoop: true,
                linebreaks: true,
                nullGetter: () => "", 
                delimiters: { start: '{{', end: '}}' }
            });
            
            // Render dokumentu
            doc.render(templateData);
            const docBuffer = doc.getZip().generate({ type: 'blob' });
            saveAs(docBuffer, `vykaz_pohotovosti_${monthNames[selectedMonth]}_${selectedYear}.docx`);
            
            zobrazOznamenie(`Súbor vykaz_pohotovosti_${monthNames[selectedMonth]}_${selectedYear}.docx bol úspešne vytvorený`);
        } catch (error) {
            console.error('Chyba pri renderovaní šablóny:', error);
            
            if (error.properties && error.properties.errors) {
                console.error('Detaily chyby:', error.properties.errors);
                const missingTags = error.properties.errors.filter(e => e.id === 'scope_parser_unresolved_tag')
                                                          .map(e => e.properties.value)
                                                          .join(', ');
                zobrazOznamenie(`Chyba v šablóne: Nenašli sa tagy: ${missingTags}`, 'error');
            } else {
                 zobrazOznamenie('Nastala chyba pri generovaní dokumentu: ' + error.message, 'error');
            }
        }
        
    } catch (error) {
        console.error('Chyba pri spracovaní súboru:', error);
        zobrazOznamenie('Nastala chyba pri spracovaní súboru: ' + error.message, 'error');
    }
});

function zobrazOznamenie(text, typ = 'success') {
    const oznamenie = document.createElement('div');
    oznamenie.className = `oznamenie ${typ}`;
    oznamenie.textContent = text;
    
    oznamenie.style.position = 'fixed';
    oznamenie.style.bottom = '20px';
    oznamenie.style.right = '20px';
    oznamenie.style.padding = '15px';
    oznamenie.style.borderRadius = '5px';
    oznamenie.style.zIndex = '1000';
    oznamenie.style.minWidth = '250px';
    oznamenie.style.maxWidth = '80%';
    oznamenie.style.boxShadow = '0 4px 8px rgba(0,0,0,0.2)';
    
    if (typ === 'success') {
        oznamenie.style.backgroundColor = '#4CAF50';
        oznamenie.style.color = 'white';
    } else if (typ === 'error') {
        oznamenie.style.backgroundColor = '#f44336';
        oznamenie.style.color = 'white';
    } else { // 'info'
        oznamenie.style.backgroundColor = '#2196F3';
        oznamenie.style.color = 'white';
    }
    
    document.body.appendChild(oznamenie);
    
    setTimeout(() => {
        oznamenie.style.opacity = '0';
        oznamenie.style.transition = 'opacity 1s ease';
        
        setTimeout(() => {
            if (document.body.contains(oznamenie)) {
                document.body.removeChild(oznamenie);
            }
        }, 1000);
    }, 2000);
}