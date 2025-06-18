
import jsPDF from 'jspdf';
import 'jspdf-autotable'; 

// --- Global state for dirty form tracking ---
let isFormDirty = false;
let saveProgressButtonElement = null;
let efetivoWasJustClearedByButton = false; // ADDED: Tracks if last empty state was due to clear button
let hasEfetivoBeenLoadedAtLeastOnce = false; // Flag for efetivo alert visibility logic

// --- Status Options ---
const statusOptions = ["Concluído", "Em Andamento", "Iniciado", "Paralisado", "Parado", ""];

// --- Localização Options ---
const localizacaoOptions = [
    "Ampliação do Sistema de Digestão de Lodo", 
    "Canteiro de Obras", 
    "Estação Elevatória de Lodo", 
    "Isolamento Térmico dos 8 Digestores", 
    "Plano de Gestão da Qualidade", 
    "Pré Operação e Operação Assistida", 
    "Sis. de Aquec. e Purif. dos Digestores", 
    "Subst. Sist. Homogeneização de Lodo", 
    "Rede Externa",
    "" // Option to clear
];

// --- Tipo de Serviço Options ---
const tipoServicoOptions = [
    "Arquitetura", "Automação", "Canteiro", "Controle Tecnológico", "Contratação", 
    "Desmontagem", "Demolição", "Estruturas", "Equipamentos", "Fornecimento", 
    "Fundações", "Geotécnico", "Hidráulico - Sanitárias", "Hidromecânico", 
    "Impermeabilização", "Instalações Auxiliares", "Instalações Elétricas", 
    "Instrumentação", "Intalações Elétricas Prediais", "Movimento de Entulho", 
    "Movimento de Solo", "Montagem", "Obra Civil", "Revestimento", 
    "Revest. Térmico", "Viário e Urbanismo", "Treinamento", "Teste",
    "" // Option to clear
];

// --- Serviço Options ---
const servicoOptions = [
    "Aço", "Alimentação", "Alvenaria", "Andaime", "Apicoamento", "Arrasamento", "Assentamento", "Aterramento", "Aterro", 
    "Base", "Binder", "Bota Espera", "Bota Fora", "Cabeamento", "CAUQ", "Cimbramento", "Compactação", "Concreto", 
    "Contenção", "Corte Verde", "Cura Concreto", "Demolição", "Desbaste", "Desforma", "Desmobilização", "Desmontagem", 
    "Drenagem", "Ensaio", "Envelope de Concreto", "Envoltória de Areia", "Equipamento", "Escada", "Escavação", 
    "Estanqueidade", "Execução", "Esgotamento", "Fabricação", "Fechamento", "Forma", "Fornecimento", "Fresagem", 
    "Furo", "Fundação", "Geogrelha", "Grauteamento", "Impermeabilização", "Imprimação", "Infraestrutura", "Injeção", 
    "Inspeção", "Instalação", "Instalação Elétrica", "Junta", "Laje Alveolar", "Lastro", "Limpeza", "Locação", "Magro", 
    "Manutenção", "Mobilização", "Montagem", "Nivelamento", "Operação", "Passivação", "Pintura", "Posto Obra", 
    "Pulverização", "Prova de Carga", "Reaterro", "Reboco", "Recup. Estrutural", "Regularização", "Remoção", "Rufos", 
    "Regularização Mecânica", "Revest. Térmico", "Segurança", "Solda", "Sondagem", "Sub-Base", "Sub-Leito", "Suporte", 
    "Telhado", "Testes", "Topografia",
    "" // Option to clear
];


// --- Helper Functions ---
function getElementByIdSafe(id, parent = document) {
    return parent.querySelector(`#${id}`);
}

function getInputValue(id, parent = document) {
    const el = parent.querySelector(`#${id}`);
    if (el && typeof el.value === 'string') { // Check if value is a string before trimming
        return el.value.trim();
    }
    return '';
}

function getTextAreaValue(id, parent = document) {
    const el = parent.querySelector(`#${id}`);
    if (el && typeof el.value === 'string') {
        return el.value; // Keep leading/trailing spaces for textareas
    }
    return '';
}


function setInputValue(id, value, parent = document) {
    const el = parent.querySelector(`#${id}`);
    if (el) {
        el.value = value;
    }
}

function parseDateDDMMYYYY(dateStr) {
    if (!dateStr || typeof dateStr !== 'string') {
        return null;
    }
    const parts = dateStr.split('/');
    if (parts.length === 3) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1; // Month is 0-indexed
        const year = parseInt(parts[2], 10);
        if (!isNaN(day) && !isNaN(month) && !isNaN(year) && year > 1000 && year < 3000 && month >= 0 && month <= 11 && day > 0 && day <= 31) {
            const d = new Date(year, month, day);
            // Verify that the parsed date components match the input, to catch invalid dates like 31/02/YYYY
            if (d.getFullYear() === year && d.getMonth() === month && d.getDate() === day) {
                return d;
            }
        }
    }
    console.warn(`Invalid date format for parseDateDDMMYYYY: ${dateStr}. Expected DD/MM/YYYY.`);
    return null;
}

function formatDateToYYYYMMDD(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) {
        console.warn("Invalid date passed to formatDateToYYYYMMDD:", date);
        return ""; 
    }
    return date.toISOString().split('T')[0];
}

function getDayOfWeekFromDate(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) {
        console.warn("Invalid date passed to getDayOfWeekFromDate:", date);
        return ""; 
    }
    const daysOfWeek = ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"];
    return daysOfWeek[date.getDay()];
}

function calculateDaysBetween(startDate, endDate) {
    if (!(startDate instanceof Date) || isNaN(startDate.getTime()) || !(endDate instanceof Date) || isNaN(endDate.getTime())) {
        console.warn("Invalid dates passed to calculateDaysBetween:", startDate, endDate);
        return NaN; 
    }
    const MS_PER_DAY = 1000 * 60 * 60 * 24;
    // Use UTC to avoid issues with daylight saving time changes
    const utc1 = Date.UTC(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
    const utc2 = Date.UTC(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
    return Math.floor((utc2 - utc1) / MS_PER_DAY);
}


function getNextDateInfo(currentDateStr) {
    // Assuming currentDateStr is in 'YYYY-MM-DD' format
    const currentDate = new Date(currentDateStr + 'T00:00:00'); // Ensure time component is neutral for date math
    currentDate.setDate(currentDate.getDate() + 1);
    const nextDate = formatDateToYYYYMMDD(currentDate);
    const dayOfWeek = getDayOfWeekFromDate(currentDate);
    return { nextDate, dayOfWeek };
}

/**
 * Calculates the number of days in a given month and year.
 * @param {number} year - The full year (e.g., 2024).
 * @param {number} month_1_indexed - The month, 1-indexed (1 for January, 12 for December).
 * @returns {number} The number of days in the month, or NaN if the year/month is invalid.
 */
function getDaysInMonth(year, month_1_indexed) {
    if (month_1_indexed < 1 || month_1_indexed > 12 || year < 1000 || year > 3000) {
        return NaN;
    }
    return new Date(year, month_1_indexed, 0).getDate();
}

// --- Dirty Form State Management ---
function updateSaveButtonState() {
    if (!saveProgressButtonElement) { // Should be cached by DOMContentLoaded
        saveProgressButtonElement = getElementByIdSafe('saveProgressButton');
    }
    if (saveProgressButtonElement) {
        saveProgressButtonElement.disabled = !isFormDirty;
        saveProgressButtonElement.setAttribute('aria-disabled', String(!isFormDirty));
    }
}

function markFormAsDirty() {
    if (!isFormDirty) {
        isFormDirty = true;
        updateSaveButtonState();
    }
}


// --- Core RDO Day Functionality ---
let rdoDayCounter = 0; 

const shiftControlOptions = {
    tempo: { options: ["", "B", "L", "F"], ariaLabelPart: 'Condição do tempo' },
    trabalho: { options: ["", "N", "P", "T"], ariaLabelPart: 'Condição de trabalho' }
};

// Bases for global master fields that propagate from Day 0 or are date-driven.
const globalMasterFieldBases = ['contratada', 'contrato_num', 'prazo', 'inicio_obra'];
// Mappings for contract header fields from page 1 master to page 2/3 slaves within the same day.
const contractFieldMappingsBase = [
    { masterIdBase: 'contratada', slaveIdBases: ['contratada_pt', 'contratada_pt2'] },
    { masterIdBase: 'data', slaveIdBases: ['data_pt', 'data_pt2'] },
    { masterIdBase: 'contrato_num', slaveIdBases: ['contrato_num_pt', 'contratada_pt2'] }, // Typo fixed: 'contratada_pt2' should likely be 'contrato_num_pt2' if it exists or used with care
    { masterIdBase: 'dia_semana', slaveIdBases: ['dia_semana_pt', 'dia_semana_pt2'] },
    { masterIdBase: 'prazo', slaveIdBases: ['prazo_pt', 'prazo_pt2'] },
    { masterIdBase: 'inicio_obra', slaveIdBases: ['inicio_obra_pt', 'inicio_obra_pt2'] },
    { masterIdBase: 'decorridos', slaveIdBases: ['decorridos_pt', 'decorridos_pt2'] },
    { masterIdBase: 'restantes', slaveIdBases: ['restantes_pt', 'restantes_pt2'] },
];

function addActivityRow(tableId, dayIndex) {
    const tableBody = getElementByIdSafe(`${tableId}_day${dayIndex}`)?.getElementsByTagName('tbody')[0];
    if (!tableBody) return;

    const newRow = tableBody.insertRow();
    newRow.style.height = '20px'; // Consistent row height
    const tableBodyRowCount = tableBody.rows.length - 1; // 0-indexed for the new row

    const cell1 = newRow.insertCell(); // Localização
    const cell2 = newRow.insertCell(); // Tipo de Serviço
    const cell3 = newRow.insertCell(); // Serviço (Descrição)
    const cell4 = newRow.insertCell(); // Status
    const cell5 = newRow.insertCell(); // Observações

    // --- Localização Cell (cell1) ---
    const localizacaoCellIdBase = `localizacao_day${dayIndex}_${tableId.replace(/Table_?/, '')}_row${tableBodyRowCount}`;
    cell1.innerHTML = ''; 
    cell1.classList.add('status-cell'); 

    const localizacaoContainer = document.createElement('div');
    localizacaoContainer.className = 'status-select-container';

    const localizacaoDisplay = document.createElement('div');
    localizacaoDisplay.className = 'status-display'; 
    localizacaoDisplay.id = `${localizacaoCellIdBase}_display`;
    localizacaoDisplay.textContent = ''; 
    localizacaoDisplay.tabIndex = 0;
    localizacaoDisplay.setAttribute('role', 'button');
    localizacaoDisplay.setAttribute('aria-haspopup', 'listbox');
    localizacaoDisplay.setAttribute('aria-expanded', 'false');
    localizacaoDisplay.setAttribute('aria-label', `Localização da atividade, Dia ${dayIndex + 1}, Linha ${tableBodyRowCount + 1}`);

    const localizacaoHiddenInput = document.createElement('input');
    localizacaoHiddenInput.type = 'hidden';
    localizacaoHiddenInput.name = `localizacao_day${dayIndex}[]`; 
    localizacaoHiddenInput.id = `${localizacaoCellIdBase}_hidden`;

    const localizacaoDropdown = document.createElement('ul');
    localizacaoDropdown.className = 'status-dropdown'; 
    localizacaoDropdown.setAttribute('role', 'listbox');
    localizacaoDropdown.id = `${localizacaoCellIdBase}_dropdown`;
    localizacaoDropdown.style.display = 'none';

    localizacaoOptions.forEach(optionText => {
        const optionElement = document.createElement('li');
        optionElement.className = 'status-option'; 
        optionElement.textContent = optionText === "" ? "(Limpar)" : optionText;
        optionElement.dataset.value = optionText;
        optionElement.setAttribute('role', 'option');
        optionElement.setAttribute('aria-selected', 'false');

        optionElement.addEventListener('click', () => {
            const selectedValue = optionElement.dataset.value || "";
            localizacaoDisplay.textContent = selectedValue === "" ? "" : selectedValue;
            localizacaoHiddenInput.value = selectedValue;
            localizacaoHiddenInput.dispatchEvent(new Event('input', { bubbles: true })); 
            markFormAsDirty();

            localizacaoDropdown.querySelectorAll('.status-option').forEach(opt => opt.setAttribute('aria-selected', 'false'));
            optionElement.setAttribute('aria-selected', 'true');

            localizacaoDropdown.style.display = 'none';
            localizacaoDisplay.setAttribute('aria-expanded', 'false');
            localizacaoDisplay.focus(); 
        });
        localizacaoDropdown.appendChild(optionElement);
    });

    localizacaoDisplay.addEventListener('click', (event) => {
        event.stopPropagation(); 
        document.querySelectorAll('.status-dropdown').forEach(dropdown => { 
            if (dropdown.id !== localizacaoDropdown.id) {
                (dropdown).style.display = 'none';
                const otherDisplayId = dropdown.id.replace('_dropdown', '_display');
                const otherDisplay = document.getElementById(otherDisplayId);
                if (otherDisplay) otherDisplay.setAttribute('aria-expanded', 'false');
            }
        });

        const isCurrentlyOpen = localizacaoDropdown.style.display === 'block';
        localizacaoDropdown.style.display = isCurrentlyOpen ? 'none' : 'block';
        localizacaoDisplay.setAttribute('aria-expanded', String(!isCurrentlyOpen));
        if (!isCurrentlyOpen) {
            const currentVal = localizacaoHiddenInput.value;
            let foundSelected = false;
            localizacaoDropdown.querySelectorAll('.status-option').forEach(opt => {
                if (opt.dataset.value === currentVal) {
                    opt.setAttribute('aria-selected', 'true');
                    foundSelected = true;
                } else {
                    opt.setAttribute('aria-selected', 'false');
                }
            });
            if (!foundSelected) {
                 const emptyOpt = Array.from(localizacaoDropdown.querySelectorAll('.status-option')).find(o => o.dataset.value === "");
                 if(emptyOpt) emptyOpt.setAttribute('aria-selected', 'true');
            }
        }
    });
    localizacaoDisplay.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); localizacaoDisplay.click(); } 
        else if (e.key === 'Escape' && localizacaoDropdown.style.display === 'block') {
            localizacaoDropdown.style.display = 'none';
            localizacaoDisplay.setAttribute('aria-expanded', 'false');
            localizacaoDisplay.focus();
        }
    });

    localizacaoContainer.appendChild(localizacaoDisplay);
    localizacaoContainer.appendChild(localizacaoHiddenInput);
    localizacaoContainer.appendChild(localizacaoDropdown);
    cell1.appendChild(localizacaoContainer);
    localizacaoHiddenInput.addEventListener('input', markFormAsDirty);

    // --- Tipo de Serviço Cell (cell2) ---
    const tipoServicoCellIdBase = `tipo_servico_day${dayIndex}_${tableId.replace(/Table_?/, '')}_row${tableBodyRowCount}`;
    cell2.innerHTML = '';
    cell2.classList.add('status-cell');

    const tipoServicoContainer = document.createElement('div');
    tipoServicoContainer.className = 'status-select-container';

    const tipoServicoDisplay = document.createElement('div');
    tipoServicoDisplay.className = 'status-display';
    tipoServicoDisplay.id = `${tipoServicoCellIdBase}_display`;
    tipoServicoDisplay.textContent = ''; // Blank by default
    tipoServicoDisplay.tabIndex = 0;
    tipoServicoDisplay.setAttribute('role', 'button');
    tipoServicoDisplay.setAttribute('aria-haspopup', 'listbox');
    tipoServicoDisplay.setAttribute('aria-expanded', 'false');
    tipoServicoDisplay.setAttribute('aria-label', `Tipo de serviço da atividade, Dia ${dayIndex + 1}, Linha ${tableBodyRowCount + 1}`);

    const tipoServicoHiddenInput = document.createElement('input');
    tipoServicoHiddenInput.type = 'hidden';
    tipoServicoHiddenInput.name = `tipo_servico_day${dayIndex}[]`;
    tipoServicoHiddenInput.id = `${tipoServicoCellIdBase}_hidden`;

    const tipoServicoDropdown = document.createElement('ul');
    tipoServicoDropdown.className = 'status-dropdown';
    tipoServicoDropdown.setAttribute('role', 'listbox');
    tipoServicoDropdown.id = `${tipoServicoCellIdBase}_dropdown`;
    tipoServicoDropdown.style.display = 'none';

    tipoServicoOptions.forEach(optionText => {
        const optionElement = document.createElement('li');
        optionElement.className = 'status-option';
        optionElement.textContent = optionText === "" ? "(Limpar)" : optionText;
        optionElement.dataset.value = optionText;
        optionElement.setAttribute('role', 'option');
        optionElement.setAttribute('aria-selected', 'false');

        optionElement.addEventListener('click', () => {
            const selectedValue = optionElement.dataset.value || "";
            tipoServicoDisplay.textContent = selectedValue === "" ? "" : selectedValue;
            tipoServicoHiddenInput.value = selectedValue;
            tipoServicoHiddenInput.dispatchEvent(new Event('input', { bubbles: true }));
            markFormAsDirty();

            tipoServicoDropdown.querySelectorAll('.status-option').forEach(opt => opt.setAttribute('aria-selected', 'false'));
            optionElement.setAttribute('aria-selected', 'true');

            tipoServicoDropdown.style.display = 'none';
            tipoServicoDisplay.setAttribute('aria-expanded', 'false');
            tipoServicoDisplay.focus();
        });
        tipoServicoDropdown.appendChild(optionElement);
    });

    tipoServicoDisplay.addEventListener('click', (event) => {
        event.stopPropagation();
        document.querySelectorAll('.status-dropdown').forEach(dropdown => {
            if (dropdown.id !== tipoServicoDropdown.id) {
                (dropdown).style.display = 'none';
                const otherDisplayId = dropdown.id.replace('_dropdown', '_display');
                const otherDisplay = document.getElementById(otherDisplayId);
                if (otherDisplay) otherDisplay.setAttribute('aria-expanded', 'false');
            }
        });

        const isCurrentlyOpen = tipoServicoDropdown.style.display === 'block';
        tipoServicoDropdown.style.display = isCurrentlyOpen ? 'none' : 'block';
        tipoServicoDisplay.setAttribute('aria-expanded', String(!isCurrentlyOpen));
        if (!isCurrentlyOpen) {
            const currentVal = tipoServicoHiddenInput.value;
            let foundSelected = false;
            tipoServicoDropdown.querySelectorAll('.status-option').forEach(opt => {
                if (opt.dataset.value === currentVal) {
                    opt.setAttribute('aria-selected', 'true');
                    foundSelected = true;
                } else {
                    opt.setAttribute('aria-selected', 'false');
                }
            });
            if (!foundSelected) {
                const emptyOpt = Array.from(tipoServicoDropdown.querySelectorAll('.status-option')).find(o => o.dataset.value === "");
                if (emptyOpt) emptyOpt.setAttribute('aria-selected', 'true');
            }
        }
    });
    tipoServicoDisplay.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); tipoServicoDisplay.click(); }
        else if (e.key === 'Escape' && tipoServicoDropdown.style.display === 'block') {
            tipoServicoDropdown.style.display = 'none';
            tipoServicoDisplay.setAttribute('aria-expanded', 'false');
            tipoServicoDisplay.focus();
        }
    });

    tipoServicoContainer.appendChild(tipoServicoDisplay);
    tipoServicoContainer.appendChild(tipoServicoHiddenInput);
    tipoServicoContainer.appendChild(tipoServicoDropdown);
    cell2.appendChild(tipoServicoContainer);
    tipoServicoHiddenInput.addEventListener('input', markFormAsDirty);

    // --- Serviço (Descrição) Cell (cell3) ---
    const servicoCellIdBase = `servico_desc_day${dayIndex}_${tableId.replace(/Table_?/, '')}_row${tableBodyRowCount}`;
    cell3.innerHTML = '';
    cell3.classList.add('status-cell');

    const servicoContainer = document.createElement('div');
    servicoContainer.className = 'status-select-container';

    const servicoDisplay = document.createElement('div');
    servicoDisplay.className = 'status-display';
    servicoDisplay.id = `${servicoCellIdBase}_display`;
    servicoDisplay.textContent = ''; // Blank by default
    servicoDisplay.tabIndex = 0;
    servicoDisplay.setAttribute('role', 'button');
    servicoDisplay.setAttribute('aria-haspopup', 'listbox');
    servicoDisplay.setAttribute('aria-expanded', 'false');
    servicoDisplay.setAttribute('aria-label', `Serviço da atividade, Dia ${dayIndex + 1}, Linha ${tableBodyRowCount + 1}`);

    const servicoHiddenInput = document.createElement('input');
    servicoHiddenInput.type = 'hidden';
    servicoHiddenInput.name = `servico_desc_day${dayIndex}[]`; // Keep existing name from textarea
    servicoHiddenInput.id = `${servicoCellIdBase}_hidden`;

    const servicoDropdown = document.createElement('ul');
    servicoDropdown.className = 'status-dropdown';
    servicoDropdown.setAttribute('role', 'listbox');
    servicoDropdown.id = `${servicoCellIdBase}_dropdown`;
    servicoDropdown.style.display = 'none';

    servicoOptions.forEach(optionText => {
        const optionElement = document.createElement('li');
        optionElement.className = 'status-option';
        optionElement.textContent = optionText === "" ? "(Limpar)" : optionText;
        optionElement.dataset.value = optionText;
        optionElement.setAttribute('role', 'option');
        optionElement.setAttribute('aria-selected', 'false');

        optionElement.addEventListener('click', () => {
            const selectedValue = optionElement.dataset.value || "";
            servicoDisplay.textContent = selectedValue === "" ? "" : selectedValue;
            servicoHiddenInput.value = selectedValue;
            servicoHiddenInput.dispatchEvent(new Event('input', { bubbles: true }));
            markFormAsDirty();

            servicoDropdown.querySelectorAll('.status-option').forEach(opt => opt.setAttribute('aria-selected', 'false'));
            optionElement.setAttribute('aria-selected', 'true');

            servicoDropdown.style.display = 'none';
            servicoDisplay.setAttribute('aria-expanded', 'false');
            servicoDisplay.focus();
        });
        servicoDropdown.appendChild(optionElement);
    });

    servicoDisplay.addEventListener('click', (event) => {
        event.stopPropagation();
        document.querySelectorAll('.status-dropdown').forEach(dropdown => {
            if (dropdown.id !== servicoDropdown.id) {
                (dropdown).style.display = 'none';
                const otherDisplayId = dropdown.id.replace('_dropdown', '_display');
                const otherDisplay = document.getElementById(otherDisplayId);
                if (otherDisplay) otherDisplay.setAttribute('aria-expanded', 'false');
            }
        });

        const isCurrentlyOpen = servicoDropdown.style.display === 'block';
        servicoDropdown.style.display = isCurrentlyOpen ? 'none' : 'block';
        servicoDisplay.setAttribute('aria-expanded', String(!isCurrentlyOpen));
        if (!isCurrentlyOpen) {
            const currentVal = servicoHiddenInput.value;
            let foundSelected = false;
            servicoDropdown.querySelectorAll('.status-option').forEach(opt => {
                if (opt.dataset.value === currentVal) {
                    opt.setAttribute('aria-selected', 'true');
                    foundSelected = true;
                } else {
                    opt.setAttribute('aria-selected', 'false');
                }
            });
            if (!foundSelected) {
                const emptyOpt = Array.from(servicoDropdown.querySelectorAll('.status-option')).find(o => o.dataset.value === "");
                if (emptyOpt) emptyOpt.setAttribute('aria-selected', 'true');
            }
        }
    });
    servicoDisplay.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); servicoDisplay.click(); }
        else if (e.key === 'Escape' && servicoDropdown.style.display === 'block') {
            servicoDropdown.style.display = 'none';
            servicoDisplay.setAttribute('aria-expanded', 'false');
            servicoDisplay.focus();
        }
    });

    servicoContainer.appendChild(servicoDisplay);
    servicoContainer.appendChild(servicoHiddenInput);
    servicoContainer.appendChild(servicoDropdown);
    cell3.appendChild(servicoContainer);
    servicoHiddenInput.addEventListener('input', markFormAsDirty);

    // --- Observações Cell (cell5) --- still a textarea
    cell5.innerHTML = `<textarea name="observacoes_atividade_day${dayIndex}[]" aria-label="Observações da atividade"></textarea>`;
    
    // --- Status Cell (cell4) ---
    const statusCellIdBase = `status_day${dayIndex}_${tableId.replace(/Table_?/, '')}_row${tableBodyRowCount}`;
    cell4.innerHTML = ''; 
    cell4.classList.add('status-cell'); 

    const statusContainer = document.createElement('div');
    statusContainer.className = 'status-select-container';

    const statusDisplay = document.createElement('div');
    statusDisplay.className = 'status-display';
    statusDisplay.id = `${statusCellIdBase}_display`;
    statusDisplay.textContent = ''; 
    statusDisplay.tabIndex = 0;
    statusDisplay.setAttribute('role', 'button');
    statusDisplay.setAttribute('aria-haspopup', 'listbox');
    statusDisplay.setAttribute('aria-expanded', 'false');
    statusDisplay.setAttribute('aria-label', `Status da atividade, Dia ${dayIndex + 1}, Linha ${tableBodyRowCount + 1}`);

    const statusHiddenInput = document.createElement('input');
    statusHiddenInput.type = 'hidden';
    statusHiddenInput.name = `status_day${dayIndex}[]`; 
    statusHiddenInput.id = `${statusCellIdBase}_hidden`;

    const statusDropdown = document.createElement('ul');
    statusDropdown.className = 'status-dropdown';
    statusDropdown.setAttribute('role', 'listbox');
    statusDropdown.id = `${statusCellIdBase}_dropdown`;
    statusDropdown.style.display = 'none';

    statusOptions.forEach(optionText => {
        const optionElement = document.createElement('li');
        optionElement.className = 'status-option';
        optionElement.textContent = optionText === "" ? "(Limpar)" : optionText;
        optionElement.dataset.value = optionText;
        optionElement.setAttribute('role', 'option');
        optionElement.setAttribute('aria-selected', 'false');


        optionElement.addEventListener('click', () => {
            const selectedValue = optionElement.dataset.value || "";
            statusDisplay.textContent = selectedValue === "" ? "" : selectedValue; 
            statusHiddenInput.value = selectedValue;
            statusHiddenInput.dispatchEvent(new Event('input', { bubbles: true })); 
            markFormAsDirty();

            statusDropdown.querySelectorAll('.status-option').forEach(opt => opt.setAttribute('aria-selected', 'false'));
            optionElement.setAttribute('aria-selected', 'true');

            statusDropdown.style.display = 'none';
            statusDisplay.setAttribute('aria-expanded', 'false');
            statusDisplay.focus(); 
        });
        statusDropdown.appendChild(optionElement);
    });

    statusDisplay.addEventListener('click', (event) => {
        event.stopPropagation(); 
        document.querySelectorAll('.status-dropdown').forEach(dropdown => {
            if (dropdown.id !== statusDropdown.id) {
                (dropdown).style.display = 'none';
                const otherDisplayId = dropdown.id.replace('_dropdown', '_display');
                const otherDisplay = document.getElementById(otherDisplayId);
                if (otherDisplay) otherDisplay.setAttribute('aria-expanded', 'false');
            }
        });

        const isCurrentlyOpen = statusDropdown.style.display === 'block';
        statusDropdown.style.display = isCurrentlyOpen ? 'none' : 'block';
        statusDisplay.setAttribute('aria-expanded', String(!isCurrentlyOpen));
        if (!isCurrentlyOpen) {
            const currentVal = statusHiddenInput.value;
            let foundSelected = false;
            statusDropdown.querySelectorAll('.status-option').forEach(opt => {
                if (opt.dataset.value === currentVal) {
                    opt.setAttribute('aria-selected', 'true');
                    foundSelected = true;
                } else {
                    opt.setAttribute('aria-selected', 'false');
                }
            });
            if (!foundSelected) { 
                 const emptyOpt = Array.from(statusDropdown.querySelectorAll('.status-option')).find(o => o.dataset.value === "");
                 if(emptyOpt) emptyOpt.setAttribute('aria-selected', 'true');
            }
        }
    });
    statusDisplay.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); statusDisplay.click(); } 
        else if (e.key === 'Escape' && statusDropdown.style.display === 'block') {
            statusDropdown.style.display = 'none';
            statusDisplay.setAttribute('aria-expanded', 'false');
            statusDisplay.focus();
        }
    });

    statusContainer.appendChild(statusDisplay);
    statusContainer.appendChild(statusHiddenInput);
    statusContainer.appendChild(statusDropdown);
    cell4.appendChild(statusContainer);
    statusHiddenInput.addEventListener('input', markFormAsDirty);

    // Add listeners to new inputs (excluding dropdowns, which are handled by their hidden inputs)
    newRow.querySelectorAll('input[type="text"], textarea').forEach(input => { // Will pick up the remaining textarea in cell5
        input.addEventListener('input', markFormAsDirty);
    });
}

const shiftControlUpdaters = new Map();


function createCombinedShiftControl(
    cellIdBase,
    turnoIdBase,
    initialTempoValue,
    initialTrabalhoValue,
    ariaTurnoSuffixBase,
    dayIndex
) {
    const cellId = `${cellIdBase}_day${dayIndex}`;
    const cell = getElementByIdSafe(cellId);
    if (!cell) return;

    const turnoId = `${turnoIdBase}_day${dayIndex}`; // e.g., t1_day0
    const ariaTurnoSuffix = `${ariaTurnoSuffixBase} Dia ${dayIndex + 1}`;

    const tempoConfig = shiftControlOptions.tempo;
    const trabalhoConfig = shiftControlOptions.trabalho;

    let currentTempoIndex = Math.max(0, tempoConfig.options.indexOf(initialTempoValue));
    let currentTrabalhoIndex = Math.max(0, trabalhoConfig.options.indexOf(initialTrabalhoValue));

    const controlWrapper = document.createElement('div');
    controlWrapper.className = 'weather-toggle-control';
    controlWrapper.setAttribute('role', 'group');
    controlWrapper.setAttribute('aria-label', `Controles de condição para ${ariaTurnoSuffix}`);

    const tempoHiddenInput = document.createElement('input');
    tempoHiddenInput.type = 'hidden';
    tempoHiddenInput.name = `tempo_${turnoId}_value`; 
    tempoHiddenInput.id = `tempo_${turnoId}_hidden`; // e.g., tempo_t1_day0_hidden
    

    const trabalhoHiddenInput = document.createElement('input');
    trabalhoHiddenInput.type = 'hidden';
    trabalhoHiddenInput.name = `trabalho_${turnoId}_value`;
    trabalhoHiddenInput.id = `trabalho_${turnoId}_hidden`; // e.g., trabalho_t1_day0_hidden
    

    const topHalf = document.createElement('div');
    topHalf.className = 'toggle-half toggle-top';
    topHalf.setAttribute('role', 'button');
    topHalf.tabIndex = 0;

    const bottomHalf = document.createElement('div');
    bottomHalf.className = 'toggle-half toggle-bottom';
    bottomHalf.setAttribute('role', 'button');
    bottomHalf.tabIndex = 0;
    
    const updateAriaLabels = () => {
        const tempoValue = tempoConfig.options[currentTempoIndex];
        const trabalhoValue = trabalhoConfig.options[currentTrabalhoIndex];
        const tempoDisplayValue = tempoValue === "" ? "(Não selecionado)" : tempoValue;
        const trabalhoDisplayValue = trabalhoValue === "" ? "(Não selecionado)" : trabalhoValue;
        topHalf.setAttribute('aria-label', `${tempoConfig.ariaLabelPart} para ${ariaTurnoSuffix}: ${tempoDisplayValue}. Clique para alterar.`);
        bottomHalf.setAttribute('aria-label', `${trabalhoConfig.ariaLabelPart} para ${ariaTurnoSuffix}: ${trabalhoDisplayValue}. Clique para alterar.`);
    };
    
    const updateVisualsAndState = (newTempoValue, newTrabalhoValue) => {
        currentTempoIndex = Math.max(0, tempoConfig.options.indexOf(newTempoValue));
        currentTrabalhoIndex = Math.max(0, trabalhoConfig.options.indexOf(newTrabalhoValue));

        const currentTempoDisplayValue = tempoConfig.options[currentTempoIndex];
        const currentTrabalhoDisplayValue = trabalhoConfig.options[currentTrabalhoIndex];

        topHalf.textContent = currentTempoDisplayValue;
        tempoHiddenInput.value = currentTempoDisplayValue;
        bottomHalf.textContent = currentTrabalhoDisplayValue;
        trabalhoHiddenInput.value = currentTrabalhoDisplayValue;
        updateAriaLabels();
    };

    shiftControlUpdaters.set(cellId, { update: updateVisualsAndState });
    updateVisualsAndState(initialTempoValue, initialTrabalhoValue); // Initialize with provided values

    topHalf.addEventListener('click', () => {
        currentTempoIndex = (currentTempoIndex + 1) % tempoConfig.options.length;
        const newTempoValue = tempoConfig.options[currentTempoIndex];
        topHalf.textContent = newTempoValue;
        tempoHiddenInput.value = newTempoValue;
        updateAriaLabels();
        markFormAsDirty();
    });
    topHalf.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); topHalf.click(); } });

    bottomHalf.addEventListener('click', () => {
        currentTrabalhoIndex = (currentTrabalhoIndex + 1) % trabalhoConfig.options.length;
        const newTrabalhoValue = trabalhoConfig.options[currentTrabalhoIndex];
        bottomHalf.textContent = newTrabalhoValue;
        trabalhoHiddenInput.value = newTrabalhoValue;
        updateAriaLabels();
        markFormAsDirty();
    });
    bottomHalf.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); bottomHalf.click(); } });

    controlWrapper.appendChild(tempoHiddenInput);
    controlWrapper.appendChild(trabalhoHiddenInput);
    controlWrapper.appendChild(topHalf);
    controlWrapper.appendChild(bottomHalf);
    
    cell.innerHTML = ''; 
    cell.appendChild(controlWrapper);
}

function applyMultiColumnLayoutIfNecessary(tableBodyElement) {
    if (!tableBodyElement) return;
    // Applies CSS class if item count exceeds a threshold (e.g., 18 items for two columns)
    if (tableBodyElement.getElementsByTagName('tr').length > 18) {
        tableBodyElement.classList.add('multi-column-items');
    }
}


function resetReportPhotoSlotState(
    photoInput, photoPreview, 
    photoPlaceholder, clearButton, 
    labelElement
) {
    if (photoPreview) {
        photoPreview.src = '#';
        photoPreview.style.display = 'none';
        photoPreview.removeAttribute('data-natural-width');
        photoPreview.removeAttribute('data-natural-height');
        photoPreview.removeAttribute('data-is-effectively-loaded');
        photoPreview.style.width = '100%';
        photoPreview.style.height = '100%';
        photoPreview.style.objectFit = 'cover';
        photoPreview.style.position = 'relative';
        photoPreview.style.transform = 'none';
        photoPreview.style.left = 'auto';
        photoPreview.style.top = 'auto';
    }
    if (photoPlaceholder) photoPlaceholder.style.display = 'block';
    if (clearButton) clearButton.style.display = 'none';
    if (labelElement) {
        labelElement.classList.remove('drag-over');
        labelElement.style.cursor = 'pointer';
    }
    if (photoInput) {
        photoInput.value = '';
        photoInput.disabled = false;
    }
}

function setReportPhotoSlotImage(
    photoInput, photoPreview, 
    photoPlaceholder, clearButton, 
    labelElement,
    imageUrl, naturalWidth, naturalHeight, file // File is optional, for new uploads
) {
    if (!photoPreview || !labelElement || !photoPlaceholder || !clearButton || !photoInput) return;

    const tempImg = new Image();
    tempImg.onload = () => {
        const nw = naturalWidth ? parseInt(naturalWidth) : tempImg.naturalWidth;
        const nh = naturalHeight ? parseInt(naturalHeight) : tempImg.naturalHeight;

        if (!nw || !nh) {
            alert("Erro ao obter dimensões da imagem.");
            resetReportPhotoSlotState(photoInput, photoPreview, photoPlaceholder, clearButton, labelElement);
            return;
        }

        photoPreview.src = imageUrl;
        photoPreview.style.display = 'block';
        photoPreview.style.width = '100%';
        photoPreview.style.height = '100%';
        photoPreview.style.objectFit = 'cover';
        photoPreview.style.position = 'relative';
        photoPreview.style.transform = 'none';
        photoPreview.style.left = 'auto';
        photoPreview.style.top = 'auto';
        photoPreview.dataset.naturalWidth = nw.toString();
        photoPreview.dataset.naturalHeight = nh.toString();
        photoPreview.dataset.isEffectivelyLoaded = 'true';


        photoPlaceholder.style.display = 'none';
        clearButton.style.display = 'block';
        
        if (file) { // If a new file was processed
            const dataTransfer = new DataTransfer();
            dataTransfer.items.add(file);
            photoInput.files = dataTransfer.files;
        } else {
             // For loaded images, we don't have the File object, so photoInput.files remains empty.
             // This is generally fine as the image is already loaded and displayed.
        }
        photoInput.disabled = true;
        labelElement.style.cursor = 'default';
    };
    tempImg.onerror = () => {
        alert("Erro ao carregar a imagem.");
        resetReportPhotoSlotState(photoInput, photoPreview, photoPlaceholder, clearButton, labelElement);
    };
    tempImg.src = imageUrl;
}


function setupReportPhotoSlot(
    inputIdBase, 
    previewIdBase, 
    placeholderIdBase, 
    clearButtonIdBase, 
    labelIdBase, 
    dayIndex,
    photoSlotNumber // For more specific error messages/logging
) {
    const daySuffix = `_day${dayIndex}`;
    const photoInput = getElementByIdSafe(`${inputIdBase}${daySuffix}`);
    const photoPreview = getElementByIdSafe(`${previewIdBase}${daySuffix}`);
    const photoPlaceholder = getElementByIdSafe(`${placeholderIdBase}${daySuffix}`);
    const clearButton = getElementByIdSafe(`${clearButtonIdBase}${daySuffix}`);
    const labelElement = getElementByIdSafe(`${labelIdBase}${daySuffix}`);

    if (!photoInput || !photoPreview || !photoPlaceholder || !clearButton || !labelElement) {
        console.warn(`Missing elements for report photo slot: ${inputIdBase}${daySuffix} (Slot ${photoSlotNumber}).`);
        return;
    }
    
    labelElement.style.cursor = 'pointer'; 
    photoInput.disabled = false; 

    const processFile = (file) => {
        if (file && file.type.startsWith('image/')) {
            const reader = new FileReader();
            reader.onload = function(e) {
                if (e.target?.result) {
                    setReportPhotoSlotImage(
                        photoInput, photoPreview, photoPlaceholder, clearButton, labelElement,
                        e.target.result, undefined, undefined, file
                    );
                    markFormAsDirty();
                }
            }
            reader.readAsDataURL(file);
        } else if (file) {
            alert("Por favor, solte apenas arquivos de imagem.");
            if (labelElement) labelElement.classList.remove('drag-over');
        }
    };

    photoInput.addEventListener('change', (event) => {
        const file = (event.target).files?.[0];
        if (file) processFile(file);
    });

    clearButton.addEventListener('click', () => {
        resetReportPhotoSlotState(photoInput, photoPreview, photoPlaceholder, clearButton, labelElement);
        markFormAsDirty();
    });

    // Drag and Drop listeners
    labelElement.addEventListener('dragover', (e) => { e.preventDefault(); e.stopPropagation(); labelElement.classList.add('drag-over'); });
    labelElement.addEventListener('dragleave', (e) => { e.preventDefault(); e.stopPropagation(); labelElement.classList.remove('drag-over'); });
    labelElement.addEventListener('drop', (e) => {
        e.preventDefault(); e.stopPropagation(); labelElement.classList.remove('drag-over');
        if (photoInput.disabled) { 
            return;
        }
        const files = e.dataTransfer?.files;
        if (files && files.length > 0) {
            processFile(files[0]); // processFile will call markFormAsDirty
        }
    });
    
    resetReportPhotoSlotState(photoInput, photoPreview, photoPlaceholder, clearButton, labelElement); // Initialize state
}


/**
 * Updates all signature preview images across all RDO days and their respective pages (Page 1, 2, 3).
 * This function is triggered when the master signature on Day 0, Page 1 (`consbemPhoto_day0`) is changed or cleared.
 * @param {string | null} imageUrl - The base64 data URL of the master signature image, or null to clear.
 */
function updateAllSignaturePreviews(imageUrl) {
    const isImageEffectivelyLoaded = imageUrl && 
                                     imageUrl !== '#' && 
                                     !imageUrl.startsWith('about:blank') &&
                                     !imageUrl.endsWith('/#'); 

    for (let d = 0; d <= rdoDayCounter; d++) {
        const daySuffix = `_day${d}`;
        const dayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        if (!dayWrapper) continue; 

        const previewIdsAndContainers = [
            { previewId: `consbemPhotoPreview${daySuffix}`, containerId: `consbemPhotoContainer${daySuffix}`, labelFor: `consbemPhoto${daySuffix}`, clearId: `clearConsbemPhoto${daySuffix}`, inputId: `consbemPhoto${daySuffix}` },
            { previewId: `consbemPhotoPreview_pt${daySuffix}`, containerId: `consbemPhotoContainer_pt${daySuffix}`, labelFor: `consbemPhoto_pt${daySuffix}`, clearId: `clearConsbemPhoto_pt${daySuffix}`, inputId: `consbemPhoto_pt${daySuffix}` },
            { previewId: `consbemPhotoPreview_pt2${daySuffix}`, containerId: `consbemPhotoContainer_pt2${daySuffix}`, labelFor: `consbemPhoto_pt2${daySuffix}`, clearId: `clearConsbemPhoto_pt2${daySuffix}`, inputId: `consbemPhoto_pt2${daySuffix}` }
        ];

        previewIdsAndContainers.forEach(pc => {
            const preview = getElementByIdSafe(pc.previewId, dayWrapper);
            const container = getElementByIdSafe(pc.containerId, dayWrapper);
            const uploadButton = dayWrapper.querySelector(`label[for="${pc.labelFor}"]`);
            const clearButton = getElementByIdSafe(pc.clearId, dayWrapper);
            const inputElement = getElementByIdSafe(pc.inputId, dayWrapper);

            if (preview && container && uploadButton && clearButton && inputElement) {
                if (isImageEffectivelyLoaded && imageUrl) { 
                    preview.src = imageUrl;
                    preview.style.display = 'block';
                    preview.dataset.isEffectivelyLoaded = 'true'; 
                    container.style.display = 'inline-block';
                    
                    uploadButton.style.display = 'none'; 
                    inputElement.disabled = true;
                    clearButton.style.display = (d === 0 && pc.inputId === `consbemPhoto${daySuffix}`) ? 'block' : 'none';
                } else { 
                    preview.src = '#'; 
                    preview.style.display = 'none';
                    preview.dataset.isEffectivelyLoaded = 'false'; 
                    container.style.display = 'none'; // Ensure container is hidden if no image
                    clearButton.style.display = 'none'; 
                    
                    if (d === 0 && pc.inputId === `consbemPhoto${daySuffix}`) {
                        uploadButton.style.display = 'inline-block'; 
                        inputElement.disabled = false;               
                    } else {
                        uploadButton.style.display = 'none'; 
                        inputElement.disabled = true;        
                    }
                }
            }
        });
    }
}

/**
 * Sets up the photo attachment functionality for the Consbem signature for a given RDO day.
 * - On Day 0: The main signature input (`consbemPhoto_day0`) is configured to be the master.
 *   Changes to it (uploading or clearing) trigger `updateAllSignaturePreviews` to propagate the signature
 *   to all other signature slots across all days and pages. Slave signature slots on Day 0 (Pages 2 & 3)
 *   are made non-interactive.
 * - On Days > 0: All signature slots (on Page 1, 2, and 3) are made non-interactive displays,
 *   as they merely reflect the master signature from Day 0.
 * @param {number} dayIndex - The index of the RDO day being set up (0-based).
 */
function setupPhotoAttachmentForDay(dayIndex) {
    const daySuffix = `_day${dayIndex}`;
    
    const masterPhotoInput = getElementByIdSafe(`consbemPhoto${daySuffix}`); 
    const masterPhotoPreview = getElementByIdSafe(`consbemPhotoPreview${daySuffix}`); 
    const masterPhotoContainer = getElementByIdSafe(`consbemPhotoContainer${daySuffix}`); 
    const masterClearPhotoButton = getElementByIdSafe(`clearConsbemPhoto${daySuffix}`); 
    const masterPhotoUploadButton = document.querySelector(`label[for="consbemPhoto${daySuffix}"]`); 

    if (dayIndex === 0) { 
        if (!masterPhotoInput || !masterPhotoPreview || !masterPhotoContainer || !masterClearPhotoButton || !masterPhotoUploadButton) {
            console.warn("Master signature elements for Day 0 not fully found. Signature functionality may be impaired.");
            return;
        }

        masterPhotoInput.addEventListener('change', function(event) {
            const file = (event.target).files?.[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = function(e) {
                    if (e.target?.result) {
                        updateAllSignaturePreviews(e.target.result);
                        markFormAsDirty();
                    }
                }
                reader.readAsDataURL(file);
            } else { 
                updateAllSignaturePreviews(null); 
                markFormAsDirty();
            }
        });

        masterClearPhotoButton.addEventListener('click', function() {
            masterPhotoInput.value = ''; 
            updateAllSignaturePreviews(null); 
            markFormAsDirty();
        });
        
        const day0Wrapper = getElementByIdSafe('rdo_day_0_wrapper');
        if (day0Wrapper) {
            ['_pt', '_pt2'].forEach(pagePrefix => { 
                const slaveUploadLabel = day0Wrapper.querySelector(`label[for="consbemPhoto${pagePrefix}${daySuffix}"]`);
                if (slaveUploadLabel) slaveUploadLabel.style.display = 'none'; 
                
                const slaveClearBtn = getElementByIdSafe(`clearConsbemPhoto${pagePrefix}${daySuffix}`, day0Wrapper);
                if (slaveClearBtn) slaveClearBtn.style.display = 'none'; 
                
                const slaveInput = getElementByIdSafe(`consbemPhoto${pagePrefix}${daySuffix}`, day0Wrapper);
                if (slaveInput) slaveInput.disabled = true; 
            });
        }

    } else { 
        ['', '_pt', '_pt2'].forEach(pagePrefix => { 
            const uploadButton = document.querySelector(`label[for="consbemPhoto${pagePrefix}${daySuffix}"]`);
            if (uploadButton) uploadButton.style.display = 'none'; 
            
            const clearButton = getElementByIdSafe(`clearConsbemPhoto${pagePrefix}${daySuffix}`);
            if (clearButton) clearButton.style.display = 'none'; 
            
            const inputElement = getElementByIdSafe(`consbemPhoto${pagePrefix}${daySuffix}`);
            if (inputElement) inputElement.disabled = true; 
        });
    }
}


function calculateSectionTotalsForDay(dayWrapperOrPageContainer, dayIndex) {
    const laborSections = [
        { columnClass: 'labor-direta', totalInputSelectorSuffix: '.total-field input.section-total-input' },
        { columnClass: 'labor-indireta', totalInputSelectorSuffix: '.total-field input.section-total-input' },
        { columnClass: 'equipamentos', totalInputSelectorSuffix: '.total-field input.section-total-input' }
    ];

    laborSections.forEach(section => {
        const columnElement = dayWrapperOrPageContainer.querySelector('.' + section.columnClass);
        if (columnElement) {
            let sectionTotal = 0;
            const quantityInputs = columnElement.querySelectorAll('.labor-table .quantity-input');
            quantityInputs.forEach(input => {
                sectionTotal += parseInt(input.value) || 0; 
            });
            const totalSectionInput = columnElement.querySelector(section.totalInputSelectorSuffix);
            if (totalSectionInput) {
                totalSectionInput.value = sectionTotal.toString();
            }
        }
    });
    checkAndToggleEfetivoAlertForDay(dayIndex);
    updateClearEfetivoButtonVisibility();
}


// --- Rainfall Data Fetching ---
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));
const REQUEST_DELAY_MS = 1000; // Increased delay between API calls for rainfall data

// Global Rate Limit State
let isGloballyRateLimited = false;
let globalRateLimitRetryAfterTimestamp = 0;
const GLOBAL_RATE_LIMIT_COOLDOWN_MS = 61000; // 61 seconds (a bit over 1 minute)



let globalFetchMessageContainer = null;
let globalSaveMessageContainer = null; // For save/load messages
let fetchMessageTimeoutId = null;
let saveProgressMessageTimeoutId = null;
const MESSAGE_AUTO_HIDE_DURATION = 7000; // 7 seconds


// Rainfall Progress Bar Elements
let rainfallProgressBarWrapper = null;
let rainfallProgressBar = null;
let rainfallProgressText = null;
let rainfallProgressDetailsText = null;


function updateFetchMessage(message, isError = false) {
    if (!globalFetchMessageContainer) {
        globalFetchMessageContainer = getElementByIdSafe('fetchRainfallMessageContainer');
    }

    if (fetchMessageTimeoutId) {
        clearTimeout(fetchMessageTimeoutId);
        fetchMessageTimeoutId = null;
    }

    if (globalFetchMessageContainer) {
        if (message && message.trim() !== "") {
            globalFetchMessageContainer.textContent = message;
            globalFetchMessageContainer.style.display = 'block';
            globalFetchMessageContainer.style.color = isError ? '#f44336' : '#212529'; // Red for errors

            if (!isError) {
                fetchMessageTimeoutId = window.setTimeout(() => {
                    updateFetchMessage(null);
                }, MESSAGE_AUTO_HIDE_DURATION);
            }
        } else {
            globalFetchMessageContainer.textContent = '';
            globalFetchMessageContainer.style.display = 'none';
        }
    }
}

function updateSaveProgressMessage(message, isError = false) {
    if (!globalSaveMessageContainer) {
        globalSaveMessageContainer = getElementByIdSafe('saveProgressMessageContainer');
    }

    if (saveProgressMessageTimeoutId) {
        clearTimeout(saveProgressMessageTimeoutId);
        saveProgressMessageTimeoutId = null;
    }

    if (globalSaveMessageContainer) {
        if (message && message.trim() !== "") {
            globalSaveMessageContainer.innerHTML = message.replace(/\n/g, '<br>'); // Allow multiline messages
            globalSaveMessageContainer.style.display = 'block';
            globalSaveMessageContainer.style.color = isError ? '#d32f2f' : '#1b5e20'; // Dark red for errors, dark green for success
            
            if (!isError) {
                saveProgressMessageTimeoutId = window.setTimeout(() => {
                    updateSaveProgressMessage(null);
                }, MESSAGE_AUTO_HIDE_DURATION);
            }
        } else {
            globalSaveMessageContainer.textContent = '';
            globalSaveMessageContainer.style.display = 'none';
        }
    }
}


function updateRainfallStatus(dayIndex, status, customMessage) {
    const statusEl = getElementByIdSafe(`indice_pluv_status_day${dayIndex}`);
    if (statusEl) {
        let messageToDisplay = customMessage;
        let tooltipText = '';

        // Determine the message to display
        if (messageToDisplay === undefined) {
            switch (status) {
                case 'loading': messageToDisplay = '...'; break;
                case 'success': messageToDisplay = '✓'; break;
                case 'error': messageToDisplay = 'X'; break;
                case 'idle': messageToDisplay = '-'; break;
                default: messageToDisplay = '-';
            }
        }
        statusEl.textContent = messageToDisplay;

        // Reset classes and add base status class
        statusEl.className = 'rainfall-status'; 
        if (status !== 'idle') { 
            statusEl.classList.add(status);
        }

        // Determine tooltip text and apply special styling for 'P' and 'RL'
        switch (messageToDisplay) {
            case 'P':
                tooltipText = "Previsão de chuva. Este valor é uma estimativa futura.";
                statusEl.classList.add('forecast'); 
                statusEl.classList.remove('success'); // Ensure 'P' is not also styled as 'success'
                break;
            case '✓':
                tooltipText = "Dados de chuva obtidos com sucesso.";
                statusEl.classList.remove('forecast'); 
                break;
            case 'X':
                tooltipText = "Erro ao buscar dados de chuva.";
                statusEl.classList.remove('forecast');
                break;
            case '-':
                tooltipText = "Dados de chuva não disponíveis ou fora do período de busca.";
                statusEl.classList.remove('forecast');
                break;
            case '...':
                tooltipText = "Buscando dados de chuva...";
                statusEl.classList.remove('forecast');
                break;
            case 'RL': 
                tooltipText = "Limite de requisições da API atingido. Tente mais tarde.";
                statusEl.classList.remove('forecast');
                statusEl.classList.add('error'); // Style 'RL' as an error for visibility
                break;
            case 'Data Inválida': 
                tooltipText = "Data inválida para busca de chuva.";
                statusEl.classList.remove('forecast');
                break;
            default:
                tooltipText = "Status dos dados pluviométricos.";
                statusEl.classList.remove('forecast');
        }
        statusEl.title = tooltipText;
    }
}

async function fetchRainfallData(dateStrYYYYMMDD, dayIndex) {
    const inputEl = getElementByIdSafe(`indice_pluv_valor_day${dayIndex}`);
    const dayWrapper = getElementByIdSafe(`rdo_day_${dayIndex}_wrapper`);
    const page1Container = dayWrapper?.querySelector('.rdo-container');


    if (!inputEl || !dayWrapper || !page1Container) {
        console.warn(`Rainfall input or day/page1 wrapper for day ${dayIndex} not found. Skipping fetch.`);
        return 'error'; 
    }

    const systemToday = new Date(); 
    systemToday.setHours(0, 0, 0, 0); 
    const requestedDateObj = new Date(dateStrYYYYMMDD + 'T00:00:00'); 

    const maxForecastDays = 15; 
    const maxForecastDate = new Date(systemToday);
    maxForecastDate.setDate(systemToday.getDate() + maxForecastDays);
    const minApiDate = new Date('2016-01-01T00:00:00'); 

    if (requestedDateObj > maxForecastDate) {
        setInputValue(`indice_pluv_valor_day${dayIndex}`, '', page1Container);
        updateRainfallStatus(dayIndex, 'idle', '-');
        return 'idle_due_to_proactive_range';
    }

    if (requestedDateObj < minApiDate) {
        setInputValue(`indice_pluv_valor_day${dayIndex}`, '', page1Container);
        updateRainfallStatus(dayIndex, 'idle', '-');
        return 'idle_due_to_proactive_range';
    }

    updateRainfallStatus(dayIndex, 'loading');

    const lat = -23.5144; // Latitude for São Paulo area
    const lon = -46.8446; // Longitude for São Paulo area
    const apiUrl = `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}&daily=precipitation_sum&timezone=America/Sao_Paulo&start_date=${dateStrYYYYMMDD}&end_date=${dateStrYYYYMMDD}`;

    try {
        const response = await fetch(apiUrl);
        if (!response.ok) {
            let apiErrorMsg = `API error: ${response.status}`;
            let isOutOfRangeErrorByApi = false;
            let isRateLimitErrorByApi = response.status === 429;

            try {
                const errorJson = await response.json();
                if (errorJson && errorJson.reason) {
                    apiErrorMsg += ` - ${errorJson.reason}`;
                    if (errorJson.reason.toLowerCase().includes("out of allowed range") || errorJson.reason.toLowerCase().includes("is out of range for daily data")) {
                        isOutOfRangeErrorByApi = true;
                    }
                }
            } catch (e) { /* ignore parsing error of error response */ }

            if (isRateLimitErrorByApi) {
                console.warn(`Rate limit error for day ${dayIndex} (${dateStrYYYYMMDD}): ${apiErrorMsg}`);
                updateRainfallStatus(dayIndex, 'error', 'RL'); 
                return 'rate_limited';
            }
            if (isOutOfRangeErrorByApi) {
                setInputValue(`indice_pluv_valor_day${dayIndex}`, '', page1Container);
                updateRainfallStatus(dayIndex, 'idle', '-');
                return 'idle_due_to_api_range';
            }
            console.error(`API error for day ${dayIndex} (${dateStrYYYYMMDD}): ${apiErrorMsg}`);
            updateRainfallStatus(dayIndex, 'error');
            return 'error';
        }
        const jsonData = await response.json();

        if (jsonData && jsonData.daily && Array.isArray(jsonData.daily.precipitation_sum) && jsonData.daily.precipitation_sum.length > 0) {
            const precipitation = jsonData.daily.precipitation_sum[0];
            if (precipitation !== null && typeof precipitation === 'number') {
                setInputValue(`indice_pluv_valor_day${dayIndex}`, precipitation.toFixed(1), page1Container);
                
                if (requestedDateObj > systemToday) {
                    updateRainfallStatus(dayIndex, 'success', 'P'); 
                } else {
                    updateRainfallStatus(dayIndex, 'success'); 
                }

                // Automate "Tempo" based on precipitation
                let newTempoValue = "";
                if (precipitation < 1) {
                    newTempoValue = "B"; // Bom
                } else if (precipitation < 20) {
                    newTempoValue = "L"; // Chuva Leve
                } else {
                    newTempoValue = "F"; // Chuva Forte
                }

                const turnosToUpdate = [
                    { turnoIdBase: 't1', cellIdBase: 'turno1_control_cell' },
                    { turnoIdBase: 't2', cellIdBase: 'turno2_control_cell' }
                    // 3rd shift (t3) is intentionally excluded from this automation
                ];

                turnosToUpdate.forEach(turnoConfig => {
                    const cellId = `${turnoConfig.cellIdBase}_day${dayIndex}`;
                    const updater = shiftControlUpdaters.get(cellId);
                    const currentTrabalhoValue = getInputValue(`trabalho_${turnoConfig.turnoIdBase}_day${dayIndex}_hidden`, page1Container);

                    if (updater) {
                        updater.update(newTempoValue, currentTrabalhoValue);
                    }
                });


                markFormAsDirty(); // Rainfall data changed and tempo automated
                return 'success';
            } else { 
                setInputValue(`indice_pluv_valor_day${dayIndex}`, '', page1Container);
                updateRainfallStatus(dayIndex, 'idle', '-'); 
                return 'idle';
            }
        } else {
            console.warn(`Precipitation data array not found or empty in API response for ${dateStrYYYYMMDD}. Response:`, jsonData);
            setInputValue(`indice_pluv_valor_day${dayIndex}`, '', page1Container);
            updateRainfallStatus(dayIndex, 'idle', '-');
            return 'idle'; 
        }
    } catch (error) {
        console.error(`Network or other error fetching rainfall data for day ${dayIndex} (${dateStrYYYYMMDD}):`, error);
        updateRainfallStatus(dayIndex, 'error');
        return 'error';
    }
}


async function fetchAllRainfallDataForAllDaysOnClick() {
    const fetchButton = getElementByIdSafe('fetchRainfallButton');
    if (!fetchButton) {
        console.error("Fetch rainfall button not found.");
        return;
    }

    if (isGloballyRateLimited && Date.now() < globalRateLimitRetryAfterTimestamp) {
        updateFetchMessage(`Limite de API atingido. Tente novamente após ${new Date(globalRateLimitRetryAfterTimestamp).toLocaleTimeString()}.`, true);
        return;
    }
    isGloballyRateLimited = false; 

    fetchButton.disabled = true;
    
    const totalDaysToFetch = rdoDayCounter + 1;
    let processedDaysCount = 0;

    if (rainfallProgressBarWrapper && rainfallProgressBar && rainfallProgressText && rainfallProgressDetailsText && totalDaysToFetch > 0) {
        rainfallProgressBar.style.width = '0%';
        rainfallProgressText.textContent = `0/${totalDaysToFetch}`;
        rainfallProgressDetailsText.textContent = 'Iniciando busca...';
        rainfallProgressBarWrapper.style.display = 'block';
        updateFetchMessage(null); // Clear the main fetch message if progress bar is shown
    } else {
        updateFetchMessage("Buscando dados... Essa ação pode demorar um pouco."); // Fallback if progress bar elements not found
    }

    let successCount = 0;
    let errorCount = 0; 
    let rateLimitHitInThisRun = false;

    for (let d = 0; d <= rdoDayCounter; d++) {
        const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_layout_wrapper_day${d}`);
        let rdoDayWrapper = null;
        if (dayWithButtonLayoutWrapper) {
            rdoDayWrapper = dayWithButtonLayoutWrapper.querySelector(`.rdo-day-wrapper`);
        } else {
            rdoDayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        }
        
        if (!rdoDayWrapper) continue;
        
        const page1Container = rdoDayWrapper.querySelector('.rdo-container');
        if (!page1Container) continue;

        const dateStr = getInputValue(`data_day${d}`, page1Container);
        if (dateStr && /^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
            try {
                const fetchStatus = await fetchRainfallData(dateStr, d);

                switch (fetchStatus) {
                    case 'success':
                        successCount++;
                        break;
                    case 'rate_limited':
                        rateLimitHitInThisRun = true;
                        isGloballyRateLimited = true;
                        globalRateLimitRetryAfterTimestamp = Date.now() + GLOBAL_RATE_LIMIT_COOLDOWN_MS;
                        errorCount++; 
                        break; 
                    case 'error':
                        errorCount++;
                        break;
                }

            } catch (e) { 
                console.error(`Unhandled exception during fetchRainfallData call for day ${d}:`, e);
                updateRainfallStatus(d, 'error'); 
                errorCount++;
            }
        } else {
            console.warn(`Invalid or missing date for day ${d}. Skipping rainfall fetch.`);
            updateRainfallStatus(d, 'idle', 'Data Inválida');
        }
        
        processedDaysCount++;
        if (rainfallProgressBarWrapper && rainfallProgressBar && rainfallProgressText && rainfallProgressDetailsText && totalDaysToFetch > 0) {
            const progressPercent = (processedDaysCount / totalDaysToFetch) * 100;
            rainfallProgressBar.style.width = `${progressPercent}%`;
            rainfallProgressText.textContent = `${processedDaysCount}/${totalDaysToFetch}`;
            
            if (rateLimitHitInThisRun) {
                 // The detail text will be updated after the loop based on this flag.
            } else if (processedDaysCount < totalDaysToFetch) {
                rainfallProgressDetailsText.textContent = `Processando dia ${processedDaysCount + 1} de ${totalDaysToFetch}...`;
            } else { 
                rainfallProgressDetailsText.textContent = 'Finalizando...';
            }
        }

        if (rateLimitHitInThisRun) {
            break; 
        }
        if (d < rdoDayCounter) { 
            await delay(REQUEST_DELAY_MS);
        }
    }

    fetchButton.disabled = false;
    let finalUserMessage = null;
    let isFinalUserMessageError = false;

    if (rateLimitHitInThisRun) {
        finalUserMessage = `Busca interrompida. Limite de API atingido. Tente o restante após ${new Date(globalRateLimitRetryAfterTimestamp).toLocaleTimeString()}.`;
        isFinalUserMessageError = true;
    } else if (errorCount > 0) {
        finalUserMessage = `Busca concluída. ${successCount} dados obtidos, ${errorCount} erros.`;
        isFinalUserMessageError = true;
    } else if (successCount > 0 || rdoDayCounter >=0 ) { 
        finalUserMessage = "Busca de dados pluviométricos concluída.";
    } else { 
        finalUserMessage = "Nenhum dado pluviométrico para buscar ou dados não disponíveis.";
    }
    updateFetchMessage(finalUserMessage, isFinalUserMessageError);


    if (rainfallProgressDetailsText) {
        if (rateLimitHitInThisRun) {
            rainfallProgressDetailsText.textContent = "Interrompido (limite API).";
        } else if (errorCount > 0) {
            rainfallProgressDetailsText.textContent = `Concluído com ${errorCount} erro(s).`;
        } else if (successCount > 0 || totalDaysToFetch > 0) {
            rainfallProgressDetailsText.textContent = "Concluído com sucesso!";
        } else { 
            rainfallProgressDetailsText.textContent = "Nenhum dado para buscar.";
        }
    }

    if (rainfallProgressBarWrapper && totalDaysToFetch > 0) {
        setTimeout(() => {
            if (rainfallProgressBarWrapper) {
                rainfallProgressBarWrapper.style.display = 'none';
            }
        }, 5000); 
    }
}

// --- Efetivo Spreadsheet Loading ---

// Global constant for words to filter during normalization and in fallback logic
const commonWordsToFilter = ['de', 'do', 'da', 'a', 'o', 'e', 'para', 'c', 's', 'geral', 'civil'];


function normalizeStringForMatching(str) {
    if (!str) return "";
    let normalized = str.toLowerCase();
    normalized = normalized.normalize("NFD").replace(/[\u0300-\u036f]/g, ""); // Decompose and remove diacritics

    // Specific fixes for corrupted characters before other pre-emptive replacements
    normalized = normalized.replace(/\bmec\s*[\ufffd]\s*nico\b/g, "mecanico"); // For mec�nico -> mecanico

    // Pre-emptive multi-word replacements for common CSV phrases
    normalized = normalized.replace(/\bop\s+retro\s+escav\b/g, "operador retroescavadeira");
    normalized = normalized.replace(/\bretro\s+escav\b/g, "retroescavadeira");
    normalized = normalized.replace(/\baux\s+serv\s+gerais\b/g, "auxiliar servico geral");
    normalized = normalized.replace(/\bserv\s+gerais\b/g, "servico geral");
    normalized = normalized.replace(/\btec\s+seg\s+trab\b/g, "tecnico seguranca trabalho");
    normalized = normalized.replace(/\bseg\s+trab\b/g, "seguranca trabalho");
    normalized = normalized.replace(/\bassist\s+adm\b/g, "assistente administrativo");
    
    // New rules based on image mapping:
    normalized = normalized.replace(/\bauxiliar\s+tec\s+eng\b/g, "auxiliar engenharia"); // AUXILIAR TEC ENG (CSV) -> Aux. de Engenharia (Site)
    normalized = normalized.replace(/\bengenheiro\s+civil\b/g, "engenheiro planejamento"); // ENGENHEIRO CIVIL (CSV) -> Eng. civil (plan.) (Site)
    normalized = normalized.replace(/\bengenheiro\s+qualidade\b/g, "engenheiro producao qualidade"); // ENGENHEIRO QUALIDADE (CSV) -> Eng. de prod. (Qual.) (Site)
    
    // Existing rules (ensure new ones are placed appropriately if there are overlaps, though these seem fine)
    normalized = normalized.replace(/\baux\s+tec\s+eng\b/g, "auxiliar tecnico engenharia"); // Keep for "AUX TEC ENG" if different from "AUXILIAR TEC ENG"
    normalized = normalized.replace(/\btec\s+eng\b/g, "tecnico engenharia");
    normalized = normalized.replace(/\bop\s+trat\s+esgoto\b/g, "operador tratamento esgoto");
    normalized = normalized.replace(/\btrat\s+esgoto\b/g, "tratamento esgoto");
    normalized = normalized.replace(/\baux\s+de\s+eng\b/g, "auxiliar engenharia"); 
    normalized = normalized.replace(/\baux\s+de\s+engenharia\b/g, "auxiliar engenharia");


    normalized = normalized.replace(/[.,\/()"'\-\ufffd]/g, " "); // Replace common punctuation & Unicode Replacement Char

    const wordMappings = {
        "enc": "encarregado", "encarr": "encarregado",
        "obras": "obra",
        "aux": "auxiliar", "auxil": "auxiliar",
        "serv": "servico", "servs": "servico", "servicos": "servico",
        "op": "operador", "operad": "operador",
        "eng": "engenheiro", "engenh": "engenheiro", "engenharia": "engenharia", // "engenharia" mapping ensures "auxiliar engenharia" becomes "auxiliar engenharia"
        "elet": "eletrico", "eletric": "eletrico",
        "mec": "mecanico", "mecanic": "mecanico", "mecanica": "mecanico",
        "tec": "tecnico", "tecn": "tecnico",
        "seg": "seguranca",
        "trab": "trabalho", "trabs": "trabalho",
        "adm": "administrativo", "admin": "administrativo",
        "prod": "producao", "produc": "producao",
        "plan": "planejamento",
        "qual": "qualidade",
        "mont": "montador", "montad": "montador",
        "est": "estrutura", "estrut": "estrutura",
        "maq": "maquina", "maqs": "maquina", "maquinas": "maquina",
        "retroescavadeira": "maquina", 
        "topog": "topografia", "topogr": "topografo",
        "cad": "cadista",
        "des": "desenhista",
        "coord": "coordenador",
        "contratos": "contrato", "contrato": "contrato",
        "contab": "contabilidade",
        "lab": "laboratorio", "laborat": "laboratorista",
        "sup": "supervisor", "superv": "supervisor",
        "aj": "ajudante", "ajud": "ajudante", "ajudante": "ajudante",
        "gerais": "geral", 
        "asg": "auxiliar servico geral",
        "gte": "gerente", "gerente": "gerente",
        "assist": "assistente",
        "i": "i", "ii": "ii", "iii": "iii", "iv": "iv", "v": "v",
        "tratamento": "tratamento",
        "esgoto": "esgoto"
    };

    let words = normalized.split(/\s+/).filter(w => w.length > 0);
    words = words.map(word => wordMappings[word] || word);
    words = words.filter(word => !commonWordsToFilter.includes(word));

    normalized = words.join(" ");
    normalized = normalized.replace(/[^a-z0-9\s]/g, ""); 
    normalized = normalized.replace(/\s+/g, " "); 
    return normalized.trim();
}


function parseCSV(csvString) {
    const lines = csvString.split(/\r\n|\n/);
    if (lines.length < 1) {
        updateSaveProgressMessage("Arquivo CSV vazio ou sem conteúdo.", true);
        return [];
    }

    const headerLine = lines[0].trim();
    if (!headerLine) {
        updateSaveProgressMessage("Cabeçalho do CSV não encontrado ou vazio.", true);
        return [];
    }

    let headers = [];
    let detectedDelimiter = '';
    
    const normalizedFuncaoVariants = ['função', 'funcao', 'fun��o', 'cargo'].map(normalizeStringForMatching);


    // Try comma first
    const commaHeadersOriginal = headerLine.split(',');
    const commaHeadersNormalized = commaHeadersOriginal.map(h => normalizeStringForMatching(h.trim()));
    const dataCommaIndex = commaHeadersNormalized.indexOf(normalizeStringForMatching('data'));
    const funcaoCommaIndex = commaHeadersNormalized.findIndex(h => normalizedFuncaoVariants.includes(h));

    if (dataCommaIndex !== -1 && funcaoCommaIndex !== -1) {
        headers = commaHeadersOriginal.map(h => h.trim()); 
        detectedDelimiter = ',';
    } else {
        // Try semicolon if comma fails
        const semicolonHeadersOriginal = headerLine.split(';');
        const semicolonHeadersNormalized = semicolonHeadersOriginal.map(h => normalizeStringForMatching(h.trim()));
        const dataSemicolonIndex = semicolonHeadersNormalized.indexOf(normalizeStringForMatching('data'));
        const funcaoSemicolonIndex = semicolonHeadersNormalized.findIndex(h => normalizedFuncaoVariants.includes(h));

        if (dataSemicolonIndex !== -1 && funcaoSemicolonIndex !== -1) {
            headers = semicolonHeadersOriginal.map(h => h.trim()); 
            detectedDelimiter = ';';
        } else {
            console.error("CSV headers 'Data' or 'Função'/'Cargo' (or variants) not found. Header line was:", headerLine);
            updateSaveProgressMessage(`Erro: Cabeçalhos 'Data' e 'Função'/'Cargo' (ou variantes normalizadas) são obrigatórios no CSV. Verifique o delimitador (',' ou ';'). Cabeçalho encontrado: ${headerLine}`, true);
            return [];
        }
    }

    const finalDataHeaderIndex = headers.map(h => normalizeStringForMatching(h)).indexOf(normalizeStringForMatching('data'));
    const finalFuncaoHeaderIndex = headers.map(h => normalizeStringForMatching(h)).findIndex(hNorm => normalizedFuncaoVariants.includes(hNorm));


    if (finalDataHeaderIndex === -1 || finalFuncaoHeaderIndex === -1) { 
        console.error("Internal error: CSV headers 'Data' or 'Função' resolved to -1 after delimiter detection. Detected Original Headers:", headers);
        updateSaveProgressMessage("Erro interno ao processar cabeçalhos do CSV.", true);
        return [];
    }
    
    const data = [];
    for (let i = 1; i < lines.length; i++) {
        const lineContent = lines[i].trim();
        if (!lineContent) continue; 

        const values = lineContent.split(detectedDelimiter); 
        
        if (values.length >= Math.max(finalDataHeaderIndex, finalFuncaoHeaderIndex) + 1) {
            const entry = {};
            entry['data'] = (values[finalDataHeaderIndex] || "").trim(); 
            entry['funcao'] = (values[finalFuncaoHeaderIndex] || "").trim();
            
            if (entry['data'] && entry['funcao']) { 
                 data.push(entry);
            }
        } else {
             console.warn(`Skipping malformed CSV line (expected at least ${Math.max(finalDataHeaderIndex, finalFuncaoHeaderIndex) + 1} values, got ${values.length} using delimiter '${detectedDelimiter}'): ${lineContent}`);
        }
    }
    return data;
}


async function populateEfetivoFromCSVData(csvData) {
    if (csvData.length === 0) {
        return;
    }

    const aggregatedData = new Map();
    const csvDatesNotFoundInRdoOrInvalid = new Set(); 
    const csvFuncoesNotFoundForProcessedDate = new Map(); 

    for (const entry of csvData) {
        const dataStrDDMMYYYY = (entry['data'] || "").trim();
        const funcaoFromCsvOriginal = (entry['funcao'] || "").trim();

        if (!dataStrDDMMYYYY || !funcaoFromCsvOriginal) {
            if (dataStrDDMMYYYY && !funcaoFromCsvOriginal) console.warn(`CSV entry for date ${dataStrDDMMYYYY} has empty função.`);
            else if (!dataStrDDMMYYYY && funcaoFromCsvOriginal) console.warn(`CSV entry for função ${funcaoFromCsvOriginal} has empty data.`);
            continue; 
        }
        
        const funcaoFromCsvNormKey = normalizeStringForMatching(funcaoFromCsvOriginal); 
        if(!funcaoFromCsvNormKey) continue; 

        const parsedDate = parseDateDDMMYYYY(dataStrDDMMYYYY);
        if (!parsedDate) {
            csvDatesNotFoundInRdoOrInvalid.add(dataStrDDMMYYYY); 
            continue;
        }
        const dateYYYYMMDD = formatDateToYYYYMMDD(parsedDate);

        if (!aggregatedData.has(dateYYYYMMDD)) {
            aggregatedData.set(dateYYYYMMDD, new Map());
        }
        const funcoesForDate = aggregatedData.get(dateYYYYMMDD);
        funcoesForDate.set(funcaoFromCsvNormKey, (funcoesForDate.get(funcaoFromCsvNormKey) || 0) + 1);
    }

    if (aggregatedData.size === 0 && csvData.length > 0) {
        let errorMsg = "Nenhuma data válida ou função encontrada no CSV para processar.";
        if (csvDatesNotFoundInRdoOrInvalid.size > 0) {
            // This specific detail might still appear if aggregatedData is empty due to all dates being invalid.
            errorMsg += ` Datas não processadas do CSV (formato inválido ou não encontradas): ${Array.from(csvDatesNotFoundInRdoOrInvalid).join(', ')}. Verifique o formato DD/MM/YYYY.`;
        }
        updateSaveProgressMessage(errorMsg, true);
        return;
    }

    let rdosEffectivelyUpdatedCount = 0;
    let totalFuncaoAssignments = 0;
    const rdoDatesProcessedThisRun = new Set(); 

    if (csvData.length > 0 || aggregatedData.size > 0) {
        hasEfetivoBeenLoadedAtLeastOnce = true;
    }


    for (let d = 0; d <= rdoDayCounter; d++) {
        let currentDayPage1Container = null;
        const dayWrapperCandidate1 = document.getElementById(`day_with_button_container_day${d}`);
        let rdoDayWrapper = null;

        if (dayWrapperCandidate1) {
            rdoDayWrapper = dayWrapperCandidate1.querySelector('.rdo-day-wrapper');
        } else {
            rdoDayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        }
        if (rdoDayWrapper) {
            currentDayPage1Container = rdoDayWrapper.querySelector('.rdo-container');
        }

        if (!currentDayPage1Container) {
            console.warn(`populateEfetivoFromCSVData: Page 1 container for RDO day index ${d} not found.`);
            continue;
        }

        const rdoDateStrYYYYMMDD = getInputValue(`data_day${d}`, currentDayPage1Container);
        if (!rdoDateStrYYYYMMDD || !aggregatedData.has(rdoDateStrYYYYMMDD)) {
            continue; 
        }
        
        csvData.forEach(csvEntry => {
            const parsedCsvDate = parseDateDDMMYYYY(csvEntry['data']);
            if (parsedCsvDate && formatDateToYYYYMMDD(parsedCsvDate) === rdoDateStrYYYYMMDD) {
                csvDatesNotFoundInRdoOrInvalid.delete(csvEntry['data']);
            }
        });
        
        const funcoesToApplyOnThisRdo = aggregatedData.get(rdoDateStrYYYYMMDD);
        if (!funcoesToApplyOnThisRdo || funcoesToApplyOnThisRdo.size === 0) continue;

        let itemsAppliedToThisRdoDay = 0;
        const laborTables = currentDayPage1Container.querySelectorAll('.labor-table');

        laborTables.forEach(table => {
            table.querySelectorAll('.quantity-input').forEach(input => input.value = '');
        });

        for (const [funcaoCsvKeyNorm, count] of funcoesToApplyOnThisRdo) {
            let foundFuncaoInRdoDay = false;
            for (const table of laborTables) {
                const rows = table.querySelectorAll('tbody tr');
                for (const row of rows) {
                    const itemNameCell = row.querySelector('.item-name-cell');
                    const quantityInput = row.querySelector('.quantity-input');

                    if (itemNameCell && quantityInput) {
                        const itemNameFromTableOriginal = (itemNameCell.textContent || "").trim();
                        const itemNameFromTableNorm = normalizeStringForMatching(itemNameFromTableOriginal);

                        if (itemNameFromTableNorm === funcaoCsvKeyNorm) { 
                            quantityInput.value = count.toString();
                            itemsAppliedToThisRdoDay++;
                            foundFuncaoInRdoDay = true;
                            break; 
                        } else { 
                            const rdoSignificantWords = itemNameFromTableNorm.split(" ").filter(w => w.length > 1 && !commonWordsToFilter.includes(w));
                            const csvSignificantWords = funcaoCsvKeyNorm.split(" ").filter(w => w.length > 1 && !commonWordsToFilter.includes(w));
                            
                            let matchedByFallback = false;
                            if (rdoSignificantWords.length > 0 && 
                                csvSignificantWords.length > 0 && 
                                rdoSignificantWords.length === csvSignificantWords.length) {
                                
                                const rdoWordSet = new Set(rdoSignificantWords);
                                const allCsvWordsFoundInRdo = csvSignificantWords.every(csvWord => rdoWordSet.has(csvWord));
                                
                                if (allCsvWordsFoundInRdo) {
                                    matchedByFallback = true;
                                }
                            }

                            if (matchedByFallback) {
                                quantityInput.value = count.toString();
                                itemsAppliedToThisRdoDay++;
                                foundFuncaoInRdoDay = true;
                                break; 
                            }
                        }
                    }
                }
                if (foundFuncaoInRdoDay) break; 
            }
            if (!foundFuncaoInRdoDay) {
                if (!csvFuncoesNotFoundForProcessedDate.has(rdoDateStrYYYYMMDD)) {
                    csvFuncoesNotFoundForProcessedDate.set(rdoDateStrYYYYMMDD, new Set());
                }
                const originalFuncaoForMessage = csvData.find(entry => {
                    const originalCsvEntryFuncao = (entry['funcao'] || "").trim();
                    return normalizeStringForMatching(originalCsvEntryFuncao) === funcaoCsvKeyNorm;
                })?.['funcao'] || funcaoCsvKeyNorm; 
                csvFuncoesNotFoundForProcessedDate.get(rdoDateStrYYYYMMDD).add(originalFuncaoForMessage);
            }
        }
        
        if (itemsAppliedToThisRdoDay > 0 || funcoesToApplyOnThisRdo.size > 0) {
            calculateSectionTotalsForDay(currentDayPage1Container, d);
            totalFuncaoAssignments += itemsAppliedToThisRdoDay; 
            if (!rdoDatesProcessedThisRun.has(rdoDateStrYYYYMMDD)) {
                rdosEffectivelyUpdatedCount++;
                rdoDatesProcessedThisRun.add(rdoDateStrYYYYMMDD);
            }
        }
    }
    
    let message = "";
    let isError = false;

    if (totalFuncaoAssignments > 0) {
        markFormAsDirty();
        message = `Efetivo carregado. ${totalFuncaoAssignments} atribuiçõe(s) de função aplicadas em ${rdosEffectivelyUpdatedCount} RDO dia(s).`;
    } else if (csvData.length > 0 && aggregatedData.size > 0 && rdosEffectivelyUpdatedCount > 0) { 
        markFormAsDirty(); 
        message = "Datas do CSV encontradas nos RDOs, mas nenhuma função do CSV correspondeu aos itens nesses RDOs. Campos de quantidade foram limpos para esses dias.";
    } else if (csvData.length > 0 && aggregatedData.size > 0 && rdosEffectivelyUpdatedCount === 0) {
        message = "Nenhuma data do CSV correspondeu às datas dos RDOs carregados.";
        isError = true;
    } else if (csvData.length > 0 && aggregatedData.size === 0 && csvDatesNotFoundInRdoOrInvalid.size === 0) {
        message = "Nenhum dado válido (Data/Função) encontrado nas linhas do CSV para processar.";
        isError = true;
    }


    const notFoundDetailsForConsole = [];
    if (csvDatesNotFoundInRdoOrInvalid.size > 0) {
        notFoundDetailsForConsole.push(`Datas do CSV não encontradas nos RDOs ou em formato inválido (DD/MM/YYYY): ${Array.from(csvDatesNotFoundInRdoOrInvalid).join(', ')}.`);
    }
    csvFuncoesNotFoundForProcessedDate.forEach((funcoesSet, dateYYYYMMDD) => {
        const originalDateForMessage = csvData.find(entry => {
            const parsed = parseDateDDMMYYYY(entry['data']);
            return parsed && formatDateToYYYYMMDD(parsed) === dateYYYYMMDD;
        })?.['data'] || dateYYYYMMDD;
        notFoundDetailsForConsole.push(`Para a data ${originalDateForMessage} (nos RDOs): Funções do CSV não encontradas: ${Array.from(funcoesSet).join(', ')}.`);
    });

    if (notFoundDetailsForConsole.length > 0) {
        console.warn("Detalhes do carregamento de efetivo (CSV):\n" + notFoundDetailsForConsole.join("\n"));
        if (totalFuncaoAssignments === 0 && !message) { 
             isError = true; 
             message = "Dados do CSV não puderam ser aplicados. Verifique se as datas e funções no CSV correspondem aos RDOs carregados. Detalhes no console.";
        }
    }
    
    // Override message for clean success case
    if (totalFuncaoAssignments > 0 && !isError && notFoundDetailsForConsole.length === 0) {
        message = "Efetivo concluído!";
    }

    if (message) {
        updateSaveProgressMessage(message, isError);
    } else if (csvData.length === 0) {
        // message handled by parseCSV or handleEfetivoFileUpload
    } else {
        if (!globalSaveMessageContainer?.textContent?.includes("Cabeçalhos")) { 
             updateSaveProgressMessage("Nenhum dado no CSV para processar ou todos os dados eram inválidos/incompletos.", true);
        }
    }
    efetivoWasJustClearedByButton = false; // Reset flag after CSV processing
    updateAllEfetivoAlerts();
    updateClearEfetivoButtonVisibility();
}


function handleEfetivoFileUpload(event) {
    updateSaveProgressMessage(null); 
    const input = event.target;
    if (!input.files || input.files.length === 0) {
        updateSaveProgressMessage("Nenhum arquivo selecionado.", true);
        return;
    }

    const file = input.files[0];
    if (file.type !== 'text/csv' && !file.name.toLowerCase().endsWith('.csv')) {
        updateSaveProgressMessage("Formato de arquivo inválido. Por favor, selecione um arquivo .csv.", true);
        input.value = ''; 
        return;
    }

    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const csvString = e.target?.result;
            if (!csvString.trim()) {
                 updateSaveProgressMessage("Arquivo CSV está vazio.", true);
                 input.value = ''; 
                 return;
            }
            const parsedData = parseCSV(csvString); 
            if (parsedData.length > 0) {
                 await populateEfetivoFromCSVData(parsedData);
            } else if (csvString.trim() && !globalSaveMessageContainer?.textContent?.includes("Cabeçalhos")) { 
                updateSaveProgressMessage("Nenhuma linha de dados válida (com Data e Função preenchidas) encontrada no CSV após o cabeçalho.", true);
            }
        } catch (error) {
            console.error("Erro ao processar arquivo CSV:", error);
            updateSaveProgressMessage(`Erro ao processar arquivo CSV: ${error.message}`, true);
        } finally {
            input.value = ''; 
        }
    };
    reader.onerror = () => {
        console.error("Erro ao ler o arquivo.");
        updateSaveProgressMessage("Erro ao ler o arquivo.", true);
        input.value = ''; 
    };
    reader.readAsText(file);
}

function triggerEfetivoFileUpload() {
    updateSaveProgressMessage(null); 
    const fileUploadElement = getElementByIdSafe('efetivoSpreadsheetUpload');
    if (fileUploadElement) {
        fileUploadElement.click();
    } else {
        console.error("Elemento de upload de arquivo de efetivo não encontrado.");
        updateSaveProgressMessage("Erro: Funcionalidade de upload não está disponível.", true);
    }
}

// --- Efetivo Clearing Logic ---
function checkIfAnyEfetivoExists() {
    for (let d = 0; d <= rdoDayCounter; d++) {
        let currentDayPage1Container = null;
        const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
        let rdoDayWrapper = null;

        if (dayWithButtonLayoutWrapper) {
            rdoDayWrapper = dayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
        } else {
            rdoDayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        }

        if (rdoDayWrapper) {
            currentDayPage1Container = rdoDayWrapper.querySelector('.rdo-container'); // Assuming first is Page 1
        }

        if (!currentDayPage1Container) {
            // console.warn(`checkIfAnyEfetivoExists: Page 1 container for RDO day index ${d} not found.`);
            continue;
        }

        const quantityInputs = currentDayPage1Container.querySelectorAll('.labor-table .quantity-input');
        for (const input of quantityInputs) {
            const value = parseInt(input.value, 10);
            if (!isNaN(value) && value > 0) {
                return true; // Efetivo found
            }
        }
    }
    return false; // No efetivo found in any RDO
}

function updateClearEfetivoButtonVisibility() {
    const clearButton = getElementByIdSafe('clearEfetivoButton');
    if (!clearButton) return;

    if (checkIfAnyEfetivoExists()) {
        clearButton.style.display = 'flex'; 
    } else {
        clearButton.style.display = 'none';
    }
}


async function clearAllEfetivo() {
    updateSaveProgressMessage(null);
    let initialCheckChangesMade = false;

    // First, check if any changes will actually be made by the clear operation.
    for (let d = 0; d <= rdoDayCounter; d++) {
        let currentDayPage1Container = null;
        const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
        let rdoDayWrapper = null;

        if (dayWithButtonLayoutWrapper) {
            rdoDayWrapper = dayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
        } else {
            rdoDayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        }
        if (rdoDayWrapper) {
            currentDayPage1Container = rdoDayWrapper.querySelector('.rdo-container');
        }

        if (currentDayPage1Container) {
            const quantityInputs = currentDayPage1Container.querySelectorAll('.labor-table .quantity-input');
            for (const input of quantityInputs) {
                if (input.value !== '') {
                    initialCheckChangesMade = true;
                    break;
                }
            }
        }
        if (initialCheckChangesMade) break;
    }

    if (initialCheckChangesMade) {
        efetivoWasJustClearedByButton = true;
        hasEfetivoBeenLoadedAtLeastOnce = true; 
    } else {
        efetivoWasJustClearedByButton = false;
        // Do not change hasEfetivoBeenLoadedAtLeastOnce if nothing was cleared; preserve its state.
    }

    let actualChangesInLoop = false;
    for (let d = 0; d <= rdoDayCounter; d++) {
        let currentDayPage1Container = null;
        const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
        let rdoDayWrapper = null;

        if (dayWithButtonLayoutWrapper) {
            rdoDayWrapper = dayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
        } else {
            rdoDayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        }

        if (rdoDayWrapper) {
            currentDayPage1Container = rdoDayWrapper.querySelector('.rdo-container');
        }

        if (!currentDayPage1Container) {
            console.warn(`clearAllEfetivo: Page 1 container for RDO day index ${d} not found.`);
            continue;
        }

        const laborTables = currentDayPage1Container.querySelectorAll('.labor-table');
        laborTables.forEach(table => {
            table.querySelectorAll('.quantity-input').forEach(input => {
                if (input.value !== '') {
                    input.value = '';
                    actualChangesInLoop = true;
                }
            });
        });
        calculateSectionTotalsForDay(currentDayPage1Container, d); // This will call checkAndToggleEfetivoAlertForDay
    }

    if (actualChangesInLoop) { // Use actualChangesInLoop for marking dirty and message
        markFormAsDirty();
        updateSaveProgressMessage("Efetivo removido de todos os RDOs.", false);
    } else {
        updateSaveProgressMessage("Nenhum efetivo para remover ou campos já estavam vazios.", false);
    }
    updateClearEfetivoButtonVisibility();
}
// --- End Efetivo Spreadsheet Loading ---


async function handleCopySaturdayEfetivo(targetSundayDayIndex) {
    try {
        updateSaveProgressMessage(null); 
        let targetSundayPage1Container = null;
        const targetSundayDayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${targetSundayDayIndex}`);
        
        if (targetSundayDayWithButtonLayoutWrapper) {
            const rdoWrapper = targetSundayDayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
            if (rdoWrapper) targetSundayPage1Container = rdoWrapper.querySelector('.rdo-container');
        } else {
            const rdoWrapper = getElementByIdSafe(`rdo_day_${targetSundayDayIndex}_wrapper`);
            if (rdoWrapper) targetSundayPage1Container = rdoWrapper.querySelector('.rdo-container');
        }

        if (!targetSundayPage1Container) {
            console.error(`handleCopySaturdayEfetivo: Target Sunday Page 1 container for day ${targetSundayDayIndex} not found.`);
            updateSaveProgressMessage("Erro: Não foi possível encontrar o RDO do domingo de destino.", true);
            return;
        }

        const sundayDateStr = getInputValue(`data_day${targetSundayDayIndex}`, targetSundayPage1Container);
        if (!sundayDateStr || !/^\d{4}-\d{2}-\d{2}$/.test(sundayDateStr)) {
            console.error(`handleCopySaturdayEfetivo: Invalid date for target Sunday day ${targetSundayDayIndex}: ${sundayDateStr}`);
            updateSaveProgressMessage("Erro: Data do domingo de destino inválida.", true);
            return;
        }

        const sundayDate = new Date(sundayDateStr + 'T00:00:00');
        if (isNaN(sundayDate.getTime())) {
            console.error(`handleCopySaturdayEfetivo: Could not parse date for target Sunday day ${targetSundayDayIndex}: ${sundayDateStr}`);
            updateSaveProgressMessage("Erro: Não foi possível processar a data do domingo de destino.", true);
            return;
        }

        const saturdayDate = new Date(sundayDate);
        saturdayDate.setDate(sundayDate.getDate() - 1);
        const saturdayDateStrYYYYMMDD = formatDateToYYYYMMDD(saturdayDate);

        let sourceSaturdayDayIndex = -1;
        let sourceSaturdayPage1Container = null;

        for (let d = 0; d <= rdoDayCounter; d++) {
            let currentDayPage1Container = null;
            const currentDayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
            if (currentDayWithButtonLayoutWrapper) {
                const rdoWrapper = currentDayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
                if (rdoWrapper) currentDayPage1Container = rdoWrapper.querySelector('.rdo-container');
            } else {
                const rdoWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
                if (rdoWrapper) currentDayPage1Container = rdoWrapper.querySelector('.rdo-container');
            }

            if (currentDayPage1Container) {
                const currentDateValue = getInputValue(`data_day${d}`, currentDayPage1Container);
                if (currentDateValue === saturdayDateStrYYYYMMDD) {
                    sourceSaturdayDayIndex = d;
                    sourceSaturdayPage1Container = currentDayPage1Container;
                    break;
                }
            }
        }

        if (sourceSaturdayDayIndex === -1 || !sourceSaturdayPage1Container) {
            updateSaveProgressMessage("RDO do sábado anterior não encontrado para cópia.", true);
            return;
        }

        const sourceLaborSection = sourceSaturdayPage1Container.querySelector('.labor-section');
        const targetLaborSection = targetSundayPage1Container.querySelector('.labor-section');

        if (!sourceLaborSection || !targetLaborSection) {
            console.error("handleCopySaturdayEfetivo: Labor sections not found in source or target.");
            updateSaveProgressMessage("Erro: Seções de efetivo não encontradas.", true);
            return;
        }

        const sourceQuantityInputs = sourceLaborSection.querySelectorAll('.labor-table .quantity-input');
        const targetQuantityInputs = targetLaborSection.querySelectorAll('.labor-table .quantity-input');

        if (sourceQuantityInputs.length !== targetQuantityInputs.length) {
            console.warn("handleCopySaturdayEfetivo: Mismatch in number of labor inputs between Saturday and Sunday. Copy might be incomplete.");
            updateSaveProgressMessage("Aviso: Discrepância no número de campos de efetivo. A cópia pode ser incompleta.", true);
        }

        let copied = false;
        sourceQuantityInputs.forEach((sourceInput, index) => {
            if (targetQuantityInputs[index]) {
                targetQuantityInputs[index].value = sourceInput.value;
                copied = true;
            }
        });

        if (copied) {
            efetivoWasJustClearedByButton = false; 
            hasEfetivoBeenLoadedAtLeastOnce = true;
            calculateSectionTotalsForDay(targetSundayPage1Container, targetSundayDayIndex);
            markFormAsDirty();
            updateSaveProgressMessage("Efetivo do sábado copiado com sucesso!");
        } else {
            updateSaveProgressMessage("Nenhum efetivo para copiar ou campos de destino não encontrados.", true);
        }
        updateAllEfetivoAlerts();
        updateClearEfetivoButtonVisibility();

    } catch (error) {
        console.error("Error in handleCopySaturdayEfetivo:", error);
        updateSaveProgressMessage("Ocorreu um erro inesperado ao copiar o efetivo.", true);
    }
}


function updateEfetivoCopyButtonVisibility(dayIndex) {
    const rdoAllDaysWrapper = getElementByIdSafe('rdo-all-days-wrapper');
    const originalRdoDayWrapper = getElementByIdSafe(`rdo_day_${dayIndex}_wrapper`);

    if (!rdoAllDaysWrapper || !originalRdoDayWrapper) {
        console.warn(`Wrappers not found for efetivo button visibility, day ${dayIndex}`);
        return;
    }

    const page1ContainerForDate = originalRdoDayWrapper.querySelector('.rdo-container');
    let dayOfWeek = '';
    if (page1ContainerForDate) {
        const dateInputForDay = getElementByIdSafe(`data_day${dayIndex}`, page1ContainerForDate);
        if (dateInputForDay && dateInputForDay.value) {
            const dateObj = new Date(dateInputForDay.value + 'T00:00:00');
            if (!isNaN(dateObj.getTime())) {
                dayOfWeek = getDayOfWeekFromDate(dateObj);
            }
        }
    }
    
    const dayWithButtonContainerId = `day_with_button_container_day${dayIndex}`;
    const sidebarContainerId = `efetivo_actions_sidebar_day${dayIndex}`;
    const buttonId = `copySaturdayEfetivoButton_day${dayIndex}`;
    
    let dayWithButtonContainer = document.getElementById(dayWithButtonContainerId);

    if (dayOfWeek === 'domingo') {
        if (!dayWithButtonContainer) {
            dayWithButtonContainer = document.createElement('div');
            dayWithButtonContainer.id = dayWithButtonContainerId;
            dayWithButtonContainer.className = 'day-with-button-container';

            const sidebarContainer = document.createElement('div');
            sidebarContainer.id = sidebarContainerId;
            sidebarContainer.className = 'efetivo-actions-sidebar';

            const copyButton = document.createElement('button');
            copyButton.id = buttonId;
            copyButton.textContent = 'Copiar Efetivo do Sábado Anterior';
            copyButton.className = 'header-action-button copy-saturday-button';
            copyButton.type = 'button';
            copyButton.addEventListener('click', () => handleCopySaturdayEfetivo(dayIndex));
            
            sidebarContainer.appendChild(copyButton);

            if (originalRdoDayWrapper.parentElement === rdoAllDaysWrapper) {
                 rdoAllDaysWrapper.insertBefore(dayWithButtonContainer, originalRdoDayWrapper);
            } else if (originalRdoDayWrapper.parentElement?.parentElement === rdoAllDaysWrapper && 
                       originalRdoDayWrapper.parentElement.id !== dayWithButtonContainerId) {
                rdoAllDaysWrapper.insertBefore(dayWithButtonContainer, originalRdoDayWrapper.parentElement);
            } else if (!rdoAllDaysWrapper.contains(dayWithButtonContainer)){ // Check if it's not already a child in some other structure
                rdoAllDaysWrapper.appendChild(dayWithButtonContainer); // Fallback append if not easily inserted before
            }
            
            dayWithButtonContainer.appendChild(originalRdoDayWrapper);
            dayWithButtonContainer.appendChild(sidebarContainer);

            originalRdoDayWrapper.style.margin = '0'; 
            originalRdoDayWrapper.style.marginBottom = '0'; 
        }
        dayWithButtonContainer.style.display = ''; 
    } else { 
        if (dayWithButtonContainer) {
            // Check if originalRdoDayWrapper is indeed inside dayWithButtonContainer before moving
            if (originalRdoDayWrapper.parentElement === dayWithButtonContainer) {
                dayWithButtonContainer.parentElement?.insertBefore(originalRdoDayWrapper, dayWithButtonContainer);
            }
            originalRdoDayWrapper.style.margin = ''; 
            originalRdoDayWrapper.style.marginBottom = ''; 
            dayWithButtonContainer.remove();
        }
    }
}


async function updateAllDateRelatedFieldsFromDay0() {
    const day0Page1Container = getElementByIdSafe('rdo_day_0_page_0');
    if (!day0Page1Container) {
        console.error("Day 0 Page 1 container (#rdo_day_0_page_0) not found. Global date synchronization cannot proceed.");
        return;
    }

    const dataDay0Input = day0Page1Container.querySelector('#data_day0');
    if (!dataDay0Input) {
        console.error("Master date input #data_day0 not found within #rdo_day_0_page_0. Date updates aborted.");
        return;
    }

    let day0DataStr = dataDay0Input.value.trim(); 
    const day0DefaultDataStr = dataDay0Input.defaultValue.trim(); 

    const dateRegex = /^\d{4}-\d{2}-\d{2}$/; 

    if (!dateRegex.test(day0DataStr)) {
        if (dateRegex.test(day0DefaultDataStr)) {
            day0DataStr = day0DefaultDataStr; 
            setInputValue('data_day0', day0DataStr, day0Page1Container); 
        } else {
            console.error(`Both current value ("${dataDay0Input.value}") and defaultValue ("${dataDay0Input.defaultValue}") for #data_day0 are invalid (not YYYY-MM-DD). Date updates aborted.`);
            return;
        }
    }
    
    const day0PrazoStr = getInputValue('prazo_day0', day0Page1Container); 
    const day0InicioObraStr = getInputValue('inicio_obra_day0', day0Page1Container); 

    const baseDateForDay0 = new Date(day0DataStr + 'T00:00:00'); 
    if (isNaN(baseDateForDay0.getTime())) {
        console.error(`Invalid base date calculated for Day 0 from string: "${day0DataStr}T00:00:00". Date updates aborted.`);
        return;
    }

    const inicioObraDateObj = parseDateDDMMYYYY(day0InicioObraStr); 
    const prazoDateObj = parseDateDDMMYYYY(day0PrazoStr);         
    
    let currentDateIterator = new Date(baseDateForDay0.getTime()); 

    for (let d = 0; d <= rdoDayCounter; d++) {
        const daySuffix = `_day${d}`;
        
        let currentRdoDayWrapper;
        const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
        if (dayWithButtonLayoutWrapper) {
            currentRdoDayWrapper = dayWithButtonLayoutWrapper.querySelector(`.rdo-day-wrapper`);
        } else {
            currentRdoDayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        }

        if (!currentRdoDayWrapper) continue; 

        const currentPage1Container = currentRdoDayWrapper.querySelectorAll('.rdo-container')[0];
        if (!currentPage1Container) continue; 

        if (d > 0) { 
            currentDateIterator.setDate(currentDateIterator.getDate() + 1);
        }
        
        const newDateStr = formatDateToYYYYMMDD(currentDateIterator); 
        const newDayOfWeek = getDayOfWeekFromDate(currentDateIterator);

        setInputValue(`data${daySuffix}`, newDateStr, currentPage1Container);
        setInputValue(`dia_semana${daySuffix}`, newDayOfWeek, currentPage1Container);
        setInputValue(`data_pt${daySuffix}`, newDateStr, currentRdoDayWrapper);
        setInputValue(`dia_semana_pt${daySuffix}`, newDayOfWeek, currentRdoDayWrapper);
        setInputValue(`data_pt2${daySuffix}`, newDateStr, currentRdoDayWrapper);
        setInputValue(`dia_semana_pt2${daySuffix}`, newDayOfWeek, currentRdoDayWrapper);

        const statusEl = getElementByIdSafe(`indice_pluv_status_day${d}`);
        if (statusEl && !statusEl.textContent?.trim()) { 
             updateRainfallStatus(d, 'idle');
        }


        let decorridosStr = "N/A"; 
        let restantesStr = "N/A";  

        if (inicioObraDateObj) { 
            const decorridosVal = calculateDaysBetween(inicioObraDateObj, currentDateIterator) + 1; 
            decorridosStr = !isNaN(decorridosVal) && decorridosVal >= 0 ? decorridosVal.toString() : "0"; 

            if (prazoDateObj) { 
                const totalPrazoDays = calculateDaysBetween(inicioObraDateObj, prazoDateObj) + 1; 
                if (!isNaN(totalPrazoDays) && totalPrazoDays >= 0) {
                    const restantesVal = totalPrazoDays - decorridosVal;
                    restantesStr = !isNaN(restantesVal) && restantesVal >= 0 ? restantesVal.toString() : "0";
                } else {
                    restantesStr = "N/A"; 
                }
            } else {
                restantesStr = "N/A"; 
            }
        } else {
            decorridosStr = "N/A"; 
            restantesStr = "N/A";
        }

        setInputValue(`decorridos${daySuffix}`, decorridosStr, currentPage1Container);
        setInputValue(`restantes${daySuffix}`, restantesStr, currentPage1Container);
        setInputValue(`decorridos_pt${daySuffix}`, decorridosStr, currentRdoDayWrapper);
        setInputValue(`restantes_pt${daySuffix}`, restantesStr, currentRdoDayWrapper);
        setInputValue(`decorridos_pt2${daySuffix}`, decorridosStr, currentRdoDayWrapper);
        setInputValue(`restantes_pt2${daySuffix}`, restantesStr, currentRdoDayWrapper);

        updateEfetivoCopyButtonVisibility(d);
    }
}

async function setupContractDataSyncForDay(dayIndex, page1Container) {
    const daySuffix = `_day${dayIndex}`;
    
    let dayWrapper;
    const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${dayIndex}`);
    if (dayWithButtonLayoutWrapper) {
        dayWrapper = dayWithButtonLayoutWrapper.querySelector(`.rdo-day-wrapper`);
    } else {
        dayWrapper = getElementByIdSafe(`rdo_day_${dayIndex}_wrapper`);
    }

    if (!dayWrapper) {
        console.error(`Day wrapper for day ${dayIndex} not found. Contract data sync aborted.`);
        return;
    }

    if (dayIndex === 0) {
        const day0DateDrivingFields = ['data_day0', 'prazo_day0', 'inicio_obra_day0'];
        day0DateDrivingFields.forEach(masterId => {
            const masterEl = getElementByIdSafe(masterId, page1Container);
            if (masterEl) {
                masterEl.readOnly = false; 
                const eventType = (masterEl.type === 'date' || masterId.includes('prazo') || masterId.includes('inicio_obra')) ? 'change' : 'input';
                masterEl.addEventListener(eventType, async () => { 
                    markFormAsDirty(); 
                    if (masterId === 'data_day0') {
                        await adjustRdoDaysForMonth(masterEl.value); 
                    }
                    await updateAllDateRelatedFieldsFromDay0(); 
                });
            }
        });
    }

    contractFieldMappingsBase.forEach(fieldPair => {
        const masterElementId = fieldPair.masterIdBase + daySuffix; 
        const masterElement = getElementByIdSafe(masterElementId, page1Container);

        if (masterElement) {
            if (dayIndex === 0) {
                const editableDay0Page1MasterFields = ['data', 'prazo', 'inicio_obra', 'contratada', 'contrato_num'];
                if (editableDay0Page1MasterFields.includes(fieldPair.masterIdBase)) {
                    masterElement.readOnly = false; 
                } else {
                    masterElement.readOnly = true;
                }
            } else {
                masterElement.readOnly = true;
            }

            const syncCurrentFieldToSlavesAndNonDateGlobals = () => {
                const masterValue = masterElement.value;

                fieldPair.slaveIdBases.forEach(slaveIdBase => {
                    // Corrected to use dayWrapper for slave elements as they might be outside page1Container
                    const slaveElement = getElementByIdSafe(slaveIdBase + daySuffix, dayWrapper);
                    if (slaveElement) {
                        slaveElement.value = masterValue;
                        slaveElement.readOnly = true; 
                    }
                });

                const isNonDateGlobalMasterField = globalMasterFieldBases.includes(fieldPair.masterIdBase) &&
                                                  !['data', 'prazo', 'inicio_obra', 'dia_semana', 'decorridos', 'restantes'].includes(fieldPair.masterIdBase);

                if (dayIndex === 0 && isNonDateGlobalMasterField) {
                    for (let d = 1; d <= rdoDayCounter; d++) { 
                        const targetDaySuffix = `_day${d}`;
                        
                        let targetDayWrapper;
                        const targetDayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
                        if(targetDayWithButtonLayoutWrapper){
                            targetDayWrapper = targetDayWithButtonLayoutWrapper.querySelector(`.rdo-day-wrapper`);
                        } else {
                            targetDayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
                        }

                        if (targetDayWrapper) {
                            const targetPage1ContainerOfSubsequentDay = targetDayWrapper.querySelectorAll('.rdo-container')[0];
                            if (targetPage1ContainerOfSubsequentDay) {
                                const targetMasterElOnSubsequentDay = getElementByIdSafe(fieldPair.masterIdBase + targetDaySuffix, targetPage1ContainerOfSubsequentDay);
                                if (targetMasterElOnSubsequentDay) {
                                    targetMasterElOnSubsequentDay.value = masterValue;
                                }
                            }
                            fieldPair.slaveIdBases.forEach(slaveIdBase => {
                                const targetSlaveElOnSubsequentDay = getElementByIdSafe(slaveIdBase + targetDaySuffix, targetDayWrapper);
                                if (targetSlaveElOnSubsequentDay) {
                                    targetSlaveElOnSubsequentDay.value = masterValue;
                                    targetSlaveElOnSubsequentDay.readOnly = true; 
                                }
                            });
                        }
                    }
                }
            };
            
            if (!masterElement.readOnly) { 
                masterElement.addEventListener('input', () => {
                    markFormAsDirty(); 
                    syncCurrentFieldToSlavesAndNonDateGlobals();
                });
                if (['contratada', 'contrato_num'].includes(fieldPair.masterIdBase)) { 
                     masterElement.addEventListener('change', () => {
                        markFormAsDirty(); 
                        syncCurrentFieldToSlavesAndNonDateGlobals();
                     }); 
                }
            }
            
            syncCurrentFieldToSlavesAndNonDateGlobals(); 
        }
    });
    
    const masterRdoInput = getElementByIdSafe(`numero_rdo${daySuffix}`, page1Container);
    if(masterRdoInput) {
        if (dayIndex > 0) { 
            masterRdoInput.readOnly = true;
        } else { 
            masterRdoInput.readOnly = false;
        }

        const syncRdoNumbers = () => {
            const rdoNumFull = masterRdoInput.value; 
            const rdoNumParts = rdoNumFull.split('-');
            const baseRdoNumStr = rdoNumParts[0]; 

            if(baseRdoNumStr) { 
                const rdoPtInput = getElementByIdSafe(`numero_rdo_pt${daySuffix}`, dayWrapper);
                if (rdoPtInput) {
                    rdoPtInput.value = `${baseRdoNumStr}-B`;
                    rdoPtInput.readOnly = true; 
                }
                const rdoPt2Input = getElementByIdSafe(`numero_rdo_pt2${daySuffix}`, dayWrapper);
                if (rdoPt2Input) {
                    rdoPt2Input.value = `${baseRdoNumStr}-C`;
                    rdoPt2Input.readOnly = true; 
                }

                if (dayIndex === 0) {
                    const day0BaseRdoNum = parseInt(baseRdoNumStr);
                    if (!isNaN(day0BaseRdoNum)) {
                        for (let d = 1; d <= rdoDayCounter; d++) { 
                            const nextDaySuffix = `_day${d}`;
                            let nextDayWrapper;
                            const nextDayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
                            if(nextDayWithButtonLayoutWrapper) {
                                nextDayWrapper = nextDayWithButtonLayoutWrapper.querySelector(`.rdo-day-wrapper`);
                            } else {
                                nextDayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
                            }

                            if (nextDayWrapper) {
                                const nextRdoBaseNum = day0BaseRdoNum + d; 

                                const nextPage1ContainerOfSubsequentDay = nextDayWrapper.querySelectorAll('.rdo-container')[0];
                                if (nextPage1ContainerOfSubsequentDay) {
                                    const nextMasterRdoInput = getElementByIdSafe(`numero_rdo${nextDaySuffix}`, nextPage1ContainerOfSubsequentDay);
                                    if (nextMasterRdoInput) {
                                        nextMasterRdoInput.value = `${nextRdoBaseNum}-A`; 
                                    }
                                }
                                setInputValue(`numero_rdo_pt${nextDaySuffix}`, `${nextRdoBaseNum}-B`, nextDayWrapper);
                                setInputValue(`numero_rdo_pt2${nextDaySuffix}`, `${nextRdoBaseNum}-C`, nextDayWrapper);
                            }
                        }
                    }
                }
            }
        };

        if (!masterRdoInput.readOnly) {
            const rdoChangeHandler = () => {
                markFormAsDirty(); 
                syncRdoNumbers();
            };
            masterRdoInput.addEventListener('input', rdoChangeHandler);
            masterRdoInput.addEventListener('change', rdoChangeHandler); 
        }
        syncRdoNumbers(); 
    }
}

// --- Day Navigation Panel ---
const dayTabsContainer = getElementByIdSafe('day-tabs-container');
const dayTabs = []; 
let navObserver = null;
let lastActivatedByScroll = -1; 
const EXTRA_SCROLL_PADDING_TOP = 11; // Pixels of extra space

function setActiveTab(activeIndex) {
    if (activeIndex < 0 || activeIndex >= dayTabs.length) {
        dayTabs.forEach(tab => {
            tab.classList.remove('active');
            tab.removeAttribute('aria-current');
        });
        if (dayTabs.length > 0 && activeIndex < 0) { 
            dayTabs[0].classList.add('active');
            dayTabs[0].setAttribute('aria-current', 'true');
        }
        return;
    }

    dayTabs.forEach((tab, index) => {
        if (index === activeIndex) {
            tab.classList.add('active');
            tab.setAttribute('aria-current', 'true');
        } else {
            tab.classList.remove('active');
            tab.removeAttribute('aria-current');
        }
    });
}

function createDayTab(dayIndex) {
    if (!dayTabsContainer) return;

    const tab = document.createElement('button');
    tab.id = `day-tab-${dayIndex}`;
    tab.className = 'day-tab';
    tab.type = 'button';
    tab.setAttribute('aria-controls', `rdo_day_${dayIndex}_wrapper`); 
    tab.setAttribute('role', 'tab');

    // Create text node for day number
    const dayNumberText = document.createTextNode(`${dayIndex + 1}`);
    tab.appendChild(dayNumberText);

    // Create span for alert icon
    const alertIconSpan = document.createElement('span');
    alertIconSpan.className = 'efetivo-alert-icon';
    alertIconSpan.setAttribute('aria-hidden', 'true');
    // alertIconSpan.style.display = 'none'; // Rely on CSS class '.efetivo-alert-icon' for default hiding
    alertIconSpan.textContent = '⚠️'; 
    tab.appendChild(alertIconSpan);


    tab.addEventListener('click', () => {
        const targetDayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${dayIndex}`);
        const targetRdoDayWrapper = getElementByIdSafe(`rdo_day_${dayIndex}_wrapper`);
        const elementToScrollTo = targetDayWithButtonLayoutWrapper || targetRdoDayWrapper;

        if (elementToScrollTo) {
            const rootStyles = getComputedStyle(document.documentElement);
            const consbemHeaderHeight = parseInt(rootStyles.getPropertyValue('--consbem-header-height').trim() || '60', 10);
            const rdoActionBarHeight = parseInt(rootStyles.getPropertyValue('--rdo-action-bar-height').trim() || '60', 10);
            const totalHeaderOffset = consbemHeaderHeight + rdoActionBarHeight;

            const elementPosition = elementToScrollTo.getBoundingClientRect().top + window.scrollY;
            const scrollToPosition = elementPosition - totalHeaderOffset - EXTRA_SCROLL_PADDING_TOP;


            window.scrollTo({
                top: scrollToPosition,
                behavior: 'smooth'
            });
        }
    });

    dayTabsContainer.appendChild(tab);
    dayTabs.push(tab);

    if (dayTabs.length === 1 && !dayTabsContainer.querySelector('.day-tab.active')) { 
        setActiveTab(0);
    }
}

function setupNavigationObserver() {
    if (navObserver) { 
        navObserver.disconnect();
    }
    const appContentCanvas = getElementByIdSafe('app-content-canvas');
    if (!appContentCanvas) {
        console.warn("App content canvas not found for navigation observer.");
        return;
    }

    const consbemHeaderHeight = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--consbem-header-height') || '60');
    const rdoActionBarHeight = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--rdo-action-bar-height') || '70');
    const topOffset = consbemHeaderHeight + rdoActionBarHeight + EXTRA_SCROLL_PADDING_TOP + 5; // Added 5px buffer to the topOffset for more reliable activation


    const observerOptions = {
        root: null, 
        rootMargin: `-${topOffset}px 0px -${window.innerHeight - topOffset - 100}px 0px`, 
        threshold: 0.01 
    };

    navObserver = new IntersectionObserver((entries) => {
        let highestVisibleEntry = null;

        entries.forEach(entry => {
            if (entry.isIntersecting) {
                if (!highestVisibleEntry || entry.boundingClientRect.top < highestVisibleEntry.boundingClientRect.top) {
                    highestVisibleEntry = entry;
                }
            }
        });

        if (highestVisibleEntry) {
            let dayIndexStr;
            const targetEl = highestVisibleEntry.target;

            if (targetEl.classList.contains('day-with-button-container')) {
                const idParts = targetEl.id.split('_'); 
                dayIndexStr = idParts[idParts.length -1]?.replace('day','');
            } else if (targetEl.classList.contains('rdo-day-wrapper')) {
                dayIndexStr = targetEl.dataset.dayIndex;
            }


            if (dayIndexStr) {
                const dayIndex = parseInt(dayIndexStr, 10);
                if (dayIndex !== lastActivatedByScroll || !dayTabs[dayIndex]?.classList.contains('active')) {
                    setActiveTab(dayIndex);
                    lastActivatedByScroll = dayIndex;
                }
            }
        }
    }, observerOptions);

    const dayWrappers = document.querySelectorAll('.rdo-day-wrapper, .day-with-button-container');
    dayWrappers.forEach(wrapper => {
        if (navObserver) { 
            navObserver.observe(wrapper);
        }
    });
}

function attachDirtyListenersToDay(dayWrapperElement) {
    const inputs = dayWrapperElement.querySelectorAll(
        'input[type="text"], input[type="date"], input[type="number"], textarea, select, input[type="file"]'
    );
    inputs.forEach(input => {
        // Exclude hidden inputs for status as they are handled explicitly or via dispatched event
        if (input.type === 'hidden' && (input.name.startsWith('status_day') || input.name.startsWith('localizacao_day') || input.name.startsWith('tipo_servico_day') || input.name.startsWith('servico_desc_day')) ) {
            return;
        }
        const eventType = (input.type === 'date' || input.tagName.toLowerCase() === 'select' || input.type === 'file') ? 'change' : 'input';
        input.addEventListener(eventType, markFormAsDirty);
    });
}

// --- Efetivo Alert Logic ---
function checkAndToggleEfetivoAlertForDay(dayIndex) {
    const tab = getElementByIdSafe(`day-tab-${dayIndex}`);
    if (!tab) return; // Tab might not exist if day was just removed

    const alertIcon = tab.querySelector('.efetivo-alert-icon');
    if (!alertIcon) return;

    const hideAlert = () => {
        alertIcon.classList.remove('visible');
    };
    const showAlert = () => {
        alertIcon.classList.add('visible');
    };

    let rdoDayWrapper = null;
    const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${dayIndex}`);
    if (dayWithButtonLayoutWrapper) {
        rdoDayWrapper = dayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
    } else {
        rdoDayWrapper = getElementByIdSafe(`rdo_day_${dayIndex}_wrapper`);
    }

    if (!rdoDayWrapper) {
        hideAlert();
        return;
    }
    const page1Container = rdoDayWrapper.querySelector('.rdo-container');
    if (!page1Container) {
        hideAlert();
        return;
    }
    
    // Condition 1: If efetivo has never been "loaded" or interacted with meaningfully, always hide.
    if (!hasEfetivoBeenLoadedAtLeastOnce) {
        hideAlert();
        return;
    }

    const quantityInputs = page1Container.querySelectorAll('.labor-table .quantity-input');
    let efetivoPresent = false;
    if (quantityInputs.length > 0) {
        for (const input of quantityInputs) {
            const value = parseInt(input.value, 10);
            if (!isNaN(value) && value > 0) {
                efetivoPresent = true;
                break;
            }
        }
    }

    if (efetivoPresent) {
        hideAlert();
    } else { // Efetivo is NOT present for this specific day
        // Condition 2: If efetivo was just successfully cleared by the global button, hide the alert
        if (efetivoWasJustClearedByButton) {
            hideAlert();
        } else {
            // Otherwise (efetivo is missing, not just cleared, and has been loaded/interacted with), show alert.
            showAlert();
        }
    }
}


function updateAllEfetivoAlerts() {
    for (let d = 0; d <= rdoDayCounter; d++) {
        checkAndToggleEfetivoAlertForDay(d);
    }
}
// --- End Efetivo Alert Logic ---


async function initializeDayInstance(dayIndex, dayWrapperElement) {
    dayWrapperElement.dataset.dayIndex = dayIndex.toString(); 
    
    const pageContainers = dayWrapperElement.querySelectorAll('.rdo-container');
    
    pageContainers.forEach((pageContainer, pageInDayIdx) => {
        let pageIndicator = pageContainer.querySelector('.page-indicator');
        if (!pageIndicator) { 
            pageIndicator = document.createElement('div');
            pageIndicator.className = 'page-indicator';
            if (pageContainer.firstChild) { 
                pageContainer.insertBefore(pageIndicator, pageContainer.firstChild);
            } else {
                pageContainer.appendChild(pageIndicator);
            }
        }
        pageIndicator.textContent = `DIA ${dayIndex + 1} / FOLHA ${pageInDayIdx + 1}`; 
    });

    const page1Container = pageContainers[0]; 
    const page2Container = pageContainers[1]; 
    const page3Container = pageContainers[2]; 

    if (!page1Container || !page2Container || !page3Container) {
        console.error(`Could not find all three page containers for day ${dayIndex}. Initialization incomplete.`);
        return;
    }
    
    createCombinedShiftControl('turno1_control_cell', 't1', 'B', 'N', '1º Turno', dayIndex);
    createCombinedShiftControl('turno2_control_cell', 't2', 'B', 'N', '2º Turno', dayIndex);
    createCombinedShiftControl('turno3_control_cell', 't3', "", "", '3º Turno', dayIndex); 

    const laborSectionPage1 = page1Container.querySelector('.labor-section');
    if (laborSectionPage1) {
        const quantityInputsP1 = laborSectionPage1.querySelectorAll('.labor-table .quantity-input');
        
        if (dayIndex === 0) { 
            quantityInputsP1.forEach(input => {
                input.value = ''; 
            });
        }
        
        quantityInputsP1.forEach(input => {
            input.addEventListener('input', () => {
                 efetivoWasJustClearedByButton = false; 
                 hasEfetivoBeenLoadedAtLeastOnce = true; 
                 markFormAsDirty(); 
                 calculateSectionTotalsForDay(page1Container, dayIndex);
            });
        });
        calculateSectionTotalsForDay(page1Container, dayIndex); 

        applyMultiColumnLayoutIfNecessary(page1Container.querySelector('.labor-direta .labor-table tbody'));
        applyMultiColumnLayoutIfNecessary(page1Container.querySelector('.labor-indireta .labor-table tbody'));
        applyMultiColumnLayoutIfNecessary(page1Container.querySelector('.equipamentos .labor-table tbody'));
    }
    
    setupPhotoAttachmentForDay(dayIndex);
    
    const activitiesTableBodyP1 = getElementByIdSafe(`activitiesTable_day${dayIndex}`, page1Container)?.getElementsByTagName('tbody')[0];
    if (activitiesTableBodyP1 && activitiesTableBodyP1.rows.length === 0) { 
        for (let i = 0; i < 24; i++) addActivityRow('activitiesTable', dayIndex); 
    }
    
    const activitiesTableBodyP2 = getElementByIdSafe(`activitiesTable_pt_day${dayIndex}`, page2Container)?.getElementsByTagName('tbody')[0];
    if (activitiesTableBodyP2 && activitiesTableBodyP2.rows.length === 0) { 
        for (let i = 0; i < 22; i++) addActivityRow('activitiesTable_pt', dayIndex); 
    }

    if (page3Container) { 
        for (let i = 1; i <= 4; i++) { 
            setupReportPhotoSlot(
                `report_photo_${i}_pt2_input`, `report_photo_${i}_pt2_preview`,
                `report_photo_${i}_pt2_placeholder`, 
                `report_photo_${i}_pt2_clear_button`,
                `report_photo_${i}_pt2_label`, dayIndex, i 
            );
            const captionInput = getElementByIdSafe(`report_photo_${i}_pt2_caption_day${dayIndex}`, page3Container);
            if (captionInput) {
                captionInput.addEventListener('input', markFormAsDirty);
            }
        }
    }
    
    await setupContractDataSyncForDay(dayIndex, page1Container);
    updateEfetivoCopyButtonVisibility(dayIndex); 
    attachDirtyListenersToDay(dayWrapperElement); 


    if (!dayTabs.find(tab => tab.id === `day-tab-${dayIndex}`)) {
        createDayTab(dayIndex);
    }
    updateRainfallStatus(dayIndex, 'idle'); 
    resetDayInstanceContent(dayWrapperElement, dayIndex); // Call this to ensure initial alert status is correct
}


function updateIdsRecursive(element, oldDaySuffix, newDaySuffix) {
    const escapedOldDaySuffixForRegex = oldDaySuffix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); 
    const regexOldSuffixGlobal = new RegExp(escapedOldDaySuffixForRegex, 'g');

    if (element.id && element.id.includes(oldDaySuffix)) {
        element.id = element.id.replace(regexOldSuffixGlobal, newDaySuffix);
    }

    const forAttribute = element.getAttribute('for');
    if (forAttribute && forAttribute.includes(oldDaySuffix)) {
        element.setAttribute('for', forAttribute.replace(regexOldSuffixGlobal, newDaySuffix));
    }

    const nameAttribute = element.getAttribute('name');
    if (nameAttribute && nameAttribute.includes(oldDaySuffix)) {
        element.setAttribute('name', nameAttribute.replace(regexOldSuffixGlobal, newDaySuffix));
    }
    
    const ariaLabelAttribute = element.getAttribute('aria-label');
    if (ariaLabelAttribute && ariaLabelAttribute.includes(`Dia ${parseInt(oldDaySuffix.replace('_day',''),10) + 1}`)) {
        element.setAttribute('aria-label', ariaLabelAttribute.replace(`Dia ${parseInt(oldDaySuffix.replace('_day',''),10) + 1}`, `Dia ${parseInt(newDaySuffix.replace('_day',''),10) + 1}`));
    }


    ['aria-controls', 'aria-labelledby', 'aria-describedby'].forEach(attr => {
        const val = element.getAttribute(attr);
        if (val && val.includes(oldDaySuffix)) {
            const updatedVal = val.split(' ')
                                .map(idPart => idPart.includes(oldDaySuffix) ? idPart.replace(regexOldSuffixGlobal, newDaySuffix) : idPart)
                                .join(' ');
            element.setAttribute(attr, updatedVal);
        }
    });

    element.childNodes.forEach(child => {
        if (child.nodeType === Node.ELEMENT_NODE) { 
            updateIdsRecursive(child, oldDaySuffix, newDaySuffix);
        }
    });
}

function resetDayInstanceContent(dayWrapperElement, dayIndex) {
    const daySuffix = `_day${dayIndex}`;

    dayWrapperElement.querySelectorAll('.labor-table .quantity-input').forEach(input => input.value = '');
    
    // Clear standard inputs and textareas in activities tables (excluding those now handled by dropdowns)
    const activityFieldsSelectors = [
        `[id^="activitiesTable_day${dayIndex}"] input[type="text"]:not([name^="localizacao_"]):not([name^="status_"]):not([name^="tipo_servico_"])`, 
        `[id^="activitiesTable_day${dayIndex}"] textarea:not([name^="servico_desc_day${dayIndex}[]"])`,
        `[id^="activitiesTable_pt_day${dayIndex}"] input[type="text"]:not([name^="localizacao_"]):not([name^="status_"]):not([name^="tipo_servico_"])`, 
        `[id^="activitiesTable_pt_day${dayIndex}"] textarea:not([name^="servico_desc_day${dayIndex}[]"])`
    ];
    dayWrapperElement.querySelectorAll(activityFieldsSelectors.join(', '))
        .forEach(el => el.value = '');

    // Reset custom dropdowns (localização, status, tipo_servico, servico)
    const dropdownContainers = dayWrapperElement.querySelectorAll('.status-select-container');
    dropdownContainers.forEach(container => {
        const display = container.querySelector('.status-display');
        const hiddenInput = container.querySelector('input[type="hidden"]');
        const dropdown = container.querySelector('.status-dropdown');

        if (display) {
            display.textContent = ''; 
            display.setAttribute('aria-expanded', 'false');
        }
        if (hiddenInput) {
            hiddenInput.value = '';
        }
        if (dropdown) {
            dropdown.style.display = 'none';
            dropdown.querySelectorAll('.status-option').forEach(opt => {
                opt.setAttribute('aria-selected', (opt.dataset.value === "") ? 'true' : 'false');
            });
        }
    });
    
    const page2Container = dayWrapperElement.querySelectorAll('.rdo-container')[1];
    if (page2Container) {
        const page2ObservationsTextarea = getElementByIdSafe(`observacoes_comentarios_pt${daySuffix}`, page2Container);
        if (page2ObservationsTextarea) page2ObservationsTextarea.value = '';
    }

    for (let i = 1; i <= 4; i++) {
        const captionInput = dayWrapperElement.querySelector(`#report_photo_${i}_pt2_caption${daySuffix}`);
        if (captionInput) captionInput.value = '';

        const photoInput = getElementByIdSafe(`report_photo_${i}_pt2_input${daySuffix}`, dayWrapperElement);
        const preview = getElementByIdSafe(`report_photo_${i}_pt2_preview${daySuffix}`, dayWrapperElement);
        const placeholder = getElementByIdSafe(`report_photo_${i}_pt2_placeholder${daySuffix}`, dayWrapperElement);
        const clearButton = getElementByIdSafe(`report_photo_${i}_pt2_clear_button${daySuffix}`, dayWrapperElement);
        const label = getElementByIdSafe(`report_photo_${i}_pt2_label${daySuffix}`, dayWrapperElement);
        resetReportPhotoSlotState(photoInput, preview, placeholder, clearButton, label);
    }
    
    const obsClima = dayWrapperElement.querySelector(`#obs_clima${daySuffix}`);
    if(obsClima) obsClima.value = '';
    
    const indicePluvInput = dayWrapperElement.querySelector(`#indice_pluv_valor${daySuffix}`);
    if (indicePluvInput) indicePluvInput.value = '';
    updateRainfallStatus(dayIndex, 'idle'); 
    
    const page1Container = dayWrapperElement.querySelectorAll('.rdo-container')[0];
    if (page1Container) {
        calculateSectionTotalsForDay(page1Container, dayIndex); 
    }
    updateEfetivoCopyButtonVisibility(dayIndex); 
    checkAndToggleEfetivoAlertForDay(dayIndex); 
}

async function addNewRdoDay() {
    const newDayIndex = ++rdoDayCounter; 
    const prevDayIndex = newDayIndex - 1; 

    const templateDayWrapper = getElementByIdSafe(`rdo_day_0_wrapper`)?.cloneNode(true);
    if (!templateDayWrapper) {
        console.error("Template for day (rdo_day_0_wrapper) not found! Cannot add new day.");
        rdoDayCounter--; 
        return;
    }

    templateDayWrapper.id = `rdo_day_${newDayIndex}_wrapper`; 
    updateIdsRecursive(templateDayWrapper, '_day0', `_day${newDayIndex}`);
    templateDayWrapper.dataset.dayIndex = newDayIndex.toString(); 

    templateDayWrapper.querySelectorAll('.rdo-container').forEach(container => {
        container.classList.add('rdo-container-subsequent-day');
    });
    
    for (let i = 1; i <= 4; i++) { 
        const captionInput = templateDayWrapper.querySelector(`#report_photo_${i}_pt2_caption_day${newDayIndex}`);
        if (captionInput) {
            captionInput.value = '';
        }
    }

    // Clear activity table bodies in the cloned template BEFORE initialization
    const newDayPageContainersInCloned = templateDayWrapper.querySelectorAll('.rdo-container');
    const newDayPage1ContainerInCloned = newDayPageContainersInCloned[0];
    const newDayPage2ContainerInCloned = newDayPageContainersInCloned[1];

    if (newDayPage1ContainerInCloned) {
        const activitiesTableBodyP1Cloned = getElementByIdSafe(`activitiesTable_day${newDayIndex}`, newDayPage1ContainerInCloned)?.getElementsByTagName('tbody')[0];
        if (activitiesTableBodyP1Cloned) {
            activitiesTableBodyP1Cloned.innerHTML = '';
        }
    }
    if (newDayPage2ContainerInCloned) {
        const activitiesTableBodyP2Cloned = getElementByIdSafe(`activitiesTable_pt_day${newDayIndex}`, newDayPage2ContainerInCloned)?.getElementsByTagName('tbody')[0];
        if (activitiesTableBodyP2Cloned) {
            activitiesTableBodyP2Cloned.innerHTML = '';
        }
    }


    const day0Page1Container = getElementByIdSafe('rdo_day_0_page_0');
    if (!day0Page1Container) {
        console.error("Critical: Day 0 Page 1 container not found. Cannot source global master values for new day. Aborting add.");
        rdoDayCounter--;
        return;
    }

    let prevDayPage1Container = null;
    const prevDayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${prevDayIndex}`);
    if(prevDayWithButtonLayoutWrapper) {
        const prevRdoDayWrapperInsideLayout = prevDayWithButtonLayoutWrapper.querySelector(`.rdo-day-wrapper`);
        if(prevRdoDayWrapperInsideLayout) {
             prevDayPage1Container = prevRdoDayWrapperInsideLayout.querySelector('.rdo-container');
        }
    } else {
        const prevRdoDayWrapperDirect = getElementByIdSafe(`rdo_day_${prevDayIndex}_wrapper`);
        if (prevRdoDayWrapperDirect) {
            prevDayPage1Container = prevRdoDayWrapperDirect.querySelector('.rdo-container');
        }
    }

    if (!prevDayPage1Container) { 
        console.error(`Critical: Previous day's page 1 container for date/RDO calculation (day ${prevDayIndex}) not found. Aborting add.`);
        rdoDayCounter--;
        return;
    }


    const prevData = getInputValue(`data_day${prevDayIndex}`, prevDayPage1Container); 
    const { nextDate, dayOfWeek } = getNextDateInfo(prevData); 
    
    const newDayPageContainers = templateDayWrapper.querySelectorAll('.rdo-container');
    const newDayPage1Container = newDayPageContainers[0]; 

    if (!newDayPage1Container) {
        console.error("Cloned day wrapper is missing its Page 1 container for new day " + newDayIndex + ". Aborting add.");
        rdoDayCounter--;
        return;
    }

    setInputValue(`data_day${newDayIndex}`, nextDate, newDayPage1Container);
    setInputValue(`dia_semana_day${newDayIndex}`, dayOfWeek, newDayPage1Container);
    
    const prevRdoNumFull = getInputValue(`numero_rdo_day${prevDayIndex}`, prevDayPage1Container); 
    const prevRdoBaseNumStr = prevRdoNumFull.split('-')[0]; 
    const prevRdoBaseNum = parseInt(prevRdoBaseNumStr || '0');
    const newRdoBaseNum = isNaN(prevRdoBaseNum) ? 1 : prevRdoBaseNum + 1; 
    setInputValue(`numero_rdo_day${newDayIndex}`, `${newRdoBaseNum}-A`, newDayPage1Container); 
    
    globalMasterFieldBases.forEach(baseId => {
        if (!['data', 'dia_semana', 'decorridos', 'restantes'].includes(baseId)) { 
            const valFromDay0 = getInputValue(`${baseId}_day0`, day0Page1Container);
            const targetInputOnNewDayPage1 = getElementByIdSafe(`${baseId}_day${newDayIndex}`, newDayPage1Container);
            if (targetInputOnNewDayPage1) {
                targetInputOnNewDayPage1.value = valFromDay0;
            }
        }
    });
    
    resetDayInstanceContent(templateDayWrapper, newDayIndex); 

    const allDaysWrapper = getElementByIdSafe('rdo-all-days-wrapper');
    if (allDaysWrapper) {
        allDaysWrapper.appendChild(templateDayWrapper); 
        await initializeDayInstance(newDayIndex, templateDayWrapper); 

        const effectiveWrapperToObserve = document.getElementById(`day_with_button_container_day${newDayIndex}`) || templateDayWrapper;

        if (navObserver) {
            navObserver.observe(effectiveWrapperToObserve);
        } else {
            setupNavigationObserver(); 
        }

        await updateAllDateRelatedFieldsFromDay0(); 
                                                    
        const masterSignaturePreview = getElementByIdSafe('consbemPhotoPreview_day0');
        let day0ImageUrl = null;
        if (masterSignaturePreview && masterSignaturePreview.dataset.isEffectivelyLoaded === 'true') {
            day0ImageUrl = masterSignaturePreview.src;
        }
        updateAllSignaturePreviews(day0ImageUrl); 
        markFormAsDirty(); 
        updateClearEfetivoButtonVisibility();
        
    } else {
        console.error("#rdo-all-days-wrapper not found. Cannot append new day.");
        rdoDayCounter--; 
    }
}


async function removeLastRdoDay() {
    if (rdoDayCounter < 1) return; 

    const lastDayIndex = rdoDayCounter;
    
    const dayWithButtonContainerToRemove = document.getElementById(`day_with_button_container_day${lastDayIndex}`);
    const rdoDayWrapperItself = getElementByIdSafe(`rdo_day_${lastDayIndex}_wrapper`);
    
    let elementToRemoveFromObserver = null;
    let elementToRemoveFromDOM = null;

    if (dayWithButtonContainerToRemove) {
        elementToRemoveFromObserver = dayWithButtonContainerToRemove;
        elementToRemoveFromDOM = dayWithButtonContainerToRemove;
    } else if (rdoDayWrapperItself) {
        elementToRemoveFromObserver = rdoDayWrapperItself;
        elementToRemoveFromDOM = rdoDayWrapperItself;
    }


    const tabToRemove = getElementByIdSafe(`day-tab-${lastDayIndex}`);
    
    if (elementToRemoveFromObserver && navObserver) {
        navObserver.unobserve(elementToRemoveFromObserver);
    }
    if (elementToRemoveFromDOM) {
        elementToRemoveFromDOM.remove();
    } else {
        console.warn(`Could not find wrapper to remove for day ${lastDayIndex}`);
    }


    if (tabToRemove) {
        tabToRemove.remove();
        dayTabs.pop(); 
    }

    rdoDayCounter--;

    if (tabToRemove && tabToRemove.classList.contains('active')) {
        setActiveTab(dayTabs.length - 1);
    } else if (dayTabs.length > 0 && !dayTabsContainer?.querySelector('.day-tab.active')) {
        setActiveTab(dayTabs.length -1);
    }
    markFormAsDirty(); 
    updateClearEfetivoButtonVisibility();
}


async function adjustRdoDaysForMonth(baseDateStrYYYYMMDD) {
    if (!baseDateStrYYYYMMDD || !/^\d{4}-\d{2}-\d{2}$/.test(baseDateStrYYYYMMDD)) {
        console.warn("adjustRdoDaysForMonth: Invalid or empty date format provided:", baseDateStrYYYYMMDD);
        return;
    }
    const dateParts = baseDateStrYYYYMMDD.split('-'); 
    const year = parseInt(dateParts[0]);
    const month_1_indexed = parseInt(dateParts[1]);

    const numDaysInTargetMonth = getDaysInMonth(year, month_1_indexed);

    if (isNaN(numDaysInTargetMonth) || numDaysInTargetMonth <= 0) {
        console.warn(`adjustRdoDaysForMonth: Could not determine days in month for ${year}-${month_1_indexed}.`);
        return;
    }

    let currentNumRdos = rdoDayCounter + 1;
    let daysAdjusted = false;

    while (currentNumRdos < numDaysInTargetMonth) {
        await addNewRdoDay(); 
        currentNumRdos = rdoDayCounter + 1;
        daysAdjusted = true;
    }

    while (currentNumRdos > numDaysInTargetMonth && currentNumRdos > 1) { 
        await removeLastRdoDay(); 
        currentNumRdos = rdoDayCounter + 1;
        daysAdjusted = true;
    }

    await updateAllDateRelatedFieldsFromDay0(); 
    
    const masterSignaturePreview = getElementByIdSafe('consbemPhotoPreview_day0');
    let day0ImageUrl = null;
    if (masterSignaturePreview && masterSignaturePreview.dataset.isEffectivelyLoaded === 'true') {
        day0ImageUrl = masterSignaturePreview.src;
    }
    updateAllSignaturePreviews(day0ImageUrl);

    setupNavigationObserver(); 

    const activeTabIndex = dayTabs.findIndex(tab => tab.classList.contains('active'));
    if (numDaysInTargetMonth > 0) {
        if (activeTabIndex >= numDaysInTargetMonth || activeTabIndex === -1) {
            setActiveTab(Math.min(numDaysInTargetMonth - 1, Math.max(0, activeTabIndex)));
        }
    } else if (dayTabs.length > 0) { 
        setActiveTab(0);
    }

    if (dayTabs.length > 0 && !dayTabsContainer?.querySelector('.day-tab.active')) {
        setActiveTab(0); 
    }
    if (daysAdjusted) {
        markFormAsDirty();
    }
    updateClearEfetivoButtonVisibility();
}


// --- PDF Generation Constants & Helpers ---
const margin = 7; 
const pageWidth = 210; 
const pageHeight = 297; 
const contentWidth = pageWidth - (2 * margin); 
const lineWeight = 0.1; 
const lineWeightStrong = 0.1; 
const photoFrameLineWeight = 0.1; 

async function imageUrlToDataUrl(url) {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Failed to fetch image ${url}: ${response.statusText}`);
    const blob = await response.blob();
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

const addImageToPdfAsync = async (pdf, selector, x, y, w, h, parent) => {
    const imgElement = parent.querySelector(selector);
    if (imgElement && imgElement.src && !imgElement.src.endsWith('#') && !imgElement.src.startsWith('about:blank')) {
        const isStaticHeaderLogo = imgElement.classList.contains('sabesp-logo') || imgElement.classList.contains('consorcio-logo');
        if (isStaticHeaderLogo || imgElement.dataset.isEffectivelyLoaded === 'true') {
            try {
                let imageDataUrl = imgElement.src;
                if (!imageDataUrl.startsWith('data:image')) {
                    try { imageDataUrl = await imageUrlToDataUrl(imgElement.src); } 
                    catch (fetchError) { console.error(`Failed to convert image URL ${imgElement.src}:`, fetchError); pdf.text(`[Img Load Err]`, x + 2, y + h / 2); return; }
                }
                
                let naturalW = 0, naturalH = 0;

                const imgNaturalWidthAttr = imgElement.dataset.naturalWidth;
                const imgNaturalHeightAttr = imgElement.dataset.naturalHeight;

                if (imgNaturalWidthAttr && imgNaturalHeightAttr) {
                    naturalW = parseFloat(imgNaturalWidthAttr);
                    naturalH = parseFloat(imgNaturalHeightAttr);
                } else { 
                    const tempImgLoad = new Image();
                    tempImgLoad.src = imageDataUrl;
                    await new Promise((resolve, reject) => { 
                         tempImgLoad.onload = () => resolve(); 
                         tempImgLoad.onerror = (err) => { console.error(`Temp Img Load Error for ${selector}`, err); reject(err); }; 
                    });
                    naturalW = tempImgLoad.naturalWidth;
                    naturalH = tempImgLoad.naturalHeight;
                    if (naturalW && naturalH && imgElement.dataset) {
                        imgElement.dataset.naturalWidth = naturalW.toString();
                        imgElement.dataset.naturalHeight = naturalH.toString();
                    }
                }
                
                if(!naturalW || !naturalH) { 
                    try {
                        const props = pdf.getImageProperties(imageDataUrl);
                        naturalW = props.width;
                        naturalH = props.height;
                    } catch (e) {
                         console.error(`Error getting image properties for ${selector} via jsPDF:`, e);
                         pdf.text(`[Img Meta Err]`, x + 2, y + h / 2); return;
                    }
                }
                 if(!naturalW || !naturalH) { 
                    console.error(`Could not determine dimensions for image ${selector}. Skipping addImage.`);
                    pdf.text(`[Img Dim Err]`, x + 2, y + h / 2); return;
                }

                let format = 'JPEG'; 
                const mimeTypeMatch = imageDataUrl.match(/^data:image\/([a-zA-Z]+);base64,/);
                if (mimeTypeMatch && mimeTypeMatch[1]) { format = mimeTypeMatch[1].toUpperCase(); } 
                else { 
                    try { 
                        const propsForFormat = pdf.getImageProperties(imageDataUrl);
                        format = propsForFormat.fileType || 'JPEG'; 
                    } catch (e) {
                        console.warn(`Could not determine image type for ${selector} via jsPDF, defaulting to JPEG. Error:`, e);
                        format = 'JPEG';
                    }
                }
                const supportedFormats = ['PNG', 'JPEG', 'WEBP']; 
                if (!supportedFormats.includes(format)) { console.warn(`Unsupported image format ${format} for ${selector}, defaulting to JPEG.`); format = 'JPEG';}


                if (imgElement.classList.contains('report-photo-image-pt2')) {
                    const htmlImg = new Image();
                    htmlImg.crossOrigin = "anonymous"; 
                    
                    const imgLoadPromise = new Promise((resolve, reject) => {
                        htmlImg.onload = () => resolve();
                        htmlImg.onerror = (err) => {
                            console.error(`Failed to load image into HTMLImageElement for canvas processing: ${selector}`, err);
                            reject(new Error(`HTMLImageElement load error for ${selector}`));
                        };
                        htmlImg.src = imageDataUrl;
                    });
                    await imgLoadPromise; 

                    const dpi = 150; 
                    const cvWidthPx = (w / 25.4) * dpi; 
                    const cvHeightPx = (h / 25.4) * dpi; 

                    const tempCanvas = document.createElement('canvas');
                    tempCanvas.width = cvWidthPx;
                    tempCanvas.height = cvHeightPx;
                    const ctx = tempCanvas.getContext('2d');

                    if (!ctx) { 
                        console.error(`Could not get 2D context for pre-cropping image ${selector}. Falling back to direct addImage with clipping (might not be perfect 'cover').`);
                        const imageAspectRatioFallback = naturalW / naturalH;
                        const boxAspectRatioFallback = w / h;
                        let drawWfb, drawHfb, drawXfb, drawYfb;
                        if (imageAspectRatioFallback > boxAspectRatioFallback) { 
                            drawHfb = h; drawWfb = h * imageAspectRatioFallback; drawXfb = x + (w - drawWfb) / 2; drawYfb = y;
                        } else { 
                            drawWfb = w; drawHfb = w / imageAspectRatioFallback; drawXfb = x; drawYfb = y + (h - drawHfb) / 2;
                        }
                        pdf.saveGraphicsState();
                        // pdf.clip(); // This was causing issues with jspdf types, and might not be available/working as expected.
                        // For a simpler 'cover' like effect with addImage, we calculate source clipping manually if possible, or adjust placement and size.
                        // The existing complex logic for canvas pre-crop is better. If it fails, a simpler contain/fit is safer than a potentially broken clip.
                        pdf.addImage(imageDataUrl, format, drawXfb, drawYfb, drawWfb, drawHfb);
                        pdf.restoreGraphicsState();
                    } else {
                        const imageAspect = naturalW / naturalH;
                        const canvasAspect = cvWidthPx / cvHeightPx; 

                        let srcX = 0, srcY = 0, srcW = naturalW, srcH = naturalH; 

                        if (imageAspect > canvasAspect) { 
                            srcH = naturalH;
                            srcW = naturalH * canvasAspect; 
                            srcX = (naturalW - srcW) / 2;   
                        } else { 
                            srcW = naturalW;
                            srcH = naturalW / canvasAspect; 
                            srcY = (naturalH - srcH) / 2;   
                        }
                        ctx.drawImage(htmlImg, srcX, srcY, srcW, srcH, 0, 0, cvWidthPx, cvHeightPx);
                        const croppedImageDataUrl = tempCanvas.toDataURL(format === 'PNG' ? 'image/png' : 'image/jpeg', 0.9); 
                        pdf.addImage(croppedImageDataUrl, format === 'PNG' ? 'PNG' : 'JPEG', x, y, w, h); 
                    }
                } else { 
                    const imageAspectRatio = naturalW / naturalH;
                    const boxAspectRatio = w / h;
                    let drawW, drawH, drawX, drawY; 
                    if (imageAspectRatio > boxAspectRatio) { 
                        drawW = w; 
                        drawH = w / imageAspectRatio;
                    } else { 
                        drawH = h; 
                        drawW = h * imageAspectRatio;
                    }
                    drawX = x + (w - drawW) / 2; 
                    drawY = y + (h - drawH) / 2; 
                    pdf.addImage(imageDataUrl, format, drawX, drawY, drawW, drawH);
                }

            } catch (e) { console.error("Error adding image " + selector + " to PDF:", e); pdf.text(`[Img Proc Err]`, x + 2, y + h / 2); }
        }
    }
};


async function drawPdfHeader(pdf, container, currentY) {
    const headerElement = container.querySelector('.header');
    if (!headerElement) return currentY; 
    const panelTopY = currentY, headerPanelHeight = 16; 
    const logoHeight = 12, logoSabespWidth = 25, logoConsorcioWidth = 55, logoConsorcioHeight = 16; 
    
    pdf.setLineWidth(lineWeightStrong);
    pdf.line(margin, panelTopY, margin + contentWidth, panelTopY); 
    pdf.line(margin, panelTopY, margin, panelTopY + headerPanelHeight); 
    pdf.line(margin + contentWidth, panelTopY, margin + contentWidth, panelTopY + headerPanelHeight); 
    
    const internalPaddingX = 1; 
    const sabespLogoY = panelTopY + 2; 
    await addImageToPdfAsync(pdf, '.sabesp-logo', margin - 3, sabespLogoY, logoSabespWidth, logoHeight, headerElement);
    
    const consorcioLogoX = pageWidth - margin - logoConsorcioWidth - internalPaddingX; 
    const consorcioLogoY = panelTopY + (headerPanelHeight - logoConsorcioHeight) / 2; 
    await addImageToPdfAsync(pdf, '.consorcio-logo', consorcioLogoX, consorcioLogoY, logoConsorcioWidth, logoConsorcioHeight, headerElement);
    
    const titleCenterX = pageWidth / 2 - 20; 
    pdf.setFontSize(10); pdf.setFont('helvetica', 'bold');
    pdf.text("RDO - RELATÓRIO DIÁRIO DE OBRAS", titleCenterX, panelTopY + 5.6, { align: 'center' }); 
    pdf.setFontSize(7);
    pdf.text("SABESP - COMPANHIA DE SANEAMENTO BÁSICO DO ESTADO DE SÃO PAULO", titleCenterX, panelTopY + 9.1, { align: 'center' }); 
    pdf.text("AMPLIAÇÃO DA CAPACIDADE DE TRATAMENTO DA FASE SÓLIDA-ETE BARUERI PARA 16M³/S", titleCenterX, panelTopY + 12.1, { align: 'center' }); 
    
    return panelTopY + headerPanelHeight; 
}

function drawContractInfo(pdf, pageContainer, currentY, dayIndex, pageInDayIdx) {
    const daySuffix = `_day${dayIndex}`;
    let idSuffixForPageFields = daySuffix; 
    let rdoNumIdForPage = `numero_rdo${daySuffix}`; 

    if (pageInDayIdx === 1) { 
        idSuffixForPageFields = `_pt${daySuffix}`;
        rdoNumIdForPage = `numero_rdo_pt${daySuffix}`;
    } else if (pageInDayIdx === 2) { 
        idSuffixForPageFields = `_pt2${daySuffix}`;
        rdoNumIdForPage = `numero_rdo_pt2${daySuffix}`;
    }

    const rdoNumero = getInputValue(rdoNumIdForPage, pageContainer);
    const contratada = getInputValue(`contratada${idSuffixForPageFields}`, pageContainer);
    const dataVal = getInputValue(`data${idSuffixForPageFields}`, pageContainer); 
    let displayData = dataVal; 
    if (dataVal.match(/^\d{4}-\d{2}-\d{2}$/)) { 
        const dateParts = dataVal.split('-');
        const dateObj = parseDateDDMMYYYY(`${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`); 
        if (dateObj) {
            displayData = `${String(dateObj.getDate()).padStart(2, '0')}/${String(dateObj.getMonth() + 1).padStart(2, '0')}/${dateObj.getFullYear()}`;
        }
    }

    const contratoNum = getInputValue(`contrato_num${idSuffixForPageFields}`, pageContainer);
    const diaSemana = getInputValue(`dia_semana${idSuffixForPageFields}`, pageContainer);
    const prazo = getInputValue(`prazo${idSuffixForPageFields}`, pageContainer); 
    const inicioObra = getInputValue(`inicio_obra${idSuffixForPageFields}`, pageContainer); 
    const decorridos = getInputValue(`decorridos${idSuffixForPageFields}`, pageContainer);
    const restantes = getInputValue(`restantes${idSuffixForPageFields}`, pageContainer);

    const boxHeight = 17;
    const mainInfoWidth = contentWidth - 33; 
    const rdoBoxWidth = 33;
    const boxTopY = currentY;
    const boxBottomY = currentY + boxHeight;

    pdf.setLineWidth(lineWeight); 
    pdf.line(margin, boxTopY, margin + contentWidth, boxTopY);

    pdf.setLineWidth(lineWeightStrong);
    pdf.line(margin, boxTopY, margin, boxBottomY); 
    pdf.line(margin + mainInfoWidth, boxTopY, margin + mainInfoWidth, boxBottomY); 
    pdf.line(margin, boxBottomY, margin + mainInfoWidth, boxBottomY); 

    pdf.setLineWidth(lineWeightStrong);
    const rdoBoxRealX = margin + mainInfoWidth;
    pdf.line(rdoBoxRealX + rdoBoxWidth, boxTopY, rdoBoxRealX + rdoBoxWidth, boxBottomY); 
    pdf.line(rdoBoxRealX, boxBottomY, rdoBoxRealX + rdoBoxWidth, boxBottomY); 
    
    pdf.setFontSize(7); pdf.setFont('helvetica', 'normal');
    const rowHeight = boxHeight / 3, textBaselineOffset = rowHeight - 1.5; 
    pdf.setLineWidth(lineWeight); 

    pdf.text(`CONTRATADA: ${contratada}`, margin + 1.5, currentY + textBaselineOffset);
    pdf.line(margin + mainInfoWidth * 0.65, currentY, margin + mainInfoWidth * 0.65, currentY + rowHeight); 
    pdf.text(`DATA: ${displayData}`, margin + mainInfoWidth * 0.65 + 1.5, currentY + textBaselineOffset);
    
    pdf.line(margin, currentY + rowHeight, margin + mainInfoWidth, currentY + rowHeight); 
    pdf.text(`CONTRATO Nº: ${contratoNum}`, margin + 1.5, currentY + rowHeight + textBaselineOffset);
    pdf.line(margin + mainInfoWidth * 0.65, currentY + rowHeight, margin + mainInfoWidth * 0.65, currentY + rowHeight * 2); 
    pdf.text(`DIA DA SEMANA: ${diaSemana}`, margin + mainInfoWidth * 0.65 + 1.5, currentY + rowHeight + textBaselineOffset);
    
    pdf.line(margin, currentY + rowHeight * 2, margin + mainInfoWidth, currentY + rowHeight * 2); 
    const textYRow3 = currentY + rowHeight * 2 + textBaselineOffset;
    const segmentWidth = (mainInfoWidth - 3) / 4; 
    pdf.text(`PRAZO: ${prazo}`, margin + 1.5, textYRow3);
    pdf.text(`INÍCIO: ${inicioObra}`, margin + 1.5 + segmentWidth, textYRow3);
    pdf.text(`DECORRIDOS: ${decorridos}`, margin + 1.5 + (2 * segmentWidth), textYRow3);
    pdf.text(`RESTANTES: ${restantes}`, margin + 1.5 + (3 * segmentWidth), textYRow3);
    
    const rdoPdfBoxX = margin + mainInfoWidth; 
    pdf.setFontSize(8); pdf.setFont('helvetica', 'bold');
    pdf.text("Número:", rdoPdfBoxX + rdoBoxWidth / 2, currentY + 4.5, { align: 'center' });
    pdf.setFontSize(20); pdf.setFont('helvetica', 'bold'); pdf.setTextColor(255, 0, 0); 
    pdf.text(rdoNumero, rdoPdfBoxX + rdoBoxWidth / 2, currentY + 11.5, { align: 'center' }); 
    pdf.setTextColor(0, 0, 0); pdf.setFont('helvetica', 'normal'); 
    
    return currentY + boxHeight;
}

function drawWeatherConditions(pdf, pageContainer, currentY, dayIndex) {
    const weatherEl = pageContainer.querySelector('.weather-conditions');
    if (!weatherEl) return currentY;
    const daySuffix = `_day${dayIndex}`;
    const sectionHeight = 25, titleBgColor = '#E0E0E0', titleHeight = 5;
    const modeloPanelWidth = 35, turnosObsCombinedWidth = contentWidth - modeloPanelWidth;
    const turnosTitleSectionWidth = turnosObsCombinedWidth * 0.45, obsTitleSectionWidth = turnosObsCombinedWidth * 0.55;
    
    const modeloPanelX = margin;
    const modeloPanelY = currentY;
    const turnosObsBaseX = modeloPanelX + modeloPanelWidth;

    // 1. Fill title bar backgrounds
    pdf.setFillColor(titleBgColor);
    pdf.rect(modeloPanelX, modeloPanelY, modeloPanelWidth, titleHeight, 'F'); // Modelo title background
    pdf.rect(turnosObsBaseX, modeloPanelY, turnosObsCombinedWidth, titleHeight, 'F'); // Turnos/Obs title background

    // 2. Draw outer borders for the entire weather section
    pdf.setLineWidth(lineWeightStrong);
    // Top outer border (covers both titles)
    pdf.line(modeloPanelX, modeloPanelY, turnosObsBaseX + turnosObsCombinedWidth, modeloPanelY);
    // Bottom outer border (covers both content areas)
    pdf.line(modeloPanelX, modeloPanelY + sectionHeight, turnosObsBaseX + turnosObsCombinedWidth, modeloPanelY + sectionHeight);
    // Left outer border (full height of section)
    pdf.line(modeloPanelX, modeloPanelY, modeloPanelX, modeloPanelY + sectionHeight);
    // Right outer border (full height of section)
    pdf.line(turnosObsBaseX + turnosObsCombinedWidth, modeloPanelY, turnosObsBaseX + turnosObsCombinedWidth, modeloPanelY + sectionHeight);

    // 3. Draw the horizontal line separating titles from content (covers both titles)
    pdf.setLineWidth(lineWeightStrong); 
    pdf.line(modeloPanelX, modeloPanelY + titleHeight, turnosObsBaseX + turnosObsCombinedWidth, modeloPanelY + titleHeight);
    
    // 4. Draw the vertical separator line ONLY for the content area (between Modelo content and Turnos/Obs content)
    pdf.setLineWidth(lineWeight); // Using normal lineWeight for internal division
    pdf.line(modeloPanelX + modeloPanelWidth, modeloPanelY + titleHeight, modeloPanelX + modeloPanelWidth, modeloPanelY + sectionHeight);
    
    // Content of Modelo Panel
    const modeloContentY = modeloPanelY + titleHeight;
    const modeloContentHeight = sectionHeight - titleHeight;
    const yShiftDown = 0.75; 
    pdf.setFont('helvetica', 'bold'); pdf.setFontSize(9); pdf.setTextColor(0,0,0);
    pdf.text("Modelo", modeloPanelX + 10, modeloContentY + modeloContentHeight / 2 + 6, { align: 'center', angle: 90 }); 
    
    const labelsX = modeloPanelX + 8; 
    pdf.setFont('helvetica', 'normal'); pdf.setFontSize(7);
    pdf.text("Tempo", labelsX, modeloContentY + (modeloContentHeight / 3) + yShiftDown, { maxWidth: 15 });
    pdf.text("Trabalho", labelsX, modeloContentY + (modeloContentHeight / 3 * 2) + yShiftDown, { maxWidth: 15 });
    
    const staticCircleRadius = 6, staticCircleCenterX = (modeloPanelX + modeloPanelWidth) - 3 - staticCircleRadius;
    const staticCircleCenterY = modeloContentY + modeloContentHeight / 2 + yShiftDown;
    const circleTextBaselineOffsetY = 1.8; 

    pdf.setLineWidth(lineWeight); // Circles use normal line weight
    pdf.circle(staticCircleCenterX, staticCircleCenterY, staticCircleRadius, 'S'); 
    pdf.line(staticCircleCenterX - staticCircleRadius, staticCircleCenterY, staticCircleCenterX + staticCircleRadius, staticCircleCenterY); 
    pdf.setFontSize(10); pdf.setFont('helvetica', 'bold'); 
    pdf.text("B", staticCircleCenterX, staticCircleCenterY - staticCircleRadius / 2 + circleTextBaselineOffsetY, { align: 'center' }); 
    pdf.text("N", staticCircleCenterX, staticCircleCenterY + staticCircleRadius / 2 + circleTextBaselineOffsetY, { align: 'center' }); 

    // Content of Turnos/Obs Panel
    pdf.setTextColor(0,0,0); pdf.setFontSize(9); pdf.setFont('helvetica', 'bold'); // Matched CSS title font size
    pdf.text('CONDIÇÕES CLIMÁTICAS E DE TRABALHO', turnosObsBaseX + turnosTitleSectionWidth / 2, modeloPanelY + 3.5, { align: 'center' });
    
    const indicePluvTitleX = turnosObsBaseX + turnosTitleSectionWidth; 
    pdf.setFont('helvetica', 'bold'); pdf.setFontSize(7); 
    const indicePluv = getInputValue(`indice_pluv_valor${daySuffix}`, weatherEl); 
    pdf.text(`Índice Pluviométrico: ${indicePluv} mm`, indicePluvTitleX + 1.5, modeloPanelY + 3.2, { maxWidth: obsTitleSectionWidth - 3 });
    pdf.setFont('helvetica', 'normal'); pdf.setFontSize(7); 
    
    const turnosObsContentY = modeloPanelY + titleHeight;
    const turnosObsContentHeight = sectionHeight - titleHeight; 
    const turnoDetailHeaderHeight = 7, turnoLabelTextYBaselineOffset = 3.0, turnoTimeTextYBaselineOffset = 6.0;  
    
    const turnoLabelCells = weatherEl.querySelectorAll('.turno-label-cell');
    const turnosDataFromDOM = Array.from(turnoLabelCells).slice(0, 3).map(thCell => {
        const label = Array.from(thCell.childNodes).find(node => node.nodeType === Node.TEXT_NODE)?.textContent?.trim() || '';
        const time = thCell.querySelector('.turno-time-cell')?.textContent?.trim() || '';
        return { label, time };
    });
    
    const turnoColWidth = turnosTitleSectionWidth / 3; 
    turnosDataFromDOM.forEach((turnoData, index) => {
        const turnoColActualX = turnosObsBaseX + index * turnoColWidth;
        pdf.setFontSize(7); pdf.setFont('helvetica', 'bold'); 
        pdf.text(turnoData.label, turnoColActualX + turnoColWidth / 2, turnosObsContentY + turnoLabelTextYBaselineOffset, { align: 'center' });
        if (turnoData.time) {
            pdf.setFontSize(6); pdf.setFont('helvetica', 'normal'); 
            pdf.text(turnoData.time, turnoColActualX + turnoColWidth / 2, turnosObsContentY + turnoTimeTextYBaselineOffset, { align: 'center' });
        }
    });
    
    const circleAreaYStart = turnosObsContentY + turnoDetailHeaderHeight;
    const circleAreaHeight = turnosObsContentHeight - turnoDetailHeaderHeight; 
    const dynamicCircleRadius = 6;
    const turnosControlValueGetters = [ 
        { cellIdBase: 'turno1_control_cell', turnoIdBase: 't1' }, 
        { cellIdBase: 'turno2_control_cell', turnoIdBase: 't2' }, 
        { cellIdBase: 'turno3_control_cell', turnoIdBase: 't3' }
    ];
    turnosControlValueGetters.forEach((turnoCtrl, index) => {
        const turnoColActualX = turnosObsBaseX + index * turnoColWidth;
        const topValue = getInputValue(`tempo_${turnoCtrl.turnoIdBase}${daySuffix}_hidden`, pageContainer);
        const bottomValue = getInputValue(`trabalho_${turnoCtrl.turnoIdBase}${daySuffix}_hidden`, pageContainer);
        
        const circleCenterX = turnoColActualX + turnoColWidth / 2;
        const circleCenterY = circleAreaYStart + (circleAreaHeight / 2);
        pdf.setLineWidth(lineWeight); // Circles use normal line weight
        pdf.circle(circleCenterX, circleCenterY, dynamicCircleRadius, 'S'); 
        pdf.line(circleCenterX - dynamicCircleRadius, circleCenterY, circleCenterX + dynamicCircleRadius, circleCenterY); 
        pdf.setFontSize(10); pdf.setFont('helvetica', 'bold'); 
        pdf.text(topValue, circleCenterX, circleCenterY - dynamicCircleRadius / 2 + circleTextBaselineOffsetY, { align: 'center' });
        pdf.text(bottomValue, circleCenterX, circleCenterY + dynamicCircleRadius / 2 + circleTextBaselineOffsetY, { align: 'center' });
    });
    
    pdf.setLineWidth(lineWeight); 
    pdf.line(indicePluvTitleX, turnosObsContentY, indicePluvTitleX, turnosObsContentY + turnosObsContentHeight); // Separator for obs_clima
    
    pdf.setFontSize(7); pdf.setFont('helvetica', 'normal');
    const obsClima = getTextAreaValue(`obs_clima${daySuffix}`, weatherEl); 
    pdf.text(`Obs: ${obsClima}`, indicePluvTitleX + 1.5, turnosObsContentY + 3.5, { maxWidth: obsTitleSectionWidth - 3, align: 'left' });
    
    return currentY + sectionHeight; 
}

function drawLegend(pdf, pageContainer, currentY) {
    const legendElement = pageContainer.querySelector('.legend');
    if (!legendElement) return currentY;
    
    const legendText = (legendElement.textContent || '').replace(/\s\s+/g, ' ').trim(); 
    const textHorizontalPadding = 2, textVerticalPadding = 1.5, legendFontSize = 6.5;
    
    pdf.setFontSize(legendFontSize); pdf.setFont('helvetica', 'normal');
    const textLines = pdf.splitTextToSize(legendText, contentWidth - (2 * textHorizontalPadding));
    const textDimensions = pdf.getTextDimensions(textLines, { fontSize: legendFontSize, maxWidth: contentWidth - (2 * textHorizontalPadding) });
    const legendTextActualHeight = textDimensions.h;
    const totalSectionHeight = legendTextActualHeight + (2 * textVerticalPadding); 
    
    pdf.setLineWidth(lineWeightStrong); pdf.rect(margin, currentY, contentWidth, totalSectionHeight, 'S'); 
    
    const ascent = (legendFontSize * (25.4 / 72)) * 0.8; 
    let firstLineBaselineY = currentY + textVerticalPadding + ascent;
    const singleLineHeight = legendTextActualHeight / textLines.length; 
    
    const prefixToBold = "Legendas:"; 
    for (let i = 0; i < textLines.length; i++) {
        const line = textLines[i], lineY = firstLineBaselineY + (i * singleLineHeight);
        let currentTextX = margin + textHorizontalPadding;
        
        if (i === 0 && line.startsWith(prefixToBold)) {
            const boldPart = prefixToBold;
            const normalPart = line.substring(boldPart.length);
            pdf.setFont('helvetica', 'bold'); pdf.setFontSize(legendFontSize);
            pdf.text(boldPart, currentTextX, lineY);
            const boldPartWidth = pdf.getTextWidth(boldPart); currentTextX += boldPartWidth;
            pdf.setFont('helvetica', 'normal'); pdf.setFontSize(legendFontSize);
            pdf.text(normalPart, currentTextX, lineY, { maxWidth: contentWidth - (2 * textHorizontalPadding) - boldPartWidth });
        } else { 
            pdf.setFont('helvetica', 'normal'); pdf.setFontSize(legendFontSize);
            pdf.text(line, currentTextX, lineY, { maxWidth: contentWidth - (2 * textHorizontalPadding) });
        }
    }
    return currentY + totalSectionHeight;
}

function drawLaborColumns(pdf, pageContainer, currentY, dayIndex) {
    const laborSectionEl = pageContainer.querySelector('.labor-section');
    if (!laborSectionEl) return currentY;
    
    const columnsConfig = [ 
        { selector: '.labor-direta', proportion: 1, isMultiSubColumn: false, totalLabel: "Total M.O.D = " },
        { selector: '.labor-indireta', proportion: 2, isMultiSubColumn: true, totalLabel: "Total M.O.I = " }, 
        { selector: '.equipamentos', proportion: 2, isMultiSubColumn: true, totalLabel: "Total Equip. = " }  
    ];
    const totalProportions = columnsConfig.reduce((sum, col) => sum + col.proportion, 0);
    
    const titleHeight = 5, itemRowHeight = 3.5, itemsPerVisualRowStack = 18; 
    const itemFontSize = 6, textBaselineInRow = itemRowHeight - 1.285; 
    const qtyDisplayWidthInPdf = 10, itemNamePaddingLeft = 2; 
    
    pdf.setFont('helvetica', 'normal');
    let columnSectionStartY = currentY;
    const colContentsHeight = itemsPerVisualRowStack * itemRowHeight; 
    const totalRowHeight = 5; 
    const totalSectionHeight = titleHeight + colContentsHeight + totalRowHeight; 
    
    pdf.setLineWidth(lineWeightStrong);
    pdf.line(margin, columnSectionStartY, margin, columnSectionStartY + totalSectionHeight); 
    pdf.line(margin + contentWidth, columnSectionStartY, margin + contentWidth, columnSectionStartY + totalSectionHeight); 
    pdf.line(margin, columnSectionStartY + totalSectionHeight, margin + contentWidth, columnSectionStartY + totalSectionHeight); 
    
    let currentCumulativeX = margin; 
    
    columnsConfig.forEach((colConfig, idx) => {
        const colEl = laborSectionEl.querySelector(colConfig.selector);
        if (!colEl) return;
        
        const currentColWidth = (contentWidth / totalProportions) * colConfig.proportion; 
        const title = (colEl.querySelector('h4')?.textContent || '').trim(); 
        
        pdf.setLineWidth(lineWeight); 
        pdf.setFillColor('#E0E0E0'); pdf.setFontSize(9); pdf.setFont('helvetica', 'bold'); // Matched CSS title font size
        pdf.rect(currentCumulativeX, columnSectionStartY, currentColWidth, titleHeight, 'FD'); 
        pdf.text(title, currentCumulativeX + currentColWidth / 2, columnSectionStartY + 3.5, { align: 'center' });
        
        let itemAreaY = columnSectionStartY + titleHeight; 
        const table = colEl.querySelector('.labor-table'); 
        pdf.setFontSize(itemFontSize); pdf.setFont('helvetica', 'normal'); pdf.setLineWidth(lineWeight);
        
        const quantityInputs = Array.from(table.querySelectorAll('tbody tr .quantity-input'));
        const itemNames = Array.from(table.querySelectorAll('tbody tr .item-name-cell')).map(el => el.textContent?.trim() || '');
        
        if (colConfig.isMultiSubColumn) { 
            const subColWidth = (currentColWidth - lineWeight) / 2; 
            
            let subColStartX = currentCumulativeX;
            for (let i = 0; i < itemsPerVisualRowStack; i++) {
                if (i >= quantityInputs.length / 2 && i >= itemNames.length / 2) break; 
                const currentItemY = itemAreaY + i * itemRowHeight;
                const qty = quantityInputs[i]?.value || "";
                const name = itemNames[i] || "";
                
                pdf.text(qty, subColStartX + qtyDisplayWidthInPdf - 2, currentItemY + textBaselineInRow, { align: 'right', maxWidth: qtyDisplayWidthInPdf - 4 });
                pdf.line(subColStartX + qtyDisplayWidthInPdf, currentItemY, subColStartX + qtyDisplayWidthInPdf, currentItemY + itemRowHeight); 
                pdf.text(name, subColStartX + qtyDisplayWidthInPdf + itemNamePaddingLeft, currentItemY + textBaselineInRow, { maxWidth: subColWidth - qtyDisplayWidthInPdf - itemNamePaddingLeft - 1 });
                
                const isLastTrueItemInSubColumn = i === Math.floor(Math.min(quantityInputs.length, itemNames.length) / 2) -1 && (Math.min(quantityInputs.length,itemNames.length)/2 <= itemsPerVisualRowStack);
                if (i < itemsPerVisualRowStack - 1 && !isLastTrueItemInSubColumn && i < (Math.min(quantityInputs.length,itemNames.length)/2) -1) {
                     pdf.line(subColStartX, currentItemY + itemRowHeight, subColStartX + subColWidth, currentItemY + itemRowHeight);
                }
            }
            
            subColStartX = currentCumulativeX + subColWidth + lineWeight; 
            for (let i = 0; i < itemsPerVisualRowStack; i++) {
                const itemGlobalIndex = itemsPerVisualRowStack + i; 
                if (itemGlobalIndex >= quantityInputs.length && itemGlobalIndex >= itemNames.length) break; 
                const currentItemY = itemAreaY + i * itemRowHeight;
                const qty = quantityInputs[itemGlobalIndex]?.value || "";
                const name = itemNames[itemGlobalIndex] || "";
                
                pdf.text(qty, subColStartX + qtyDisplayWidthInPdf - 2, currentItemY + textBaselineInRow, { align: 'right', maxWidth: qtyDisplayWidthInPdf - 4 });
                pdf.line(subColStartX + qtyDisplayWidthInPdf, currentItemY, subColStartX + qtyDisplayWidthInPdf, currentItemY + itemRowHeight); 
                pdf.text(name, subColStartX + qtyDisplayWidthInPdf + itemNamePaddingLeft, currentItemY + textBaselineInRow, { maxWidth: subColWidth - qtyDisplayWidthInPdf - itemNamePaddingLeft - 1 });

                const isLastTrueItemInThisSubColumn = (itemsPerVisualRowStack + i) === (Math.min(quantityInputs.length, itemNames.length) -1) && (itemsPerVisualRowStack + i < itemsPerVisualRowStack * 2);

                if (i < itemsPerVisualRowStack - 1 && !isLastTrueItemInThisSubColumn && (itemsPerVisualRowStack + i < (Math.min(quantityInputs.length, itemNames.length) -1 ))) {
                     pdf.line(subColStartX, currentItemY + itemRowHeight, subColStartX + subColWidth, currentItemY + itemRowHeight);
                }
            }
            pdf.line(currentCumulativeX + subColWidth, itemAreaY, currentCumulativeX + subColWidth, itemAreaY + colContentsHeight); 
        } else { 
            for (let i = 0; i < itemsPerVisualRowStack; i++) {
                if (i >= quantityInputs.length && i >= itemNames.length) break; 
                const currentItemY = itemAreaY + i * itemRowHeight;
                const qty = quantityInputs[i]?.value || "";
                const name = itemNames[i] || "";

                pdf.text(qty, currentCumulativeX + qtyDisplayWidthInPdf - 2, currentItemY + textBaselineInRow, { align: 'right', maxWidth: qtyDisplayWidthInPdf - 4 });
                pdf.line(currentCumulativeX + qtyDisplayWidthInPdf, currentItemY, currentCumulativeX + qtyDisplayWidthInPdf, currentItemY + itemRowHeight); 
                pdf.text(name, currentCumulativeX + qtyDisplayWidthInPdf + itemNamePaddingLeft, currentItemY + textBaselineInRow, { maxWidth: currentColWidth - qtyDisplayWidthInPdf - itemNamePaddingLeft - 1 });
                
                const isLastTrueItemOverall = i === Math.min(quantityInputs.length, itemNames.length) -1 && Math.min(quantityInputs.length, itemNames.length) <= itemsPerVisualRowStack;
                if (i < itemsPerVisualRowStack - 1 && !isLastTrueItemOverall && i < Math.min(quantityInputs.length, itemNames.length) -1 ) {
                    pdf.line(currentCumulativeX, currentItemY + itemRowHeight, currentCumulativeX + currentColWidth, currentItemY + itemRowHeight);
                }
            }
        }
        
        const totalRowY = columnSectionStartY + titleHeight + colContentsHeight;
        pdf.setFontSize(7); // Reduced font size for totals
        pdf.setFont('helvetica', 'bold');
        pdf.setLineWidth(lineWeight); 

        pdf.rect(currentCumulativeX, totalRowY, currentColWidth, totalRowHeight, 'S');

        const totalValue = (colEl.querySelector('.section-total-input'))?.value || "0";
        const totalText = colConfig.totalLabel + totalValue;
        
        const textPaddingRight = 2; 
        pdf.text(totalText, currentCumulativeX + currentColWidth - textPaddingRight, totalRowY + 3.5, { 
            align: 'right', 
            maxWidth: currentColWidth - (textPaddingRight * 2) 
        });
        
        if (idx < columnsConfig.length - 1) { 
            pdf.setLineWidth(lineWeightStrong); 
            pdf.line(currentCumulativeX + currentColWidth, columnSectionStartY, currentCumulativeX + currentColWidth, columnSectionStartY + totalSectionHeight);
        }
        currentCumulativeX += currentColWidth;
    });
    return columnSectionStartY + totalSectionHeight; 
}


function drawActivityTable(pdf, pageContainer, currentY, dayIndex, pageInDayIdx, targetTotalHeightForSection) {
    let tableIdBase = "activitiesTable";
    let numRowsToDraw = 24; 
    const sectionTitle = "DESCRIÇÃO DAS ATIVIDADES EXECUTADAS";
    
    if (pageInDayIdx === 1) { 
        tableIdBase = "activitiesTable_pt";
        numRowsToDraw = 22; 
    }
    
    const tableElement = getElementByIdSafe(`${tableIdBase}_day${dayIndex}`, pageContainer);
    if (!tableElement) return currentY;
    
    const tableHead = tableElement.tHead;
    const tableBody = tableElement.tBodies[0];
    if (!tableHead || !tableBody) return currentY;
    
    const titleHeight = 5, headerRowHeight = 5;
    let activityRowHeight = 5.25; 
    const MIN_ACTIVITY_ROW_PDF_HEIGHT = 3.5; 

    if (pageInDayIdx === 0 && targetTotalHeightForSection !== undefined && targetTotalHeightForSection > 0) {
        const heightForRowsArea = targetTotalHeightForSection - titleHeight - headerRowHeight;
        if (heightForRowsArea > 0 && numRowsToDraw > 0) {
            activityRowHeight = Math.max(heightForRowsArea / numRowsToDraw, MIN_ACTIVITY_ROW_PDF_HEIGHT);
        } else if (heightForRowsArea <=0) { 
            activityRowHeight = MIN_ACTIVITY_ROW_PDF_HEIGHT; 
        }
    }
    
    const fontSizeHeader = 7, fontSizeBody = 6.5; 
    
    const colWidthsConfig = [
        { proportion: 24.14, title: 'Localização', dataKey: `localizacao_day${dayIndex}[]` },
        { proportion: 13.79, title: 'Tipo', dataKey: `tipo_servico_day${dayIndex}[]` },
        { proportion: 13.79, title: 'Serviço', dataKey: `servico_desc_day${dayIndex}[]` },
        { proportion: 10.34, title: 'Status', dataKey: `status_day${dayIndex}[]` },
        { proportion: 37.93, title: 'Observações', dataKey: `observacoes_atividade_day${dayIndex}[]` },
    ];

    const totalProportions = colWidthsConfig.reduce((sum, col) => sum + col.proportion, 0);
    const calculatedColWidths = colWidthsConfig.map(c => (contentWidth / totalProportions) * c.proportion);
    
    const initialCurrentY = currentY; 
    
    pdf.setLineWidth(lineWeightStrong); 
    pdf.setFillColor('#E0E0E0'); pdf.setFontSize(8); pdf.setFont('helvetica', 'bold'); 
    pdf.rect(margin, currentY, contentWidth, titleHeight, 'FD'); 
    pdf.text(sectionTitle, margin + contentWidth / 2, currentY + 3.5, { align: 'center' });
    currentY += titleHeight;
    
    const tableHeaderY = currentY;
    pdf.setLineWidth(lineWeight); pdf.setFillColor('#E0E0E0'); pdf.setFontSize(fontSizeHeader); pdf.setFont('helvetica', 'bold');
    pdf.rect(margin, currentY, contentWidth, headerRowHeight, 'FD'); 
    let currentX = margin;
    colWidthsConfig.forEach((col, idx) => {
        pdf.text(col.title, currentX + calculatedColWidths[idx] / 2, currentY + 3.5, { align: 'center', maxWidth: calculatedColWidths[idx] - 2 });
        if (idx < colWidthsConfig.length -1) pdf.line(currentX + calculatedColWidths[idx], currentY, currentX + calculatedColWidths[idx], currentY + headerRowHeight);
        currentX += calculatedColWidths[idx];
    });
    currentY += headerRowHeight;
    
    pdf.setFontSize(fontSizeBody); pdf.setFont('helvetica', 'normal');
    pdf.setLineWidth(lineWeight);
    const rows = Array.from(tableBody.rows).slice(0, numRowsToDraw);
    
    const pointToMm = 0.352778;
    const lineHeightFactor = 1.15; 
    const singleLineRenderHeight = fontSizeBody * pointToMm * lineHeightFactor; 

    rows.forEach((row, rowIndex) => {
        const rowY = currentY + rowIndex * activityRowHeight;
        currentX = margin;
        
        // Iterate through cells based on colWidthsConfig to ensure correct order and handling
        colWidthsConfig.forEach((colConfig, cellIndex) => {
            const cellWidth = calculatedColWidths[cellIndex];
            let text = "";

            if (colConfig.dataKey.startsWith(`status_day${dayIndex}`) || 
                colConfig.dataKey.startsWith(`localizacao_day${dayIndex}`) ||
                colConfig.dataKey.startsWith(`tipo_servico_day${dayIndex}`) ||
                colConfig.dataKey.startsWith(`servico_desc_day${dayIndex}`)
                ) {
                const hiddenInput = row.cells[cellIndex]?.querySelector('input[type="hidden"]');
                if (hiddenInput) {
                    text = hiddenInput.value;
                }
            } else { // For remaining textareas like Observações
                const inputEl = row.cells[cellIndex]?.querySelector('input[type="text"], textarea');
                if (inputEl) {
                    text = inputEl.value;
                }
            }

            const textLines = pdf.splitTextToSize(text, cellWidth - 2); 
            
            const textBlockRenderHeight = (textLines.length > 0) ? (textLines.length - 1) * singleLineRenderHeight + (fontSizeBody * pointToMm) : 0;
            const yForFirstLine = rowY + (activityRowHeight - textBlockRenderHeight) / 2 + (fontSizeBody * pointToMm * 0.8); 

            let currentDrawingY = yForFirstLine;
            for (const line of textLines) {
                 if (currentDrawingY < rowY + activityRowHeight - (fontSizeBody*pointToMm*0.2) ) { 
                    pdf.text(line, currentX + 1, currentDrawingY, { maxWidth: cellWidth - 2 });
                 }
                currentDrawingY += singleLineRenderHeight;
            }
            
            if (cellIndex < colWidthsConfig.length - 1) {
                pdf.line(currentX + cellWidth, rowY, currentX + cellWidth, rowY + activityRowHeight);
            }
            currentX += cellWidth;
        });
        
        if (!(pageInDayIdx === 0 && rowIndex === rows.length - 1)) {
            pdf.line(margin, rowY + activityRowHeight, margin + contentWidth, rowY + activityRowHeight);
        }
    });
    
    const finalTableY = currentY + (numRowsToDraw * activityRowHeight);

    pdf.setLineWidth(lineWeightStrong); 
    pdf.line(margin, tableHeaderY, margin, finalTableY); 
    pdf.line(margin + contentWidth, tableHeaderY, margin + contentWidth, finalTableY); 
    
    return initialCurrentY + (titleHeight + headerRowHeight + (numRowsToDraw * activityRowHeight));
}


function drawObservationsSection(
    pdf, 
    pageContainer, 
    currentY, 
    dayIndex,
    targetTotalHeightForSection
) {
    const daySuffix = `_pt_day${dayIndex}`; 
    const obsTextarea = getElementByIdSafe(`observacoes_comentarios${daySuffix}`, pageContainer);
    if (!obsTextarea) return currentY;

    const sectionTitle = "OUTRAS OBSERVAÇÕES E COMENTÁRIOS";
    const titleHeight = 5;
    const fontSizeTitle = 9, fontSizeBody = 7;
    const padding = 1.5; 
    const pdfLineHeightMm = 6; 
    const MIN_OBSERVATIONS_CONTENT_BOX_HEIGHT_PDF = 65; 

    let actualSectionHeight;
    let contentBoxHeight;
    const initialCurrentY = currentY;

    if (targetTotalHeightForSection !== undefined && targetTotalHeightForSection > titleHeight) {
        actualSectionHeight = targetTotalHeightForSection;
        contentBoxHeight = actualSectionHeight - titleHeight;
    } else {
        const obsText = obsTextarea.value;
        pdf.setFontSize(fontSizeBody);
        pdf.setFont('helvetica', 'normal');
        
        const pointToMm = 0.352778;
        const lineHeightFactorObs = 1.15; 
        const singleLineHeightWithFactorObs = fontSizeBody * pointToMm * lineHeightFactorObs;
        const textLines = pdf.splitTextToSize(obsText, contentWidth - (2 * padding));
        const requiredHeightForTextActual = (textLines.length > 0) ? (textLines.length -1) * singleLineHeightWithFactorObs + (fontSizeBody * pointToMm) : 0;
        
        const calculatedContentHeightFromText = requiredHeightForTextActual + (2 * padding);
        contentBoxHeight = Math.max(calculatedContentHeightFromText, MIN_OBSERVATIONS_CONTENT_BOX_HEIGHT_PDF);
        actualSectionHeight = contentBoxHeight + titleHeight;
    }

    pdf.setLineWidth(lineWeightStrong); 
    pdf.setFillColor('#E0E0E0'); 
    pdf.setFontSize(fontSizeTitle); 
    pdf.setFont('helvetica', 'bold');
    pdf.rect(margin, currentY, contentWidth, titleHeight, 'FD'); 
    pdf.text(sectionTitle, margin + contentWidth / 2, currentY + 3.5, { align: 'center' });
    currentY += titleHeight;

    const contentBoxTopY = currentY;

    pdf.setLineWidth(lineWeight); 
    pdf.rect(margin, contentBoxTopY, contentWidth, contentBoxHeight, 'S'); 

    pdf.setLineWidth(lineWeight); 
    pdf.setDrawColor(0, 0, 0); 

    let lineRenderY = contentBoxTopY + pdfLineHeightMm; 
    const contentBoxBottomYForLines = contentBoxTopY + contentBoxHeight;

    while (lineRenderY < contentBoxBottomYForLines - 0.5) { 
        pdf.line(margin, lineRenderY, margin + contentWidth, lineRenderY);
        lineRenderY += pdfLineHeightMm;
    }

    pdf.setDrawColor(0, 0, 0); 
    pdf.setLineWidth(lineWeight); 

    const obsText = obsTextarea.value;
    pdf.setFontSize(fontSizeBody); 
    pdf.setFont('helvetica', 'normal');
    const textLines = pdf.splitTextToSize(obsText, contentWidth - (2 * padding));
    
    const pointToMm = 0.352778;
    const lineHeightFactorObs = 1.15; 
    const singleLineHeightWithFactorObs = fontSizeBody * pointToMm * lineHeightFactorObs;
    const textPrintStartY = contentBoxTopY + padding + (fontSizeBody * pointToMm * 0.8); 

    let currentDrawingY = textPrintStartY;
    for (const line of textLines) {
        pdf.text(line, margin + padding, currentDrawingY, {
            maxWidth: contentWidth - (2 * padding)
        });
        currentDrawingY += singleLineHeightWithFactorObs;
        if (currentDrawingY > contentBoxTopY + contentBoxHeight - padding) break; 
    }

    return initialCurrentY + actualSectionHeight;
}


async function drawReportPhotosSection(
    pdf,
    pageContainer,
    currentY, 
    dayIndex
) {
    const daySuffix = `_day${dayIndex}`;
    const sectionTitle = "RELATÓRIO FOTOGRÁFICO";
    const titleHeight = 5;
    const photoGridMarginBelowTitle = 2; 

    const titleActualStartY = currentY; 

    pdf.setLineWidth(lineWeightStrong);
    pdf.setFillColor('#E0E0E0');
    pdf.setFontSize(9);
    pdf.setFont('helvetica', 'bold');
    pdf.rect(margin, titleActualStartY, contentWidth, titleHeight, 'FD');
    pdf.text(sectionTitle, margin + contentWidth / 2, titleActualStartY + 3.5, { align: 'center' });
    
    const photoGridContentStartY = titleActualStartY + titleHeight + photoGridMarginBelowTitle;

    const photoGridPadding = 2; 
    const photoFrameImgHeight = 70; 
    const captionAreaHeight = 12; 
    const interPhotoColumnGap = 10; 
    const interPhotoRowGap = 0; 
    const imageDisplayPadding = 0.5; 

    const photoGridEffectiveWidth = contentWidth - (2 * photoGridPadding);
    const photoCellWidth = (photoGridEffectiveWidth - interPhotoColumnGap) / 2; 
    const photoCellHeight = photoFrameImgHeight + captionAreaHeight; 
    const photoGridEffectiveHeight = (photoCellHeight * 2) + interPhotoRowGap;

    const gridStartX = margin + photoGridPadding; 
    const panelLineStartY = titleActualStartY + titleHeight; 
    
    const signatureSectionHeightConstant = 25; 
    const panelLineEndY = pageHeight - margin - signatureSectionHeightConstant;

    pdf.setLineWidth(lineWeightStrong);
    pdf.line(margin, panelLineStartY, margin, panelLineEndY);
    pdf.line(margin + contentWidth, panelLineStartY, margin + contentWidth, panelLineEndY);


    for (let rowIndex = 0; rowIndex < 2; rowIndex++) {
        for (let colIndex = 0; colIndex < 2; colIndex++) {
            const photoNum = rowIndex * 2 + colIndex + 1;

            const cellStartX = gridStartX + (colIndex * (photoCellWidth + interPhotoColumnGap));
            const cellStartY = photoGridContentStartY + (rowIndex * (photoCellHeight + interPhotoRowGap));

            const imageX = cellStartX + imageDisplayPadding;
            const imageY = cellStartY + imageDisplayPadding;
            const imageW = photoCellWidth - (2 * imageDisplayPadding);
            const imageH = photoFrameImgHeight - (2 * imageDisplayPadding);

            const imgPreviewId = `#report_photo_${photoNum}_pt2_preview${daySuffix}`;
            await addImageToPdfAsync(pdf, imgPreviewId, imageX, imageY, imageW, imageH, pageContainer);

            pdf.setLineWidth(photoFrameLineWeight);
            pdf.rect(imageX, imageY, imageW, imageH, 'S'); 

            const captionContentRenderAreaY = cellStartY + photoFrameImgHeight + 1.5;
            const captionFontSize = 6.5;
            const pointToMm = 0.352778;
            const captionFontSizeInMm = captionFontSize * pointToMm;
            const captionLineHeightFactor = 1.1;
            const captionBaselineToBaselineSeparation = captionFontSizeInMm * captionLineHeightFactor;

            const captionLabel = `Foto ${String(photoNum).padStart(2, '0')}:`;
            pdf.setFontSize(captionFontSize);
            pdf.setFont('helvetica', 'bold'); 

            const labelBaselineY = captionContentRenderAreaY + captionFontSizeInMm * 0.8;
            pdf.text(captionLabel, cellStartX + 1.5, labelBaselineY);

            pdf.setFont('helvetica', 'normal'); 
            const captionLabelWidth = pdf.getTextWidth(captionLabel); 
            const captionTextValue = getInputValue(`report_photo_${photoNum}_pt2_caption${daySuffix}`, pageContainer);
            const captionTextStartX = cellStartX + 1.5 + captionLabelWidth + 1; 
            const captionTextMaxWidth = photoCellWidth - (1.5 + captionLabelWidth + 1 + 1.5); 

            if (captionTextValue.trim() !== "") {
                const textLines = pdf.splitTextToSize(captionTextValue, captionTextMaxWidth);
                let currentDrawingTextY = labelBaselineY; 
                for (const line of textLines) {
                    if (currentDrawingTextY < (cellStartY + photoFrameImgHeight + captionAreaHeight) - (captionFontSizeInMm * 0.1)) {
                        pdf.text(line, captionTextStartX, currentDrawingTextY, { maxWidth: captionTextMaxWidth });
                    }
                    currentDrawingTextY += captionBaselineToBaselineSeparation;
                }
            }
        }
    }
    const finalYAfterSection = titleActualStartY + titleHeight + photoGridMarginBelowTitle + photoGridEffectiveHeight;
    return finalYAfterSection;
}


async function drawSignatureSection(pdf, pageContainer, currentY, dayIndex, pageInDayIdx) {
    const daySuffix = `_day${dayIndex}`;
    const sectionTopY = currentY;
    const sectionHeight = 25; 
    const titleHeight = 5;    
    const contentAreaHeight = sectionHeight - titleHeight;
    const sectionBottomY = sectionTopY + sectionHeight;
    const sectionRightX = margin + contentWidth;

    const singleSignatureAreaWidth = contentWidth / 2;
    const actualSignatureLineWidth = singleSignatureAreaWidth * 0.7;

    pdf.setLineWidth(lineWeight);
    pdf.line(margin, sectionTopY, sectionRightX, sectionTopY);

    pdf.setLineWidth(lineWeightStrong);
    pdf.line(margin, sectionTopY, margin, sectionBottomY); 
    pdf.line(sectionRightX, sectionTopY, sectionRightX, sectionBottomY); 
    pdf.line(margin, sectionBottomY, sectionRightX, sectionBottomY); 

    pdf.setLineWidth(lineWeight); 
    pdf.setFillColor('#E0E0E0'); 
    pdf.rect(margin, sectionTopY, contentWidth, titleHeight, 'FD'); 

    pdf.setFontSize(8);
    pdf.setFont('helvetica', 'bold');
    pdf.setTextColor(0, 0, 0); 
    pdf.text("ASSINATURAS", margin + contentWidth / 2, sectionTopY + titleHeight / 2 + 1.5, { align: 'center' });

    pdf.setLineWidth(lineWeight); 

    const contentTopY = sectionTopY + titleHeight;

    const leftSigAreaX = margin;
    const imgMaxH = contentAreaHeight * 0.55; 
    const imgMaxW = actualSignatureLineWidth; 
    const imgX = leftSigAreaX + (singleSignatureAreaWidth - imgMaxW) / 2;
    const imgY = contentTopY + 1.5; 

    let consbemPreviewSelector = `#consbemPhotoPreview${daySuffix}`;
    if (pageInDayIdx === 1) {
        consbemPreviewSelector = `#consbemPhotoPreview_pt${daySuffix}`;
    } else if (pageInDayIdx === 2) {
        consbemPreviewSelector = `#consbemPhotoPreview_pt2${daySuffix}`;
    }
    
    await addImageToPdfAsync(pdf, consbemPreviewSelector, imgX, imgY, imgMaxW, imgMaxH, pageContainer);

    const leftLineY = imgY + imgMaxH + 2; 
    const leftLineXStart = leftSigAreaX + (singleSignatureAreaWidth - actualSignatureLineWidth) / 2;
    const leftLineXEnd = leftLineXStart + actualSignatureLineWidth;
    pdf.line(leftLineXStart, leftLineY, leftLineXEnd, leftLineY); 

    const leftLabelY = leftLineY + 3.5; 
    pdf.setFontSize(7);
    pdf.setFont('helvetica', 'normal');
    pdf.text("RESPONSAVEL CONSBEM", leftSigAreaX + singleSignatureAreaWidth / 2, leftLabelY, { align: 'center' });

    const rightSigAreaX = margin + singleSignatureAreaWidth;
    const rightLineY = leftLineY; 
    const rightLineXStart = rightSigAreaX + (singleSignatureAreaWidth - actualSignatureLineWidth) / 2;
    const rightLineXEnd = rightLineXStart + actualSignatureLineWidth;
    pdf.line(rightLineXStart, rightLineY, rightLineXEnd, rightLineY); 

    const rightLabelY = leftLabelY; 
    pdf.text("FISCALIZAÇÃO SABESP", rightSigAreaX + singleSignatureAreaWidth / 2, rightLabelY, { align: 'center' });
    
    pdf.setTextColor(0, 0, 0);

    return sectionTopY + sectionHeight;
}

async function generateSingleDayPdf(pdf, dayIndex, dayWrapperElement, isFirstDay) {
    const pageContainers = Array.from(dayWrapperElement.querySelectorAll('.rdo-container'));
    
    for (let pageIdx = 0; pageIdx < pageContainers.length; pageIdx++) {
        const pageContainer = pageContainers[pageIdx];
        if (!isFirstDay || pageIdx > 0) { 
            pdf.addPage();
        }
        
        let currentY = margin;
        currentY = await drawPdfHeader(pdf, pageContainer, currentY);
        currentY = drawContractInfo(pdf, pageContainer, currentY, dayIndex, pageIdx);
        
        const signatureSectionHeight = 25; 

        if (pageIdx === 0) { 
            currentY = drawWeatherConditions(pdf, pageContainer, currentY, dayIndex);
            currentY = drawLegend(pdf, pageContainer, currentY);
            currentY = drawLaborColumns(pdf, pageContainer, currentY, dayIndex);

            const signatureSectionTargetY = pageHeight - margin - signatureSectionHeight;
            const availableHeightForFullActivitySection = signatureSectionTargetY - currentY;
            
            currentY = drawActivityTable(pdf, pageContainer, currentY, dayIndex, 0, availableHeightForFullActivitySection);
            currentY = signatureSectionTargetY;

        } else if (pageIdx === 1) { 
             currentY = drawActivityTable(pdf, pageContainer, currentY, dayIndex, 1); 
             
             const signatureSectionTargetY = pageHeight - margin - signatureSectionHeight;
             const availableHeightForObservationsSection = signatureSectionTargetY - currentY;
             currentY = drawObservationsSection(pdf, pageContainer, currentY, dayIndex, availableHeightForObservationsSection);
             currentY = signatureSectionTargetY; 

        } else if (pageIdx === 2) { 
            currentY = await drawReportPhotosSection(pdf, pageContainer, currentY, dayIndex);
            currentY = pageHeight - margin - signatureSectionHeight; 
        }
        
        await drawSignatureSection(pdf, pageContainer, currentY, dayIndex, pageIdx);
    }
}


async function generatePdf() {
    const generatePdfButton = getElementByIdSafe('generatePdfButton');
    if (generatePdfButton) generatePdfButton.disabled = true;
    updateSaveProgressMessage("Gerando PDF... Por favor, aguarde.", false);

    const pdf = new jsPDF('p', 'mm', 'a4');
    pdf.setProperties({ title: 'Relatório Diário de Obras - RDO' });

    try {
        for (let d = 0; d <= rdoDayCounter; d++) {
            let dayWrapper;
            const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
            if (dayWithButtonLayoutWrapper) {
                dayWrapper = dayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
            } else {
                dayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
            }

            if (dayWrapper) {
                await generateSingleDayPdf(pdf, d, dayWrapper, d === 0);
            } else {
                console.warn(`generatePdf: Wrapper for day ${d} not found. Skipping this day.`);
            }
             if (d < rdoDayCounter) { 
                await delay(100);
            }
        }

        const day0Page1Container = getElementByIdSafe('rdo_day_0_page_0');
        let filename = "RDO_Completo.pdf";
        if (day0Page1Container) {
            const rdoNum = getInputValue('numero_rdo_day0', day0Page1Container);
            const dataVal = getInputValue('data_day0', day0Page1Container);
            if (rdoNum && dataVal) {
                filename = `RDO_${rdoNum.replace('-A', '')}_${dataVal}.pdf`;
            } else if (rdoNum) {
                 filename = `RDO_${rdoNum.replace('-A', '')}.pdf`;
            } else if (dataVal) {
                 filename = `RDO_Ref_${dataVal}.pdf`;
            }
        }
        
        pdf.save(filename);
        updateSaveProgressMessage("PDF gerado com sucesso!", false);

    } catch (error) {
        console.error("Error generating PDF:", error);
        updateSaveProgressMessage("Erro ao gerar PDF. Verifique o console para detalhes.", true);
    } finally {
        if (generatePdfButton) generatePdfButton.disabled = false;
    }
}

// --- Save/Load Progress Functionality ---
const CURRENT_SAVE_VERSION = 2; // Incremented version for new save structure


function collectDayData(dayIndex, dayWrapperElement) {
    const dayData = {};
    const inputs = dayWrapperElement.querySelectorAll(
        'input[type="text"], input[type="date"], input[type="number"], textarea, select, input[type="hidden"]' // Added input[type="hidden"]
    );
    inputs.forEach(input => {
        if (input.id) {
            dayData[input.id] = (input).value; // Cast to HTMLInputElement for .value
        }
    });
    
    // Collect report photo data (base64 and dimensions)
    const page3Container = dayWrapperElement.querySelectorAll('.rdo-container')[2];
    if (page3Container) {
        for (let i = 1; i <= 4; i++) {
            const photoInputId = `report_photo_${i}_pt2_input_day${dayIndex}`;
            const photoPreview = getElementByIdSafe(`report_photo_${i}_pt2_preview_day${dayIndex}`, page3Container);
            const captionId = `report_photo_${i}_pt2_caption_day${dayIndex}`;
            
            if (photoPreview && photoPreview.dataset.isEffectivelyLoaded === 'true' && photoPreview.src.startsWith('data:image')) {
                dayData[photoInputId] = photoPreview.src; // Save base64
                if (photoPreview.dataset.naturalWidth) dayData[`${photoInputId}_naturalWidth`] = photoPreview.dataset.naturalWidth;
                if (photoPreview.dataset.naturalHeight) dayData[`${photoInputId}_naturalHeight`] = photoPreview.dataset.naturalHeight;
            }
            const captionInput = getElementByIdSafe(captionId, page3Container);
            if (captionInput) {
                dayData[captionId] = captionInput.value;
            }
        }
    }
    return dayData;
}


function saveProgress() {
    updateSaveProgressMessage(null); 

    const formData = {
        version: CURRENT_SAVE_VERSION,
        rdoDayCounter: rdoDayCounter,
        days: [],
        masterSignatureBase64: null
    };

    for (let d = 0; d <= rdoDayCounter; d++) {
        let dayWrapper = null;
        const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
        if (dayWithButtonLayoutWrapper) {
            dayWrapper = dayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
        } else {
            dayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        }

        if (dayWrapper) {
            formData.days.push(collectDayData(d, dayWrapper));
        } else {
            console.warn(`Save Progress: Day wrapper for day index ${d} not found. Data for this day will not be saved.`);
        }
    }

    const masterSignaturePreview = getElementByIdSafe('consbemPhotoPreview_day0');
    if (masterSignaturePreview && masterSignaturePreview.dataset.isEffectivelyLoaded === 'true' && masterSignaturePreview.src) {
        if (masterSignaturePreview.src.startsWith('data:image')) {
            formData.masterSignatureBase64 = masterSignaturePreview.src;
        } else {
            console.warn("Master signature source is not a data URL. It might not be correctly restored from save file.");
        }
    }

    try {
        const jsonString = JSON.stringify(formData, null, 2);
        const blob = new Blob([jsonString], { type: 'application/json' });
        
        const day0Page1Container = getElementByIdSafe('rdo_day_0_page_0');
        let filename = "RDO_Progresso.json";

        if (day0Page1Container) {
            const rdoNum = getInputValue('numero_rdo_day0', day0Page1Container);
            const dataVal = getInputValue('data_day0', day0Page1Container);
            if (rdoNum && dataVal) {
                filename = `RDO_${rdoNum.replace('-A', '')}_${dataVal}_Progresso.json`;
            } else if (rdoNum) {
                 filename = `RDO_Progresso_${rdoNum.replace('-A', '')}.json`;
            } else if (dataVal) {
                 filename = `RDO_Progresso_Ref_${dataVal}.json`;
            }
        }

        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(link.href);

        isFormDirty = false;
        updateSaveButtonState();
        updateSaveProgressMessage("Progresso salvo com sucesso!", false);

    } catch (error) {
        console.error("Erro ao salvar progresso:", error);
        updateSaveProgressMessage(`Erro ao salvar progresso: ${error.message}`, true);
    }
}

async function loadProgress(event) {
    updateSaveProgressMessage(null);
    const input = event.target;
    if (!input.files || input.files.length === 0) {
        updateSaveProgressMessage("Nenhum arquivo de progresso selecionado.", true);
        return;
    }

    const file = input.files[0];
    if (!file.name.toLowerCase().endsWith('.json')) {
        updateSaveProgressMessage("Formato de arquivo inválido. Por favor, selecione um arquivo .json.", true);
        input.value = '';
        return;
    }

    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const jsonString = e.target?.result;
            if (!jsonString.trim()) {
                updateSaveProgressMessage("Arquivo de progresso está vazio.", true);
                input.value = '';
                return;
            }
            const data = JSON.parse(jsonString);

            if (!data || typeof data.version !== 'number' || data.version > CURRENT_SAVE_VERSION) {
                updateSaveProgressMessage("Arquivo de progresso inválido ou de uma versão incompatível.", true);
                return;
            }
            
            if (data.days && data.days.length > 0) {
                hasEfetivoBeenLoadedAtLeastOnce = true;
            }
            efetivoWasJustClearedByButton = false; 

            while (rdoDayCounter > 0) {
                await removeLastRdoDay();
            }
            const day0Wrapper = getElementByIdSafe('rdo_day_0_wrapper');
            if (day0Wrapper) {
                 resetDayInstanceContent(day0Wrapper, 0); 
            } else {
                updateSaveProgressMessage("Erro crítico: RDO Dia 0 não encontrado. Não é possível carregar.", true);
                return;
            }

            rdoDayCounter = -1; 

            if (data.masterSignatureBase64) {
                updateAllSignaturePreviews(data.masterSignatureBase64);
            } else {
                updateAllSignaturePreviews(null);
            }

            if (data.days && data.days.length > 0) {
                const day0Data = data.days[0];
                const day0Page1Container = getElementByIdSafe('rdo_day_0_page_0');
                if (day0Data && day0Wrapper && day0Page1Container) {
                    Object.keys(day0Data).forEach(key => {
                        const value = day0Data[key];
                         if (typeof value === 'string' && !key.startsWith('photo_') && !key.startsWith('tempo_') && !key.startsWith('trabalho_') && !key.includes('_naturalWidth') && !key.includes('_naturalHeight') ) {
                             setInputValue(key, value, day0Wrapper);
                        }
                    });
                    const dataDay0InputVal = day0Data.data_day0;
                    if (dataDay0InputVal && typeof dataDay0InputVal === 'string') {
                        setInputValue('data_day0', dataDay0InputVal, day0Page1Container); 
                        await adjustRdoDaysForMonth(dataDay0InputVal); 
                    } else {
                        await updateAllDateRelatedFieldsFromDay0();
                    }
                }
            } else {
                updateSaveProgressMessage("Arquivo de progresso não contém dados de dias.", true);
                return;
            }
            
            const targetRdoDayCountFromFile = data.rdoDayCounter; 
            
            while(rdoDayCounter < targetRdoDayCountFromFile) {
                await addNewRdoDay();
            }
            while(rdoDayCounter > targetRdoDayCountFromFile && rdoDayCounter > 0) {
                await removeLastRdoDay();
            }

            for (let d = 0; d <= targetRdoDayCountFromFile; d++) {
                const dayData = data.days[d];
                if (!dayData) continue;

                let currentDayWrapper = null;
                const dayWithButtonWrapper = document.getElementById(`day_with_button_container_day${d}`);
                if (dayWithButtonWrapper) {
                    currentDayWrapper = dayWithButtonWrapper.querySelector('.rdo-day-wrapper');
                } else {
                    currentDayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
                }
                
                if (!currentDayWrapper) continue;
                
                Object.keys(dayData).forEach(key => {
                    const value = dayData[key];
                     if (typeof value === 'string' && !key.startsWith('photo_') && !key.startsWith('tempo_') && !key.startsWith('trabalho_')  && !key.includes('_naturalWidth') && !key.includes('_naturalHeight')) {
                        setInputValue(key, value, currentDayWrapper);
                    }
                });

                ['t1', 't2', 't3'].forEach(turnoIdBase => {
                    const tempoValue = dayData[`tempo_${turnoIdBase}_day${d}_hidden`];
                    const trabalhoValue = dayData[`trabalho_${turnoIdBase}_day${d}_hidden`];
                    const cellId = `${turnoIdBase === 't1' ? 'turno1' : turnoIdBase === 't2' ? 'turno2' : 'turno3'}_control_cell_day${d}`;
                    const updater = shiftControlUpdaters.get(cellId);
                    if (updater && typeof tempoValue === 'string' && typeof trabalhoValue === 'string') {
                        updater.update(tempoValue, trabalhoValue);
                    }
                });
                
                // Update status, localização, tipo_servico, and servico_desc displays based on loaded hidden input values
                const dropdownHiddenInputs = currentDayWrapper.querySelectorAll(
                    'input[type="hidden"][id^="status_day'+ d +'_"], input[type="hidden"][id^="localizacao_day'+ d +'_"], input[type="hidden"][id^="tipo_servico_day'+ d +'_"], input[type="hidden"][id^="servico_desc_day'+ d +'_"]'
                );
                dropdownHiddenInputs.forEach(hiddenInput => {
                    if (dayData[hiddenInput.id] !== undefined) { 
                        hiddenInput.value = dayData[hiddenInput.id]; 
                    }
                    const displayId = hiddenInput.id.replace('_hidden', '_display');
                    const displayElement = getElementByIdSafe(displayId, currentDayWrapper);
                    if (displayElement) {
                        const currentValue = hiddenInput.value;
                        displayElement.textContent = currentValue === "" ? "" : currentValue;

                        const dropdownId = hiddenInput.id.replace('_hidden', '_dropdown');
                        const dropdownElement = getElementByIdSafe(dropdownId, currentDayWrapper);
                        if (dropdownElement) {
                            dropdownElement.querySelectorAll('.status-option').forEach(opt => {
                                opt.setAttribute('aria-selected', (opt.dataset.value === currentValue) ? 'true' : 'false');
                            });
                        }
                    }
                });


                const page3Container = currentDayWrapper.querySelectorAll('.rdo-container')[2];
                if (page3Container) {
                    for (let i = 1; i <= 4; i++) {
                        const photoKey = `report_photo_${i}_pt2_input_day${d}`;
                        const captionKey = `report_photo_${i}_pt2_caption_day${d}`;
                        const photoDataUrl = dayData[photoKey];
                        const captionValue = dayData[captionKey];

                        const photoInput = getElementByIdSafe(`report_photo_${i}_pt2_input_day${d}`, page3Container);
                        const photoPreview = getElementByIdSafe(`report_photo_${i}_pt2_preview_day${d}`, page3Container);
                        const photoPlaceholder = getElementByIdSafe(`report_photo_${i}_pt2_placeholder_day${d}`, page3Container);
                        const clearButton = getElementByIdSafe(`report_photo_${i}_pt2_clear_button_day${d}`, page3Container);
                        const labelElement = getElementByIdSafe(`report_photo_${i}_pt2_label_day${d}`, page3Container);

                        if (photoInput && photoPreview && photoPlaceholder && clearButton && labelElement) {
                            if (typeof photoDataUrl === 'string' && photoDataUrl.startsWith('data:image')) {
                                const nw = dayData[`${photoKey}_naturalWidth`];
                                const nh = dayData[`${photoKey}_naturalHeight`];
                                setReportPhotoSlotImage(photoInput, photoPreview, photoPlaceholder, clearButton, labelElement, photoDataUrl, nw, nh);
                            } else {
                                resetReportPhotoSlotState(photoInput, photoPreview, photoPlaceholder, clearButton, labelElement);
                            }
                        }
                        if (typeof captionValue === 'string') {
                            setInputValue(`report_photo_${i}_pt2_caption_day${d}`, captionValue, page3Container);
                        }
                    }
                }
                const page1Container = currentDayWrapper.querySelectorAll('.rdo-container')[0];
                if (page1Container) {
                     calculateSectionTotalsForDay(page1Container, d);
                }
                updateRainfallStatus(d, 'idle');
            }

            await updateAllDateRelatedFieldsFromDay0(); 
            updateAllSignaturePreviews(data.masterSignatureBase64); 
            updateAllEfetivoAlerts(); 
            updateClearEfetivoButtonVisibility();

            isFormDirty = false;
            updateSaveButtonState();
            updateSaveProgressMessage("Progresso carregado com sucesso!", false);
            setActiveTab(0);
             if(dayTabs.length > 0) {
                const firstTabTargetEl = document.getElementById(`day_with_button_container_day0`) || getElementByIdSafe(`rdo_day_0_wrapper`);
                if(firstTabTargetEl) firstTabTargetEl.scrollIntoView({behavior: "auto", block: "start"});
            }

        } catch (error) {
            console.error("Erro ao carregar progresso:", error);
            updateSaveProgressMessage(`Erro ao carregar progresso: ${error.message}`, true);
        } finally {
            input.value = ''; 
        }
    };
    reader.onerror = () => {
        console.error("Erro ao ler o arquivo de progresso.");
        updateSaveProgressMessage("Erro ao ler o arquivo de progresso.", true);
        input.value = '';
    };
    reader.readAsText(file);
}

function setupSaveLoadListeners() {
    const saveBtn = getElementByIdSafe('saveProgressButton');
    const loadBtn = getElementByIdSafe('loadProgressButton');
    const loadInpt = getElementByIdSafe('loadProgressInput');

    if (saveBtn) {
        saveBtn.addEventListener('click', saveProgress);
    }
    if (loadBtn && loadInpt) {
        loadBtn.addEventListener('click', () => loadInpt.click());
        loadInpt.addEventListener('change', loadProgress);
    } else {
        console.warn("Save/Load UI elements not fully found. Progress saving/loading might not work.");
    }
}


// --- Main Initialization ---
document.addEventListener('DOMContentLoaded', async () => {
    // Cache progress bar elements
    rainfallProgressBarWrapper = getElementByIdSafe('rainfallProgressBarWrapper');
    rainfallProgressBar = getElementByIdSafe('rainfallProgressBar');
    rainfallProgressText = getElementByIdSafe('rainfallProgressText');
    rainfallProgressDetailsText = getElementByIdSafe('rainfallProgressDetailsText');
    globalFetchMessageContainer = getElementByIdSafe('fetchRainfallMessageContainer');
    globalSaveMessageContainer = getElementByIdSafe('saveProgressMessageContainer');


    const initialDay0Wrapper = getElementByIdSafe('rdo_day_0_wrapper');
    if (initialDay0Wrapper) {
        await initializeDayInstance(0, initialDay0Wrapper);
    } else {
        console.error("Critical: Initial RDO Day 0 wrapper ('rdo_day_0_wrapper') not found in HTML. App cannot initialize.");
        updateFetchMessage("Erro crítico na inicialização. Verifique o console.", true);
        return;
    }
    
    const day0Page1Container = getElementByIdSafe('rdo_day_0_page_0');
    if(day0Page1Container) {
        const dataDay0Input = getElementByIdSafe('data_day0', day0Page1Container);
        if(dataDay0Input && dataDay0Input.value) {
            await adjustRdoDaysForMonth(dataDay0Input.value);
        } else if (dataDay0Input) { // Set to today if empty and then adjust
            const today = new Date();
            dataDay0Input.value = formatDateToYYYYMMDD(today);
            await adjustRdoDaysForMonth(dataDay0Input.value);
        }
    } else {
         console.warn("Day 0 Page 1 container (rdo_day_0_page_0) not found during initial month adjustment.");
    }
    
    await updateAllDateRelatedFieldsFromDay0();
    updateAllSignaturePreviews(null); // Initialize all signatures as cleared
    updateAllEfetivoAlerts(); 
    updateClearEfetivoButtonVisibility(); // Initial check for clear efetivo button

    setupNavigationObserver();
    setActiveTab(0);

    // Global click listener to close status dropdowns
    document.addEventListener('click', (event) => {
        document.querySelectorAll('.status-dropdown').forEach(dropdown => { // This now correctly targets all
            const container = (dropdown).closest('.status-select-container');
            if (container && !container.contains(event.target)) {
                (dropdown).style.display = 'none';
                const display = container.querySelector('.status-display');
                if (display) display.setAttribute('aria-expanded', 'false');
            }
        });
    });

    const addDayButton = getElementByIdSafe('addRdoDayButton');
    const removeDayButton = getElementByIdSafe('removeRdoDayButton');
    const generatePdfButton = getElementByIdSafe('generatePdfButton');
    const fetchRainfallButton = getElementByIdSafe('fetchRainfallButton');
    const loadEfetivoButton = getElementByIdSafe('loadEfetivoSpreadsheetButton');
    const clearEfetivoButton = getElementByIdSafe('clearEfetivoButton');


    if (addDayButton) addDayButton.addEventListener('click', addNewRdoDay);
    if (removeDayButton) removeDayButton.addEventListener('click', removeLastRdoDay);
    if (generatePdfButton) generatePdfButton.addEventListener('click', generatePdf);
    if (fetchRainfallButton) fetchRainfallButton.addEventListener('click', fetchAllRainfallDataForAllDaysOnClick);
    if (loadEfetivoButton) loadEfetivoButton.addEventListener('click', triggerEfetivoFileUpload);
    if (clearEfetivoButton) clearEfetivoButton.addEventListener('click', clearAllEfetivo);
    
    const efetivoSpreadsheetUpload = getElementByIdSafe('efetivoSpreadsheetUpload');
    if (efetivoSpreadsheetUpload) {
        efetivoSpreadsheetUpload.addEventListener('change', handleEfetivoFileUpload);
    }

    // Initialize Save/Load functionality
    saveProgressButtonElement = getElementByIdSafe('saveProgressButton');
    updateSaveButtonState(); // Initialize save button state
    setupSaveLoadListeners();


    // Prevent unsaved changes loss
    window.addEventListener('beforeunload', (event) => {
        if (isFormDirty) {
            event.preventDefault();
            event.returnValue = ''; // Standard for most browsers
        }
    });

    isFormDirty = false; 
    updateSaveButtonState();
});

// Ensure global type for jsPDF AutoTable extension if not implicitly available
// TypeScript 'declare module' is not valid in JS.
// If jspdf-autotable extends jsPDF correctly, this might not be needed,
// or a different approach for JS type hinting/checking would be used if desired (e.g., JSDoc).
// For direct execution, we assume the import handles the extension.
// declare module 'jspdf' {
//   interface jsPDF {
//     autoTable: (options: any) => jsPDF;
//   }
// }