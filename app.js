let periodCount = 0;
let periods = [];
let detentionPeriodCount = 0;
let detentionPeriods = [];
let detentionInitialized = false;
let lastResult = null;
let lastBatchResult = null;
let batchRows = [];
let batchRowCount = 0;
let printCleanupTimeout = null;

const NAVEX_LOGO_SRC = 'assets/Navex_VH (3).png';
const PAGE_MODE = document.body?.dataset.page || 'single';
const BATCH_FIELDS = ['ref', 'releaseDate', 'pickupDate', 'freeDays', 'returnDate', 'detentionFreeDays'];
const BATCH_DATE_FIELDS = new Set(['releaseDate', 'pickupDate', 'returnDate']);
const BATCH_NUMERIC_FIELDS = new Set(['freeDays', 'detentionFreeDays']);
const COLUMN_MAP = {
    'referencia': 'ref',
    'referencia contentor': 'ref',
    'ref': 'ref',
    'contentor': 'ref',
    'container': 'ref',
    'descarga': 'releaseDate',
    'descarregamento': 'releaseDate',
    'release': 'releaseDate',
    'release_date': 'releaseDate',
    'levantamento': 'pickupDate',
    'pickup': 'pickupDate',
    'pickup_date': 'pickupDate',
    'dias_livres_storage': 'freeDays',
    'dias livres storage': 'freeDays',
    'free_days_storage': 'freeDays',
    'free_days': 'freeDays',
    'devolucao': 'returnDate',
    'devolucao contentor': 'returnDate',
    'return': 'returnDate',
    'return_date': 'returnDate',
    'dias_livres_detention': 'detentionFreeDays',
    'dias livres detention': 'detentionFreeDays',
    'free_days_detention': 'detentionFreeDays',
};

function isSinglePage() {
    return PAGE_MODE === 'single';
}

function isBatchPage() {
    return PAGE_MODE === 'batch';
}

function getConfigStorageKey() {
    return isBatchPage() ? 'calcConfigBatch' : 'calcConfigSingle';
}

function getPricingEngine() {
    if (!window.StorageDetentionPricing) {
        throw new Error('Pricing engine is not loaded.');
    }
    return window.StorageDetentionPricing;
}

function ensureBatchPrintArea() {
    let area = document.getElementById('batchPrintArea');
    if (!area) {
        area = document.createElement('div');
        area.id = 'batchPrintArea';
        area.className = 'batch-print-area';
        document.body.appendChild(area);
    }
    return area;
}

function clearPrintModes() {
    document.body.classList.remove('print-mode-single', 'print-mode-batch');
    const area = document.getElementById('batchPrintArea');
    if (area) area.innerHTML = '';
    if (printCleanupTimeout) {
        clearTimeout(printCleanupTimeout);
        printCleanupTimeout = null;
    }
}

function triggerPrint(mode, filename) {
    document.body.classList.remove('print-mode-single', 'print-mode-batch');
    document.body.classList.add(mode === 'batch' ? 'print-mode-batch' : 'print-mode-single');
    if (printCleanupTimeout) clearTimeout(printCleanupTimeout);
    printCleanupTimeout = setTimeout(clearPrintModes, 30000);
    const originalTitle = document.title;
    if (filename) document.title = filename;
    window.print();
    document.title = originalTitle;
}

window.addEventListener('afterprint', clearPrintModes);

function parseDate(value) {
    if (!value) return null;
    const trimmed = String(value).trim();
    if (!trimmed) return null;

    let match = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (match) {
        const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
        return Number.isNaN(date.getTime()) ? null : date;
    }

    match = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (match) {
        const date = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
        return Number.isNaN(date.getTime()) ? null : date;
    }

    return null;
}

function addCalendarDays(date, days) {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
}

function formatDate(date) {
    return date.toLocaleDateString('pt-PT', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

function formatInputDate(value) {
    const date = parseDate(value);
    return date ? formatDate(date) : '-';
}

function escapeHTML(value) {
    return String(value ?? '').replace(/[&<>"']/g, (char) => (
        { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]
    ));
}

function ensureXLSXLoaded() {
    if (typeof XLSX !== 'undefined') return true;
    alert('As funcionalidades Excel precisam da biblioteca XLSX.');
    return false;
}

function createNode(tag, options = {}) {
    const el = document.createElement(tag);
    if (options.className) el.className = options.className;
    if (options.text != null) el.textContent = options.text;
    if (options.html != null) el.innerHTML = options.html;
    if (options.attrs) {
        Object.entries(options.attrs).forEach(([key, value]) => {
            if (value != null) el.setAttribute(key, String(value));
        });
    }
    if (options.style) Object.assign(el.style, options.style);
    return el;
}

function getDetentionStartMode(context = 'single') {
    const selector = context === 'batch'
        ? 'input[name="batchDetentionStart"]:checked'
        : 'input[name="detentionStart"]:checked';
    const input = document.querySelector(selector);
    return input ? input.value : 'release';
}

function initializePeriods() {
    periods = [
        { id: 0, days: 10, rate: 50 },
        { id: 1, days: 0, rate: 100 }
    ];
    periodCount = 2;
    renderPeriods();
}

function initializeDetentionPeriods() {
    detentionPeriods = [
        { id: 0, days: 10, rate: 50 },
        { id: 1, days: 0, rate: 100 }
    ];
    detentionPeriodCount = 2;
    renderDetentionPeriods();
}

function createPeriodMarkup(period, index, totalCount, kind) {
    const isLast = index === totalCount - 1;
    const idAttr = kind === 'detention' ? 'data-det-period-id' : 'data-period-id';
    const inputClass = kind === 'detention' ? 'det-period-input' : 'period-input';
    const removeCall = kind === 'detention' ? `removeDetentionPeriod(${period.id})` : `removePeriod(${period.id})`;
    const removeButton = totalCount > 1
        ? `<button type="button" class="remove-btn" onclick="${removeCall}" title="Remover periodo">x</button>`
        : '';

    return `
        ${removeButton}
        <div class="period-header">
            <span class="period-title">Periodo ${index + 1}${isLast ? ' (em diante)' : ''}</span>
        </div>
        <div class="period-inputs" style="${isLast ? 'grid-template-columns: 1fr;' : ''}">
            ${!isLast ? `
                <div class="form-group">
                    <label>Duracao (dias)</label>
                    <input type="number" ${idAttr}="${period.id}" data-field="days" value="${period.days}" min="0" class="${inputClass}">
                </div>
            ` : ''}
            <div class="form-group">
                <label>Tarifa (EUR/dia)</label>
                <input type="number" ${idAttr}="${period.id}" data-field="rate" value="${period.rate}" step="0.01" min="0" class="${inputClass}">
            </div>
        </div>
    `;
}

function renderPeriods() {
    const container = document.getElementById('periodsContainer');
    if (!container) return;

    container.innerHTML = '';
    periods.forEach((period, index) => {
        const periodDiv = document.createElement('div');
        periodDiv.className = 'period-item';
        periodDiv.id = `period-${period.id}`;
        periodDiv.innerHTML = createPeriodMarkup(period, index, periods.length, 'storage');
        container.appendChild(periodDiv);
    });

    container.querySelectorAll('.period-input').forEach((input) => {
        input.addEventListener('change', updatePeriodValue);
    });
}

function renderDetentionPeriods() {
    const container = document.getElementById('detentionPeriodsContainer');
    if (!container) return;

    container.innerHTML = '';
    detentionPeriods.forEach((period, index) => {
        const periodDiv = document.createElement('div');
        periodDiv.className = 'period-item';
        periodDiv.id = `det-period-${period.id}`;
        periodDiv.innerHTML = createPeriodMarkup(period, index, detentionPeriods.length, 'detention');
        container.appendChild(periodDiv);
    });

    container.querySelectorAll('.det-period-input').forEach((input) => {
        input.addEventListener('change', updateDetentionPeriodValue);
    });
}

function updatePeriodValue(event) {
    const periodId = Number(event.target.dataset.periodId);
    const field = event.target.dataset.field;
    const period = periods.find((item) => item.id === periodId);
    if (!period) return;
    period[field] = field === 'days'
        ? (event.target.value === '' ? 0 : Number(event.target.value) || 0)
        : Number(event.target.value) || 0;
}

function updateDetentionPeriodValue(event) {
    const periodId = Number(event.target.dataset.detPeriodId);
    const field = event.target.dataset.field;
    const period = detentionPeriods.find((item) => item.id === periodId);
    if (!period) return;
    period[field] = field === 'days'
        ? (event.target.value === '' ? 0 : Number(event.target.value) || 0)
        : Number(event.target.value) || 0;
}

function addPeriod() {
    periods.push({ id: periodCount++, days: 0, rate: 50 });
    renderPeriods();
    saveConfig();
    recalculateSingleIfNeeded();
}

function removePeriod(id) {
    periods = periods.filter((period) => period.id !== id);
    if (periods.length === 0) initializePeriods();
    else renderPeriods();
    saveConfig();
    recalculateSingleIfNeeded();
}

function addDetentionPeriod() {
    detentionPeriods.push({ id: detentionPeriodCount++, days: 0, rate: 50 });
    renderDetentionPeriods();
    saveConfig();
    recalculateSingleIfNeeded();
}

function removeDetentionPeriod(id) {
    detentionPeriods = detentionPeriods.filter((period) => period.id !== id);
    if (detentionPeriods.length === 0) initializeDetentionPeriods();
    else renderDetentionPeriods();
    saveConfig();
    recalculateSingleIfNeeded();
}

function saveConfig() {
    const freeDaysInput = document.getElementById('freeDays');
    const detentionFreeDaysInput = document.getElementById('detentionFreeDays');
    if (!freeDaysInput || !detentionFreeDaysInput) return;

    const config = {
        freeDays: freeDaysInput.value,
        detentionFreeDays: detentionFreeDaysInput.value,
        periods,
        detentionPeriods,
        detentionStartModeSingle: getDetentionStartMode('single'),
        detentionStartModeBatch: getDetentionStartMode('batch'),
    };

    try {
        localStorage.setItem(getConfigStorageKey(), JSON.stringify(config));
    } catch (error) {
        console.error(error);
    }
}

function getSavedConfig() {
    const candidates = [getConfigStorageKey()];
    if (isSinglePage()) candidates.push('calcConfig');

    for (const key of candidates) {
        try {
            const config = JSON.parse(localStorage.getItem(key));
            if (config) return config;
        } catch (error) {
            console.error(error);
        }
    }
    return null;
}

function loadConfig() {
    const freeDaysInput = document.getElementById('freeDays');
    const detentionFreeDaysInput = document.getElementById('detentionFreeDays');
    const detentionSection = document.getElementById('detentionSection');
    if (!freeDaysInput || !detentionFreeDaysInput || !detentionSection) return false;

    const config = getSavedConfig();
    if (!config) return false;

    if (config.freeDays != null) freeDaysInput.value = config.freeDays;
    if (config.detentionFreeDays != null) detentionFreeDaysInput.value = config.detentionFreeDays;

    if (config.detentionStartModeSingle) {
        const singleStart = document.querySelector(`input[name="detentionStart"][value="${config.detentionStartModeSingle}"]`);
        if (singleStart) singleStart.checked = true;
    }

    if (config.detentionStartModeBatch) {
        const batchStart = document.querySelector(`input[name="batchDetentionStart"][value="${config.detentionStartModeBatch}"]`);
        if (batchStart) batchStart.checked = true;
    }

    if (Array.isArray(config.periods) && config.periods.length > 0) {
        periods = config.periods;
        periodCount = Math.max(...periods.map((period) => period.id)) + 1;
        renderPeriods();
    } else {
        initializePeriods();
    }

    if (Array.isArray(config.detentionPeriods) && config.detentionPeriods.length > 0) {
        detentionPeriods = config.detentionPeriods;
        detentionPeriodCount = Math.max(...detentionPeriods.map((period) => period.id)) + 1;
        renderDetentionPeriods();
        detentionInitialized = true;
    } else {
        initializeDetentionPeriods();
        detentionInitialized = true;
    }

    toggleDetention();

    return true;
}

function toggleDetention() {
    const detentionSection = document.getElementById('detentionSection');
    if (!detentionSection) return;

    detentionSection.classList.add('show');

    if (!detentionInitialized) {
        initializeDetentionPeriods();
        detentionInitialized = true;
    }
}

function calculateDays() {
    const releaseDate = parseDate(document.getElementById('releaseDate')?.value);
    const pickupDate = parseDate(document.getElementById('pickupDate')?.value);
    const returnDate = parseDate(document.getElementById('returnDate')?.value);
    const totalDaysInput = document.getElementById('totalDays');
    const containerDaysInput = document.getElementById('containerDays');
    if (!totalDaysInput && !containerDaysInput) return;

    if (totalDaysInput) {
        if (!releaseDate || !pickupDate) {
            totalDaysInput.value = '';
        } else {
            const storageDays = getPricingEngine().calculateStorage({
                releaseDate,
                pickupDate,
                freeDays: 0,
                periods: []
            });

            totalDaysInput.value = storageDays.error ? '' : String(storageDays.totalDays);
        }
    }

    if (containerDaysInput) {
        const startMode = getDetentionStartMode('single');
        const startDate = startMode === 'release' ? releaseDate : pickupDate;
        if (!startDate || !returnDate) {
            containerDaysInput.value = '';
        } else {
            const containerDays = getPricingEngine().calculateDetention({
                releaseDate,
                pickupDate,
                returnDate,
                detentionFreeDays: 0,
                startMode,
                periods: []
            });
            containerDaysInput.value = containerDays ? String(containerDays.totalDays) : '';
        }
    }
}

function clampDateYear(id) {
    const input = document.getElementById(id);
    if (!input?.value) return;
    const parts = input.value.split('-');
    if (parts[0] && parts[0].length > 4) {
        parts[0] = parts[0].slice(0, 4);
        input.value = parts.join('-');
    }
}

function calculateSingleDetention(releaseDate, pickupDate) {
    const returnDate = parseDate(document.getElementById('returnDate')?.value);
    const detentionFreeDays = Number(document.getElementById('detentionFreeDays')?.value) || 0;
    if (!returnDate) return null;

    const startMode = getDetentionStartMode('single');
    const detention = getPricingEngine().calculateDetention({
        releaseDate,
        pickupDate,
        returnDate,
        detentionFreeDays,
        startMode,
        periods: detentionPeriods,
    });

    if (!detention) return null;

    return {
        ...detention,
        freeDatesLabel: detention.freeStartDate && detention.freeEndDate
            ? `${formatDate(detention.freeStartDate)} - ${formatDate(detention.freeEndDate)}`
            : '',
        startLabel: startMode === 'release' ? 'descarga' : 'levantamento',
    };
}

function buildResultPeriodElement(periodName, datesLabel, calcLabel, subtotal) {
    const period = createNode('div', { className: 'result-period' });
    const top = createNode('div', { className: 'result-period-top' });
    const bottom = createNode('div', { className: 'result-period-bottom' });

    top.appendChild(createNode('span', { className: 'result-period-name', text: periodName }));
    top.appendChild(createNode('span', { className: 'result-period-subtotal', text: `EUR ${subtotal.toFixed(2)}` }));
    bottom.appendChild(createNode('span', { className: 'result-period-dates', text: datesLabel }));
    bottom.appendChild(createNode('span', { className: 'result-period-calc', text: calcLabel }));

    period.appendChild(top);
    period.appendChild(bottom);
    return period;
}

function buildFreeDaysElement(labelText, datesText) {
    const row = createNode('div', { className: 'result-free-days' });
    row.appendChild(createNode('span', { className: 'free-label', text: labelText }));
    row.appendChild(createNode('span', { className: 'free-dates', text: datesText || '' }));
    return row;
}

function renderSingleResult(resultDiv, resultData) {
    const {
        today,
        cost,
        freeDays,
        freeDatesLabel,
        breakdown,
        detention,
        grandTotal
    } = resultData;

    resultDiv.replaceChildren();

    const printHeader = createNode('div', { className: 'print-header' });
    const printBrand = createNode('div', { className: 'print-header-brand' });
    const printBrandText = createNode('div');
    printBrand.appendChild(createNode('img', {
        className: 'print-header-logo',
        attrs: { src: NAVEX_LOGO_SRC, alt: 'Simbolo' }
    }));
    printBrandText.appendChild(createNode('div', { className: 'print-header-title', text: 'Storage e Detention' }));
    printBrand.appendChild(printBrandText);
    printHeader.appendChild(printBrand);
    printHeader.appendChild(createNode('div', { className: 'print-header-meta', text: `Emitido em ${today}` }));

    const header = createNode('div', { className: 'result-header' });
    header.appendChild(createNode('div', { className: 'result-label', text: 'Total a Pagar' }));
    header.appendChild(createNode('div', { className: 'result-total', text: `EUR ${grandTotal.toFixed(2)}` }));

    const body = createNode('div', { className: 'result-body' });
    body.appendChild(createNode('div', { className: 'result-section-divider', text: 'Storage' }));
    body.appendChild(buildFreeDaysElement(`Dias Livres - ${freeDays} dias`, freeDatesLabel));

    breakdown.forEach((item, index) => {
        const periodName = `Periodo ${index + 1}${item.isLast ? ' - em diante' : ''}`;
        const datesLabel = item.calStart && item.calEnd
            ? `${formatDate(item.calStart)} - ${formatDate(item.calEnd)}`
            : `${item.days} dias`;
        const calcLabel = `${item.days} dias x EUR ${item.rate.toFixed(2)}/dia`;
        body.appendChild(buildResultPeriodElement(periodName, datesLabel, calcLabel, item.subtotal));
    });

    if (detention) {
        body.appendChild(createNode('div', {
            className: 'result-section-divider',
            text: `Detention desde ${detention.startLabel}`
        }));
        body.appendChild(buildFreeDaysElement(
            `Dias Livres - ${detention.detentionFreeDays} dias`,
            detention.freeDatesLabel
        ));

        detention.breakdown.forEach((item, index) => {
            const periodName = `Periodo ${index + 1}${item.isLast ? ' - em diante' : ''}`;
            const datesLabel = item.calStart && item.calEnd
                ? `${formatDate(item.calStart)} - ${formatDate(item.calEnd)}`
                : `${item.days} dias`;
            const calcLabel = `${item.days} dias x EUR ${item.rate.toFixed(2)}/dia`;
            body.appendChild(buildResultPeriodElement(periodName, datesLabel, calcLabel, item.subtotal));
        });
    }

    const actions = createNode('div', { className: 'result-btn-group' });
    const copyBtn = createNode('button', { className: 'result-copy-btn', text: 'Copiar resumo' });
    const pdfBtn = createNode('button', { className: 'result-pdf-btn', text: 'Exportar PDF' });
    copyBtn.addEventListener('click', copyResult);
    pdfBtn.addEventListener('click', exportPDF);
    actions.appendChild(copyBtn);
    actions.appendChild(pdfBtn);

    resultDiv.appendChild(printHeader);
    resultDiv.appendChild(header);
    resultDiv.appendChild(body);
    resultDiv.appendChild(actions);
    resultDiv.classList.add('show');
}

function calculate() {
    const freeDays = Number(document.getElementById('freeDays')?.value) || 0;
    const releaseDate = parseDate(document.getElementById('releaseDate')?.value);
    const pickupDate = parseDate(document.getElementById('pickupDate')?.value);

    const storage = getPricingEngine().calculateStorage({
        releaseDate,
        pickupDate,
        freeDays,
        periods,
    });

    const totalDaysInput = document.getElementById('totalDays');
    if (totalDaysInput) totalDaysInput.value = storage.error ? '' : String(storage.totalDays);

    const cost = storage.error ? 0 : storage.cost;
    const breakdown = storage.error ? [] : storage.breakdown;
    const freeDatesLabel = storage.freeStartDate && storage.freeEndDate
        ? `${formatDate(storage.freeStartDate)} - ${formatDate(storage.freeEndDate)}`
        : '';
    const detention = calculateSingleDetention(releaseDate, pickupDate);
    const grandTotal = cost + (detention ? detention.cost : 0);

    const resultDiv = document.getElementById('result');
    if (!resultDiv) return;

    renderSingleResult(resultDiv, {
        today: new Date().toLocaleDateString('pt-PT', { day: '2-digit', month: '2-digit', year: 'numeric' }),
        cost,
        freeDays,
        freeDatesLabel,
        breakdown,
        detention,
        grandTotal
    });

    lastResult = { cost, freeDays, freeDatesLabel, breakdown, detention, grandTotal };
}

function copyResult() {
    if (!lastResult) return;

    const { cost, freeDays, freeDatesLabel, breakdown, detention, grandTotal } = lastResult;
    const lines = [];

    lines.push('Storage');
    lines.push('-'.repeat(42));
    lines.push(`Dias Livres: ${freeDays} dias${freeDatesLabel ? ` (${freeDatesLabel})` : ''}`);
    lines.push('');

    breakdown.forEach((item, index) => {
        lines.push(`Periodo ${index + 1}${item.isLast ? ' (em diante)' : ''}`);
        if (item.calStart && item.calEnd) {
            lines.push(`  ${formatDate(item.calStart)} - ${formatDate(item.calEnd)}`);
        }
        lines.push(`  ${item.days} dias x EUR ${item.rate.toFixed(2)}/dia = EUR ${item.subtotal.toFixed(2)}`);
        lines.push('');
    });

    if (detention) {
        lines.push('Detention');
        lines.push(`Dias Livres: ${detention.detentionFreeDays} dias${detention.freeDatesLabel ? ` (${detention.freeDatesLabel})` : ''}`);
        lines.push('');
        detention.breakdown.forEach((item, index) => {
            lines.push(`Periodo ${index + 1}${item.isLast ? ' (em diante)' : ''}`);
            if (item.calStart && item.calEnd) {
                lines.push(`  ${formatDate(item.calStart)} - ${formatDate(item.calEnd)}`);
            }
            lines.push(`  ${item.days} dias x EUR ${item.rate.toFixed(2)}/dia = EUR ${item.subtotal.toFixed(2)}`);
            lines.push('');
        });
        lines.push(`Storage: EUR ${cost.toFixed(2)}`);
        lines.push(`Detention: EUR ${detention.cost.toFixed(2)}`);
    }

    lines.push(`Total: EUR ${grandTotal.toFixed(2)}`);
    navigator.clipboard.writeText(lines.join('\n')).catch(() => {});
}

function exportPDF() {
    if (!lastResult) return;
    const now = new Date();
    const fileDate = now.toISOString().slice(0, 10);
    const fileTime = `${String(now.getHours()).padStart(2,'0')}h${String(now.getMinutes()).padStart(2,'0')}`;
    triggerPrint('single', `Storage_${fileDate}_${fileTime}`);
}

function recalculateSingleIfNeeded() {
    const resultDiv = document.getElementById('result');
    if (isSinglePage() && resultDiv?.classList.contains('show')) {
        calculate();
    }
}

function resetForm() {
    const form = document.getElementById('calculatorForm');
    if (!form) return;

    form.reset();
    const totalDaysInput = document.getElementById('totalDays');
    if (totalDaysInput) totalDaysInput.value = '';
    const containerDaysInput = document.getElementById('containerDays');
    if (containerDaysInput) containerDaysInput.value = '';

    const freeDaysInput = document.getElementById('freeDays');
    const detentionFreeDaysInput = document.getElementById('detentionFreeDays');
    if (freeDaysInput) freeDaysInput.value = '7';
    if (detentionFreeDaysInput) detentionFreeDaysInput.value = '7';

    initializePeriods();
    initializeDetentionPeriods();
    detentionInitialized = true;

    const resultDiv = document.getElementById('result');
    if (resultDiv) {
        resultDiv.classList.remove('show');
        resultDiv.innerHTML = '';
    }

    saveConfig();
}

function normalizeBatchDateInput(value) {
    const trimmed = String(value ?? '').trim();
    if (!trimmed) return '';

    if (/^\d{4}[/-]\d{1,2}[/-]\d{1,2}$/.test(trimmed)) {
        const parts = trimmed.replace(/\//g, '-').split('-');
        return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
    }

    if (/^\d{1,2}[-.]\d{1,2}[-.]\d{4}$/.test(trimmed)) {
        return trimmed.replace(/[-.]/g, '/');
    }

    return trimmed;
}

function sanitizeBatchCellValue(field, value) {
    const trimmed = String(value ?? '').trim();
    if (!trimmed) return '';
    if (BATCH_DATE_FIELDS.has(field)) return normalizeBatchDateInput(trimmed);
    if (BATCH_NUMERIC_FIELDS.has(field)) {
        const numeric = Number(trimmed.replace(',', '.'));
        return Number.isFinite(numeric) ? String(numeric) : trimmed;
    }
    return trimmed;
}

function createBatchIndexCell(index) {
    const td = document.createElement('td');
    td.className = 'batch-row-index';
    td.textContent = String(index);
    return td;
}

function createBatchRemoveCell(id) {
    const td = document.createElement('td');
    td.className = 'batch-remove';
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'batch-remove-btn';
    button.textContent = 'x';
    button.addEventListener('click', () => removeBatchRow(id));
    td.appendChild(button);
    return td;
}

function createBatchInputCell(id, field, type, value, extra = {}) {
    const td = document.createElement('td');
    td.className = `batch-cell-${field}`;

    const input = document.createElement('input');
    input.type = type;
    input.value = value ?? '';
    input.dataset.batchId = String(id);
    input.dataset.batchField = field;

    if (extra.placeholder) input.placeholder = extra.placeholder;
    if (extra.inputMode) input.inputMode = extra.inputMode;
    if (extra.min != null) input.min = String(extra.min);
    if (extra.className) input.className = extra.className;
    if (extra.style) input.style.cssText = extra.style;

    td.appendChild(input);
    return td;
}

function addBatchRow(data) {
    const id = batchRowCount++;
    batchRows.push({ id, ...(data || {}) });
    renderBatchTable();
}

function removeBatchRow(id) {
    batchRows = batchRows.filter((row) => row.id !== id);
    renderBatchTable();
}

function updateBatchRowValue(event) {
    const id = Number(event.target.dataset.batchId);
    const field = event.target.dataset.batchField;
    const row = batchRows.find((item) => item.id === id);
    if (!row) return;

    const sanitized = sanitizeBatchCellValue(field, event.target.value);
    row[field] = sanitized;
    if (event.type === 'change') event.target.value = sanitized;
}

function syncBatchRowsFromDOM() {
    document.querySelectorAll('[data-batch-id]').forEach((input) => {
        const id = Number(input.dataset.batchId);
        const field = input.dataset.batchField;
        const row = batchRows.find((item) => item.id === id);
        if (!row) return;
        row[field] = sanitizeBatchCellValue(field, input.value);
    });
}

function setBatchCellValue(startRowIndex, startFieldIndex, valueMatrix) {
    for (let rowOffset = 0; rowOffset < valueMatrix.length; rowOffset += 1) {
        const targetRowIndex = startRowIndex + rowOffset;
        while (targetRowIndex >= batchRows.length) addBatchRow();

        const row = batchRows[targetRowIndex];
        for (let colOffset = 0; colOffset < valueMatrix[rowOffset].length; colOffset += 1) {
            const fieldIndex = startFieldIndex + colOffset;
            if (fieldIndex >= BATCH_FIELDS.length) continue;
            const field = BATCH_FIELDS[fieldIndex];
            row[field] = sanitizeBatchCellValue(field, valueMatrix[rowOffset][colOffset]);
        }
    }
}

function handleBatchCellPaste(event) {
    const text = event.clipboardData?.getData('text');
    if (!text) return;

    const startId = Number(event.target.dataset.batchId);
    const startField = event.target.dataset.batchField;
    const startRowIndex = batchRows.findIndex((row) => row.id === startId);
    const startFieldIndex = BATCH_FIELDS.indexOf(startField);
    if (startRowIndex < 0 || startFieldIndex < 0) return;

    const rows = text
        .replace(/\r\n/g, '\n')
        .split('\n')
        .filter((line) => line.length > 0)
        .map((line) => line.split('\t'));

    if (rows.length === 0) return;

    event.preventDefault();
    setBatchCellValue(startRowIndex, startFieldIndex, rows);
    renderBatchTable();
}

function renderBatchTable() {
    const tbody = document.getElementById('batchTableBody');
    const meta = document.getElementById('batchTableMeta');
    if (!tbody) return;

    tbody.innerHTML = '';
    const defaultFreeDays = document.getElementById('freeDays')?.value || '7';
    const defaultDetentionFreeDays = document.getElementById('detentionFreeDays')?.value || '7';

    batchRows.forEach((row, index) => {
        const tr = document.createElement('tr');
        tr.appendChild(createBatchIndexCell(index + 1));
        tr.appendChild(createBatchInputCell(row.id, 'ref', 'text', row.ref || '', { placeholder: 'CONT123' }));
        tr.appendChild(createBatchInputCell(row.id, 'releaseDate', 'text', row.releaseDate || '', {
            placeholder: 'dd/mm/yyyy',
            inputMode: 'numeric',
            className: 'batch-date-input'
        }));
        tr.appendChild(createBatchInputCell(row.id, 'pickupDate', 'text', row.pickupDate || '', {
            placeholder: 'dd/mm/yyyy',
            inputMode: 'numeric',
            className: 'batch-date-input'
        }));
        tr.appendChild(createBatchInputCell(row.id, 'freeDays', 'number', row.freeDays != null ? String(row.freeDays) : defaultFreeDays, {
            min: 0,
            className: 'batch-number-input'
        }));
        tr.appendChild(createBatchInputCell(row.id, 'returnDate', 'text', row.returnDate || '', {
            placeholder: 'dd/mm/yyyy',
            inputMode: 'numeric',
            className: 'batch-date-input'
        }));
        tr.appendChild(createBatchInputCell(
            row.id,
            'detentionFreeDays',
            'number',
            row.detentionFreeDays != null ? String(row.detentionFreeDays) : defaultDetentionFreeDays,
            { min: 0, className: 'batch-number-input' }
        ));
        tr.appendChild(createBatchRemoveCell(row.id));
        tbody.appendChild(tr);
    });

    tbody.querySelectorAll('[data-batch-id]').forEach((input) => {
        input.addEventListener('change', updateBatchRowValue);
        input.addEventListener('input', updateBatchRowValue);
        input.addEventListener('paste', handleBatchCellPaste);
        input.addEventListener('focus', function onFocus() {
            if (typeof this.select === 'function') this.select();
        });
    });

    if (meta) {
        const count = batchRows.length;
        meta.textContent = count === 0
            ? 'Sem contentores carregados'
            : `${count} contentor${count === 1 ? '' : 'es'} carregado${count === 1 ? '' : 's'}`;
    }
}

function batchRowHasReturnDate(row) {
    return !!parseDate(row.returnDate);
}

function batchHasAnyDetentionRows() {
    return batchRows.some(batchRowHasReturnDate);
}

function calcStorageForRow(row) {
    return getPricingEngine().calculateStorage({
        releaseDate: parseDate(row.releaseDate),
        pickupDate: parseDate(row.pickupDate),
        freeDays: Number(row.freeDays) || 0,
        periods,
    });
}

function calcDetentionForRow(row) {
    return getPricingEngine().calculateDetention({
        releaseDate: parseDate(row.releaseDate),
        pickupDate: parseDate(row.pickupDate),
        returnDate: parseDate(row.returnDate),
        detentionFreeDays: Number(row.detentionFreeDays) || 0,
        startMode: getDetentionStartMode('batch'),
        periods: detentionPeriods,
    });
}

function renderBatchResults(container, batchData, batchCount) {
    const { resultRows, grandStorage, grandDetention, grandTotal, detentionActive } = batchData;
    container.replaceChildren();

    const panel = createNode('div', { className: 'batch-results-panel' });
    const summary = createNode('div', { className: 'batch-results-summary' });
    summary.appendChild(createNode('span', {
        text: `Total Geral - ${batchCount} contentor${batchCount !== 1 ? 'es' : ''}`
    }));
    summary.appendChild(createNode('strong', { text: `EUR ${grandTotal.toFixed(2)}` }));

    const wrap = createNode('div', { className: 'batch-results-table-wrap' });
    const table = createNode('table', { className: 'batch-results-table' });
    const thead = createNode('thead');
    const headRow = createNode('tr');

    ['Referencia', 'Descarga', 'Levantamento', 'Devolucao', 'Storage'].forEach((label, index) => {
        const th = createNode('th', { text: label });
        if (index === 4) th.style.textAlign = 'right';
        headRow.appendChild(th);
    });

    if (detentionActive) {
        const detTh = createNode('th', { text: 'Detention' });
        detTh.style.textAlign = 'right';
        headRow.appendChild(detTh);
    }

    const totalTh = createNode('th', { text: 'Total' });
    totalTh.style.textAlign = 'right';
    headRow.appendChild(totalTh);
    thead.appendChild(headRow);

    const tbody = createNode('tbody');
    resultRows.forEach((row) => {
        const tr = createNode('tr', { className: row.error ? 'batch-results-row-error' : '' });
        tr.appendChild(createNode('td', { className: 'ref', text: row.ref }));
        tr.appendChild(createNode('td', { text: row.releaseDateLabel }));
        tr.appendChild(createNode('td', { text: row.pickupDateLabel }));
        tr.appendChild(createNode('td', { text: row.returnDateLabel }));

        if (row.error) {
            const errorCell = createNode('td', { className: 'num', text: `ERRO: ${row.error}` });
            errorCell.colSpan = detentionActive ? 3 : 2;
            tr.appendChild(errorCell);
            tbody.appendChild(tr);
            return;
        }

        tr.appendChild(createNode('td', {
            className: 'num',
            text: `EUR ${row.storageCost.toFixed(2)}`
        }));

        if (detentionActive) {
            tr.appendChild(createNode('td', {
                className: 'num',
                text: row.detCost != null ? `EUR ${row.detCost.toFixed(2)}` : '-'
            }));
        }

        tr.appendChild(createNode('td', {
            className: 'num total',
            text: `EUR ${row.total.toFixed(2)}`
        }));

        tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);
    wrap.appendChild(table);
    panel.appendChild(summary);
    panel.appendChild(wrap);

    if (detentionActive) {
        const subtotals = createNode('div', { className: 'batch-results-subtotals' });
        subtotals.appendChild(createNode('span', { text: `Storage: EUR ${grandStorage.toFixed(2)}` }));
        subtotals.appendChild(createNode('span', { text: `Detention: EUR ${grandDetention.toFixed(2)}` }));
        panel.appendChild(subtotals);
    }

    container.appendChild(panel);
}

function calculateBatch() {
    syncBatchRowsFromDOM();

    const batchResults = document.getElementById('batchResults');
    if (!batchResults) return;

    if (batchRows.length === 0) {
        batchResults.innerHTML = '<p style="color:#888;font-size:13px;">Adicione pelo menos um contentor.</p>';
        return;
    }

    if (detentionPeriods.length === 0) {
        initializeDetentionPeriods();
        detentionInitialized = true;
    }

    const detentionActive = batchHasAnyDetentionRows();
    let grandStorage = 0;
    let grandDetention = 0;
    const resultRows = [];

    batchRows.forEach((row) => {
        const ref = row.ref || '(sem referencia)';
        const releaseDateLabel = formatInputDate(row.releaseDate);
        const pickupDateLabel = formatInputDate(row.pickupDate);
        const returnDateLabel = formatInputDate(row.returnDate);
        const storage = calcStorageForRow(row);

        if (storage.error) {
            resultRows.push({ ref, releaseDateLabel, pickupDateLabel, returnDateLabel, error: storage.error });
            return;
        }

        grandStorage += storage.cost;
        let detCost = null;

        if (batchRowHasReturnDate(row)) {
            const detention = calcDetentionForRow(row);
            if (detention) {
                detCost = detention.cost;
                grandDetention += detention.cost;
            }
        }

        resultRows.push({
            ref,
            releaseDateLabel,
            pickupDateLabel,
            returnDateLabel,
            storageCost: storage.cost,
            detCost,
            total: storage.cost + (detCost ?? 0)
        });
    });

    lastBatchResult = {
        resultRows,
        grandStorage,
        grandDetention,
        grandTotal: grandStorage + grandDetention,
        detentionActive
    };

    renderBatchResults(batchResults, lastBatchResult, batchRows.length);
    const exportButtons = document.getElementById('batchExportBtns');
    if (exportButtons) exportButtons.style.display = '';
}

function clearBatch() {
    batchRows = [];
    batchRowCount = 0;
    lastBatchResult = null;
    renderBatchTable();
    const batchResults = document.getElementById('batchResults');
    const exportButtons = document.getElementById('batchExportBtns');
    if (batchResults) batchResults.innerHTML = '';
    if (exportButtons) exportButtons.style.display = 'none';
}

function copyBatchResult() {
    if (!lastBatchResult) return;

    const { resultRows, grandStorage, grandDetention, grandTotal, detentionActive } = lastBatchResult;
    const lines = [];

    lines.push('RESUMO - MULTIPLOS CONTENTORES');
    lines.push('-'.repeat(52));

    resultRows.forEach((row) => {
        if (row.error) {
            lines.push(`${row.ref} -> ERRO: ${row.error}`);
            return;
        }
        const detStr = detentionActive
            ? ` | Detention: ${row.detCost != null ? `EUR ${row.detCost.toFixed(2)}` : 'sem data'}`
            : '';
        lines.push(`${row.ref} -> Storage: EUR ${row.storageCost.toFixed(2)}${detStr} | TOTAL: EUR ${row.total.toFixed(2)}`);
    });

    lines.push('-'.repeat(52));
    if (detentionActive) {
        lines.push(`Storage total: EUR ${grandStorage.toFixed(2)}`);
        lines.push(`Detention total: EUR ${grandDetention.toFixed(2)}`);
        lines.push('-'.repeat(52));
    }
    lines.push(`TOTAL GERAL: EUR ${grandTotal.toFixed(2)}`);

    navigator.clipboard.writeText(lines.join('\n')).catch(() => {});
}

function exportBatchPDF() {
    if (!lastBatchResult) return;

    const { resultRows, grandStorage, grandDetention, grandTotal, detentionActive } = lastBatchResult;
    const today = new Date().toLocaleDateString('pt-PT', { day: '2-digit', month: '2-digit', year: 'numeric' });
    const totalColumns = detentionActive ? 7 : 6;
    const detentionHeader = detentionActive ? '<th style="text-align:right;">Detention (EUR)</th>' : '';

    let rowsHTML = '';
    resultRows.forEach((row) => {
        if (row.error) {
            rowsHTML += `
                <tr class="batch-report-error">
                    <td>${escapeHTML(row.ref)}</td>
                    <td>${escapeHTML(row.releaseDateLabel)}</td>
                    <td>${escapeHTML(row.pickupDateLabel)}</td>
                    <td>${escapeHTML(row.returnDateLabel)}</td>
                    <td colspan="${totalColumns - 4}">ERRO: ${escapeHTML(row.error)}</td>
                </tr>`;
            return;
        }

        const detentionCell = detentionActive
            ? `<td style="text-align:right;">${row.detCost != null ? `EUR ${row.detCost.toFixed(2)}` : '-'}</td>`
            : '';

        rowsHTML += `
            <tr>
                <td>${escapeHTML(row.ref)}</td>
                <td>${escapeHTML(row.releaseDateLabel)}</td>
                <td>${escapeHTML(row.pickupDateLabel)}</td>
                <td>${escapeHTML(row.returnDateLabel)}</td>
                <td style="text-align:right;">EUR ${row.storageCost.toFixed(2)}</td>
                ${detentionCell}
                <td style="text-align:right;font-weight:700;color:#0f2744;">EUR ${row.total.toFixed(2)}</td>
            </tr>`;
    });

    const splitHTML = detentionActive
        ? `
            <div class="batch-report-split">
                <div class="batch-report-split-item">
                    <div class="batch-report-split-label">Storage</div>
                    <div class="batch-report-split-value">EUR ${grandStorage.toFixed(2)}</div>
                </div>
                <div class="batch-report-split-item">
                    <div class="batch-report-split-label">Detention</div>
                    <div class="batch-report-split-value">EUR ${grandDetention.toFixed(2)}</div>
                </div>
            </div>`
        : '';

    const htmlContent = `
        <style>
            @page { size: A4 portrait; margin: 7mm; }
            html, body { margin: 0; padding: 0; width: 100%; }
            .batch-report-wrap { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; width: 100%; max-width: none; margin: 0; color: #1a1a1a; box-sizing: border-box; }
            .batch-report-head { display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; border-bottom: 2px solid #1e3a5f; padding-bottom: 8px; margin-bottom: 10px; gap: 8px; }
            .batch-report-brand { display: flex; align-items: center; gap: 10px; }
            .batch-report-logo { width: 42px; height: 42px; object-fit: contain; border: 1px solid #d7dee8; border-radius: 8px; padding: 4px; background: white; }
            .batch-report-title { margin: 0; font-size: 16px; color: #1e3a5f; }
            .batch-report-sub { margin: 2px 0 0 0; font-size: 10px; color: #667; }
            .batch-report-meta { text-align: right; font-size: 10px; color: #667; line-height: 1.35; }
            .batch-report-grand { border: 1px solid #d7dee8; border-radius: 8px; padding: 10px 12px; margin-bottom: 8px; text-align: center; background: #f6f9fc; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
            .batch-report-grand-label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.8px; color: #4f5b66; margin-bottom: 4px; font-weight: 700; }
            .batch-report-grand-value { font-size: 26px; font-weight: 800; color: #3f4a54; }
            .batch-report-split { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; margin-bottom: 10px; }
            .batch-report-split-item { border: 1px solid #dbe3ec; border-radius: 8px; padding: 7px 8px; background: #f7fafc; }
            .batch-report-split-label { font-size: 10px; text-transform: uppercase; color: #68788b; letter-spacing: 0.4px; margin-bottom: 3px; }
            .batch-report-split-value { font-size: 16px; font-weight: 700; color: #1e3a5f; }
            .batch-report-table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 9px; }
            .batch-report-table th { background: #f0f4f8; color: #1e3a5f; padding: 5px 4px; border: 1px solid #dbe3ec; font-size: 8px; text-transform: uppercase; letter-spacing: 0.2px; }
            .batch-report-table td { border: 1px solid #e5eaf0; padding: 5px 4px; word-break: break-word; line-height: 1.25; }
            .batch-report-table tr:nth-child(even):not(.batch-report-error) { background: #fbfdff; }
            .batch-report-error td { background: #fff5f5; color: #a12323; font-weight: 600; }
            .batch-report-table th:nth-child(1), .batch-report-table td:nth-child(1) { width: 15%; }
            .batch-report-table th:nth-child(2), .batch-report-table td:nth-child(2),
            .batch-report-table th:nth-child(3), .batch-report-table td:nth-child(3),
            .batch-report-table th:nth-child(4), .batch-report-table td:nth-child(4) { width: 14%; }
            .batch-report-table th:nth-child(5), .batch-report-table td:nth-child(5),
            .batch-report-table th:nth-child(6), .batch-report-table td:nth-child(6),
            .batch-report-table th:nth-child(7), .batch-report-table td:nth-child(7) { width: 15%; }
            .batch-report-foot { margin-top: 8px; color: #667; font-size: 10px; text-align: right; }
        </style>
        <div class="batch-report-wrap">
            <div class="batch-report-head">
                <div class="batch-report-brand">
                    <img src="${NAVEX_LOGO_SRC}" alt="Simbolo" class="batch-report-logo">
                    <div>
                        <h2 class="batch-report-title">Multiplos Contentores</h2>
                    </div>
                </div>
                <div class="batch-report-meta">
                    <div>Emitido em ${today}</div>
                    <div>Contentores: ${resultRows.length}</div>
                </div>
            </div>
            <div class="batch-report-grand">
                <div class="batch-report-grand-label">Total a Pagar</div>
                <div class="batch-report-grand-value">EUR ${grandTotal.toFixed(2)}</div>
            </div>
            ${splitHTML}
            <table class="batch-report-table">
                <thead>
                    <tr>
                        <th>Referencia</th>
                        <th>Descarga</th>
                        <th>Levantamento</th>
                        <th>Devolucao</th>
                        <th style="text-align:right;">Storage (EUR)</th>
                        ${detentionHeader}
                        <th style="text-align:right;">Total (EUR)</th>
                    </tr>
                </thead>
                <tbody>${rowsHTML}</tbody>
            </table>
            <div class="batch-report-foot">Documento gerado pela calculadora</div>
        </div>`;

    const area = ensureBatchPrintArea();
    area.innerHTML = htmlContent;
    const n = resultRows.length;
    const now = new Date();
    const fileDate = now.toISOString().slice(0, 10);
    const fileTime = `${String(now.getHours()).padStart(2,'0')}h${String(now.getMinutes()).padStart(2,'0')}`;
    triggerPrint('batch', `Storage_Batch_${n}_contentores_${fileDate}_${fileTime}`);
}

function excelDateToInputValue(value) {
    if (value == null || value === '') return '';
    if (value instanceof Date) return formatDate(value);

    const str = String(value).trim();
    const ddmmyyyy = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (ddmmyyyy) return `${ddmmyyyy[1].padStart(2, '0')}/${ddmmyyyy[2].padStart(2, '0')}/${ddmmyyyy[3]}`;
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
    return str;
}

function normalizeHeader(value) {
    return String(value)
        .toLowerCase()
        .trim()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\s+/g, ' ');
}

function processExcelFile(file) {
    if (!ensureXLSXLoaded()) return;

    const reader = new FileReader();
    reader.onload = function onLoad(event) {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

            if (rows.length < 2) {
                alert('O ficheiro nao tem dados suficientes.');
                return;
            }

            const headerRow = rows[0].map(normalizeHeader);
            const fieldMap = {};
            headerRow.forEach((header, index) => {
                if (COLUMN_MAP[header]) fieldMap[index] = COLUMN_MAP[header];
            });

            if (Object.keys(fieldMap).length === 0) {
                alert('Nenhuma coluna reconhecida no ficheiro.');
                return;
            }

            const importedRows = [];
            for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
                const row = rows[rowIndex];
                if (!row || row.every((cell) => cell == null || cell === '')) continue;

                const mapped = {};
                Object.entries(fieldMap).forEach(([columnIndex, field]) => {
                    const value = row[columnIndex];
                    mapped[field] = BATCH_DATE_FIELDS.has(field)
                        ? excelDateToInputValue(value)
                        : (value != null ? String(value).trim() : '');
                });
                importedRows.push(mapped);
            }

            if (importedRows.length === 0) {
                alert('Nao foram encontradas linhas com dados.');
                return;
            }

            batchRows = [];
            batchRowCount = 0;
            importedRows.forEach((row) => addBatchRow(row));
            renderBatchTable();

            const fileInput = document.getElementById('excelFileInput');
            if (fileInput) fileInput.value = '';
        } catch (error) {
            console.error(error);
            alert(`Erro ao ler o ficheiro: ${error.message}`);
        }
    };

    reader.readAsArrayBuffer(file);
}

function loadExcelFile(event) {
    const file = event.target.files?.[0];
    if (file) processExcelFile(file);
}

function downloadTemplate() {
    if (!ensureXLSXLoaded()) return;

    const headers = ['Referencia', 'Descarga', 'Levantamento', 'Dias_Livres_Storage', 'Devolucao', 'Dias_Livres_Detention'];
    const example = ['CONT001', '01/03/2025', '15/03/2025', '7', '30/03/2025', '7'];
    const worksheet = XLSX.utils.aoa_to_sheet([headers, example]);
    worksheet['!cols'] = [{ wch: 16 }, { wch: 18 }, { wch: 18 }, { wch: 20 }, { wch: 18 }, { wch: 22 }];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Contentores');
    XLSX.writeFile(workbook, 'template_contentores.xlsx');
}

document.addEventListener('change', function onChange(event) {
    if (event.target.matches('.period-input, .det-period-input, #freeDays, #detentionFreeDays, input[name="detentionStart"], input[name="batchDetentionStart"]')) {
        saveConfig();
    }

    if (isSinglePage() && event.target.matches('input[name="detentionStart"]')) {
        calculateDays();
        recalculateSingleIfNeeded();
    } else if (isSinglePage() && event.target.matches('.period-input, .det-period-input, #freeDays, #detentionFreeDays')) {
        recalculateSingleIfNeeded();
    }

    if (isBatchPage() && event.target.matches('#freeDays, #detentionFreeDays')) {
        renderBatchTable();
    }
});

window.addEventListener('load', function onLoad() {
    const restored = loadConfig();
    if (!restored) {
        initializePeriods();
        if (isBatchPage()) {
            initializeDetentionPeriods();
            detentionInitialized = true;
        }
    }

    const releaseDateInput = document.getElementById('releaseDate');
    const pickupDateInput = document.getElementById('pickupDate');
    const returnDateInput = document.getElementById('returnDate');

    if (releaseDateInput && pickupDateInput) {
        releaseDateInput.addEventListener('change', function handleReleaseChange() {
            clampDateYear('releaseDate');
            calculateDays();
            recalculateSingleIfNeeded();
        });
        pickupDateInput.addEventListener('change', function handlePickupChange() {
            clampDateYear('pickupDate');
            calculateDays();
            recalculateSingleIfNeeded();
        });
        if (returnDateInput) {
            returnDateInput.addEventListener('change', function handleReturnChange() {
                clampDateYear('returnDate');
                calculateDays();
                recalculateSingleIfNeeded();
            });
        }
        calculateDays();
    }

    const calculatorForm = document.getElementById('calculatorForm');
    if (calculatorForm) {
        calculatorForm.addEventListener('keypress', function handleKeyPress(event) {
            if (event.key === 'Enter') calculate();
        });
    }

    toggleDetention();
    if (isBatchPage() && batchRows.length === 0) addBatchRow();

    const dropZone = document.getElementById('dropZone');
    if (dropZone) {
        dropZone.addEventListener('dragover', function onDragOver(event) {
            event.preventDefault();
            dropZone.classList.add('drag-over');
        });
        dropZone.addEventListener('dragleave', function onDragLeave(event) {
            if (!dropZone.contains(event.relatedTarget)) {
                dropZone.classList.remove('drag-over');
            }
        });
        dropZone.addEventListener('drop', function onDrop(event) {
            event.preventDefault();
            dropZone.classList.remove('drag-over');
            const file = event.dataTransfer.files?.[0];
            if (file) processExcelFile(file);
        });
    }
});
