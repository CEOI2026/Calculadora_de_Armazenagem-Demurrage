        let periodCount = 0;
        let periods = [];
        let detentionPeriodCount = 0;
        let detentionPeriods = [];
        let lastResult = null;
        let printCleanupTimeout = null;
        const NAVEX_LOGO_SRC = 'assets/Navex_VH (3).png';

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

        function triggerPrint(mode) {
            document.body.classList.remove('print-mode-single', 'print-mode-batch');
            document.body.classList.add(mode === 'batch' ? 'print-mode-batch' : 'print-mode-single');
            if (printCleanupTimeout) clearTimeout(printCleanupTimeout);
            // Fallback cleanup in case the browser doesn't fire afterprint.
            printCleanupTimeout = setTimeout(clearPrintModes, 30000);
            window.print();
        }

        window.addEventListener('afterprint', clearPrintModes);

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

        function renderDetentionPeriods() {
            const container = document.getElementById('detentionPeriodsContainer');
            if (!container) return;
            container.innerHTML = '';

            detentionPeriods.forEach((period, index) => {
                const isLast = index === detentionPeriods.length - 1;
                const periodDiv = document.createElement('div');
                periodDiv.className = 'period-item';
                periodDiv.id = `det-period-${period.id}`;
                periodDiv.innerHTML = `
                    ${detentionPeriods.length > 1 ? `<button type="button" class="remove-btn" onclick="removeDetentionPeriod(${period.id})" title="Remover período">×</button>` : ''}
                    <div class="period-header">
                        <span class="period-title">Período ${index + 1}${isLast ? ' (em diante)' : ''}</span>
                    </div>
                    <div class="period-inputs" style="${isLast ? 'grid-template-columns: 1fr;' : ''}">
                        ${!isLast ? `
                        <div class="form-group">
                            <label>
                                Duração (dias)
                                <span>Quantos dias dura este período</span>
                            </label>
                            <input type="number" data-det-period-id="${period.id}" data-field="days" value="${period.days}" min="0" class="det-period-input">
                        </div>` : ''}
                        <div class="form-group">
                            <label>
                                Tarifa (€/dia)
                                <span>Custo por dia neste período</span>
                            </label>
                            <input type="number" data-det-period-id="${period.id}" data-field="rate" value="${period.rate}" step="0.01" min="0" class="det-period-input">
                        </div>
                    </div>
                `;
                container.appendChild(periodDiv);
            });

            document.querySelectorAll('.det-period-input').forEach(input => {
                input.addEventListener('change', updateDetentionPeriodValue);
            });
        }

        function updateDetentionPeriodValue(e) {
            const periodId = parseInt(e.target.dataset.detPeriodId);
            const field = e.target.dataset.field;
            const value = field === 'days' ? parseInt(e.target.value) || 0 : parseFloat(e.target.value) || 0;
            const period = detentionPeriods.find(p => p.id === periodId);
            if (period) period[field] = value;
        }

        function addDetentionPeriod() {
            detentionPeriods.push({ id: detentionPeriodCount++, days: 0, rate: 50 });
            renderDetentionPeriods();
            saveConfig();
        }

        function removeDetentionPeriod(id) {
            detentionPeriods = detentionPeriods.filter(p => p.id !== id);
            renderDetentionPeriods();
            saveConfig();
        }

        function renderPeriods() {
            const container = document.getElementById('periodsContainer');
            container.innerHTML = '';

            periods.forEach((period, index) => {
                const isLast = index === periods.length - 1;
                const periodDiv = document.createElement('div');
                periodDiv.className = 'period-item';
                periodDiv.id = `period-${period.id}`;
                periodDiv.innerHTML = `
                    ${periods.length > 1 ? `<button type="button" class="remove-btn" onclick="removePeriod(${period.id})" title="Remover período">×</button>` : ''}
                    <div class="period-header">
                        <span class="period-title">Período ${index + 1}${isLast ? ' (em diante)' : ''}</span>
                    </div>
                    <div class="period-inputs" style="${isLast ? 'grid-template-columns: 1fr;' : ''}">
                        ${!isLast ? `
                        <div class="form-group">
                            <label>
                                Duração (dias)
                                <span>Quantos dias dura este período</span>
                            </label>
                            <input type="number" data-period-id="${period.id}" data-field="days" value="${period.days}" min="0" class="period-input">
                        </div>` : ''}
                        <div class="form-group">
                            <label>
                                Tarifa (€/dia)
                                <span>Custo por dia neste período</span>
                            </label>
                            <input type="number" data-period-id="${period.id}" data-field="rate" value="${period.rate}" step="0.01" min="0" class="period-input">
                        </div>
                    </div>
                `;
                container.appendChild(periodDiv);
            });

            document.querySelectorAll('.period-input').forEach(input => {
                input.addEventListener('change', updatePeriodValue);
            });
        }

        function updatePeriodValue(e) {
            const periodId = parseInt(e.target.dataset.periodId);
            const field = e.target.dataset.field;
            const value = field === 'days' ? parseInt(e.target.value) || 0 : parseFloat(e.target.value) || 0;

            const period = periods.find(p => p.id === periodId);
            if (period) {
                period[field] = value;
            }
        }

        function addPeriod() {
            periods.push({ id: periodCount++, days: 0, rate: 50 });
            renderPeriods();
            saveConfig();
        }

        function removePeriod(id) {
            periods = periods.filter(p => p.id !== id);
            renderPeriods();
            saveConfig();
        }

        // Aceita YYYY-MM-DD (type=date) ou DD/MM/AAAA (escrito à mão)
        function parseDate(str) {
            if (!str) return null;
            // formato do input type=date: YYYY-MM-DD
            let m = str.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
            if (m) {
                const d = new Date(+m[1], +m[2] - 1, +m[3]);
                return isNaN(d) ? null : d;
            }
            // formato escrito à mão: DD/MM/AAAA
            m = str.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
            if (m) {
                const d = new Date(+m[3], +m[2] - 1, +m[1]);
                return isNaN(d) ? null : d;
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
            return date ? formatDate(date) : '—';
        }

        function escapeHTML(value) {
            return String(value ?? '').replace(/[&<>"']/g, (char) => (
                { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]
            ));
        }

        function ensureXLSXLoaded() {
            if (typeof XLSX !== 'undefined') return true;
            alert('As funcionalidades Excel precisam da biblioteca XLSX. Verifique a ligacao a internet ou use uma copia local da biblioteca.');
            return false;
        }

        function createNode(tag, options = {}) {
            const el = document.createElement(tag);
            if (options.className) el.className = options.className;
            if (options.text != null) el.textContent = options.text;
            if (options.attrs) {
                Object.entries(options.attrs).forEach(([key, value]) => {
                    if (value != null) el.setAttribute(key, String(value));
                });
            }
            if (options.style) Object.assign(el.style, options.style);
            return el;
        }

        function buildResultPeriodElement(periodName, datesLabel, calcLabel, subtotal) {
            const period = createNode('div', { className: 'result-period' });
            const top = createNode('div', { className: 'result-period-top' });
            const bottom = createNode('div', { className: 'result-period-bottom' });

            top.appendChild(createNode('span', { className: 'result-period-name', text: periodName }));
            top.appendChild(createNode('span', { className: 'result-period-subtotal', text: `€${subtotal.toFixed(2)}` }));
            bottom.appendChild(createNode('span', { className: 'result-period-dates', text: datesLabel }));
            bottom.appendChild(createNode('span', { className: 'result-period-calc', text: calcLabel }));

            period.appendChild(top);
            period.appendChild(bottom);
            return period;
        }

        function buildFreeDaysElement(labelText, datesText) {
            const row = createNode('div', { className: 'result-free-days' });
            row.appendChild(createNode('span', { className: 'free-label', text: labelText }));
            row.appendChild(createNode('span', { className: 'free-dates', text: datesText }));
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
                attrs: { src: NAVEX_LOGO_SRC, alt: 'Símbolo' }
            }));
            printBrandText.appendChild(createNode('div', { className: 'print-header-title', text: 'Storage e Detention' }));
            printBrandText.appendChild(createNode('div', { className: 'print-header-sub', text: 'Resumo de Custos' }));
            printBrand.appendChild(printBrandText);
            printHeader.appendChild(printBrand);
            printHeader.appendChild(createNode('div', { className: 'print-header-meta', text: `Emitido em ${today}` }));

            const header = createNode('div', { className: 'result-header' });
            header.appendChild(createNode('div', { className: 'result-label', text: 'Total a Pagar' }));
            header.appendChild(createNode('div', { className: 'result-total', text: `€${grandTotal.toFixed(2)}` }));

            if (detention) {
                const split = createNode('div', {
                    style: {
                        display: 'flex',
                        justifyContent: 'center',
                        gap: '24px',
                        marginTop: '10px'
                    }
                });

                const storageBlock = createNode('div', { style: { textAlign: 'center' } });
                storageBlock.appendChild(createNode('div', {
                    text: 'Storage',
                    style: {
                        fontSize: '10px',
                        color: 'rgba(255,255,255,0.55)',
                        textTransform: 'uppercase',
                        letterSpacing: '0.8px',
                        fontWeight: '600'
                    }
                }));
                storageBlock.appendChild(createNode('div', {
                    text: `€${cost.toFixed(2)}`,
                    style: {
                        fontSize: '18px',
                        fontWeight: '700',
                        color: 'rgba(255,255,255,0.85)'
                    }
                }));

                const divider = createNode('div', {
                    style: {
                        width: '1px',
                        background: 'rgba(255,255,255,0.2)'
                    }
                });

                const detentionBlock = createNode('div', { style: { textAlign: 'center' } });
                detentionBlock.appendChild(createNode('div', {
                    text: 'Detention',
                    style: {
                        fontSize: '10px',
                        color: 'rgba(255,255,255,0.55)',
                        textTransform: 'uppercase',
                        letterSpacing: '0.8px',
                        fontWeight: '600'
                    }
                }));
                detentionBlock.appendChild(createNode('div', {
                    text: `€${detention.cost.toFixed(2)}`,
                    style: {
                        fontSize: '18px',
                        fontWeight: '700',
                        color: 'rgba(255,255,255,0.85)'
                    }
                }));

                split.appendChild(storageBlock);
                split.appendChild(divider);
                split.appendChild(detentionBlock);
                header.appendChild(split);
            }

            const body = createNode('div', { className: 'result-body' });
            body.appendChild(createNode('div', { className: 'result-section-divider', text: 'Storage' }));
            body.appendChild(buildFreeDaysElement(`Dias Livres — ${freeDays} dias`, freeDatesLabel));

            breakdown.forEach((item, idx) => {
                const isLastItem = idx === breakdown.length - 1 && periods[periods.length - 1].days === 0;
                const periodName = `Período ${idx + 1}${isLastItem ? ' — em diante' : ''}`;
                const datesLabel = item.calStart
                    ? `${formatDate(item.calStart)} – ${formatDate(item.calEnd)}`
                    : isLastItem
                        ? `Dia ${item.startDay} em diante`
                        : `Dias ${item.startDay}–${item.endDay}`;
                const calcLabel = `${item.days} dias × €${item.rate.toFixed(2)}/dia`;
                body.appendChild(buildResultPeriodElement(periodName, datesLabel, calcLabel, item.subtotal));
            });

            if (detention) {
                const section = createNode('div', { className: 'result-section-divider', text: 'Detention ' });
                const hint = createNode('span', {
                    text: `desde ${detention.startLabel}`,
                    style: {
                        fontWeight: '400',
                        fontSize: '10px',
                        opacity: '0.7',
                        marginLeft: '4px'
                    }
                });
                section.appendChild(hint);
                body.appendChild(section);
                body.appendChild(buildFreeDaysElement(
                    `Dias Livres — ${detention.detentionFreeDays} dias`,
                    detention.freeDatesLabel
                ));

                detention.breakdown.forEach((item, idx) => {
                    const periodName = `Período ${idx + 1}${item.isLast ? ' — em diante' : ''}`;
                    const datesLabel = item.calStart
                        ? `${formatDate(item.calStart)} – ${formatDate(item.calEnd)}`
                        : `${item.days} dias`;
                    const calcLabel = `${item.days} dias × €${item.rate.toFixed(2)}/dia`;
                    body.appendChild(buildResultPeriodElement(periodName, datesLabel, calcLabel, item.subtotal));
                });
            }

            const actions = createNode('div', { className: 'result-btn-group' });
            const copyBtn = createNode('button', { className: 'result-copy-btn', text: 'Copiar resumo' });
            copyBtn.addEventListener('click', copyResult);
            const pdfBtn = createNode('button', { className: 'result-pdf-btn', text: 'Exportar PDF' });
            pdfBtn.addEventListener('click', exportPDF);
            actions.appendChild(copyBtn);
            actions.appendChild(pdfBtn);

            resultDiv.appendChild(printHeader);
            resultDiv.appendChild(header);
            resultDiv.appendChild(body);
            resultDiv.appendChild(actions);
            resultDiv.classList.add('show');
        }

        function renderBatchResults(container, batchData, batchCount) {
            const { resultRows, grandStorage, grandDetention, grandTotal, detentionActive } = batchData;
            container.replaceChildren();

            const panel = createNode('div', { className: 'batch-results-panel' });
            const summary = createNode('div', { className: 'batch-results-summary' });
            summary.appendChild(createNode('span', {
                text: `Total Geral — ${batchCount} contentor${batchCount !== 1 ? 'es' : ''}`
            }));
            summary.appendChild(createNode('strong', { text: `€${grandTotal.toFixed(2)}` }));

            const wrap = createNode('div', { className: 'batch-results-table-wrap' });
            const table = createNode('table', { className: 'batch-results-table' });
            const thead = createNode('thead');
            const headRow = createNode('tr');
            ['Referência', 'Descarga', 'Levantamento', 'Devolução', 'Storage'].forEach((label, idx) => {
                const th = createNode('th', { text: label });
                if (idx === 4) th.style.textAlign = 'right';
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
            resultRows.forEach(row => {
                const tr = createNode('tr', { className: row.error ? 'batch-results-row-error' : '' });
                const refCell = createNode('td', { className: 'ref', text: row.ref });
                tr.appendChild(refCell);
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

                const storageCell = createNode('td', { className: 'num', text: `€${row.storageCost.toFixed(2)}` });
                tr.appendChild(storageCell);

                if (detentionActive) {
                    const detentionCell = createNode('td', {
                        className: 'num',
                        text: row.detCost != null ? `€${row.detCost.toFixed(2)}` : '—'
                    });
                    tr.appendChild(detentionCell);
                }

                tr.appendChild(createNode('td', {
                    className: 'num total',
                    text: `€${row.total.toFixed(2)}`
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
                const storageSpan = createNode('span');
                storageSpan.appendChild(document.createTextNode('Storage: '));
                storageSpan.appendChild(createNode('strong', { text: `€${grandStorage.toFixed(2)}` }));
                const detentionSpan = createNode('span');
                detentionSpan.appendChild(document.createTextNode('Detention: '));
                detentionSpan.appendChild(createNode('strong', { text: `€${grandDetention.toFixed(2)}` }));
                subtotals.appendChild(storageSpan);
                subtotals.appendChild(detentionSpan);
                panel.appendChild(subtotals);
            }

            container.appendChild(panel);
        }

        let detentionInitialized = false;

        function toggleDetention() {
            const enabled = document.getElementById('detentionEnabled').checked;
            document.getElementById('detentionSection').classList.toggle('show', enabled);
            if (enabled && !detentionInitialized) {
                initializeDetentionPeriods();
                detentionInitialized = true;
            }
        }

        function getDetentionStartMode(context = 'single') {
            const selector = context === 'batch'
                ? 'input[name="batchDetentionStart"]:checked'
                : 'input[name="detentionStart"]:checked';
            const el = document.querySelector(selector);
            return el ? el.value : 'release';
        }

        function calculateDetention(pickupDate, releaseDate) {
            const enabled = document.getElementById('detentionEnabled').checked;
            if (!enabled) return null;

            const returnDateStr = document.getElementById('returnDate').value;
            const returnDate = parseDate(returnDateStr);
            const detentionFreeDays = parseInt(document.getElementById('detentionFreeDays').value) || 0;

            if (!returnDate) return null;

            // Início da contagem: levantamento ou descarga
            const startMode = getDetentionStartMode('single');
            const startDate = startMode === 'release' ? releaseDate : pickupDate;

            if (!startDate) return null;

            const baseDays = Math.ceil((returnDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
            if (baseDays <= 0) return null;

            let remaining = Math.max(0, baseDays - detentionFreeDays);
            let cost = 0;
            let breakdown = [];
            let daysFromStart = detentionFreeDays;

            for (let period of detentionPeriods) {
                if (remaining <= 0) break;
                const daysInPeriod = Math.min(remaining, period.days || Infinity);
                const periodCost = daysInPeriod * period.rate;
                cost += periodCost;
                breakdown.push({
                    days: daysInPeriod,
                    rate: period.rate,
                    subtotal: periodCost,
                    calStart: addCalendarDays(startDate, daysFromStart),
                    calEnd: addCalendarDays(startDate, daysFromStart + daysInPeriod - 1),
                    isLast: period.days === 0,
                });
                remaining -= daysInPeriod;
                daysFromStart += daysInPeriod;
            }

            const freeDatesLabel = detentionFreeDays > 0
                ? `${formatDate(startDate)} – ${formatDate(addCalendarDays(startDate, detentionFreeDays - 1))}`
                : '';

            const startLabel = startMode === 'release' ? 'descarga' : 'levantamento';
            return { cost, breakdown, detentionFreeDays, freeDatesLabel, returnDate, startLabel };
        }

        function calculate() {
            const freeDays = parseInt(document.getElementById('freeDays').value) || 0;
            const releaseDateStr = document.getElementById('releaseDate').value;
            const releaseDate = parseDate(releaseDateStr);
            const pickupDateStr = document.getElementById('pickupDate').value;
            const pickupDate = parseDate(pickupDateStr);
            const computedTotalDays = (releaseDate && pickupDate)
                ? Math.ceil((pickupDate - releaseDate) / (1000 * 60 * 60 * 24)) + 1
                : 0;
            const totalDays = computedTotalDays > 0 ? computedTotalDays : 0;
            document.getElementById('totalDays').value = totalDays > 0 ? String(totalDays) : '';

            let cost = 0;
            let chargedDays = Math.max(0, totalDays - freeDays);
            let breakdown = [];

            let daysFromStart = freeDays;

            for (let period of periods) {
                if (chargedDays <= 0) break;

                const daysInPeriod = Math.min(chargedDays, period.days || Infinity);
                const periodCost = daysInPeriod * period.rate;
                cost += periodCost;

                breakdown.push({
                    startDay: daysFromStart + 1,
                    endDay: daysFromStart + daysInPeriod,
                    days: daysInPeriod,
                    rate: period.rate,
                    subtotal: periodCost,
                    calStart: releaseDate ? addCalendarDays(releaseDate, daysFromStart) : null,
                    calEnd: releaseDate ? addCalendarDays(releaseDate, daysFromStart + daysInPeriod - 1) : null,
                });

                chargedDays -= daysInPeriod;
                daysFromStart += daysInPeriod;
            }

            const freeDatesLabel = releaseDate && freeDays > 0
                ? `${formatDate(releaseDate)} – ${formatDate(addCalendarDays(releaseDate, freeDays - 1))}`
                : '';

            const detention = calculateDetention(pickupDate, releaseDate);
            const grandTotal = cost + (detention ? detention.cost : 0);

            const resultDiv = document.getElementById('result');
            const today = new Date().toLocaleDateString('pt-PT', { day: '2-digit', month: '2-digit', year: 'numeric' });
            renderSingleResult(resultDiv, {
                today,
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

            lines.push(`Storage — Resumo de Custos`);
            lines.push(`${'─'.repeat(42)}`);
            lines.push(`Dias Livres: ${freeDays} dias${freeDatesLabel ? '  (' + freeDatesLabel + ')' : ''}`);
            lines.push('');

            breakdown.forEach((item, idx) => {
                const isLastItem = idx === breakdown.length - 1 && periods[periods.length - 1].days === 0;
                const periodName = `Período ${idx + 1}${isLastItem ? ' (em diante)' : ''}`;
                const datesLabel = item.calStart
                    ? `${formatDate(item.calStart)} – ${formatDate(item.calEnd)}`
                    : isLastItem
                        ? `Dia ${item.startDay} em diante`
                        : `Dias ${item.startDay}–${item.endDay}`;
                lines.push(`${periodName}`);
                lines.push(`  ${datesLabel}`);
                lines.push(`  ${item.days} dias × €${item.rate.toFixed(2)}/dia = €${item.subtotal.toFixed(2)}`);
                if (idx < breakdown.length - 1) lines.push('');
            });

            if (detention) {
                lines.push('');
                lines.push(`${'─'.repeat(42)}`);
                lines.push(`DETENTION`);
                lines.push(`Dias Livres: ${detention.detentionFreeDays} dias${detention.freeDatesLabel ? '  (' + detention.freeDatesLabel + ')' : ''}`);
                lines.push('');
                detention.breakdown.forEach((item, idx) => {
                    const periodName = `Período ${idx + 1}${item.isLast ? ' (em diante)' : ''}`;
                    const datesLabel = item.calStart
                        ? `${formatDate(item.calStart)} – ${formatDate(item.calEnd)}`
                        : `${item.days} dias`;
                    lines.push(periodName);
                    lines.push(`  ${datesLabel}`);
                    lines.push(`  ${item.days} dias × €${item.rate.toFixed(2)}/dia = €${item.subtotal.toFixed(2)}`);
                    if (idx < detention.breakdown.length - 1) lines.push('');
                });
                lines.push('');
                lines.push(`${'─'.repeat(42)}`);
                lines.push(`Storage: €${cost.toFixed(2)}`);
                lines.push(`Detention: €${detention.cost.toFixed(2)}`);
                lines.push(`${'─'.repeat(42)}`);
                lines.push(`TOTAL GERAL: €${grandTotal.toFixed(2)}`);
            } else {
                lines.push('');
                lines.push(`${'─'.repeat(42)}`);
                lines.push(`TOTAL A PAGAR: €${cost.toFixed(2)}`);
            }

            navigator.clipboard.writeText(lines.join('\n')).then(() => {
                const btn = document.querySelector('.result-copy-btn');
                const original = btn.textContent;
                btn.textContent = 'Copiado!';
                btn.style.background = '#1e3a5f';
                btn.style.color = 'white';
                setTimeout(() => {
                    btn.textContent = original;
                    btn.style.background = '';
                    btn.style.color = '';
                }, 2000);
            });
        }

        function exportPDF() {
            if (!lastResult) return;
            triggerPrint('single');
        }

        function resetForm() {
            document.getElementById('calculatorForm').reset();
            document.getElementById('releaseDate').value = '';
            document.getElementById('pickupDate').value = '';
            document.getElementById('freeDays').value = '7';
            document.getElementById('totalDays').value = '';
            document.getElementById('detentionEnabled').checked = false;
            document.getElementById('detentionSection').classList.remove('show');
            document.getElementById('returnDate').value = '';
            document.getElementById('detentionFreeDays').value = '7';
            detentionInitialized = false;
            detentionPeriods = [];
            const detCont = document.getElementById('detentionPeriodsContainer');
            if (detCont) detCont.innerHTML = '';
            initializePeriods();
            document.getElementById('result').classList.remove('show');
            saveConfig();
        }

        document.getElementById('calculatorForm').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                calculate();
            }
        });

        function clampDateYear(id) {
            const input = document.getElementById(id);
            if (!input.value) return;
            const parts = input.value.split('-');
            if (parts[0] && parts[0].length > 4) {
                parts[0] = parts[0].slice(0, 4);
                input.value = parts.join('-');
            }
        }

        // Inicializar ao carregar a página
        window.addEventListener('load', function() {
            const restored = loadConfig();
            if (!restored) initializePeriods();
            addBatchRow();
            document.getElementById('releaseDate').addEventListener('change', function() {
                clampDateYear('releaseDate');
                calculateDays();
            });
            document.getElementById('pickupDate').addEventListener('change', function() {
                clampDateYear('pickupDate');
                calculateDays();
            });
            document.getElementById('returnDate').addEventListener('change', function() {
                clampDateYear('returnDate');
            });
            calculateDays();

            // Drag and drop on the drop zone
            const zone = document.getElementById('dropZone');
            zone.addEventListener('dragover', function(e) {
                e.preventDefault();
                zone.classList.add('drag-over');
            });
            zone.addEventListener('dragleave', function(e) {
                if (!zone.contains(e.relatedTarget)) zone.classList.remove('drag-over');
            });
            zone.addEventListener('drop', function(e) {
                e.preventDefault();
                zone.classList.remove('drag-over');
                const file = e.dataTransfer.files[0];
                if (file) processExcelFile(file);
            });
        });

        function calculateDays() {
            const start = parseDate(document.getElementById('releaseDate').value);
            const end = parseDate(document.getElementById('pickupDate').value);
            const totalDaysInput = document.getElementById('totalDays');

            if (start && end) {
                const diffDays = Math.ceil((end - start) / (1000 * 60 * 60 * 24)) + 1;
                if (diffDays > 0) {
                    totalDaysInput.value = String(diffDays);
                    return;
                }
            }
            totalDaysInput.value = '';
        }

        // ── localStorage persistence ──────────────────────────────────────────

        function saveConfig() {
            const config = {
                freeDays: document.getElementById('freeDays').value,
                detentionFreeDays: document.getElementById('detentionFreeDays').value,
                detentionEnabled: document.getElementById('detentionEnabled').checked,
                periods: periods,
                detentionPeriods: detentionPeriods,
            };
            try { localStorage.setItem('calcConfig', JSON.stringify(config)); } catch(e) {}
        }

        function loadConfig() {
            let config;
            try { config = JSON.parse(localStorage.getItem('calcConfig')); } catch(e) {}
            if (!config) return false;

            if (config.freeDays != null) document.getElementById('freeDays').value = config.freeDays;
            if (config.periods && config.periods.length > 0) {
                periods = config.periods;
                periodCount = Math.max(...periods.map(p => p.id)) + 1;
                renderPeriods();
            }
            if (config.detentionEnabled) {
                document.getElementById('detentionEnabled').checked = true;
                document.getElementById('detentionSection').classList.add('show');
                if (config.detentionFreeDays != null) document.getElementById('detentionFreeDays').value = config.detentionFreeDays;
                if (config.detentionPeriods && config.detentionPeriods.length > 0) {
                    detentionPeriods = config.detentionPeriods;
                    detentionPeriodCount = Math.max(...detentionPeriods.map(p => p.id)) + 1;
                    renderDetentionPeriods();
                    detentionInitialized = true;
                }
            }
            return true;
        }

        // Auto-save on any tariff/config input change
        document.addEventListener('change', function(e) {
            if (e.target.matches('.period-input, .det-period-input, #freeDays, #detentionFreeDays, #detentionEnabled')) {
                saveConfig();
            }
        });

        // ── Batch Calculator ──────────────────────────────────────────────────

        let batchRows = [];
        let batchRowCount = 0;

        function addBatchRow(data) {
            const id = batchRowCount++;
            batchRows.push({ id, ...(data || {}) });
            renderBatchTable();
        }

        function removeBatchRow(id) {
            batchRows = batchRows.filter(r => r.id !== id);
            renderBatchTable();
        }

        function createBatchInputCell(rowId, field, type, value, options = {}) {
            const td = document.createElement('td');
            const input = document.createElement('input');
            input.type = type;
            input.dataset.batchId = rowId;
            input.dataset.batchField = field;
            input.value = value ?? '';

            if (options.placeholder) input.placeholder = options.placeholder;
            if (options.min != null) input.min = String(options.min);
            if (options.step != null) input.step = String(options.step);
            if (options.style) input.style.cssText = options.style;

            td.appendChild(input);
            return td;
        }

        function createBatchRemoveCell(rowId) {
            const td = document.createElement('td');
            td.className = 'batch-remove';

            const button = document.createElement('button');
            button.type = 'button';
            button.className = 'batch-remove-btn';
            button.title = 'Remover';
            button.textContent = '\u00D7';
            button.addEventListener('click', () => removeBatchRow(rowId));

            td.appendChild(button);
            return td;
        }

        function renderBatchTable() {
            const tbody = document.getElementById('batchTableBody');
            const meta = document.getElementById('batchTableMeta');
            tbody.innerHTML = '';
            batchRows.forEach(row => {
                const tr = document.createElement('tr');
                tr.appendChild(createBatchInputCell(row.id, 'ref', 'text', row.ref || '', { placeholder: 'CONT123' }));
                tr.appendChild(createBatchInputCell(row.id, 'releaseDate', 'date', row.releaseDate || ''));
                tr.appendChild(createBatchInputCell(row.id, 'pickupDate', 'date', row.pickupDate || ''));
                tr.appendChild(createBatchInputCell(
                    row.id,
                    'freeDays',
                    'number',
                    row.freeDays != null ? String(row.freeDays) : document.getElementById('freeDays').value,
                    { min: 0, style: 'min-width:60px;' }
                ));
                tr.appendChild(createBatchInputCell(row.id, 'returnDate', 'date', row.returnDate || ''));
                tr.appendChild(createBatchInputCell(
                    row.id,
                    'detentionFreeDays',
                    'number',
                    row.detentionFreeDays != null ? String(row.detentionFreeDays) : document.getElementById('detentionFreeDays').value,
                    { min: 0, style: 'min-width:60px;' }
                ));
                tr.appendChild(createBatchRemoveCell(row.id));
                tbody.appendChild(tr);
            });
            tbody.querySelectorAll('[data-batch-id]').forEach(input => {
                input.addEventListener('change', updateBatchRowValue);
                input.addEventListener('input', updateBatchRowValue);
            });

            if (meta) {
                const count = batchRows.length;
                const label = `${count} contentor${count === 1 ? '' : 'es'} carregado${count === 1 ? '' : 's'}`;
                meta.textContent = count > 8 ? `${label} · use scroll para navegar na tabela` : label;
            }
        }

        function updateBatchRowValue(e) {
            const id = parseInt(e.target.dataset.batchId);
            const field = e.target.dataset.batchField;
            const row = batchRows.find(r => r.id === id);
            if (row) row[field] = e.target.value;
        }

        function calcStorageForRow(row) {
            const releaseDate = parseDate(row.releaseDate);
            const pickupDate = parseDate(row.pickupDate);
            const freeDays = parseInt(row.freeDays) || 0;

            if (!releaseDate || !pickupDate) return { error: 'Datas de descarregamento/levantamento em falta.' };

            const totalDays = Math.ceil((pickupDate - releaseDate) / (1000 * 60 * 60 * 24)) + 1;
            if (totalDays <= 0) return { error: 'Data de levantamento deve ser após descarregamento.' };

            let chargedDays = Math.max(0, totalDays - freeDays);
            let cost = 0;

            for (let period of periods) {
                if (chargedDays <= 0) break;
                const daysInPeriod = Math.min(chargedDays, period.days || Infinity);
                cost += daysInPeriod * period.rate;
                chargedDays -= daysInPeriod;
            }
            return { cost };
        }

        function calcDetentionForRow(row) {
            const pickupDate = parseDate(row.pickupDate);
            const releaseDate = parseDate(row.releaseDate);
            const returnDate = parseDate(row.returnDate);
            const detentionFreeDays = parseInt(row.detentionFreeDays) || 0;

            if (!returnDate) return null;

            const startMode = getDetentionStartMode('batch');
            const startDate = startMode === 'release' ? releaseDate : pickupDate;
            if (!startDate) return null;

            const baseDays = Math.ceil((returnDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
            if (baseDays <= 0) return null;

            let remaining = Math.max(0, baseDays - detentionFreeDays);
            let cost = 0;

            for (let period of detentionPeriods) {
                if (remaining <= 0) break;
                const daysInPeriod = Math.min(remaining, period.days || Infinity);
                cost += daysInPeriod * period.rate;
                remaining -= daysInPeriod;
            }
            return { cost };
        }

        let lastBatchResult = null;

        function calculateBatch() {
            // Sync any unsaved input changes from DOM
            document.querySelectorAll('[data-batch-id]').forEach(input => {
                const id = parseInt(input.dataset.batchId);
                const field = input.dataset.batchField;
                const row = batchRows.find(r => r.id === id);
                if (row) row[field] = input.value;
            });

            if (batchRows.length === 0) {
                document.getElementById('batchResults').innerHTML = '<p style="color:#888;font-size:13px;">Adicione pelo menos um contentor.</p>';
                return;
            }

            const detentionActive = document.getElementById('detentionEnabled').checked && detentionPeriods.length > 0;
            let grandStorage = 0, grandDetention = 0;
            const resultRows = []; // for copy/export

            batchRows.forEach(row => {
                const ref = row.ref || '(sem referência)';
                const releaseDateLabel = formatInputDate(row.releaseDate);
                const pickupDateLabel = formatInputDate(row.pickupDate);
                const returnDateLabel = formatInputDate(row.returnDate);
                const storage = calcStorageForRow(row);

                if (storage.error) {
                    resultRows.push({ ref, releaseDateLabel, pickupDateLabel, returnDateLabel, error: storage.error });
                    return;
                }

                grandStorage += storage.cost;
                let detCost = null; // null = não calculado (sem data)

                if (detentionActive) {
                    const det = calcDetentionForRow(row);
                    if (det) {
                        detCost = det.cost;
                        grandDetention += detCost;
                    }
                }

                const total = storage.cost + (detCost ?? 0);
                resultRows.push({ ref, releaseDateLabel, pickupDateLabel, returnDateLabel, storageCost: storage.cost, detCost, total });
            });

            const grandTotal = grandStorage + grandDetention;
            lastBatchResult = { resultRows, grandStorage, grandDetention, grandTotal, detentionActive };
            const batchResults = document.getElementById('batchResults');
            renderBatchResults(batchResults, lastBatchResult, batchRows.length);
            document.getElementById('batchExportBtns').style.display = '';
            batchResults.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }

        function clearBatch() {
            batchRows = [];
            batchRowCount = 0;
            lastBatchResult = null;
            renderBatchTable();
            document.getElementById('batchResults').innerHTML = '';
            document.getElementById('batchExportBtns').style.display = 'none';
        }

        function copyBatchResult() {
            if (!lastBatchResult) return;
            const { resultRows, grandStorage, grandDetention, grandTotal, detentionActive } = lastBatchResult;
            const sep = '─'.repeat(52);
            const lines = [];
            lines.push('RESUMO — MÚLTIPLOS CONTENTORES');
            lines.push(sep);

            resultRows.forEach(row => {
                if (row.error) {
                    lines.push(`${row.ref}  →  ERRO: ${row.error}`);
                    return;
                }
                const detStr = detentionActive
                    ? (row.detCost != null ? `  |  Detention: €${row.detCost.toFixed(2)}` : '  |  Detention: sem data')
                    : '';
                lines.push(`${row.ref}  →  Storage: €${row.storageCost.toFixed(2)}${detStr}  |  TOTAL: €${row.total.toFixed(2)}`);
            });

            lines.push(sep);
            if (detentionActive) {
                lines.push(`Storage total:   €${grandStorage.toFixed(2)}`);
                lines.push(`Detention total: €${grandDetention.toFixed(2)}`);
                lines.push(sep);
            }
            lines.push(`TOTAL GERAL: €${grandTotal.toFixed(2)}`);

            navigator.clipboard.writeText(lines.join('\n')).then(() => {
                const btn = document.querySelector('#batchExportBtns .result-copy-btn');
                const orig = btn.textContent;
                btn.textContent = 'Copiado!';
                btn.style.background = '#1e3a5f';
                btn.style.color = 'white';
                setTimeout(() => { btn.textContent = orig; btn.style.background = ''; btn.style.color = ''; }, 2000);
            });
        }

        function exportBatchPDF() {
            if (!lastBatchResult) return;
            
            const { resultRows, grandStorage, grandDetention, grandTotal, detentionActive } = lastBatchResult;
            const today = new Date().toLocaleDateString('pt-PT', { day: '2-digit', month: '2-digit', year: 'numeric' });

            const totalColumns = detentionActive ? 7 : 6;
            const detentionHeader = detentionActive ? '<th style="text-align:right;">Detention (€)</th>' : '';

            let rowsHTML = '';
            resultRows.forEach(row => {
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
                    ? `<td style="text-align:right;">${row.detCost != null ? `€${row.detCost.toFixed(2)}` : '—'}</td>`
                    : '';

                rowsHTML += `
                    <tr>
                        <td>${escapeHTML(row.ref)}</td>
                        <td>${escapeHTML(row.releaseDateLabel)}</td>
                        <td>${escapeHTML(row.pickupDateLabel)}</td>
                        <td>${escapeHTML(row.returnDateLabel)}</td>
                        <td style="text-align:right;">€${row.storageCost.toFixed(2)}</td>
                        ${detentionCell}
                        <td style="text-align:right;font-weight:700;color:#0f2744;">€${row.total.toFixed(2)}</td>
                    </tr>`;
            });

            const breakdownTotals = detentionActive
                ? `
                    <div class="batch-report-split">
                        <div class="batch-report-split-item">
                            <div class="batch-report-split-label">Storage</div>
                            <div class="batch-report-split-value">€${grandStorage.toFixed(2)}</div>
                        </div>
                        <div class="batch-report-split-item">
                            <div class="batch-report-split-label">Detention</div>
                            <div class="batch-report-split-value">€${grandDetention.toFixed(2)}</div>
                        </div>
                    </div>`
                : '';

            const htmlContent = `
                <style>
                    .batch-report-wrap {
                        font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
                        max-width: 980px;
                        margin: 0 auto;
                        color: #1a1a1a;
                    }
                    .batch-report-head {
                        display: flex;
                        justify-content: space-between;
                        align-items: flex-start;
                        border-bottom: 2px solid #1e3a5f;
                        padding-bottom: 12px;
                        margin-bottom: 16px;
                        gap: 12px;
                    }
                    .batch-report-brand {
                        display: flex;
                        align-items: center;
                        gap: 12px;
                    }
                    .batch-report-logo {
                        width: 52px;
                        height: 52px;
                        object-fit: contain;
                        border: 1px solid #d7dee8;
                        border-radius: 8px;
                        padding: 5px;
                        background: white;
                    }
                    .batch-report-title {
                        margin: 0;
                        font-size: 18px;
                        color: #1e3a5f;
                    }
                    .batch-report-sub {
                        margin: 3px 0 0 0;
                        font-size: 12px;
                        color: #667;
                    }
                    .batch-report-meta {
                        text-align: right;
                        font-size: 12px;
                        color: #667;
                        line-height: 1.4;
                    }
                    .batch-report-grand {
                        background: #1e3a5f;
                        color: white;
                        border-radius: 8px;
                        padding: 14px 16px;
                        margin-bottom: 12px;
                        text-align: center;
                    }
                    .batch-report-grand-label {
                        font-size: 11px;
                        text-transform: uppercase;
                        letter-spacing: 1px;
                        opacity: 0.85;
                        margin-bottom: 4px;
                    }
                    .batch-report-grand-value {
                        font-size: 34px;
                        font-weight: 700;
                    }
                    .batch-report-split {
                        display: grid;
                        grid-template-columns: 1fr 1fr;
                        gap: 10px;
                        margin-bottom: 14px;
                    }
                    .batch-report-split-item {
                        border: 1px solid #dbe3ec;
                        border-radius: 8px;
                        padding: 10px 12px;
                        background: #f7fafc;
                    }
                    .batch-report-split-label {
                        font-size: 11px;
                        text-transform: uppercase;
                        color: #68788b;
                        letter-spacing: 0.6px;
                        margin-bottom: 4px;
                    }
                    .batch-report-split-value {
                        font-size: 20px;
                        font-weight: 700;
                        color: #1e3a5f;
                    }
                    .batch-report-table {
                        width: 100%;
                        border-collapse: collapse;
                        font-size: 12px;
                    }
                    .batch-report-table th {
                        background: #f0f4f8;
                        color: #1e3a5f;
                        padding: 9px 10px;
                        border: 1px solid #dbe3ec;
                        font-size: 11px;
                        text-transform: uppercase;
                        letter-spacing: 0.5px;
                    }
                    .batch-report-table td {
                        border: 1px solid #e5eaf0;
                        padding: 8px 10px;
                    }
                    .batch-report-table tr:nth-child(even):not(.batch-report-error) {
                        background: #fbfdff;
                    }
                    .batch-report-error td {
                        background: #fff5f5;
                        color: #a12323;
                        font-weight: 600;
                    }
                    .batch-report-foot {
                        margin-top: 10px;
                        color: #667;
                        font-size: 11px;
                        text-align: right;
                    }
                    @media print {
                        @page {
                            size: A4;
                            margin: 10mm;
                        }

                        body {
                            margin: 0;
                        }

                        .batch-report-wrap {
                            width: 100%;
                            max-width: 100%;
                        }

                        .batch-report-head,
                        .batch-report-grand,
                        .batch-report-split,
                        .batch-report-foot {
                            page-break-inside: avoid;
                            break-inside: avoid;
                        }

                        .batch-report-grand {
                            padding: 10px 12px;
                            margin-bottom: 10px;
                        }

                        .batch-report-grand-value {
                            font-size: 28px;
                        }

                        .batch-report-split {
                            margin-bottom: 10px;
                        }

                        .batch-report-split-item {
                            padding: 8px 10px;
                        }

                        .batch-report-table tr {
                            page-break-inside: avoid;
                            break-inside: avoid;
                        }

                        .batch-report-table {
                            width: 100%;
                            table-layout: fixed;
                            page-break-inside: auto;
                            break-inside: auto;
                            font-size: 10px;
                        }

                        .batch-report-table thead {
                            display: table-header-group;
                        }

                        .batch-report-table th,
                        .batch-report-table td {
                            padding: 6px 6px;
                            font-size: 10px;
                            word-break: break-word;
                        }

                        .batch-report-table th:nth-child(1),
                        .batch-report-table td:nth-child(1) {
                            width: 13%;
                        }

                        .batch-report-table th:nth-child(2),
                        .batch-report-table td:nth-child(2),
                        .batch-report-table th:nth-child(3),
                        .batch-report-table td:nth-child(3),
                        .batch-report-table th:nth-child(4),
                        .batch-report-table td:nth-child(4) {
                            width: 15%;
                        }

                        .batch-report-table th:nth-child(5),
                        .batch-report-table td:nth-child(5),
                        .batch-report-table th:nth-child(6),
                        .batch-report-table td:nth-child(6),
                        .batch-report-table th:nth-child(7),
                        .batch-report-table td:nth-child(7) {
                            width: 9%;
                        }
                    }
                </style>
                <div class="batch-report-wrap">
                    <div class="batch-report-head">
                        <div class="batch-report-brand">
                            <img src="${NAVEX_LOGO_SRC}" alt="Símbolo" class="batch-report-logo">
                            <div>
                                <h2 class="batch-report-title">Múltiplos Contentores</h2>
                                <p class="batch-report-sub">Resumo de Custos</p>
                            </div>
                        </div>
                        <div class="batch-report-meta">
                            <div>Emitido em ${today}</div>
                            <div>Contentores: ${resultRows.length}</div>
                        </div>
                    </div>

                    <div class="batch-report-grand">
                        <div class="batch-report-grand-label">Total a Pagar</div>
                        <div class="batch-report-grand-value">€${grandTotal.toFixed(2)}</div>
                    </div>

                    ${breakdownTotals}

                    <table class="batch-report-table">
                        <thead>
                            <tr>
                                <th>Referência</th>
                                <th>Descarga</th>
                                <th>Levantamento</th>
                                <th>Devolução</th>
                                <th style="text-align:right;">Storage (€)</th>
                                ${detentionHeader}
                                <th style="text-align:right;">Total (€)</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${rowsHTML}
                        </tbody>
                    </table>

                    <div class="batch-report-foot">Documento gerado pela calculadora</div>
                </div>`;

            const area = ensureBatchPrintArea();
            area.innerHTML = htmlContent;
            triggerPrint('batch');
        }

        // ── Excel template download ───────────────────────────────────────────

        function downloadTemplate() {
            if (!ensureXLSXLoaded()) return;
            const headers = ['Referencia', 'Descarga', 'Levantamento', 'Dias_Livres_Storage', 'Devolucao', 'Dias_Livres_Detention'];
            const example = ['CONT001', '01/03/2025', '15/03/2025', '7', '30/03/2025', '7'];
            const ws = XLSX.utils.aoa_to_sheet([headers, example]);

            // Column widths
            ws['!cols'] = [{ wch: 16 }, { wch: 18 }, { wch: 18 }, { wch: 20 }, { wch: 18 }, { wch: 22 }];

            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Contentores');
            XLSX.writeFile(wb, 'template_contentores.xlsx');
        }

        // ── Excel file upload ─────────────────────────────────────────────────

        // Maps common column name variants to our internal field names
        const COLUMN_MAP = {
            'referencia': 'ref',
            'referência': 'ref',
            'ref': 'ref',
            'contentor': 'ref',
            'container': 'ref',
            'descarregamento': 'releaseDate',
            'release': 'releaseDate',
            'release_date': 'releaseDate',
            'levantamento': 'pickupDate',
            'pickup': 'pickupDate',
            'pickup_date': 'pickupDate',
            'dias_livres_storage': 'freeDays',
            'dias livres storage': 'freeDays',
            'dias_livres': 'freeDays',
            'free_days': 'freeDays',
            'free_days_storage': 'freeDays',
            'devolucao': 'returnDate',
            'devolução': 'returnDate',
            'return': 'returnDate',
            'return_date': 'returnDate',
            'dias_livres_detention': 'detentionFreeDays',
            'dias livres detention': 'detentionFreeDays',
            'free_days_detention': 'detentionFreeDays',
        };

        function normalizeHeader(h) {
            return String(h).toLowerCase().trim().replace(/\s+/g, ' ');
        }

        // ── Batch section toggle ──────────────────────────────────────────────

        function toggleBatchSection() {
            const body = document.getElementById('batchSectionBody');
            const arrow = document.getElementById('batchToggleArrow');
            const isOpen = body.classList.toggle('open');
            arrow.classList.toggle('open', isOpen);
        }

        // ── Excel date conversion ─────────────────────────────────────────────

        // SheetJS with cellDates:true returns JS Date objects for date cells.
        // <input type="date"> needs YYYY-MM-DD.
        function excelDateToInputValue(val) {
            if (val == null || val === '') return '';
            if (val instanceof Date) {
                const y = val.getFullYear();
                const m = String(val.getMonth() + 1).padStart(2, '0');
                const d = String(val.getDate()).padStart(2, '0');
                return `${y}-${m}-${d}`;
            }
            // String typed by user: could be DD/MM/YYYY or YYYY-MM-DD
            const s = String(val).trim();
            // DD/MM/YYYY → YYYY-MM-DD
            const dmY = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
            if (dmY) return `${dmY[3]}-${dmY[2].padStart(2,'0')}-${dmY[1].padStart(2,'0')}`;
            // Already YYYY-MM-DD
            if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
            return '';
        }

        function processExcelFile(file) {
            if (!ensureXLSXLoaded()) return;
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const wb = XLSX.read(data, { type: 'array', cellDates: true });
                    const ws = wb.Sheets[wb.SheetNames[0]];
                    // raw:true so we get Date objects for date cells (not formatted strings)
                    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });

                    if (rows.length < 2) {
                        alert('O ficheiro não tem dados suficientes (mínimo 1 linha de cabeçalho + 1 linha de dados).');
                        return;
                    }

                    // Map headers
                    const headerRow = rows[0].map(normalizeHeader);
                    const fieldMap = {};
                    headerRow.forEach((h, i) => {
                        const field = COLUMN_MAP[h];
                        if (field) fieldMap[i] = field;
                    });

                    if (Object.keys(fieldMap).length === 0) {
                        alert('Nenhuma coluna reconhecida. Verifique se o cabeçalho corresponde ao template.');
                        return;
                    }

                    const dateFields = new Set(['releaseDate', 'pickupDate', 'returnDate']);
                    const newRows = [];
                    for (let r = 1; r < rows.length; r++) {
                        const row = rows[r];
                        if (!row || row.every(cell => cell == null || cell === '')) continue;
                        const obj = {};
                        Object.entries(fieldMap).forEach(([colIdx, field]) => {
                            const val = row[colIdx];
                            obj[field] = dateFields.has(field)
                                ? excelDateToInputValue(val)
                                : (val != null ? String(val).trim() : '');
                        });
                        newRows.push(obj);
                    }

                    if (newRows.length === 0) {
                        alert('Não foram encontradas linhas com dados.');
                        return;
                    }

                    batchRows = [];
                    batchRowCount = 0;
                    newRows.forEach(row => {
                        const id = batchRowCount++;
                        batchRows.push({ id, ...row });
                    });
                    renderBatchTable();
                    document.getElementById('excelFileInput').value = '';

                } catch (err) {
                    console.error(err);
                    alert('Erro ao ler o ficheiro: ' + err.message);
                }
            };
            reader.readAsArrayBuffer(file);
        }

        function loadExcelFile(event) {
            const file = event.target.files[0];
            if (file) processExcelFile(file);
        }

