document.addEventListener('DOMContentLoaded', function () {
    'use strict';
    const els = {
        fileInput: document.getElementById('fileInput'),
        dropZone: document.getElementById('dropZone'),
        loadBtn: document.getElementById('loadBtn'),
        fileDisplay: document.getElementById('fileDisplayName'),
        fileStatus: document.getElementById('fileStatus'),
        results: document.getElementById('results'),
        filter: document.getElementById('filter'),
        dateFilter: document.getElementById('dateFilter'),
        periodFilter: document.getElementById('periodFilter'),
        techFilter: document.getElementById('techFilter'),
        regionFilter: document.getElementById('regionFilter'),
        moreBtn: document.getElementById('moreBtn'),
        toolbar: document.getElementById('toolbarSection'),
        loader: document.getElementById('globalLoader'),
        loaderBar: document.getElementById('loaderBar'),
        loaderText: document.querySelector('.loader-text'),
        toastContainer: document.getElementById('toastContainer'),
        overlay: document.getElementById('overlay'),
        modalBody: document.getElementById('modalBody'),
        closeModal: document.getElementById('closeModal'),
        templatesModal: document.getElementById('templatesModal'),
        numberEntryModal: document.getElementById('numberEntryModal'),
        inlinePhoneInput: document.getElementById('inlinePhoneInput'),
        savePhoneBtn: document.getElementById('savePhoneBtn'),
        manualPhoneTriggerBtn: document.getElementById('manualPhoneTriggerBtn'),
        copyListBtn: document.getElementById('copyLinkListBtn'),
        openRouteBtn: document.getElementById('openRouteBtn'),
        statsDashboard: document.getElementById('statsDashboard'),
        statTotal: document.getElementById('totalCount'),
        statBalsa: document.getElementById('balsaCount'),
        statOverdue: document.getElementById('overdueCount'),
        techWrapper: null,
        techChipsContainer: null
    };
    let state = {
        items: [],
        filtered: [],
        renderedCount: 0,
        currentFile: null,
        activeFilter: 'all',
        itemForTemplate: null,
        PAGE_SIZE: 50,
        todayStr: '',
        debounceTimer: null,
        selectedTechs: []
    };
    const REGEX = {
        BALSA: /balsa|itaquacetuba|boror[eé]|curucutu|tatetos|agua limpa|jardim santa tereza|jardim borba gato|p[oó]s balsa|capivari|santa cruz/i,
        RIACHO: /fincos|tup[ãa]|rio grande|boa vista|capelinha|cocaia|zanzala|vila pele|riacho grande|arei[aã]o|jussara|balne[aá]ria/i,
        BLOCKED: /\b(retirada|reagend|tentativa|^re$|removido)\b/i
    };
    function toLocalIsoDate(d) {
        if (!(d instanceof Date) || isNaN(d)) return '';
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + day;
    }
    function initToday() {
        state.todayStr = toLocalIsoDate(new Date());
    }
    function getMapsLink(q) {
        if (!q) return '#';
        return 'https://www.google.com/maps/search/' + encodeURIComponent(q);
    }
    function getRouteLink(destinations) {
        if (!destinations || destinations.length === 0) return '#';
        const originPlaceholder = 'Meu Local';
        const path = [originPlaceholder, ...destinations.slice(0, 10)].map(encodeURIComponent).join('/');
        return 'https://www.google.com/maps/dir/' + path;
    }
    function showToast(msg, type) {
        const div = document.createElement('div');
        div.className = 'toast';
        const span = document.createElement('span');
        span.textContent = msg;
        div.appendChild(span);
        let color = 'var(--primary)';
        if (type === 'error') color = 'var(--danger)';
        if (type === 'success') color = 'var(--success)';
        div.style.borderLeftColor = color;
        els.toastContainer.appendChild(div);
        setTimeout(function () {
            div.style.opacity = '0';
            div.style.transform = 'translateX(100%)';
            setTimeout(function () { div.remove(); }, 300);
        }, 3000);
    }
    function setLoader(percent, text) {
        if (percent === 0) {
            els.loader.classList.add('hidden');
            els.loaderBar.style.width = '0%';
            if (els.loaderText) els.loaderText.textContent = '';
            return;
        }
        els.loader.classList.remove('hidden');
        els.loaderBar.style.width = percent + '%';
        if (text && els.loaderText) els.loaderText.textContent = text;
    }
    function normalizeKey(key) {
        return String(key).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '');
    }
    function convertExcelDate(serial) {
        if (typeof serial !== 'number') return null;
        const excelEpoch = Date.UTC(1899, 11, 30);
        const ms = Math.round(serial * 86400 * 1000);
        return new Date(excelEpoch + ms);
    }
    function formatDateTime(val) {
        if (!val) return '';
        if (val instanceof Date && !isNaN(val)) {
            const d = val;
            const day = String(d.getDate()).padStart(2, '0');
            const month = String(d.getMonth() + 1).padStart(2, '0');
            const hours = String(d.getHours()).padStart(2, '0');
            const minutes = String(d.getMinutes()).padStart(2, '0');
            return day + '/' + month + ' às ' + hours + ':' + minutes;
        }
        const s = String(val).trim();
        let datePart = s;
        let timePart = '';
        if (s.includes(' ')) { const parts = s.split(' '); datePart = parts[0]; timePart = parts.slice(1).join(' '); }
        timePart = timePart ? timePart.substring(0, 5) : '';
        if (datePart.match(/^\d{4}-\d{2}-\d{2}/)) { const p = datePart.split('-'); datePart = p[2] + '/' + p[1]; }
        else if (datePart.match(/^\d{2}\/\d{2}\/\d{4}/)) { const p = datePart.split('/'); datePart = p[0] + '/' + p[1]; }
        else { const parsed = Date.parse(s); if (!isNaN(parsed)) { const d = new Date(parsed); const day = String(d.getDate()).padStart(2, '0'); const month = String(d.getMonth() + 1).padStart(2, '0'); datePart = day + '/' + month; } }
        return timePart ? (datePart + ' às ' + timePart) : datePart;
    }
    function getDateOnly(val) {
        if (!val) return '';
        if (val instanceof Date && !isNaN(val)) return toLocalIsoDate(val);
        if (typeof val === 'number') {
            const d = convertExcelDate(val);
            if (d) return toLocalIsoDate(d);
        }
        const s = String(val).trim();

        // yyyy-mm-dd (com ou sem hora)
        let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (m) return m[1] + '-' + m[2] + '-' + m[3];

        // dd/mm/yyyy (com ou sem hora)
        m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
        if (m) return m[3] + '-' + m[2] + '-' + m[1];

        // dd/mm (sem ano): assume ano atual
        m = s.match(/^(\d{2})\/(\d{2})\b/);
        if (m) {
            const y = String(new Date().getFullYear());
            return y + '-' + m[2] + '-' + m[1];
        }
        const parsed = Date.parse(s);
        if (!isNaN(parsed)) return toLocalIsoDate(new Date(parsed));
        return '';
    }
    function buildAddress(it) {
        const parts = [];
        if (it.endereco) parts.push(it.endereco);
        if (it.numero) parts.push(it.numero);
        if (it.complemento) parts.push('(Comp: ' + it.complemento + ')');
        if (it.bairro) parts.push(it.bairro);
        if (it.cidade && String(it.cidade).toUpperCase() !== 'SAO PAULO') parts.push(it.cidade);
        return parts.filter(Boolean).join(', ');
    }
    function buildCleanAddress(it) {
        const parts = [];
        if (it.endereco) parts.push(it.endereco);
        if (it.numero) parts.push(it.numero);
        let main = parts.join(', ');
        if (it.bairro) main += ' - ' + it.bairro;
        if (it.cidade) main += ' - ' + it.cidade;
        return main.replace(/, ,/g, ',').trim();
    }
    function mapData(rows) {
        if (!rows || rows.length < 2) return [];
        const headers = rows[0];
        const mapIdx = {};
        headers.forEach(function (h, index) {
            if (!h) return;
            const norm = normalizeKey(h);
            if (norm.includes('cliente') || norm.includes('nome')) mapIdx.cliente = index;
            else if (norm.includes('filial')) mapIdx.filial = index;
            else if (norm.includes('assunto') || norm.includes('motivo')) mapIdx.assunto = index;
            else if (norm.includes('tecnico') || norm.includes('colaborador')) mapIdx.tecnico = index;
            else if (norm.includes('endereco') || norm.includes('logradouro')) mapIdx.endereco = index;
            else if (norm.includes('numero')) mapIdx.numero = index;
            else if (norm.includes('complemento')) mapIdx.complemento = index;
            else if (norm.includes('bairro')) mapIdx.bairro = index;
            else if (norm.includes('cidade')) mapIdx.cidade = index;
            else if (norm.includes('referencia') || norm.includes('obs')) mapIdx.referencia = index;
            else if (norm.includes('login') || norm.includes('usuario')) mapIdx.login = index;
            else if (norm.includes('senha') || norm.includes('password')) mapIdx.senha = index;
            else if (norm.includes('agendamento')) mapIdx.agendamento = index;
            else if (norm.includes('melhor') && norm.includes('horario')) mapIdx.horario = index;
            else if (norm.includes('contrato') || norm.includes('plano')) mapIdx.contrato = index;
            else if (norm.includes('data') && norm.includes('reservada')) mapIdx.dataReserva = index;
            else if (norm.includes('telefone') || norm.includes('celular') || norm.includes('whatsapp')) {
                if (!mapIdx.telefones) mapIdx.telefones = [];
                mapIdx.telefones.push(index);
            }
        });
        const data = [];
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;
            const tels = [];
            if (mapIdx.telefones) {
                mapIdx.telefones.forEach(function (idx) { if (row[idx]) tels.push(String(row[idx]).trim()); });
            }
            const rawAgenda = (mapIdx.agendamento !== undefined) ? row[mapIdx.agendamento] : null;
            const horarioTexto = (mapIdx.horario !== undefined && row[mapIdx.horario]) ? String(row[mapIdx.horario]).toLowerCase() : '';
            let periodoClass = '';
            if (horarioTexto.includes('manh')) periodoClass = 'slot-manha';
            else if (horarioTexto.includes('tarde')) periodoClass = 'slot-tarde';
            const item = {
                cliente: (mapIdx.cliente !== undefined && row[mapIdx.cliente]) ? String(row[mapIdx.cliente]) : 'Cliente Desconhecido',
                filial: (mapIdx.filial !== undefined) ? String(row[mapIdx.filial]) : '',
                assunto: (mapIdx.assunto !== undefined) ? String(row[mapIdx.assunto]) : '',
                tecnico: (mapIdx.tecnico !== undefined && row[mapIdx.tecnico]) ? String(row[mapIdx.tecnico]) : '',
                endereco: (mapIdx.endereco !== undefined) ? String(row[mapIdx.endereco]) : '',
                numero: (mapIdx.numero !== undefined) ? String(row[mapIdx.numero]) : '',
                complemento: (mapIdx.complemento !== undefined) ? String(row[mapIdx.complemento]) : '',
                bairro: (mapIdx.bairro !== undefined) ? String(row[mapIdx.bairro]) : '',
                cidade: (mapIdx.cidade !== undefined) ? String(row[mapIdx.cidade]) : '',
                referencia: (mapIdx.referencia !== undefined) ? String(row[mapIdx.referencia]) : '',
                login: (mapIdx.login !== undefined) ? String(row[mapIdx.login]) : '',
                senha: (mapIdx.senha !== undefined) ? String(row[mapIdx.senha]) : '',
                horario: (mapIdx.horario !== undefined) ? String(row[mapIdx.horario]) : '',
                contrato: (mapIdx.contrato !== undefined) ? String(row[mapIdx.contrato]) : '',
                agendamento: rawAgenda === null ? '' : rawAgenda,
                agendamentoIso: getDateOnly(rawAgenda),
                periodoClass: periodoClass,
                telefone: tels.filter(Boolean).join(' / '),
                raw: row
            };
            item.isToday = (item.agendamentoIso === state.todayStr);
            item.isOverdue = (item.agendamentoIso && item.agendamentoIso < state.todayStr);
            const fullAddr = (item.endereco + ' ' + item.bairro + ' ' + item.cidade).toLowerCase();
            const cepMatch = String(item.endereco).match(/\d{5}-?\d{3}/);
            const cep = cepMatch ? cepMatch[0].replace(/\D/g, '').substring(0, 3) : '';
            if (cep === '078' || fullAddr.includes('franco')) item.region = '078-franco';
            else if (cep === '048' || fullAddr.includes('graja')) item.region = '048-grajau';
            else if (cep === '041' || fullAddr.includes('moraes') || fullAddr.includes('mv') || fullAddr.includes('mundo virtua')) item.region = '041-mv';
            else if (cep === '098' || fullAddr.includes('bernardo') || fullAddr.includes('riacho')) {
                if (REGEX.BALSA.test(fullAddr)) item.region = '098-balsa';
                else if (REGEX.RIACHO.test(fullAddr)) item.region = '098-riacho';
                else item.region = '098-sbc';
            } else item.region = 'other';
            if (!REGEX.BLOCKED.test(item.assunto)) data.push(item);
        }
        return data.sort(function (a, b) {
            if (a.isOverdue && !b.isOverdue) return -1;
            if (!a.isOverdue && b.isOverdue) return 1;
            if (a.isToday && !b.isToday) return -1;
            if (!a.isToday && b.isToday) return 1;
            return 0;
        });
    }
    function createCard(it) {
        const template = document.getElementById('cardTemplate');
        if (!template) return document.createElement('div');
        const clone = template.content.cloneNode(true);
        const card = clone.querySelector('.result-card');
        const displayAddr = buildAddress(it);
        const cleanAddr = buildCleanAddress(it);
        const hasAddr = cleanAddr.length > 10 && !/n[ãa]o (informado|consta|nao consta)/i.test(cleanAddr);
        card.classList.add('region-' + it.region);
        if (it.periodoClass) card.classList.add(it.periodoClass);
        if (it.isToday) card.classList.add('is-today');
        if (it.isOverdue) card.classList.add('is-overdue');
        if (!hasAddr) card.classList.add('no-addr');
        const clientEl = card.querySelector('.card-client');
        if (clientEl) clientEl.textContent = it.cliente;
        let subjectText = it.assunto || '';
        if (it.agendamento) {
            const fmtDate = formatDateTime(it.agendamento);
            if (fmtDate.length > 2) subjectText = 'Agendamento: ' + fmtDate + ' - ' + it.assunto;
        }
        const subjEl = card.querySelector('.card-subject');
        if (subjEl) subjEl.textContent = subjectText;
        const techEl = card.querySelector('.card-tech');
        if (techEl) techEl.textContent = it.tecnico || 'Sem Técnico';
        const addrEl = card.querySelector('.card-address');
        if (addrEl) addrEl.textContent = hasAddr ? displayAddr : 'Endereço não identificado';
        const badgeBox = card.querySelector('.card-badges');
        while (badgeBox && badgeBox.firstChild) badgeBox.removeChild(badgeBox.firstChild);
        if (it.region === '098-balsa') {
            const badge = document.createElement('span');
            badge.className = 'badge-count';
            badge.style.cssText = 'background:var(--warning);color:#000';
            badge.textContent = 'Balsa';
            if (badgeBox) badgeBox.appendChild(badge);
        }
        if (!hasAddr) {
            const badge = document.createElement('span');
            badge.className = 'badge-count';
            badge.style.cssText = 'background:var(--danger);color:#fff';
            badge.textContent = 'Sem Endereço';
            if (badgeBox) badgeBox.appendChild(badge);
        }
        if (it.contrato && String(it.contrato).includes('MB')) {
            const speed = String(it.contrato).split(' ').pop();
            const badge = document.createElement('span');
            badge.className = 'badge-count';
            badge.style.cssText = 'background:var(--bg-element);border:1px solid var(--border)';
            badge.textContent = speed;
            if (badgeBox) badgeBox.appendChild(badge);
        }
        if (it.horario) {
            const badge = document.createElement('span');
            badge.className = 'badge-count';
            badge.style.cssText = 'background:var(--bg-element);border:1px solid var(--border)';
            badge.textContent = it.horario;
            if (badgeBox) badgeBox.appendChild(badge);
        }
        const btnMaps = card.querySelector('.open-maps');
        if (btnMaps) {
            btnMaps.href = getMapsLink(cleanAddr);
            if (!hasAddr) { btnMaps.style.opacity = '0.5'; btnMaps.setAttribute('aria-disabled', 'true'); }
            else { btnMaps.style.opacity = ''; btnMaps.removeAttribute('aria-disabled'); }
        }
        const btnZap = card.querySelector('.open-zap');
        if (btnZap) btnZap.onclick = function (e) { e.stopPropagation(); state.itemForTemplate = it; els.templatesModal.classList.remove('hidden'); };
        const btnCreds = card.querySelector('.copy-creds');
        if (btnCreds) btnCreds.onclick = function (e) {
            e.stopPropagation();
            const txt = it.login + '\n \n' + it.senha + '\n \nVlan 500';
            navigator.clipboard.writeText(txt).then(function () { showToast('Login copiado!', 'success'); }).catch(function () { showToast('Erro ao copiar credenciais', 'error'); });
        };
        const btnDetails = card.querySelector('.details-btn');
        if (btnDetails) btnDetails.onclick = function (e) { e.stopPropagation(); openDetails(it); };
        card.onclick = function (e) {
            if (e.target === card || e.target.closest('.card-main-info')) openDetails(it);
        };
        return clone;
    }
    function renderEmptyState(title, description) {
        els.results.innerHTML =
            '<div class="empty-state-card">' +
            '<h3 class="empty-state-title">' + title + '</h3>' +
            '<p class="empty-state-text">' + description + '</p>' +
            '</div>';
        els.moreBtn.classList.add('hidden');
    }
    function render() {
        const filtered = state.filtered;
        const container = els.results;
        if (state.renderedCount === 0) container.innerHTML = '';
        if (!filtered || filtered.length === 0) {
            renderEmptyState('Nenhum resultado encontrado', 'Ajuste os filtros ou carregue outra planilha para continuar.');
            return;
        }
        const nextBatch = filtered.slice(state.renderedCount, state.renderedCount + state.PAGE_SIZE);
        const frag = document.createDocumentFragment();
        nextBatch.forEach(function (it) { frag.appendChild(createCard(it)); });
        container.appendChild(frag);
        state.renderedCount += nextBatch.length;
        if (state.renderedCount >= filtered.length) els.moreBtn.classList.add('hidden'); else els.moreBtn.classList.remove('hidden');
        updateCounters();
    }
    function filterData() {
        const term = (els.filter.value || '').toLowerCase();
        const region = els.regionFilter.value;
        const dateVal = els.dateFilter ? els.dateFilter.value : '';
        const period = els.periodFilter ? els.periodFilter.value : '';
        const quick = state.activeFilter;
        const techVals = state.selectedTechs.slice();
        state.filtered = state.items.filter(function (it) {
            const searchString = (String(it.cliente) + ' ' + String(it.endereco) + ' ' + String(it.login) + ' ' + String(it.assunto) + ' ' + String(it.contrato)).toLowerCase();
            const matchText = searchString.includes(term);
            const matchTech = techVals.length === 0 || techVals.indexOf(it.tecnico) >= 0;
            const matchRegion = region === 'all' || it.region === region;
            const matchDate = !dateVal || it.agendamentoIso === dateVal;
            let matchPeriod = true;
            if (period === 'manha') matchPeriod = it.periodoClass === 'slot-manha';
            else if (period === 'tarde') matchPeriod = it.periodoClass === 'slot-tarde';
            let matchQuick = true;
            if (quick === 'balsa') matchQuick = it.region === '098-balsa';
            else if (quick === 'overdue') matchQuick = it.isOverdue;
            else if (quick === 'today') matchQuick = it.isToday;
            else if (quick === 'noaddr') matchQuick = buildAddress(it).length < 10;
            return matchText && matchTech && matchRegion && matchDate && matchPeriod && matchQuick;
        });
        state.renderedCount = 0;
        render();
    }
    function updateCounters() {
        const balsaCount = state.items.filter(function (it) { return it.region === '098-balsa'; }).length;
        const overdueCount = state.items.filter(function (it) { return it.isOverdue; }).length;
        const badge = document.querySelector('.chip[data-filter="balsa"] .badge-count');
        if (badge) { badge.textContent = balsaCount; badge.classList.toggle('hidden', balsaCount === 0); }
        if (els.statTotal) els.statTotal.textContent = state.items.length;
        if (els.statBalsa) els.statBalsa.textContent = balsaCount;
        if (els.statOverdue) els.statOverdue.textContent = overdueCount;
        if (els.statsDashboard && state.items.length > 0) els.statsDashboard.classList.remove('hidden'); else if (els.statsDashboard) els.statsDashboard.classList.add('hidden');
    }
    function populateSelects() {
        const techs = [...new Set(state.items.map(function (i) { return i.tecnico; }).filter(Boolean))].sort();
        els.techFilter.innerHTML = '<option value="">Todos os Técnicos</option>';
        techs.forEach(function (t) { const opt = document.createElement('option'); opt.value = t; opt.textContent = t; els.techFilter.appendChild(opt); });
        renderTechChips(techs);
    }
    function renderTechChips(list) {
        if (!els.techFilter) return;
        if (els.techChipsContainer) {
            els.techChipsContainer.remove();
            els.techChipsContainer = null;
        }
        if (els.techFilter.style.display === 'none') {
            els.techFilter.style.display = '';
        }
    }
    function openDetails(it) {
        const tpl = document.getElementById('modalTemplate');
        if (!tpl) { showToast('Template de modal não encontrado', 'error'); return; }
        const clone = tpl.content.cloneNode(true);
        const addr = buildAddress(it);
        const cleanAddr = buildCleanAddress(it);
        const mapQ = encodeURIComponent(cleanAddr);
        const hasAddr = cleanAddr.length > 10;
        const clientEl = clone.querySelector('.modal-client'); if (clientEl) clientEl.textContent = it.cliente;
        const branchEl = clone.querySelector('.modal-branch'); if (branchEl) branchEl.textContent = it.filial || '-';
        const techEl = clone.querySelector('.modal-tech'); if (techEl) techEl.textContent = it.tecnico || '-';
        const loginEl = clone.querySelector('.modal-login'); if (loginEl) loginEl.textContent = (it.login || '-') + ' / ' + (it.senha || '-');
        const phoneEl = clone.querySelector('.modal-phone'); if (phoneEl) phoneEl.textContent = it.telefone || '-';
        const subjectEl = clone.querySelector('.modal-subject'); if (subjectEl) subjectEl.textContent = it.assunto || '-';
        const addressEl = clone.querySelector('.modal-address'); if (addressEl) addressEl.textContent = addr || 'Endereço não cadastrado';
        const refEl = clone.querySelector('.modal-ref'); if (refEl) refEl.textContent = it.referencia || '-';
        const contractEl = clone.querySelector('.modal-contract'); if (contractEl) contractEl.textContent = it.contrato || 'Não informado';
        const agendaEl = clone.querySelector('.modal-agenda'); if (agendaEl) agendaEl.textContent = formatDateTime(it.agendamento) || 'Sem agendamento';
        const periodoEl = clone.querySelector('.modal-periodo'); if (periodoEl) periodoEl.textContent = it.horario || '-';
        const mapContainer = clone.querySelector('#mapContainer');
        if (hasAddr) {
            const embedUrl = 'https://www.google.com/maps?q=' + mapQ + '&output=embed';
            const existingIframe = mapContainer.querySelector('.map-iframe');
            if (existingIframe) { existingIframe.src = embedUrl; } else {
                mapContainer.innerHTML = '';
                const newIframe = document.createElement('iframe');
                newIframe.className = 'map-iframe';
                newIframe.src = embedUrl;
                newIframe.loading = 'lazy';
                newIframe.allowFullscreen = true;
                mapContainer.appendChild(newIframe);
            }
        } else {
            if (mapContainer) mapContainer.innerHTML = '<div style="text-align:center;padding:20px;color:var(--text-secondary)">Endereço insuficiente para exibir mapa.</div>';
        }
        const btns = clone.querySelectorAll('.provider-btn');
        btns.forEach(function (b) {
            b.onclick = function () {
                btns.forEach(function (x) { x.classList.remove('active'); });
                b.classList.add('active');
                if (b.dataset.provider === 'waze') window.open('https://waze.com/ul?q=' + mapQ + '&navigate=yes', '_blank');
                else if (b.dataset.provider === 'google' && hasAddr) window.open(getMapsLink(cleanAddr), '_blank');
            };
        });
        const copyCredsBtn = clone.querySelector('.copy-creds-btn');
        if (copyCredsBtn) copyCredsBtn.onclick = function (e) { const txt = it.login + '\n \n' + it.senha + '\n \nVlan 500'; navigator.clipboard.writeText(txt).then(function () { e.target.textContent = 'Copiado!'; setTimeout(function () { e.target.textContent = 'Copiar'; }, 1500); }).catch(function () { showToast('Erro ao copiar credenciais', 'error'); }); };
        const linkBtn = clone.querySelector('.copy-link-btn');
        if (linkBtn && hasAddr) linkBtn.onclick = function (e) { const link = getMapsLink(cleanAddr); navigator.clipboard.writeText(link).then(function () { e.target.textContent = 'Copiado!'; setTimeout(function () { e.target.textContent = 'Copiar Link'; }, 1500); }).catch(function () { showToast('Erro ao copiar link', 'error'); }); };
        else if (linkBtn) linkBtn.disabled = true;
        els.modalBody.innerHTML = '';
        els.modalBody.appendChild(clone);
        els.overlay.classList.remove('hidden');
    }
    initToday();
    renderEmptyState(
        'Carregue a planilha para começar',
        'Selecione um arquivo .xls, .xlsx ou .csv para visualizar e filtrar as ordens de serviço.'
    );
    els.loadBtn.onclick = function () {
        if (!state.currentFile) { showToast('Nenhum arquivo selecionado.', 'info'); return; }
        setLoader(30, 'Lendo Arquivo...');
        const reader = new FileReader();
        reader.onload = function (e) {
            try {
                if (typeof XLSX === 'undefined') throw new Error('Biblioteca XLSX não carregada.');
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true, raw: false, defval: '' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: false, defval: '' });
                if (rows.length < 2) throw new Error('Planilha vazia ou sem cabeçalho');
                state.items = mapData(rows);
                populateSelects();
                setLoader(100, 'Finalizando...');
                setTimeout(function () { setLoader(0); els.toolbar.classList.remove('hidden'); filterData(); showToast(state.items.length + ' O.S. carregadas!', 'success'); }, 500);
            } catch (err) {
                console.error(err);
                setLoader(0);
                showToast('Erro: ' + (err.message || 'Verifique o formato do arquivo.'), 'error');
            }
        };
        reader.readAsArrayBuffer(state.currentFile);
    };
    els.fileInput.onchange = function (e) {
        if (!e.target.files.length) return;
        state.currentFile = e.target.files[0];
        els.fileDisplay.textContent = state.currentFile.name;
        if (els.fileStatus) els.fileStatus.textContent = 'Carregando planilha...';
        els.loadBtn.disabled = false;
        els.loadBtn.click();
    };
    els.dropZone.ondragover = function (e) { e.preventDefault(); els.dropZone.classList.add('dragover'); };
    els.dropZone.ondragleave = function () { els.dropZone.classList.remove('dragover'); };
    els.dropZone.ondrop = function (e) {
        e.preventDefault();
        els.dropZone.classList.remove('dragover');
        if (!e.dataTransfer.files.length) return;
        state.currentFile = e.dataTransfer.files[0];
        els.fileDisplay.textContent = state.currentFile.name;
        if (els.fileStatus) els.fileStatus.textContent = 'Carregando planilha...';
        els.loadBtn.disabled = false;
        els.loadBtn.click();
    };
    els.filter.oninput = function () { clearTimeout(state.debounceTimer); state.debounceTimer = setTimeout(filterData, 300); };
    if (els.dateFilter) els.dateFilter.onchange = filterData;
    if (els.periodFilter) els.periodFilter.onchange = filterData;
    if (els.techFilter) {
        els.techFilter.onchange = function () {
            const selected = Array.from(els.techFilter.selectedOptions || []);
            state.selectedTechs = selected.map(function (opt) { return opt.value; }).filter(Boolean);
            filterData();
        };
    }
    if (els.regionFilter) els.regionFilter.onchange = filterData;
    document.addEventListener('click', function (e) { if (e.target.classList && e.target.classList.contains('chip')) { document.querySelectorAll('.chip').forEach(function (c) { c.classList.remove('active'); c.setAttribute('aria-pressed', 'false'); }); e.target.classList.add('active'); e.target.setAttribute('aria-pressed', 'true'); state.activeFilter = e.target.dataset.filter; filterData(); } });
    els.moreBtn.onclick = render;
    els.manualPhoneTriggerBtn.onclick = function () { els.templatesModal.classList.add('hidden'); els.inlinePhoneInput.value = ''; els.numberEntryModal.classList.remove('hidden'); };
    document.addEventListener('click', function (e) { const btn = e.target.closest('.template-option'); if (!btn) return; const it = state.itemForTemplate; if (!it) { showToast('Selecione uma O.S. primeiro.', 'error'); return; } let text = btn.dataset.template; const hour = new Date().getHours(); const saudacao = hour < 12 ? 'Bom dia' : hour < 18 ? 'Boa tarde' : 'Boa noite'; const nome = (it.cliente || '').split(' ')[0]; text = text.replace('{{saudacao}}', saudacao).replace('{{nome}}', nome).replace('{{tratamento}}', 'Sr(a).'); const rawPhone = it.telefone ? it.telefone.split('/')[0] : ''; const phone = rawPhone.replace(/\D/g, '').replace(/^55/, ''); if (phone.length >= 10) { window.open('https://wa.me/55' + phone + '?text=' + encodeURIComponent(text), '_blank'); els.templatesModal.classList.add('hidden'); } else { els.templatesModal.classList.add('hidden'); els.inlinePhoneInput.value = ''; els.numberEntryModal.classList.remove('hidden'); } });
    els.savePhoneBtn.onclick = function () { const num = els.inlinePhoneInput.value.replace(/\D/g, '').replace(/^55/, ''); if (num.length < 10) { showToast('Número inválido', 'error'); return; } state.itemForTemplate.telefone = num; els.numberEntryModal.classList.add('hidden'); els.templatesModal.classList.remove('hidden'); showToast('Número salvo, selecione o template novamente.', 'info'); };
    els.closeModal.onclick = function () { els.overlay.classList.add('hidden'); };
    els.overlay.onclick = function (e) { if (e.target === els.overlay) { els.overlay.classList.add('hidden'); els.templatesModal.classList.add('hidden'); els.numberEntryModal.classList.add('hidden'); } };
els.copyListBtn.onclick = function () {
        if (!state.filtered.length) { showToast('Nenhuma O.S. para copiar.', 'info'); return; }
        const txt = state.filtered.slice(0, 50).map(function (i) { const clean = buildCleanAddress(i); const link = getMapsLink(clean); return '*' + i.cliente + '*\n' + link; }).join('\n\n');
        navigator.clipboard.writeText(txt).then(function () { showToast('Lista de ' + Math.min(state.filtered.length, 50) + ' O.S. copiada!', 'success'); }).catch(function () { showToast('Erro ao copiar lista', 'error'); });
    };
    els.openRouteBtn.onclick = function () {
        if (!state.filtered.length) { showToast('Nenhuma O.S. para rota.', 'info'); return; }
        const validItems = state.filtered.filter(function (it) { return buildCleanAddress(it).length > 5; }).slice(0, 10);
        if (validItems.length === 0) { showToast('Nenhum endereço válido para rota.', 'error'); return; }
        const dests = validItems.map(function (it) { return buildCleanAddress(it); });
        const routeLink = getRouteLink(dests);
        window.open(routeLink, '_blank');
    };
});

