(() => {
  const DIAGS = [
    "REAGENDADO PELO CLIENTE",
    "REAGENDADO DEVIDO HORARIO",
    "CLIENTE AUSENTE",
    "ENDERECO NAO LOCALIZADO"
  ];

  const OPEN_STATUSES = ["ABERTA", "ABERTO", "AGENDADA", "AGENDADO", "ASSUMIDA", "ASSUMIDO"];
  const FIELD_MAP = {
    diagnostico: ["DIAGNOSTICO", "DIAGNOSTICO", "DIAG"],
    cliente: ["CLIENTE", "NOME", "NAME"],
    login: ["LOGIN", "USUARIO", "USUARIO", "USERNAME", "USER", "EMAIL"],
    os: ["ORDEM DE SERVICO", "ORDEM SERVICO", "NUMERO OS", "NUM OS", "OS ID", "ID OS", "ID O.S", "ORDEM", "NUM_OS", "OS_ID"],
    status: ["STATUS", "SITUACAO", "SITUACAO", "STATUS_OS", "SITUACAO_OS"],
    assunto: ["ASSUNTO", "SUBJECT"],
    setor: ["SETOR"]
  };

  const MAX_RENDER_PER_REGION = 300;
  let rawRows = [];
  let displayRows = [];
  let currentRegion = "SBC";
  let copyMode = "nome";

  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("fileInput");
  const resultsDiv = document.getElementById("results");
  const actionsBar = document.getElementById("actionsBar");
  const regionBtns = document.querySelectorAll(".pill-btn");
  const btnCopyAll = document.getElementById("btnCopyAll");
  const btnCopyIdsFree = document.getElementById("btnCopyIdsFree");
  const btnReset = document.getElementById("btnReset");
  const loaderEl = document.getElementById("loader");
  const rowCountEl = document.getElementById("rowCount");
  const osPanel = document.getElementById("osPanel");
  const osPanelContent = document.getElementById("osPanelContent");
  const btnCopyOs = document.getElementById("btnCopyOs");

  function normalizeStr(v) {
    return String(v || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .trim();
  }

  function findFieldKey(keys, logical) {
    const candidates = FIELD_MAP[logical];
    if (!candidates) return null;

    // Para O.S., evita falso positivo com "DIAGNOSTICO" por causa de "OS".
    if (logical === "os") {
      for (const k of keys) {
        const nk = normalizeStr(k)
          .replace(/[^A-Z0-9]+/g, " ")
          .trim();
        if (
          nk === "OS" ||
          nk === "ID OS" ||
          nk === "OS ID" ||
          nk === "NUM OS" ||
          nk === "NUMERO OS" ||
          nk.includes("ORDEM DE SERVICO") ||
          nk.includes("ORDEM SERVICO")
        ) {
          return k;
        }
      }
    }

    for (const k of keys) {
      const nk = normalizeStr(k);
      for (const c of candidates) {
        if (logical === "os" && normalizeStr(c) === "OS") continue;
        if (nk.includes(normalizeStr(c))) return k;
      }
    }

    return null;
  }

  function mapRow(sheetRow) {
    const keys = Object.keys(sheetRow);
    const diagKey = findFieldKey(keys, "diagnostico");
    const clienteKey = findFieldKey(keys, "cliente");
    const loginKey = findFieldKey(keys, "login");
    const osKey = findFieldKey(keys, "os");
    const statusKey = findFieldKey(keys, "status");
    const assuntoKey = findFieldKey(keys, "assunto");
    const setorKey = findFieldKey(keys, "setor");

    const firstKey = keys[0] || null;
    const diagnostico = diagKey ? normalizeStr(sheetRow[diagKey]) : "";
    const cliente = clienteKey ? String(sheetRow[clienteKey]).trim() : "";
    const login = loginKey ? String(sheetRow[loginKey]).trim() : "";
    const osRaw = osKey ? String(sheetRow[osKey]).trim() : "";
    const firstColRaw = firstKey ? String(sheetRow[firstKey] || "").trim() : "";
    const osLooksWrong = !osRaw || DIAGS.includes(normalizeStr(osRaw));
    const os = osLooksWrong ? firstColRaw : osRaw;
    const status = statusKey ? String(sheetRow[statusKey]).trim() : "";
    const assunto = assuntoKey ? String(sheetRow[assuntoKey]).trim().replace(/^\s*\d+\s*-\s*/, "") : "";
    const setor = setorKey ? String(sheetRow[setorKey]).trim() : "";
    const region = detectRegionFrom(setor, assunto);

    return { diagnostico, cliente, login, os, status, assunto, setor, region };
  }

  function detectRegionFrom(setor, assunto) {
    const s = normalizeStr(setor || "");
    if (s.includes("SBC") || s.includes("REDES SBC")) return "SBC";
    if (s.includes("GRAJA") || s.includes("GRAJAU")) return "GRAJAÚ";
    if (s.includes("FRANCO")) return "FRANCO";

    const m = String(assunto || "").match(/^(\d+)\s*-/);
    if (m) {
      const n = Number(m[1]);
      if (n === 1) return "SBC";
      if (n === 21 || n === 23) return "SP";
    }

    return "OUTROS";
  }

  function computeFlags() {
    const byLoginAll = {};

    rawRows.forEach(r => {
      const lk = normalizeStr(r.login);
      if (!lk) return;
      byLoginAll[lk] ??= [];
      byLoginAll[lk].push(r);
    });

    displayRows = rawRows.filter(r => DIAGS.includes(normalizeStr(r.diagnostico)));

    const loginHasFlag = {};
    Object.keys(byLoginAll).forEach(lk => {
      const rows = byLoginAll[lk];
      const hasDiag = rows.some(rr => DIAGS.includes(normalizeStr(rr.diagnostico)));
      const hasOpenOther = rows.some(rr => {
        const st = normalizeStr(rr.status || "");
        return OPEN_STATUSES.some(os => st.includes(os));
      });
      const distinctOs = Array.from(new Set(rows.map(rr => String(rr.os || "")))).filter(Boolean);
      const hasAnother = distinctOs.length >= 2 || rows.length >= 2;
      loginHasFlag[lk] = Boolean(hasDiag && hasOpenOther && hasAnother);
    });

    displayRows = displayRows.map(r => {
      const lk = normalizeStr(r.login);
      return { ...r, displayOpen: Boolean(loginHasFlag[lk]) };
    });
  }

  function render() {
    resultsDiv.innerHTML = "";
    const filtered = displayRows.filter(r => r.region === currentRegion);

    if (!filtered.length) {
      resultsDiv.innerHTML = `<div class="empty-msg">Nenhum reagendamento em ${currentRegion}.</div>`;
      updateOsPanel();
      return;
    }

    DIAGS.forEach(cat => {
      const items = filtered.filter(i => normalizeStr(i.diagnostico) === normalizeStr(cat));
      if (!items.length) return;

      const block = document.createElement("div");
      block.className = "group-block";

      const header = document.createElement("div");
      header.className = "group-header";
      header.textContent = cat;
      block.appendChild(header);

      items.slice(0, MAX_RENDER_PER_REGION).forEach(it => {
        const el = document.createElement("div");
        el.className = "list-item";
        if (it.displayOpen) el.classList.add("os-open");

        const nameEl = document.createElement("div");
        nameEl.className = "client-name";
        if (it.displayOpen) nameEl.classList.add("os-open");
        const idPrefix = it.os ? `${it.os} - ` : "";
        nameEl.textContent = `${idPrefix}${it.cliente || it.login || ""}`;
        nameEl.title = `${idPrefix}${it.cliente || it.login || ""}`;
        nameEl.style.cursor = "pointer";
        nameEl.addEventListener("click", async () => {
          const payload = copyMode === "login" && it.login
            ? it.login
            : copyMode === "id" && it.os
              ? it.os
              : (it.cliente || it.login);
          await writeClipboard(payload);
          nameEl.animate([{ opacity: 0.6 }, { opacity: 1 }], { duration: 160 });
        });

        const loginEl = document.createElement("div");
        loginEl.className = "client-subject";
        loginEl.textContent = it.login ? `${it.login}` : "";
        loginEl.title = it.login || "";
        loginEl.style.cursor = it.login ? "pointer" : "default";
        if (it.login) {
          loginEl.addEventListener("click", async ev => {
            ev.stopPropagation();
            await writeClipboard(it.login);
            loginEl.animate([{ opacity: 0.6 }, { opacity: 1 }], { duration: 160 });
          });
        }

        const subjEl = document.createElement("div");
        subjEl.className = "client-subject";
        subjEl.textContent = it.assunto || "";

        el.appendChild(nameEl);
        el.appendChild(loginEl);
        el.appendChild(subjEl);
        block.appendChild(el);
      });

      resultsDiv.appendChild(block);
    });

    updateOsPanel();
  }

  function updateOsPanel() {
    const grouped = {};
    const unique = {};

    displayRows.forEach(it => {
      if (!it.displayOpen) return;
      const key = it.login || it.cliente;
      const label = it.login ? `${it.cliente} (${it.login}) - ${it.assunto}` : `${it.cliente} - ${it.assunto}`;
      grouped[it.region] ??= new Set();
      if (!unique[key]) {
        grouped[it.region].add(label);
        unique[key] = true;
      }
    });

    const regions = Object.keys(grouped);
    if (!regions.length) {
      osPanel.hidden = true;
      osPanelContent.innerHTML = "";
      return;
    }

    osPanel.hidden = false;
    osPanelContent.innerHTML = "";
    regions.forEach(r => {
      const block = document.createElement("div");
      block.className = "os-region-block";

      const title = document.createElement("div");
      title.className = "os-region-title";
      title.textContent = `${r} (${grouped[r].size})`;
      block.appendChild(title);

      grouped[r].forEach(lbl => {
        const row = document.createElement("div");
        row.className = "os-item";
        row.textContent = lbl;
        block.appendChild(row);
      });

      osPanelContent.appendChild(block);
    });
  }

  async function writeClipboard(text) {
    if (!text) return;
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text);
    }

    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.left = "-9999px";
    document.body.appendChild(ta);
    ta.select();
    try {
      document.execCommand("copy");
    } finally {
      ta.remove();
    }
  }

  function buildCopyAll() {
    if (!displayRows.length) return "";

    const groups = {};
    displayRows.forEach(r => {
      groups[r.region] ??= [];
      const payload = copyMode === "login" && r.login
        ? r.login
        : copyMode === "id" && r.os
          ? r.os
          : (r.cliente || r.login);
      groups[r.region].push({ diag: r.diagnostico, payload, assunto: r.assunto, os: r.os });
    });

    const regions = Object.keys(groups);
    let out = "";

    regions.forEach((reg, ri) => {
      out += `REAGENDAMENTO ${reg}\n\n`;
      DIAGS.forEach(cat => {
        const items = groups[reg].filter(x => normalizeStr(x.diag) === normalizeStr(cat));
        if (!items.length) return;
        out += `${cat}:\n`;
        items.forEach(it => {
          const withId = it.os
            ? String(it.payload) === String(it.os)
              ? `${it.payload}`
              : `${it.os} - ${it.payload}`
            : it.payload;
          out += `- ${withId} - ${it.assunto}\n`;
        });
        out += "\n";
      });
      if (ri !== regions.length - 1) out += "-------------------\n\n";
    });

    return out.trim();
  }

  async function copyAllHandler() {
    const text = buildCopyAll();
    if (!text) return alert("Nada para copiar.");
    try {
      await writeClipboard(text);
      alert("Copiado!");
    } catch {
      alert("Erro ao copiar.");
    }
  }

  async function copyOsHandler() {
    const groups = {};
    displayRows.forEach(it => {
      if (!it.displayOpen) return;
      groups[it.region] ??= [];
      groups[it.region].push(it);
    });

    const regs = Object.keys(groups);
    if (!regs.length) return alert("Nenhum cliente com O.S. em aberto.");

    let out = "";
    regs.forEach((r, idx) => {
      out += `REGIAO ${r}\n`;
      groups[r].forEach(it => {
        out += `- ${it.cliente}${it.login ? ` (${it.login})` : ""} - ${it.assunto}\n`;
      });
      if (idx !== regs.length - 1) out += "\n";
    });

    try {
      await writeClipboard(out);
      alert("Lista de clientes com O.S. copiada!");
    } catch {
      alert("Erro ao copiar lista de O.S.");
    }
  }

  async function copyFreeIdsHandler() {
    const ids = Array.from(
      new Set(
        displayRows
          .filter(it => it.region === currentRegion)
          .filter(it => !it.displayOpen)
          .map(it => String(it.os || "").trim())
          .filter(Boolean)
      )
    );

    if (!ids.length) return alert(`Nenhum ID sem outra O.S. em aberto/agendada em ${currentRegion}.`);

    try {
      await writeClipboard(ids.join("\n"));
      alert(`${ids.length} ID(s) copiado(s) de ${currentRegion}.`);
    } catch {
      alert("Erro ao copiar IDs.");
    }
  }

  function processWorkbook(file) {
    loaderEl.hidden = false;
    const reader = new FileReader();

    reader.onload = e => {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: "array" });
      const sheetName = wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      rawRows = json.map(raw => ({ ...raw, ...mapRow(raw) }));
      computeFlags();

      rowCountEl.hidden = false;
      rowCountEl.textContent = `${displayRows.length} registro(s) válidos encontrados`;
      actionsBar.hidden = !displayRows.length;
      loaderEl.hidden = true;
      render();
    };

    reader.readAsArrayBuffer(file);
  }

  dropzone.addEventListener("click", () => fileInput.click());
  dropzone.addEventListener("keydown", e => {
    if (e.key === "Enter" || e.key === " ") fileInput.click();
  });

  fileInput.addEventListener("change", e => {
    if (e.target.files[0]) processWorkbook(e.target.files[0]);
    fileInput.value = "";
  });

  dropzone.addEventListener("dragover", e => {
    e.preventDefault();
    dropzone.classList.add("drag-active");
  });

  dropzone.addEventListener("dragleave", () => dropzone.classList.remove("drag-active"));

  dropzone.addEventListener("drop", e => {
    e.preventDefault();
    dropzone.classList.remove("drag-active");
    if (e.dataTransfer.files[0]) processWorkbook(e.dataTransfer.files[0]);
  });

  regionBtns.forEach(btn => btn.addEventListener("click", () => {
    regionBtns.forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    currentRegion = btn.dataset.target;
    render();
  }));

  document.querySelectorAll('input[name="copyMode"]').forEach(r => {
    r.addEventListener("change", e => {
      const nextMode = String(e.target.value || "nome");
      copyMode = ["nome", "login", "id"].includes(nextMode) ? nextMode : "nome";
    });
  });

  btnCopyAll.addEventListener("click", copyAllHandler);
  if (btnCopyIdsFree) btnCopyIdsFree.addEventListener("click", copyFreeIdsHandler);
  if (btnCopyOs) btnCopyOs.addEventListener("click", copyOsHandler);

  if (btnReset) {
    btnReset.addEventListener("click", () => {
      rawRows = [];
      displayRows = [];
      resultsDiv.innerHTML = '<div class="empty-msg">Aguardando arquivo...</div>';
      actionsBar.hidden = true;
      rowCountEl.hidden = true;
      osPanel.hidden = true;
      osPanelContent.innerHTML = "";
    });
  }

  window.__debug = { rawRows, displayRows, computeFlags };
})();
