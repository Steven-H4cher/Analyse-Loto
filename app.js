// ============================================================
// Registre des Tirages — analyse locale de fichiers Excel 01–90
// Tout se passe dans le navigateur : rien n'est envoyé en ligne.
// ============================================================

// ---------- Helpers numériques ----------

function pad2(n) {
  return String(n).padStart(2, "0");
}

function normalizeHeader(h) {
  return String(h ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function extractNumbers(cell) {
  return String(cell ?? "")
    .split(/[^0-9]+/)
    .filter(Boolean)
    .map(Number)
    .filter((n) => n >= 1 && n <= 90);
}

function combinations2(arr) {
  const out = [];
  for (let i = 0; i < arr.length; i++) {
    for (let j = i + 1; j < arr.length; j++) {
      const a = Math.min(arr[i], arr[j]);
      const b = Math.max(arr[i], arr[j]);
      if (a !== b) out.push(`${pad2(a)}-${pad2(b)}`);
    }
  }
  return out;
}

// Lot 0: 01–09, Lot 1: 10–19, ... Lot 8: 80–90
const LOT_RANGES = [
  { id: 0, min: 1, max: 9, label: "01–09" },
  { id: 1, min: 10, max: 19, label: "10–19" },
  { id: 2, min: 20, max: 29, label: "20–29" },
  { id: 3, min: 30, max: 39, label: "30–39" },
  { id: 4, min: 40, max: 49, label: "40–49" },
  { id: 5, min: 50, max: 59, label: "50–59" },
  { id: 6, min: 60, max: 69, label: "60–69" },
  { id: 7, min: 70, max: 79, label: "70–79" },
  { id: 8, min: 80, max: 90, label: "80–90" },
];

function lotOf(n) {
  const found = LOT_RANGES.find((l) => n >= l.min && n <= l.max);
  return found ? found.id : null;
}

// Un lot compte comme "sorti" sur un tirage seulement s'il récolte au moins
// 2 de ses numéros parmi les 5 résultats — un seul numéro ne suffit pas.
function lotsHit(numbers) {
  const counts = new Map();
  for (const n of numbers) {
    const lot = lotOf(n);
    counts.set(lot, (counts.get(lot) || 0) + 1);
  }
  return Array.from(counts.entries())
    .filter(([, c]) => c >= 2)
    .map(([lot]) => lot);
}

// ---------- Lecture du fichier Excel ----------

function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (!rows.length) throw new Error("Feuille vide.");

  const header = rows[0].map(normalizeHeader);
  let colIdx = header.findIndex((h) => h.includes("resultat"));
  let dataRows = rows.slice(1);

  if (colIdx === -1) {
    // Pas de colonne "Résultats" trouvée -> on cherche la colonne qui contient
    // le plus de cellules au format "nn-nn-nn..."
    const width = Math.max(...rows.map((r) => r.length));
    let bestCol = -1;
    let bestScore = 0;
    for (let c = 0; c < width; c++) {
      let score = 0;
      for (const r of rows) {
        if (extractNumbers(r[c]).length >= 2) score++;
      }
      if (score > bestScore) {
        bestScore = score;
        bestCol = c;
      }
    }
    if (bestCol === -1 || bestScore === 0) {
      throw new Error(
        'Aucune colonne "Résultats" trouvée, et aucune colonne ne ressemble à des numéros de tirage.'
      );
    }
    colIdx = bestCol;
    dataRows = rows; // pas d'en-tête fiable, on garde tout
  }

  const draws = dataRows.map((r) => extractNumbers(r[colIdx])).filter((n) => n.length > 0);

  if (!draws.length) {
    throw new Error("Aucun numéro valide (1–90) trouvé dans le fichier.");
  }
  return draws;
}

// ---------- Calcul de l'analyse complète ----------

function computeAnalysis(draws) {
  const freq = new Map();
  for (let n = 1; n <= 90; n++) freq.set(n, 0);
  let totalNumbers = 0;
  for (const draw of draws) {
    for (const n of draw) {
      freq.set(n, freq.get(n) + 1);
      totalNumbers++;
    }
  }

  const freqEntries = Array.from(freq.entries());
  const maxFreq = Math.max(...freqEntries.map(([, c]) => c));
  const top10 = [...freqEntries].sort((a, b) => b[1] - a[1]).slice(0, 10);
  const missingNumbers = freqEntries.filter(([, c]) => c === 0).map(([n]) => n);

  const terminationSet = new Set();
  freqEntries.forEach(([n, c]) => {
    if (c > 0) terminationSet.add(n % 10);
  });
  const missingTerminations = Array.from({ length: 10 }, (_, i) => i).filter(
    (t) => !terminationSet.has(t)
  );

  const moyenne = totalNumbers / 90;

  // Paires
  const pairFreq = new Map();
  for (const draw of draws) {
    for (const p of combinations2(draw)) pairFreq.set(p, (pairFreq.get(p) || 0) + 1);
  }
  const totalPossiblePairs = (90 * 89) / 2; // 4005
  const pairsSeen = pairFreq.size;
  const pairsMissingCount = totalPossiblePairs - pairsSeen;

  const missingPairs = [];
  for (let a = 1; a <= 90; a++) {
    for (let b = a + 1; b <= 90; b++) {
      const key = `${pad2(a)}-${pad2(b)}`;
      if (!pairFreq.has(key)) missingPairs.push(key);
    }
  }
  const topPairs = Array.from(pairFreq.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 12);

  // Lots (dizaines) — un lot compte pour un tirage seulement s'il récolte
  // au moins 2 des 5 numéros tirés.
  const lotFreq = new Map(LOT_RANGES.map((l) => [l.id, 0]));
  let satureDraws = 0;
  for (const draw of draws) {
    const hits = lotsHit(draw);
    if (hits.length > 0) satureDraws++;
    for (const lot of hits) lotFreq.set(lot, lotFreq.get(lot) + 1);
  }
  const lotEntries = LOT_RANGES.map((l) => ({ ...l, count: lotFreq.get(l.id) }));
  const maxLotFreq = Math.max(...lotEntries.map((l) => l.count));
  const hottestLots = [...lotEntries].sort((a, b) => b.count - a.count).slice(0, 3);
  const coldestLots = [...lotEntries].sort((a, b) => a.count - b.count).slice(0, 3);

  return {
    totalDraws: draws.length,
    totalNumbers,
    moyenne,
    freqEntries,
    maxFreq,
    top10,
    missingNumbers,
    missingTerminations,
    totalPossiblePairs,
    pairsSeen,
    pairsMissingCount,
    missingPairs,
    topPairs,
    lotEntries,
    maxLotFreq,
    hottestLots,
    coldestLots,
    satureDraws,
  };
}

// ---------- Rendu ----------

function ballClass(count) {
  return count === 0 ? "ball empty" : "ball filled";
}

function ballStyle(count, maxFreq) {
  if (count === 0) return "";
  const ratio = maxFreq > 0 ? count / maxFreq : 0;
  const alpha = (0.18 + ratio * 0.72).toFixed(2);
  const textColor = ratio > 0.45 ? "#FBECEA" : "#1B2A3D";
  return `background: rgba(181, 52, 42, ${alpha}); color: ${textColor};`;
}

function renderFreq(a) {
  const grid = a.freqEntries
    .map(
      ([n, c]) =>
        `<div class="${ballClass(c)}" style="${ballStyle(c, a.maxFreq)}" title="${pad2(n)} — ${c} fois">${pad2(n)}</div>`
    )
    .join("");

  const top10 = a.top10
    .map(
      ([n, c], i) =>
        `<div class="chip"><span class="rank">${i + 1}.</span> <span class="num">${pad2(n)}</span> <span class="count">· ${c}x</span></div>`
    )
    .join("");

  const missingNums = a.missingNumbers.length
    ? a.missingNumbers.map(pad2).join("  ·  ")
    : "Tous les numéros sont sortis au moins une fois.";

  const missingTerms = a.missingTerminations.length
    ? a.missingTerminations.join("  ·  ")
    : "Toutes les terminaisons (0–9) sont représentées.";

  return `
    <div class="section-title">Grille 01–90</div>
    <div class="ball-grid">${grid}</div>
    <div class="hint">Intensité = fréquence d'apparition. Contour pointillé = jamais sorti.</div>

    <div class="section-title">Top 10</div>
    <div class="chip-row">${top10}</div>

    <div class="section-title">Numéros jamais sortis (${a.missingNumbers.length})</div>
    <div class="mono-list">${missingNums}</div>

    <div class="section-title">Terminaisons jamais sorties</div>
    <div class="mono-list">${missingTerms}</div>
  `;
}

function renderPairs(a) {
  const coverage = ((a.pairsSeen / a.totalPossiblePairs) * 100).toFixed(1);

  const summary = [
    ["Paires possibles", a.totalPossiblePairs],
    ["Paires vues", a.pairsSeen],
    ["Paires manquantes", a.pairsMissingCount],
    ["Couverture", `${coverage}%`],
  ]
    .map(
      ([label, val]) =>
        `<div class="summary-cell"><div class="summary-label">${label}</div><div class="summary-value">${val}</div></div>`
    )
    .join("");

  const topPairs = a.topPairs
    .map(
      ([p, c]) => `<div class="chip"><span class="num">${p}</span> <span class="count">· ${c}x</span></div>`
    )
    .join("");

  const missingPairs = a.missingPairs.map((p) => `<div>${p}</div>`).join("");

  return `
    <div class="section-title">Résumé</div>
    <div class="summary-grid">${summary}</div>

    <div class="section-title">Paires les plus fréquentes</div>
    <div class="chip-row">${topPairs}</div>

    <div class="section-title">Paires jamais sorties (${a.missingPairs.length.toLocaleString("fr-FR")})</div>
    <div class="missing-pairs-grid">${missingPairs}</div>
  `;
}

function lotBarColor(count, maxCount) {
  const ratio = maxCount > 0 ? count / maxCount : 0;
  const alpha = (0.2 + ratio * 0.7).toFixed(2);
  return `rgba(168, 123, 30, ${alpha})`;
}

function renderLots(a) {
  const rows = a.lotEntries
    .map((lot) => {
      const width = a.maxLotFreq ? (lot.count / a.maxLotFreq) * 100 : 0;
      return `
      <div class="lot-row">
        <div class="lot-label">Lot ${lot.id} <span>(${lot.label})</span></div>
        <div class="lot-bar-track">
          <div class="lot-bar-fill" style="width:${width}%; min-width:${lot.count > 0 ? 4 : 0}px; background:${lotBarColor(lot.count, a.maxLotFreq)};"></div>
        </div>
        <div class="lot-count">${lot.count}</div>
      </div>`;
    })
    .join("");

  const hot = a.hottestLots
    .map((l) => `<span class="lot-chip-hot">Lot ${l.id} (${l.label}) · ${l.count}</span>`)
    .join("");
  const cold = a.coldestLots
    .map((l) => `<span class="lot-chip-cold">Lot ${l.id} (${l.label}) · ${l.count}</span>`)
    .join("");

  return `
    <div class="section-title">Les 9 lots (dizaines)</div>
    <div class="hint-block">
      Grille des 90 numéros divisée en 9 blocs de 10 (Lot 0 : 01–09 … Lot 8 : 80–90). Un lot
      compte comme <strong>sorti</strong> sur un tirage seulement si au moins 2 de ses numéros
      figurent parmi les 5 résultats — un seul numéro dans le lot ne suffit pas.
    </div>
    <div class="hint-block">
      ${a.satureDraws} tirage${a.satureDraws > 1 ? "s" : ""} sur ${a.totalDraws}
      ${a.satureDraws > 1 ? "ont" : "a"} au moins un lot saturé (2+ numéros).
    </div>

    <div>${rows}</div>

    <div class="lot-groups">
      <div>
        <div class="lot-group-label">Lots chauds 🔥</div>
        <div class="chip-row">${hot}</div>
      </div>
      <div>
        <div class="lot-group-label">Lots froids ❄️</div>
        <div class="chip-row">${cold}</div>
      </div>
    </div>
  `;
}

function renderStatsBar(a) {
  const stats = [
    ["Tirages", a.totalDraws],
    ["Numéros analysés", a.totalNumbers],
    ["Moyenne / numéro", a.moyenne.toFixed(2)],
    ["Numéros absents", a.missingNumbers.length],
    ["Paires vues", `${a.pairsSeen} / ${a.totalPossiblePairs}`],
  ];
  return stats
    .map(
      ([label, val]) =>
        `<div class="stat-cell"><div class="stat-label">${label}</div><div class="stat-value">${val}</div></div>`
    )
    .join("");
}

// ---------- Câblage UI ----------

const fileInput = document.getElementById("file-input");
const errorBox = document.getElementById("error-box");
const errorText = document.getElementById("error-text");
const emptyState = document.getElementById("empty-state");
const results = document.getElementById("results");
const fileLabel = document.getElementById("file-label");
const statsBar = document.getElementById("stats-bar");
const tabFreq = document.getElementById("tab-freq");
const tabPairs = document.getElementById("tab-pairs");
const tabLots = document.getElementById("tab-lots");
const tabButtons = document.querySelectorAll(".tab-btn");
const tabContents = {
  freq: tabFreq,
  pairs: tabPairs,
  lots: tabLots,
};

function showError(message) {
  errorText.textContent = message;
  errorBox.hidden = false;
  results.hidden = true;
  emptyState.hidden = false;
}

function clearError() {
  errorBox.hidden = true;
}

function switchTab(name) {
  tabButtons.forEach((btn) => btn.classList.toggle("active", btn.dataset.tab === name));
  Object.entries(tabContents).forEach(([key, el]) => {
    el.hidden = key !== name;
  });
}

tabButtons.forEach((btn) => {
  btn.addEventListener("click", () => switchTab(btn.dataset.tab));
});

fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;
  clearError();

  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const draws = parseWorkbook(evt.target.result);
      const analysis = computeAnalysis(draws);

      fileLabel.textContent = `FICHIER : ${file.name}`;
      statsBar.innerHTML = renderStatsBar(analysis);
      tabFreq.innerHTML = renderFreq(analysis);
      tabPairs.innerHTML = renderPairs(analysis);
      tabLots.innerHTML = renderLots(analysis);

      emptyState.hidden = true;
      results.hidden = false;
      switchTab("freq");
    } catch (err) {
      showError(err.message || "Erreur de lecture du fichier.");
    }
  };
  reader.onerror = () => showError("Impossible de lire le fichier.");
  reader.readAsArrayBuffer(file);
});
