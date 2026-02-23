(function () {
  const inputConfig = {
    "ambient-inputs": [
      { key: "outsideTdb", name: "Outside conditions - TDB", value: 38.0, unit: "deg C", step: "0.001" },
      { key: "outsideRhPct", name: "Outside conditions - RH", value: 60.0, unit: "%", step: "0.001" },
      { key: "insideTdb", name: "Inside conditions - TDB", value: 25.0, unit: "deg C", step: "0.001" },
      { key: "insideRhPct", name: "Inside conditions - RH", value: 58.0, unit: "%", step: "0.001" },
      { key: "insideWdb", name: "Inside conditions - WDB", value: 16.0, unit: "deg C", step: "0.001" }
    ],
    "coach-heat-inputs": [
      { key: "floorTemp", name: "Floor Temp", value: 45.0, unit: "deg C", step: "0.001" },
      { key: "roofTemp", name: "Roof Temp", value: 45.0, unit: "deg C", step: "0.001", readOnly: true, derivedFrom: "floorTemp" },
      { key: "wallTemp", name: "Wall Temp", value: 45.0, unit: "deg C", step: "0.001", readOnly: true, derivedFrom: "floorTemp" },
      { key: "shW", name: "SH", value: 74.0, unit: "W", step: "0.001" },
      { key: "lhW", name: "LH", value: 60.0, unit: "W", step: "0.001" }
    ],
    "rmpu-load-inputs": [
      { key: "numberOfCoaches", name: "Number of Coaches", value: 1, unit: "Nos", step: "1" },
      { key: "hvacPerCoach", name: "Number of HVAC per Coaches", value: 2, unit: "Nos", step: "1" },
      { key: "totalPersons", name: "Total Number of Person", value: 560, unit: "Nos", step: "1" },
      { key: "freshAirCmmPerPerson", name: "Fresh Air Flow (CMM /Person)", value: 0.14, unit: "CMM/Person", step: "0.001" },
      { key: "extraHeatW", name: "Extra Heat Addition", value: 1000, unit: "W", step: "0.001" },
      { key: "byPassFactorPct", name: "By Pass Factor (BF%)", value: 10, unit: "%", step: "0.001" }
    ],
    "rmpu-dimension-inputs": [
      { key: "lengthM", name: "Length (L)", value: 21.337, unit: "m", step: "0.001" },
      { key: "heightM", name: "Height (H)", value: 2.1, unit: "m", step: "0.001" },
      { key: "widthM", name: "Width (W)", value: 3.245, unit: "m", step: "0.001" },
      { key: "windowAreaM2", name: "Window Area", value: 8.64, unit: "m2", step: "0.001" },
      { key: "solarRadWm2", name: "Solar Rad.", value: 1000, unit: "W/m2", step: "0.001" }
    ],
    "heat-transfer-inputs": [
      { key: "roofU", name: "Roof", value: 2.5, unit: "W/m2-K", step: "0.001" },
      { key: "sideWallU", name: "Side Wall", value: 2.5, unit: "W/m2-K", step: "0.001" },
      { key: "floorU", name: "Floor", value: 2.5, unit: "W/m2-K", step: "0.001" },
      { key: "glassU", name: "Glass Window", value: 2.25622, unit: "W/m2-K", step: "0.00001" }
    ]
  };

  const state = buildInitialState(inputConfig);

  function buildInitialState(config) {
    const next = {};
    Object.values(config).forEach(function (rows) {
      rows.forEach(function (row) {
        next[row.key] = row.value;
      });
    });
    return next;
  }

  function renderInputTable(tbodyId, rows) {
    const tbody = document.getElementById(tbodyId);
    if (!tbody) {
      return;
    }

    tbody.innerHTML = rows
      .map(function (row) {
        const readOnlyAttr = row.readOnly ? " readonly" : "";
        return [
          "<tr>",
          "<td>" + row.name + "</td>",
          "<td>",
          '<input type="number" step="' + row.step + '" data-key="' + row.key + '" value="' + row.value + '"' + readOnlyAttr + " />",
          "</td>",
          "<td>" + row.unit + "</td>",
          "</tr>"
        ].join("");
      })
      .join("");
  }

  function safeDivide(numerator, denominator, label) {
    if (!Number.isFinite(numerator) || !Number.isFinite(denominator) || Math.abs(denominator) < 1e-12) {
      throw new Error("Invalid denominator in formula: " + label);
    }
    return numerator / denominator;
  }

  function workbookFormula(data) {
    const roofTemp = data.floorTemp;
    const wallTemp = roofTemp;
    const floorTemp = data.floorTemp;

    const rhOutside = data.outsideRhPct / 100;
    const rhInside = data.insideRhPct / 100;

    const S10 = 0.62069;
    const S11 = 760;
    const trDivisor = 3517;

    const C6 = data.outsideTdb;
    const D6 = rhOutside;
    const C7 = data.insideTdb;
    const D7 = rhInside;

    const U7 = 0.61078 * 7.501 * Math.exp((17.2694 * C6) / (238.3 + C6));
    const U8 = 0.61078 * 7.501 * Math.exp((17.2694 * C7) / (238.3 + C7));

    const V7 = 100000 * S10 * (1 / (100 * S11 / D6 / U7 - 1));
    const V8 = 100000 * S10 * (1 / (100 * S11 / D7 / U8 - 1));

    const F6 = V7;
    const F7 = V8;

    const C8 = C6 - C7;
    const F8 = F6 - F7;

    const H11 = data.numberOfCoaches;
    const H12 = data.hvacPerCoach;
    const H13 = data.totalPersons;
    const H14 = data.freshAirCmmPerPerson;
    const H15 = data.extraHeatW;

    const D17 = safeDivide(H13, H11, "Heat Load!D17 = H13/H11");
    const D18 = H14 * 60;
    const D19 = 15;
    const D20 = D17 * D18;

    const D23 = data.lengthM;
    const D24 = data.heightM;
    const D25 = data.widthM;
    const D26 = data.windowAreaM2;
    const D27 = D23 * D24;
    const D28 = D23 * D25;
    const D29 = D24 * D25;
    const D30 = D27 - D26;

    const J18 = data.roofU;
    const J19 = data.sideWallU;
    const J20 = data.floorU;
    const J21 = data.glassU;

    const C34 = D26 * 2;
    const H34 = ((C34 * data.solarRadWm2) * 40) / 100;

    const C35 = D30 * 2;
    const F35 = wallTemp - C7;
    const H35 = C35 * J19 * F35;

    const C36 = D29 * 2;
    const H36 = C36 * J19 * F35;

    const C37 = D28;
    const F37 = roofTemp - C7;
    const H37 = ((C37 * J18 * F37) * 80) / 100;

    const H38 = H34 + H35 + H36 + H37;

    const C42 = D30 * 2;
    const F42 = F35;
    const H42 = C42 * J19 * F42;

    const C43 = D29 * 2;
    const H43 = C43 * 2.5 * F42;

    const C44 = D28;
    const H44 = C44 * 2.5 * F37;

    const C45 = C34 * 3;
    const H45 = C45 * J21 * F42;

    const C46 = D28;
    const F46 = floorTemp - C7;
    const H46 = C46 * J20 * F46;

    const H47 = H42 + H43 + H44 + H45 + H46;

    const E51 = H15;
    const G51 = E51;

    const C52 = D17;
    const E52 = data.shW;
    const F52 = data.lhW;

    const G52 = C52 * E52;
    const H52 = C52 * F52;

    const G53 = G51 + G52;
    const H53 = H52;

    const C57 = D20;
    const D57 = C8;
    const E57 = F8;
    const F57 = D19;

    const G57 = (((1210 * C57 * D57) / 3600) * (F57 / 100));
    const H57 = (((3010 * C57 * E57) / 3600) * (F57 / 100));

    const F62 = 100 - F57;
    const G62 = (((1210 * C57 * D57) / 3600) * (F62 / 100));
    const H62 = (((3010 * C57 * E57) / 3600) * (F62 / 100));

    const D67 = H38;
    const E67 = 0;
    const F67 = D67 + E67;

    const D68 = H47;
    const E68 = 0;
    const F68 = D68 + E68;

    const D69 = G53;
    const E69 = H53;
    const F69 = D69 + E69;

    const D70 = G57 + G62;
    const E70 = H57 + H62;
    const F70 = D70 + E70;

    const D72 = D67 + D68 + D69 + D70;
    const E72 = E67 + E68 + E69 + E70;
    const F72 = D72 + E72;

    const D74 = safeDivide(D72, H12, "Heat Load!D74 = D72/H12");
    const E74 = safeDivide(E72, H12, "Heat Load!E74 = E72/H12");
    const F74 = D74 + E74;

    const D76 = D74 * 1.05;
    const E76 = E74 * 1.05;
    const F76 = D76 + E76;

    const H67 = D67 / trDivisor;
    const I67 = E67 / trDivisor;
    const J67 = H67 + I67;

    const H68 = D68 / trDivisor;
    const I68 = E68 / trDivisor;
    const J68 = H68 + I68;

    const H69 = D69 / trDivisor;
    const I69 = E69 / trDivisor;
    const J69 = H69 + I69;

    const H70 = D70 / trDivisor;
    const I70 = E70 / trDivisor;
    const J70 = H70 + I70;

    const H72 = D72 / trDivisor;
    const I72 = E72 / trDivisor;
    const J72 = H72 + I72;

    const H74 = D74 / trDivisor;
    const I74 = E74 / trDivisor;
    const J74 = H74 + I74;

    const H76 = D76 / trDivisor;
    const I76 = E76 / trDivisor;
    const J76 = H76 + I76;

    const F35Summary = (data.freshAirCmmPerPerson * 60 * data.totalPersons) / 2;
    const J35Summary = F35Summary * 3;

    return {
      overview: [
        { category: "Solar heat", sensibleKw: D67 / 1000, latentKw: E67 / 1000, totalKw: F67 / 1000, sensibleTr: H67, latentTr: I67, totalTr: J67 },
        { category: "Transmission", sensibleKw: D68 / 1000, latentKw: E68 / 1000, totalKw: F68 / 1000, sensibleTr: H68, latentTr: I68, totalTr: J68 },
        { category: "Internal", sensibleKw: D69 / 1000, latentKw: E69 / 1000, totalKw: F69 / 1000, sensibleTr: H69, latentTr: I69, totalTr: J69 },
        { category: "Ventilation & Infiltration", sensibleKw: D70 / 1000, latentKw: E70 / 1000, totalKw: F70 / 1000, sensibleTr: H70, latentTr: I70, totalTr: J70 }
      ],
      summary: [
        {
          category: "Total Load per RMPU",
          sensibleKw: D72 / 1000,
          latentKw: E72 / 1000,
          totalKw: F72 / 1000,
          sensibleTr: H72,
          latentTr: I72,
          totalTr: J72,
          rowClass: "row-strong"
        },
        {
          category: "Total Load per HVAC",
          sensibleKw: D74 / 1000,
          latentKw: E74 / 1000,
          totalKw: F74 / 1000,
          sensibleTr: H74,
          latentTr: I74,
          totalTr: J74,
          rowClass: "row-highlight"
        },
        {
          category: "Total Load per HVAC (With 5% Safety)",
          sensibleKw: D76 / 1000,
          latentKw: E76 / 1000,
          totalKw: F76 / 1000,
          sensibleTr: H76,
          latentTr: I76,
          totalTr: J76,
          rowClass: "row-strong"
        }
      ],
      airflow: [
        { name: "Fresh Air Flow per HVAC", value: F35Summary, unit: "CMH" },
        { name: "Supply Air Flow per HVAC", value: J35Summary, unit: "CMH" }
      ]
    };
  }

  function syncDerivedInputs() {
    state.roofTemp = state.floorTemp;
    state.wallTemp = state.floorTemp;

    document.querySelectorAll('input[data-key="roofTemp"], input[data-key="wallTemp"]').forEach(function (input) {
      if (input) {
        input.value = String(state.floorTemp);
      }
    });
  }

  function formatNumber(value, digits) {
    if (!Number.isFinite(value)) {
      return "--";
    }
    return Number(value).toFixed(digits);
  }

  function renderOutputTable(tbodyId, rows) {
    const tbody = document.getElementById(tbodyId);
    if (!tbody) {
      return;
    }

    tbody.innerHTML = rows
      .map(function (row) {
        const className = row.rowClass ? ' class="' + row.rowClass + '"' : "";
        return [
          "<tr" + className + ">",
          "<td>" + row.category + "</td>",
          "<td>" + formatNumber(row.sensibleKw, 2) + "</td>",
          "<td>" + formatNumber(row.latentKw, 2) + "</td>",
          "<td>" + formatNumber(row.totalKw, 2) + "</td>",
          "<td>" + formatNumber(row.sensibleTr, 3) + "</td>",
          "<td>" + formatNumber(row.latentTr, 3) + "</td>",
          "<td>" + formatNumber(row.totalTr, 3) + "</td>",
          "</tr>"
        ].join("");
      })
      .join("");
  }

  function renderAirflowTable(rows) {
    const tbody = document.getElementById("airflow-outputs");
    if (!tbody) {
      return;
    }

    tbody.innerHTML = rows
      .map(function (row) {
        return [
          "<tr>",
          "<td>" + row.name + "</td>",
          "<td>" + formatNumber(row.value, 0) + "</td>",
          "<td>" + row.unit + "</td>",
          "</tr>"
        ].join("");
      })
      .join("");
  }

  function recalculateAndRender() {
    try {
      const result = workbookFormula(state);
      renderOutputTable("overview-outputs", result.overview);
      renderOutputTable("summary-outputs", result.summary);
      renderAirflowTable(result.airflow);
    } catch (error) {
      renderOutputTable("overview-outputs", []);
      renderOutputTable("summary-outputs", []);
      renderAirflowTable([]);
    }
  }

  function bindInputListeners() {
    document.addEventListener("input", function (event) {
      const target = event.target;
      if (!target || target.tagName !== "INPUT") {
        return;
      }

      const key = target.getAttribute("data-key");
      if (!key || !Object.prototype.hasOwnProperty.call(state, key)) {
        return;
      }

      if (!target.hasAttribute("readonly")) {
        target.classList.add("edited");
      }

      const next = Number(target.value);
      if (!Number.isFinite(next)) {
        return;
      }

      state[key] = next;
      syncDerivedInputs();
      recalculateAndRender();
    });
  }

  function init() {
    Object.entries(inputConfig).forEach(function (entry) {
      renderInputTable(entry[0], entry[1]);
    });
    syncDerivedInputs();
    bindInputListeners();
    recalculateAndRender();
  }

  init();
})();
