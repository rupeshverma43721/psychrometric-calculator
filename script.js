(function () {
  const form = document.getElementById("psychrometric-form");
  const messageEl = document.getElementById("form-message");
  const calculateBtn = document.getElementById("calculate-result");
  const exportBtn = document.getElementById("export-pdf");
  const resultSection = document.getElementById("result-section");
  const resultMetaEl = document.getElementById("result-meta");

  const inputTableBody = document.getElementById("input-table-body");
  const resultsTableBody = document.getElementById("results-table-body");
  const previewCompanyEl = document.getElementById("preview-company");
  const previewProjectEl = document.getElementById("preview-project");
  const previewRevisionEl = document.getElementById("preview-revision");
  const previewReportDateEl = document.getElementById("preview-report-date");
  const metricPageViewsEl = document.getElementById("metric-page-views");
  const metricCalculateClicksEl = document.getElementById("metric-calculate-clicks");
  const metricExportClicksEl = document.getElementById("metric-export-clicks");

  const counterApiBase = "https://api.countapi.xyz";
  const counterNamespace = `heatload-${(window.location.hostname || "local")
    .toLowerCase()
    .replace(/[^a-z0-9-]+/g, "-")}`;

  let lastRun = null;

  function todayAsInputDate() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const day = String(now.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function formatCounter(value) {
    if (!Number.isFinite(value)) {
      return "--";
    }
    return value.toLocaleString();
  }

  function renderCounter(element, value) {
    if (!element) {
      return;
    }
    element.textContent = formatCounter(value);
  }

  function counterUrl(mode, key) {
    return `${counterApiBase}/${mode}/${encodeURIComponent(counterNamespace)}/${encodeURIComponent(key)}`;
  }

  async function getCounterValue(key) {
    try {
      const response = await fetch(counterUrl("get", key), { cache: "no-store" });
      if (!response.ok) {
        return 0;
      }
      const payload = await response.json();
      return Number.isFinite(payload.value) ? payload.value : 0;
    } catch (error) {
      return null;
    }
  }

  async function hitCounterValue(key) {
    try {
      const response = await fetch(counterUrl("hit", key), { cache: "no-store" });
      if (!response.ok) {
        return null;
      }
      const payload = await response.json();
      return Number.isFinite(payload.value) ? payload.value : null;
    } catch (error) {
      return null;
    }
  }

  async function loadUsageCounters() {
    const [calculateClicks, exportClicks, pageViews] = await Promise.all([
      getCounterValue("calculate_clicks"),
      getCounterValue("export_clicks"),
      hitCounterValue("page_views")
    ]);

    renderCounter(metricCalculateClicksEl, calculateClicks);
    renderCounter(metricExportClicksEl, exportClicks);
    renderCounter(metricPageViewsEl, pageViews);
  }

  async function trackClickCounter(counterKey, metricElement) {
    const value = await hitCounterValue(counterKey);
    if (value !== null) {
      renderCounter(metricElement, value);
    }
  }

  function setMessage(text, tone) {
    messageEl.textContent = text;
    messageEl.classList.remove("ok", "warn", "error");

    if (tone === "ok") {
      messageEl.classList.add("ok");
    } else if (tone === "warn") {
      messageEl.classList.add("warn");
    } else if (tone === "error") {
      messageEl.classList.add("error");
    }
  }

  function clearMessage() {
    setMessage("", "");
  }

  function sanitize(text) {
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function formatNumber(value, digits) {
    return Number(value).toFixed(digits);
  }

  function formatDatePretty(dateStr) {
    const date = new Date(`${dateStr}T00:00:00`);
    if (Number.isNaN(date.getTime())) {
      return dateStr;
    }
    return date.toLocaleDateString();
  }

  function parseRequiredNumber(fieldName, label) {
    const raw = form.elements[fieldName].value.trim();
    if (!raw) {
      throw new Error(`Please enter ${label}.`);
    }

    const value = Number(raw);
    if (!Number.isFinite(value)) {
      throw new Error(`${label} must be a valid number.`);
    }

    return value;
  }

  function validateRange(label, value, min, max) {
    if (value < min || value > max) {
      throw new Error(`${label} must be between ${min} and ${max}.`);
    }
  }

  function collectInput() {
    const input = {
      companyName: form.elements.companyName.value.trim(),
      projectName: form.elements.projectName.value.trim(),
      revisionNumber: form.elements.revisionNumber.value.trim(),
      reportDate: form.elements.reportDate.value || todayAsInputDate(),

      b9: parseRequiredNumber("b9AmbientDbt", "Ambient dry bulb temperature (B9)"),
      f9: parseRequiredNumber("f9Passengers", "No of passengers (F9)"),
      b10: parseRequiredNumber("b10AmbientRh", "Ambient Relative Humidity (B10)"),
      f10: parseRequiredNumber("f10FreshAir", "Fresh air (F10)"),
      b11: parseRequiredNumber("b11Altitude", "Altitude (B11)"),
      f11: parseRequiredNumber("f11SolarRadiation", "Solar radiation (F11)"),
      b12: parseRequiredNumber("b12InternalSensible", "Internal Load Sensible (B12)"),
      f12: parseRequiredNumber("f12TravelSpeed", "Travel speed (F12)"),
      b13: parseRequiredNumber("b13InternalLatent", "Internal Load Latent (B13)"),
      f13: parseRequiredNumber("f13Co2Outside", "CO2 concentration outside (F13)"),
      b14: parseRequiredNumber("b14DuctLoss", "Pressure losses duct (B14)"),
      f14: parseRequiredNumber("f14Co2Emission", "CO2 emission per passenger (F14)"),

      b18: parseRequiredNumber("b18Length", "Length (B18)"),
      f18: parseRequiredNumber("f18KRoof", "Roof k-value (F18)"),
      b19: parseRequiredNumber("b19Width", "Width (B19)"),
      f19: parseRequiredNumber("f19KSideWall", "Side wall k-value (F19)"),
      b20: parseRequiredNumber("b20Height", "Height (B20)"),
      f20: parseRequiredNumber("f20KFloor", "Floor k-value (F20)"),
      b21: parseRequiredNumber("b21WindowArea", "Window area per side (B21)"),
      f21: parseRequiredNumber("f21KWindow", "Window k-value (F21)"),
      b22: parseRequiredNumber("b22DoorArea", "Door area per side (B22)"),
      f22: parseRequiredNumber("f22KFront", "Front k-value (F22)"),
      b23: parseRequiredNumber("b23AbsSideWall", "Absorption side wall (B23)"),
      f23: parseRequiredNumber("f23KDoor", "Door k-value (F23)"),
      b24: parseRequiredNumber("b24AbsRoof", "Absorption roof (B24)"),

      f27: parseRequiredNumber("f27HvacUnits", "No of HVAC units (F27)"),
      b28: form.elements.b28Refrigerant.value.trim(),
      b29: parseRequiredNumber("b29Frequency", "Operating frequency (B29)"),
      b30: parseRequiredNumber("b30CondenserAir", "Condenser air volume (B30)"),
      b32: parseRequiredNumber("b32SupplyAirVolume", "Supply air volume (B32)"),
      b33: parseRequiredNumber("b33SupplyAirTemp", "Supply air temperature (B33)"),
      b42: parseRequiredNumber("b42VentSensible", "Ventilation performance sensible (B42)"),
      f42: parseRequiredNumber("f42VentLatent", "Ventilation performance latent (F42)"),
      b46: parseRequiredNumber("b46FreshAirRate", "Fresh air rate (B46)"),
      b47: parseRequiredNumber("b47InternalRh", "Internal RH (B47)"),
      b48: parseRequiredNumber("b48SaloonTemp", "Saloon temperature (B48)")
    };

    if (!input.b28) {
      throw new Error("Please enter Refrigerant (B28).");
    }

    validateRange("Ambient RH (B10)", input.b10, 0, 100);
    validateRange("Internal RH (B47)", input.b47, 0, 100);
    validateRange("Passengers (F9)", input.f9, 0, 10000);
    validateRange("Fresh air (F10)", input.f10, 1, 200000);
    validateRange("Supply air volume (B32)", input.b32, 1, 200000);

    if (input.b32 < input.f10) {
      throw new Error("Supply air volume (B32) should be >= fresh air (F10) to match workbook assumptions.");
    }

    return input;
  }

  function pvsAirFormula(tempC) {
    return Math.exp(16.6536 - 4030.183 / (tempC + 235)) * 1000;
  }

  function pvsWorkbookFormula(tempC) {
    return Math.pow(10, 5) * Math.exp(Math.log(140974) - 3928.5 / (231.667 + tempC));
  }

  function clamp(value, min, max) {
    return Math.min(Math.max(value, min), max);
  }

  function calculateAirSheet(i) {
    const D4 = 97800;

    const D6 = i.b9;
    const E6 = D6;
    const F6 = i.b48;

    const D8 = i.b10;
    const E8 = D8;
    const F8 = i.b47;

    const D7 = 273.15 + D6;
    const E7 = D7;
    const F7 = 273.15 + F6;

    const D10 = pvsAirFormula(D6);
    const E10 = pvsAirFormula(E6);
    const F10 = pvsAirFormula(F6);

    const safeDen = function (value, label) {
      if (Math.abs(value) < 1e-9) {
        throw new Error(`${label} produced unstable denominator.`);
      }
      return value;
    };

    const D11 = (0.622 * D8 * D10) / safeDen(D4 * 100 - D8 * D10, "Air D11");
    const E11 = (0.622 * E8 * E10) / safeDen(D4 * 100 - E8 * E10, "Air E11");
    const F11 = (0.622 * F8 * F10) / safeDen(D4 * 100 - F8 * F10, "Air F11");

    const D13 = (462 * (0.622 + D11)) * (D7 / D4);
    const E13 = (462 * (0.622 + E11)) * (E7 / D4);
    const F13 = (462 * (0.622 + F11)) * (F7 / D4);

    const D14 = 1.007 * D6 + D11 * (2500.8 + 1.846 * D6);
    const E14 = 1.007 * E6 + E11 * (2500.8 + 1.846 * E6);
    const F14 = 1.007 * F6 + F11 * (2500.8 + 1.846 * F6);

    const N4 = E6;
    const N6 = i.f10;
    const N7 = F6;
    const N8 = F8;
    const N9 = i.b32 - i.f10;
    const N10 = D4;

    const P4 = pvsWorkbookFormula(N4);
    const P6 = P4 * 0.01 * N5(D8);
    const Q4 = 287.05 * (273 + N4) / safeDen(N10 - P6, "Air Q4");

    const P7 = pvsWorkbookFormula(N7);
    const P9 = P7 * 0.01 * N8;
    const Q7 = 287.05 * (273 + N7) / safeDen(N10 - P9, "Air Q7");

    function N5(value) {
      return value;
    }

    const P5 = 0.622 * 0.01 * D8 * P4 / safeDen(N10 - 0.01 * D8 * P4, "Air P5");
    const P8 = 0.622 * 0.01 * N8 * P7 / safeDen(N10 - 0.01 * N8 * P7, "Air P8");

    const P5g = P5 * 1000;
    const P8g = P8 * 1000;

    const Q5 = (1 + P5g / 1000) / Q4;
    const Q8 = (1 + P8g / 1000) / Q7;

    const N11 = (N6 * Q5) / safeDen(N6 * Q5 + N9 * Q8, "Air N11") * 100;
    const N13 = (N11 * N4 + (100 - N11) * N7) / 100;
    const N14 = pvsWorkbookFormula(N13);
    const N15 = (N11 * P5g + (100 - N11) * P8g) / 100;
    const N16 = (0.001 * N15) / safeDen(0.622 + 0.001 * N15, "Air N16") * N10;
    const N20 = N16 / safeDen(N14, "Air N20") * 100;

    return {
      D6,
      E6,
      F6,
      D13,
      E13,
      F13,
      E14,
      F14,
      N13,
      N20
    };
  }

  function calculateHeatLoadSheet(i, air) {
    const B6 = i.f9;
    const C6 = air.F6;

    const D6 = clamp(-0.0069 * Math.pow(C6, 3) + 0.3664 * Math.pow(C6, 2) - 9.6959 * C6 + 194.84, 0, 120);
    const E6 = clamp(-0.0085 * Math.pow(C6, 3) + 0.8026 * Math.pow(C6, 2) - 19.824 * C6 + 170.27, 0, 110);

    const F6 = D6 * B6;
    const G6 = E6 * B6;

    const B13 = i.f10;
    const C13 = 1 / Bounded(air.F13, "Heat Load C13");
    const D13 = air.E6;
    const E13 = air.F6;

    const F13 = ((B13 * C13) / 3600) * (D13 - E13) * 1005;
    const H13 = (air.E14 - air.F14) * B13 * C13 / 3.6;
    const G13 = H13 - F13;

    const K25 = i.b9 - i.b48 < -5 ? i.f12 : 10;
    const E20 = K25 < 11 ? 17 : 7.12 * Math.pow(K25 / 3.6, 0.78);

    const D34 = i.f11;
    const E34 = D34 * 0.2;

    const S = {
      6: 0,
      7: 0,
      8: i.b18 * i.b20 - i.b21 - i.b22,
      9: i.b22,
      10: i.b21,
      11: 0,
      12: i.b18 * i.b19,
      13: 0
    };

    const V = {
      6: i.b18 * i.b19,
      7: 0,
      8: i.b18 * i.b20 - i.b21 - i.b22,
      9: i.b22,
      10: 0,
      11: i.b21,
      12: 0,
      13: 0
    };

    const N = {
      6: i.b23,
      7: i.b24,
      8: i.b23,
      9: i.b23,
      10: 0,
      11: 0,
      12: i.b24,
      13: i.b23
    };

    const P = {
      6: i.f20,
      7: i.f18,
      8: i.f19,
      9: i.f23,
      10: i.f21,
      11: i.f21,
      12: i.f18,
      13: i.f18
    };

    const U = {};
    const X = {};

    for (let r = 6; r <= 13; r += 1) {
      const t = S[r] === 0 ? null : air.D6 + (N[r] * E34) / Bounded(E20, "Heat Load T");
      const w = V[r] === 0 ? null : air.D6 + (N[r] * E34) / Bounded(E20, "Heat Load W");
      U[r] = S[r] === 0 ? 0 : P[r] * S[r] * (t - air.F6);
      X[r] = V[r] === 0 ? 0 : P[r] * V[r] * (w - air.F6);
    }

    const sum = function (obj) {
      return Object.keys(obj).reduce((acc, key) => acc + (Number(obj[key]) || 0), 0);
    };

    const B20 = sum(S) + sum(V);
    const C20 =
      P[6] * (S[6] + V[6]) +
      P[7] * (S[7] + V[7]) +
      P[8] * (S[8] + V[8]) +
      P[9] * (S[9] + V[9]) +
      P[10] * (S[10] + V[10]) +
      P[11] * (S[11] + V[11]) +
      P[12] * (S[12] + V[12]) +
      P[13] * (S[13] + V[13]);

    const D20 = C20 / Bounded(B20, "Heat Load D20");
    const F20 = C20 * (D13 - E13);
    const G20 = 0;

    const F27 = sum(U) + sum(X);
    const G27 = 0;

    const lookupValues = {
      "01_DP1": (2 / 3) * i.f11,
      "02_DP2": 533.3333333333334,
      "03_DP3": 400,
      "04_DP4": 266.6666666666667,
      "05_DP4": 266.6666666666667,
      "06_DP4": 591,
      "07_DP5": 41.6666666666667,
      "08_DP6": 41.6666666666667,
      "09_DP7": 41.6666666666667
    };

    const selectedCondition = "01_DP1";
    const lookupVal = lookupValues[selectedCondition];

    const B34 = 0.45;
    const C34 = D34 * 0.8;
    const K34 = 0.75;
    const S10 = i.b21;
    const S11 = 0;
    const V11 = i.b21;

    const F34 = B34 * S11 * C34 + V11 * B34 * E34 + lookupVal * K34 * S10;

    return {
      F6,
      G6,
      F13,
      G13,
      B20,
      C20,
      D20,
      F20,
      G20,
      F27,
      G27,
      F34
    };
  }

  function Bounded(value, label) {
    if (!Number.isFinite(value) || Math.abs(value) < 1e-12) {
      throw new Error(`${label} became invalid during formula evaluation.`);
    }
    return value;
  }

  function calculateConsolidatedOutputs(i, heat) {
    const B37 = heat.F13 / 1000;
    const B38 = heat.F6 / 1000;
    const B39 = heat.F20 / 1000;
    const B40 = (heat.F27 + heat.F34) / 1000;
    const B41 = i.b12;
    const B42 = i.b42;
    const B43 = (B37 + B38 + B39 + B40 + B41 + B42) * 0.1;

    const F37 = heat.G13 / 1000;
    const F38 = heat.G6 / 1000;
    const F39 = 0;
    const F40 = 0;
    const F41 = i.b13;
    const F42 = i.f42;
    const F43 = (F37 + F38 + F39 + F40 + F41 + F42) * 0.2;

    const coolingMode = !(i.b9 - i.b48 < -5);

    const B27 = coolingMode
      ? B37 + B38 + B39 + B40 + B41 + B42 + F37 + F38 + F39 + F40 + F41 + F42 + B43 + F43
      : -(B37 + B39 + B40 + F37 + F41) + B43 + F43;

    const J27 = B27 / 3.52;
    const J28 = (400 / 0.6) * J27;
    const J29 = (700 / 0.6) * J27;

    const B31 = i.f10;
    const B45 = ((i.f13 * B31) + i.f14 * (i.f9 / 2) * (i.b32 - B31)) / Bounded(i.b32, "Consolidated B45");

    return {
      modeLabel: coolingMode ? "Cooling capacity" : "Heating capacity",
      B27,
      J27,
      J28,
      J29,
      B31,
      B37,
      B38,
      B39,
      B40,
      B41,
      B42,
      B43,
      F37,
      F38,
      F41,
      F42,
      F43,
      B45
    };
  }

  function runWorkbookCalculation(input) {
    const air = calculateAirSheet(input);
    const heat = calculateHeatLoadSheet(input, air);
    const consolidated = calculateConsolidatedOutputs(input, heat);

    return { air, heat, consolidated };
  }

  function buildPreviewData(input, result) {
    const reportHeader = {
      companyName: input.companyName || "Not provided",
      projectName: input.projectName || "Not provided",
      revisionNumber: input.revisionNumber || "Not provided",
      reportDate: formatDatePretty(input.reportDate)
    };

    const inputRows = [
      ["Ambient dry bulb temperature (B9)", `${formatNumber(input.b9, 2)} deg C`],
      ["Ambient Relative Humidity (B10)", `${formatNumber(input.b10, 2)} %`],
      ["Altitude (B11)", `${formatNumber(input.b11, 2)} m`],
      ["Internal Load Sensible (B12)", `${formatNumber(input.b12, 3)} kW`],
      ["Internal Load Latent (B13)", `${formatNumber(input.b13, 3)} kW`],
      ["Pressure losses duct (B14)", `${formatNumber(input.b14, 2)} Pa`],
      ["No of passengers (F9)", `${formatNumber(input.f9, 0)} Nos`],
      ["Fresh air (F10)", `${formatNumber(input.f10, 2)} CMH`],
      ["Solar radiation (F11)", `${formatNumber(input.f11, 2)} W/m2`],
      ["Travel speed (F12)", `${formatNumber(input.f12, 2)} km/h`],
      ["CO2 concentration outside (F13)", `${formatNumber(input.f13, 2)} ppm`],
      ["CO2 emission per passenger (F14)", `${formatNumber(input.f14, 2)} ppm`],
      ["Length x Width x Height (B18:B20)", `${formatNumber(input.b18, 2)} x ${formatNumber(input.b19, 2)} x ${formatNumber(input.b20, 2)} m`],
      ["Window / Door area per side (B21:B22)", `${formatNumber(input.b21, 2)} / ${formatNumber(input.b22, 2)} m2`],
      ["Absorption side wall / roof (B23:B24)", `${formatNumber(input.b23, 2)} / ${formatNumber(input.b24, 2)}`],
      ["k-values roof/side/floor/window/front/door (F18:F23)", `${formatNumber(input.f18, 2)} / ${formatNumber(input.f19, 2)} / ${formatNumber(input.f20, 2)} / ${formatNumber(input.f21, 2)} / ${formatNumber(input.f22, 2)} / ${formatNumber(input.f23, 2)}`],
      ["No of HVAC units (F27)", `${formatNumber(input.f27, 0)}`],
      ["Refrigerant (B28)", input.b28],
      ["Operating frequency (B29)", `${formatNumber(input.b29, 2)} Hz`],
      ["Condenser air volume (B30)", `${formatNumber(input.b30, 2)} m3/h`],
      ["Supply air volume (B32)", `${formatNumber(input.b32, 2)} m3/h`],
      ["Supply air temperature (B33)", `${formatNumber(input.b33, 2)} deg C`],
      ["Ventilation performance sensible / latent (B42/F42)", `${formatNumber(input.b42, 3)} / ${formatNumber(input.f42, 3)} kW`],
      ["Fresh air rate / Internal RH / Saloon temp (B46:B48)", `${formatNumber(input.b46, 2)} m3/h/person, ${formatNumber(input.b47, 2)} %, ${formatNumber(input.b48, 2)} deg C`],
      ["Formula engine", "Consolidated + Air + Heat Load cal"]
    ];

    const c = result.consolidated;
    const h = result.heat;
    const a = result.air;

    const resultRows = [
      { parameter: "Capacity Mode (A27)", symbol: "A27", value: c.modeLabel, unit: "text" },
      { parameter: "Load per car (B27)", symbol: "B27", value: formatNumber(c.B27, 3), unit: "kW" },
      { parameter: "TR per car (J27)", symbol: "J27", value: formatNumber(c.J27, 3), unit: "TR" },
      { parameter: "SAF (J28)", symbol: "J28", value: formatNumber(c.J28, 3), unit: "m3/h" },
      { parameter: "CAF (J29)", symbol: "J29", value: formatNumber(c.J29, 3), unit: "m3/h" },
      { parameter: "Fresh air per HVAC (B31)", symbol: "B31", value: formatNumber(c.B31, 3), unit: "m3/h" },
      { parameter: "CO2 concentration inside (B45)", symbol: "B45", value: formatNumber(c.B45, 3), unit: "ppm" },

      { parameter: "Fresh air sensible load (B37)", symbol: "B37", value: formatNumber(c.B37, 3), unit: "kW" },
      { parameter: "Passenger sensible load (B38)", symbol: "B38", value: formatNumber(c.B38, 3), unit: "kW" },
      { parameter: "Transmission windows sensible (B39)", symbol: "B39", value: formatNumber(c.B39, 3), unit: "kW" },
      { parameter: "Radiation S+W sensible (B40)", symbol: "B40", value: formatNumber(c.B40, 3), unit: "kW" },
      { parameter: "Other sensible load (B41)", symbol: "B41", value: formatNumber(c.B41, 3), unit: "kW" },
      { parameter: "Ventilation performance sensible (B42)", symbol: "B42", value: formatNumber(c.B42, 3), unit: "kW" },
      { parameter: "Infiltration sensible (B43)", symbol: "B43", value: formatNumber(c.B43, 3), unit: "kW" },

      { parameter: "Fresh air latent load (F37)", symbol: "F37", value: formatNumber(c.F37, 3), unit: "kW" },
      { parameter: "Passenger latent load (F38)", symbol: "F38", value: formatNumber(c.F38, 3), unit: "kW" },
      { parameter: "Other latent load (F41)", symbol: "F41", value: formatNumber(c.F41, 3), unit: "kW" },
      { parameter: "Ventilation performance latent (F42)", symbol: "F42", value: formatNumber(c.F42, 3), unit: "kW" },
      { parameter: "Infiltration latent (F43)", symbol: "F43", value: formatNumber(c.F43, 3), unit: "kW" },

      { parameter: "Heat Load cal F20", symbol: "F20", value: formatNumber(h.F20, 3), unit: "W" },
      { parameter: "Heat Load cal F27", symbol: "F27", value: formatNumber(h.F27, 3), unit: "W" },
      { parameter: "Heat Load cal F34", symbol: "F34", value: formatNumber(h.F34, 3), unit: "W" },
      { parameter: "Air mixed temperature (N13)", symbol: "N13", value: formatNumber(a.N13, 3), unit: "deg C" },
      { parameter: "Air mixed RH (N20)", symbol: "N20", value: formatNumber(a.N20, 3), unit: "%" }
    ];

    return { reportHeader, inputRows, resultRows };
  }

  function renderKeyValueTable(tableBody, rows) {
    tableBody.innerHTML = rows
      .map(
        (row, index) =>
          `<tr>
            <td class="sno-cell">${index + 1}</td>
            <td>${sanitize(row[0])}</td>
            <td>${sanitize(row[1])}</td>
          </tr>`
      )
      .join("");
  }

  function renderReportHeader(header) {
    previewCompanyEl.textContent = header.companyName;
    previewProjectEl.textContent = header.projectName;
    previewRevisionEl.textContent = header.revisionNumber;
    previewReportDateEl.textContent = header.reportDate;
  }

  function renderResultsTable(rows) {
    resultsTableBody.innerHTML = rows
      .map(
        (row, index) =>
          `<tr>
            <td class="sno-cell">${index + 1}</td>
            <td>${sanitize(row.parameter)}</td>
            <td>${row.symbolHtml || sanitize(row.symbol)}</td>
            <td>${sanitize(row.value)}</td>
            <td>${row.unitHtml || sanitize(row.unit)}</td>
          </tr>`
      )
      .join("");
  }

  function resetView() {
    clearMessage();
    resultSection.classList.add("hidden");
    resultMetaEl.textContent = "--";
    inputTableBody.innerHTML = "";
    resultsTableBody.innerHTML = "";
    previewCompanyEl.textContent = "--";
    previewProjectEl.textContent = "--";
    previewRevisionEl.textContent = "--";
    previewReportDateEl.textContent = "--";
    exportBtn.disabled = true;
    lastRun = null;
  }

  function safePdfText(text) {
    return String(text)
      .normalize("NFKD")
      .replace(/[^\x20-\x7E]/g, "?")
      .replace(/\\/g, "\\\\")
      .replace(/\(/g, "\\(")
      .replace(/\)/g, "\\)");
  }

  function trimToCell(text, width) {
    const maxChars = Math.max(4, Math.floor(width / 5.2));
    if (text.length <= maxChars) {
      return text;
    }
    if (maxChars < 4) {
      return text.slice(0, maxChars);
    }
    return `${text.slice(0, maxChars - 3)}...`;
  }

  function drawText(commands, x, y, text, size, gray) {
    const fontSize = size || 10;
    const grayValue = gray === undefined ? 0 : gray;
    commands.push("BT");
    commands.push(`/F1 ${fontSize} Tf`);
    commands.push(`${grayValue} g`);
    commands.push(`1 0 0 1 ${x} ${y} Tm`);
    commands.push(`(${safePdfText(text)}) Tj`);
    commands.push("ET");
  }

  function drawRichPdfSegments(commands, x, y, segments, baseFontSize, gray, maxWidth) {
    const items = Array.isArray(segments) ? segments : [{ text: String(segments || ""), style: "normal" }];
    let cursorX = x;
    const rightLimit = x + maxWidth;
    const textGray = gray === undefined ? 0.08 : gray;

    const metricsForStyle = function (style) {
      if (style === "sup" || style === "sub") {
        return { widthFactor: 0.6, extraGap: 0.65 };
      }
      return { widthFactor: 0.53, extraGap: 0.35 };
    };

    const advanceFor = function (txt, size, style) {
      const metric = metricsForStyle(style);
      return txt.length * size * metric.widthFactor + metric.extraGap;
    };

    for (const item of items) {
      const style = item && item.style ? item.style : "normal";
      const fontSize = style === "normal" ? baseFontSize : Math.max(6, baseFontSize * 0.74);
      const yOffset = style === "sup" ? baseFontSize * 0.2 : style === "sub" ? -baseFontSize * 0.12 : 0;
      const raw = item && item.text ? String(item.text) : "";

      if (!raw) {
        continue;
      }

      const remaining = rightLimit - cursorX;
      if (remaining <= 2) {
        break;
      }

      const maxChars = Math.max(1, Math.floor(remaining / advanceFor("W", fontSize, style)));
      const text = raw.slice(0, maxChars);
      if (!text) {
        continue;
      }

      drawText(commands, cursorX, y + yOffset, text, fontSize, textGray);
      cursorX += advanceFor(text, fontSize, style);

      if (text.length < raw.length) {
        break;
      }
    }
  }

  function drawRoundedRect(commands, x, y, width, height, radius, mode) {
    const safeRadius = Math.max(0, Math.min(radius, width / 2, height / 2));
    const kappa = 0.5522847498;
    const c = safeRadius * kappa;
    const right = x + width;
    const top = y + height;

    commands.push(`${x + safeRadius} ${y} m`);
    commands.push(`${right - safeRadius} ${y} l`);
    commands.push(`${right - safeRadius + c} ${y} ${right} ${y + safeRadius - c} ${right} ${y + safeRadius} c`);
    commands.push(`${right} ${top - safeRadius} l`);
    commands.push(
      `${right} ${top - safeRadius + c} ${right - safeRadius + c} ${top} ${right - safeRadius} ${top} c`
    );
    commands.push(`${x + safeRadius} ${top} l`);
    commands.push(`${x + safeRadius - c} ${top} ${x} ${top - safeRadius + c} ${x} ${top - safeRadius} c`);
    commands.push(`${x} ${y + safeRadius} l`);
    commands.push(`${x} ${y + safeRadius - c} ${x + safeRadius - c} ${y} ${x + safeRadius} ${y} c`);
    commands.push(mode || "S");
  }

  function drawTable(commands, config) {
    const x = config.x;
    const top = config.top;
    const colWidths = config.colWidths;
    const headers = config.headers;
    const rows = config.rows;
    const rowHeight = config.rowHeight || 18;
    const fontSize = config.fontSize || 9.5;

    const width = colWidths.reduce((sum, w) => sum + w, 0);
    const totalRows = rows.length + 1;
    const height = totalRows * rowHeight;
    const bottom = top - height;

    commands.push("0.93 0.95 0.99 rg");
    commands.push(`${x} ${top - rowHeight} ${width} ${rowHeight} re f`);
    commands.push("0 g");
    drawRoundedRect(commands, x, bottom, width, height, 7, "S");

    for (let i = 1; i < totalRows; i += 1) {
      const y = top - i * rowHeight;
      commands.push(`${x} ${y} m ${x + width} ${y} l S`);
    }

    let cx = x;
    for (let i = 0; i < colWidths.length - 1; i += 1) {
      cx += colWidths[i];
      commands.push(`${cx} ${bottom} m ${cx} ${top} l S`);
    }

    let textX = x;
    headers.forEach((header, index) => {
      drawText(commands, textX + 4, top - rowHeight + 5.5, trimToCell(header, colWidths[index]), fontSize, 0);
      textX += colWidths[index];
    });

    rows.forEach((row, rowIndex) => {
      let colX = x;
      row.forEach((cell, colIndex) => {
        const y = top - rowHeight * (rowIndex + 2) + 5.5;
        if (cell && typeof cell === "object" && Array.isArray(cell.segments)) {
          drawRichPdfSegments(commands, colX + 4, y, cell.segments, fontSize, 0.08, colWidths[colIndex] - 8);
        } else {
          drawText(commands, colX + 4, y, trimToCell(String(cell), colWidths[colIndex]), fontSize, 0.08);
        }
        colX += colWidths[colIndex];
      });
    });

    return bottom;
  }

  function buildPdfBlob(report) {
    const commands = [];
    const pageWidth = 595;
    const tableWidth = 520;
    const tableX = (pageWidth - tableWidth) / 2;

    commands.push("0.95 0.97 1 rg");
    drawRoundedRect(commands, 42, 760, 511, 52, 8, "f");
    commands.push("0 g");
    drawRoundedRect(commands, 42, 760, 511, 52, 8, "S");

    drawText(commands, 48, 798, `Company: ${report.reportHeader.companyName}`, 9.5, 0);
    drawText(commands, 300, 798, `Revision: ${report.reportHeader.revisionNumber}`, 9.5, 0);
    drawText(commands, 48, 780, `Project: ${report.reportHeader.projectName}`, 9.5, 0);
    drawText(commands, 300, 780, `Report Date: ${report.reportHeader.reportDate}`, 9.5, 0);

    let cursorY = 740;

    drawText(commands, tableX + 2, cursorY - 2, "INPUT", 10.5, 0);
    cursorY -= 12;

    cursorY = drawTable(commands, {
      x: tableX,
      top: cursorY,
      headers: ["S.No", "Field", "Value"],
      colWidths: [56, 174, 290],
      rows: report.inputRows.map((row, index) => [index + 1, row[0], row[1]]),
      rowHeight: 19,
      fontSize: 9.5
    });

    cursorY -= 18;
    drawText(commands, tableX + 2, cursorY - 2, "OUTPUT", 10.5, 0);
    cursorY -= 12;

    const resultRowsForPdf = report.resultRows.map((row, index) => [
      index + 1,
      row.parameter,
      { segments: row.symbolPdf || [{ text: row.symbol, style: "normal" }] },
      row.value,
      { segments: row.unitPdf || [{ text: row.unit, style: "normal" }] }
    ]);

    cursorY = drawTable(commands, {
      x: tableX,
      top: cursorY,
      headers: ["S.No", "Parameter", "Symbol", "Value", "Unit"],
      colWidths: [50, 170, 70, 130, 100],
      rows: resultRowsForPdf,
      rowHeight: 19,
      fontSize: 9
    });

    if (cursorY < 40) {
      drawText(commands, 42, 28, "Note: Report content exceeded one page view area.", 8.5, 0.35);
    }

    const contentStream = commands.join("\n");

    const objects = {
      1: "<< /Type /Catalog /Pages 2 0 R >>",
      2: "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      3:
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
      4: "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
      5: `<< /Length ${contentStream.length} >>\nstream\n${contentStream}\nendstream`
    };

    const encoder = new TextEncoder();
    const byteLength = (text) => encoder.encode(text).length;

    let pdf = "%PDF-1.4\n";
    const offsets = [0];

    for (let i = 1; i <= 5; i += 1) {
      offsets[i] = byteLength(pdf);
      pdf += `${i} 0 obj\n${objects[i]}\nendobj\n`;
    }

    const xrefPos = byteLength(pdf);
    pdf += "xref\n0 6\n0000000000 65535 f \n";

    for (let i = 1; i <= 5; i += 1) {
      pdf += `${String(offsets[i]).padStart(10, "0")} 00000 n \n`;
    }

    pdf += `trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xrefPos}\n%%EOF`;

    return new Blob([pdf], { type: "application/pdf" });
  }

  function downloadPdf(reportData) {
    const blob = buildPdfBlob(reportData);
    const url = URL.createObjectURL(blob);

    const cleanFilePart = function (value, fallback) {
      const base = (value || "").trim() || fallback;
      return base
        .replace(/[\\/:*?"<>|]+/g, "-")
        .replace(/\s+/g, " ")
        .trim();
    };

    const projectName = cleanFilePart(reportData.reportHeader.projectName, "Project");
    const revision = cleanFilePart(reportData.reportHeader.revisionNumber, "Rev");
    const fileName = `${projectName}_${revision}.pdf`;

    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();

    setTimeout(() => URL.revokeObjectURL(url), 500);
  }

  function runCalculation() {
    const input = collectInput();
    const result = runWorkbookCalculation(input);
    const previewData = buildPreviewData(input, result);

    renderReportHeader(previewData.reportHeader);
    renderKeyValueTable(inputTableBody, previewData.inputRows);
    renderResultsTable(previewData.resultRows);

    const now = new Date();
    resultSection.classList.remove("hidden");
    resultMetaEl.textContent = `Calculated at ${now.toLocaleString()} with Consolidated + Air + Heat Load cal formulas`;
    exportBtn.disabled = false;

    const reportData = {
      generatedAt: now.toLocaleString(),
      reportHeader: previewData.reportHeader,
      inputRows: previewData.inputRows,
      resultRows: previewData.resultRows
    };

    lastRun = {
      input,
      result,
      reportData
    };

    setMessage("Calculation complete. Workbook-linked outputs are ready.", "ok");
  }

  form.addEventListener("submit", function (event) {
    event.preventDefault();

    try {
      runCalculation();
      void trackClickCounter("calculate_clicks", metricCalculateClicksEl);
    } catch (error) {
      exportBtn.disabled = true;
      setMessage(error.message, "error");
    }
  });

  exportBtn.addEventListener("click", function () {
    try {
      if (!lastRun) {
        runCalculation();
      }

      downloadPdf(lastRun.reportData);
      setMessage("PDF exported successfully.", "ok");
      void trackClickCounter("export_clicks", metricExportClicksEl);
    } catch (error) {
      setMessage(error.message, "error");
    }
  });

  form.addEventListener("reset", function () {
    setTimeout(function () {
      form.elements.reportDate.value = todayAsInputDate();
      resetView();
    }, 0);
  });

  form.elements.reportDate.value = todayAsInputDate();
  resetView();

  if (calculateBtn && exportBtn) {
    void loadUsageCounters();
  }
})();
