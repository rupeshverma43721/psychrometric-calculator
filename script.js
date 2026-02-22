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
  const counterNamespace = `psychrometric-${(window.location.hostname || "local")
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

  function parseOptionalNumber(fieldName) {
    const raw = form.elements[fieldName].value.trim();
    if (!raw) {
      return null;
    }

    const value = Number(raw);
    if (!Number.isFinite(value)) {
      throw new Error(`Invalid number in ${fieldName}.`);
    }

    return value;
  }

  function parseRequiredNumber(fieldName) {
    const value = parseOptionalNumber(fieldName);
    if (value === null) {
      throw new Error(`Please enter ${fieldName}.`);
    }
    return value;
  }

  function clamp(value, min, max) {
    return Math.min(Math.max(value, min), max);
  }

  function saturationPressureKPa(tempC) {
    return 0.61078 * Math.exp((17.2694 * tempC) / (tempC + 237.3));
  }

  function pressureFromAltitudeKPa(altitudeM) {
    return 101.325 * Math.pow(1 - 2.25577e-5 * altitudeM, 5.2559);
  }

  function dewpointFromPw(pwKPa) {
    const ratio = Math.max(pwKPa, 0.000001) / 0.61078;
    const factor = Math.log(ratio);
    return (237.3 * factor) / (17.2694 - factor);
  }

  function wetBulbApproxFromDryRh(tempC, rhPercent) {
    const rh = clamp(rhPercent, 1, 100);
    return (
      tempC * Math.atan(0.151977 * Math.sqrt(rh + 8.313659)) +
      Math.atan(tempC + rh) -
      Math.atan(rh - 1.676331) +
      0.00391838 * Math.pow(rh, 1.5) * Math.atan(0.023101 * rh) -
      4.686035
    );
  }

  function validateInputRange(name, value, min, max) {
    if (value < min || value > max) {
      throw new Error(`${name} must be between ${min} and ${max}.`);
    }
  }

  function collectInput() {
    const dryBulbTemp = parseRequiredNumber("dryBulbTemp");
    const wetBulbTemp = parseOptionalNumber("wetBulbTemp");
    const relativeHumidity = parseOptionalNumber("relativeHumidity");
    const dewPointTemp = parseOptionalNumber("dewPointTemp");
    const altitudeRaw = parseOptionalNumber("altitude");

    const input = {
      companyName: form.elements.companyName.value.trim(),
      projectName: form.elements.projectName.value.trim(),
      revisionNumber: form.elements.revisionNumber.value.trim(),
      reportDate: form.elements.reportDate.value || todayAsInputDate(),
      dryBulbTemp,
      wetBulbTemp,
      relativeHumidity,
      dewPointTemp,
      altitude: altitudeRaw === null ? 0 : altitudeRaw
    };

    validateInputRange("Dry Bulb Temp", input.dryBulbTemp, -50, 80);
    validateInputRange("Altitude", input.altitude, -500, 10000);

    const providedMoistureCount = [
      input.wetBulbTemp !== null,
      input.relativeHumidity !== null,
      input.dewPointTemp !== null
    ].filter(Boolean).length;

    if (providedMoistureCount === 0) {
      throw new Error("Enter at least one moisture input: Wet Bulb, Relative Humidity, or Dewpoint.");
    }

    if (input.wetBulbTemp !== null) {
      validateInputRange("Wet Bulb Temp", input.wetBulbTemp, -50, 80);
      if (input.wetBulbTemp > input.dryBulbTemp) {
        throw new Error("Wet Bulb Temp must be less than or equal to Dry Bulb Temp.");
      }
    }

    if (input.relativeHumidity !== null) {
      validateInputRange("Relative Humidity", input.relativeHumidity, 0.1, 100);
    }

    if (input.dewPointTemp !== null) {
      validateInputRange("Dewpoint Temp", input.dewPointTemp, -80, 60);
      if (input.dewPointTemp > input.dryBulbTemp) {
        throw new Error("Dewpoint Temp should be less than or equal to Dry Bulb Temp.");
      }
    }

    return input;
  }

  function resolveVapourPressure(input, pressureKPa) {
    if (input.wetBulbTemp !== null) {
      const saturationAtWetBulb = saturationPressureKPa(input.wetBulbTemp);
      const gamma = 0.00066 * (1 + 0.00115 * input.wetBulbTemp) * pressureKPa;
      const pw = saturationAtWetBulb - gamma * (input.dryBulbTemp - input.wetBulbTemp);

      if (pw <= 0 || pw >= pressureKPa) {
        throw new Error("Wet bulb input produced invalid vapour pressure.");
      }

      return {
        source: "Wet Bulb Temp",
        pwKPa: pw,
        warning:
          input.relativeHumidity !== null || input.dewPointTemp !== null
            ? "Multiple moisture fields were filled. Wet Bulb Temp was used for calculation priority."
            : ""
      };
    }

    if (input.relativeHumidity !== null) {
      const pw = (input.relativeHumidity / 100) * saturationPressureKPa(input.dryBulbTemp);

      if (pw <= 0 || pw >= pressureKPa) {
        throw new Error("Relative humidity input produced invalid vapour pressure.");
      }

      return {
        source: "Relative Humidity",
        pwKPa: pw,
        warning:
          input.dewPointTemp !== null
            ? "Multiple moisture fields were filled. Relative Humidity was used for calculation priority."
            : ""
      };
    }

    const pw = saturationPressureKPa(input.dewPointTemp);

    if (pw <= 0 || pw >= pressureKPa) {
      throw new Error("Dewpoint input produced invalid vapour pressure.");
    }

    return {
      source: "Dewpoint Temp",
      pwKPa: pw,
      warning: ""
    };
  }

  function calculateProperties(input) {
    const pressureKPa = pressureFromAltitudeKPa(input.altitude);
    if (pressureKPa <= 0) {
      throw new Error("Altitude generated non-physical atmospheric pressure.");
    }

    const saturationAtDryBulb = saturationPressureKPa(input.dryBulbTemp);
    const moisture = resolveVapourPressure(input, pressureKPa);
    const pwKPa = moisture.pwKPa;

    const derivedRh = clamp((pwKPa / saturationAtDryBulb) * 100, 0, 100);
    const humidityRatio = 0.62198 * (pwKPa / (pressureKPa - pwKPa));
    const enthalpy = 1.006 * input.dryBulbTemp + humidityRatio * (2501 + 1.86 * input.dryBulbTemp);
    const specificVolume =
      (0.287042 * (input.dryBulbTemp + 273.15) * (1 + 1.607858 * humidityRatio)) / pressureKPa;
    const density = (1 + humidityRatio) / specificVolume;

    return {
      source: moisture.source,
      sourceWarning: moisture.warning,
      enthalpy,
      humidityRatio,
      density,
      partialVapourPressurePa: pwKPa * 1000,
      specificVolume,
      saturatedVapourPressurePa: saturationAtDryBulb * 1000,
      atmosphericPressurePa: pressureKPa * 1000,
      derivedRh,
      derivedDewpoint: dewpointFromPw(pwKPa),
      derivedWetBulb: wetBulbApproxFromDryRh(input.dryBulbTemp, derivedRh)
    };
  }

  function formatDatePretty(dateStr) {
    const date = new Date(`${dateStr}T00:00:00`);
    if (Number.isNaN(date.getTime())) {
      return dateStr;
    }
    return date.toLocaleDateString();
  }

  function buildPreviewData(input, result) {
    const reportHeader = {
      companyName: input.companyName || "Not provided",
      projectName: input.projectName || "Not provided",
      revisionNumber: input.revisionNumber || "Not provided",
      reportDate: formatDatePretty(input.reportDate)
    };

    const inputRows = [
      ["Dry Bulb Temp", `${formatNumber(input.dryBulbTemp, 2)} deg C`],
      ["Wet Bulb Temp", input.wetBulbTemp === null ? "Not provided" : `${formatNumber(input.wetBulbTemp, 2)} deg C`],
      [
        "Relative Humidity",
        input.relativeHumidity === null ? "Not provided" : `${formatNumber(input.relativeHumidity, 2)} %`
      ],
      ["Dewpoint Temp", input.dewPointTemp === null ? "Not provided" : `${formatNumber(input.dewPointTemp, 2)} deg C`],
      ["Altitude Above Sea Level", `${formatNumber(input.altitude, 0)} m`],
      ["Calculation Source", result.source]
    ];

    const resultRows = [
      {
        parameter: "Enthalpy",
        symbol: "h",
        symbolHtml: "h",
        symbolPdf: [{ text: "h", style: "normal" }],
        value: formatNumber(result.enthalpy, 2),
        unit: "kJ/kg",
        unitHtml: "kJ/kg",
        unitPdf: [{ text: "kJ/kg", style: "normal" }]
      },
      {
        parameter: "Humidity Ratio",
        symbol: "w",
        symbolHtml: "w",
        symbolPdf: [{ text: "w", style: "normal" }],
        value: formatNumber(result.humidityRatio, 5),
        unit: "kg/kg",
        unitHtml: "kg/kg",
        unitPdf: [{ text: "kg/kg", style: "normal" }]
      },
      {
        parameter: "Density",
        symbol: "rho",
        symbolHtml: "&rho;",
        symbolPdf: [{ text: "rho", style: "normal" }],
        value: formatNumber(result.density, 3),
        unit: "kg/m^3",
        unitHtml: "kg/m<sup>3</sup>",
        unitPdf: [
          { text: "kg/m", style: "normal" },
          { text: "3", style: "sup" }
        ]
      },
      {
        parameter: "Partial Vapour Pressure",
        symbol: "Pv",
        symbolHtml: "P<sub>v</sub>",
        symbolPdf: [
          { text: "P", style: "normal" },
          { text: "v", style: "sub" }
        ],
        value: formatNumber(result.partialVapourPressurePa, 2),
        unit: "Pa",
        unitHtml: "Pa",
        unitPdf: [{ text: "Pa", style: "normal" }]
      },
      {
        parameter: "Specific Volume",
        symbol: "v",
        symbolHtml: "v",
        symbolPdf: [{ text: "v", style: "normal" }],
        value: formatNumber(result.specificVolume, 3),
        unit: "m^3/kg",
        unitHtml: "m<sup>3</sup>/kg",
        unitPdf: [
          { text: "m", style: "normal" },
          { text: "3", style: "sup" },
          { text: "/kg", style: "normal" }
        ]
      },
      {
        parameter: "Saturated Vapour Pressure",
        symbol: "Pws",
        symbolHtml: "P<sub>ws</sub>",
        symbolPdf: [
          { text: "P", style: "normal" },
          { text: "ws", style: "sub" }
        ],
        value: formatNumber(result.saturatedVapourPressurePa, 2),
        unit: "Pa",
        unitHtml: "Pa",
        unitPdf: [{ text: "Pa", style: "normal" }]
      },
      {
        parameter: "Atmospheric Pressure",
        symbol: "Patm",
        symbolHtml: "P<sub>atm</sub>",
        symbolPdf: [
          { text: "P", style: "normal" },
          { text: "atm", style: "sub" }
        ],
        value: formatNumber(result.atmosphericPressurePa, 2),
        unit: "Pa",
        unitHtml: "Pa",
        unitPdf: [{ text: "Pa", style: "normal" }]
      },
      {
        parameter: "Derived Relative Humidity",
        symbol: "RH",
        symbolHtml: "RH",
        symbolPdf: [{ text: "RH", style: "normal" }],
        value: formatNumber(result.derivedRh, 2),
        unit: "%",
        unitHtml: "%",
        unitPdf: [{ text: "%", style: "normal" }]
      },
      {
        parameter: "Derived Dewpoint",
        symbol: "Tdp",
        symbolHtml: "T<sub>dp</sub>",
        symbolPdf: [
          { text: "T", style: "normal" },
          { text: "dp", style: "sub" }
        ],
        value: formatNumber(result.derivedDewpoint, 2),
        unit: "deg C",
        unitHtml: "&deg;C",
        unitPdf: [{ text: "deg C", style: "normal" }]
      },
      {
        parameter: "Derived Wet Bulb",
        symbol: "Twb",
        symbolHtml: "T<sub>wb</sub>",
        symbolPdf: [
          { text: "T", style: "normal" },
          { text: "wb", style: "sub" }
        ],
        value: formatNumber(result.derivedWetBulb, 2),
        unit: "deg C",
        unitHtml: "&deg;C",
        unitPdf: [{ text: "deg C", style: "normal" }]
      }
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

      const maxChars = Math.max(1, Math.floor(remaining / (advanceFor("W", fontSize, style))));
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
    const result = calculateProperties(input);
    const previewData = buildPreviewData(input, result);

    renderReportHeader(previewData.reportHeader);
    renderKeyValueTable(inputTableBody, previewData.inputRows);
    renderResultsTable(previewData.resultRows);

    const now = new Date();
    resultSection.classList.remove("hidden");
    resultMetaEl.textContent = `Calculated at ${now.toLocaleString()} using ${result.source}`;
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

    if (result.sourceWarning) {
      setMessage(`Calculation complete. ${result.sourceWarning}`, "warn");
    } else {
      setMessage("Calculation complete. Preview and PDF are now ready.", "ok");
    }
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
