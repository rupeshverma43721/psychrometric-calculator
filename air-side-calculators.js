(function () {
  const isEmbed = new URLSearchParams(window.location.search).get("embed") === "1";
  if (isEmbed) {
    document.body.classList.add("embed");
  }

  const rhDbt = document.getElementById("rh-dbt");
  const rhWbt = document.getElementById("rh-wbt");
  const rhOutput = document.getElementById("rh-output");

  const bfSuction = document.getElementById("bf-suction");
  const bfFresh = document.getElementById("bf-fresh");
  const bfRecirculation = document.getElementById("bf-recirculation");
  const bfOutput = document.getElementById("bf-output");

  function postEmbedSize() {
    if (!isEmbed || !window.parent) {
      return;
    }

    const container = document.querySelector(".mini-card") || document.querySelector(".air-page");
    if (!container) {
      return;
    }

    const rect = container.getBoundingClientRect();
    const width = Math.ceil(rect.width);
    const height = Math.ceil(rect.height);

    window.parent.postMessage(
      {
        type: "calculator-embed-size",
        height: height,
        width: width
      },
      "*"
    );
  }

  function formatFixed(value, digits) {
    if (!Number.isFinite(value)) {
      return "--";
    }
    return Number(value).toFixed(digits);
  }

  function formatPercent(value) {
    if (!Number.isFinite(value)) {
      return "--";
    }

    const pct = value * 100;
    const rounded = Math.round(pct);
    if (Math.abs(pct - rounded) < 0.05) {
      return String(rounded) + "%";
    }
    return pct.toFixed(1) + "%";
  }

  function calculateRelativeHumidity() {
    if (!rhDbt || !rhWbt || !rhOutput) {
      return;
    }

    const db = Number(rhDbt.value);
    const wb = Number(rhWbt.value);

    if (!Number.isFinite(db) || !Number.isFinite(wb)) {
      rhOutput.textContent = "--";
      return;
    }

    const o10 = 0.6687451584;
    const o8 = 6.112 * Math.exp((17.502 * db) / (240.97 + db));
    const o9 = 6.112 * Math.exp((17.502 * wb) / (240.97 + wb));

    if (!Number.isFinite(o8) || Math.abs(o8) < 1e-12) {
      rhOutput.textContent = "--";
      return;
    }

    const rh = (100 / o8) * (o9 - o10 * (1 + 0.00115 * wb) * (db - wb));
    rhOutput.textContent = formatFixed(rh, 1);
  }

  function calculateByPassFactor() {
    if (!bfSuction || !bfFresh || !bfRecirculation || !bfOutput) {
      return;
    }

    const suction = Number(bfSuction.value);
    const fresh = Number(bfFresh.value);
    const recirculation = Number(bfRecirculation.value);

    if (!Number.isFinite(suction) || !Number.isFinite(fresh) || !Number.isFinite(recirculation)) {
      bfOutput.textContent = "--";
      return;
    }

    const denominator = fresh - recirculation;
    if (Math.abs(denominator) < 1e-12) {
      bfOutput.textContent = "--";
      return;
    }

    const bf = (suction - recirculation) / denominator;
    bfOutput.textContent = formatPercent(bf);
  }

  function recalculate() {
    calculateRelativeHumidity();
    calculateByPassFactor();
    postEmbedSize();
  }

  [rhDbt, rhWbt, bfSuction, bfFresh, bfRecirculation].forEach(function (input) {
    if (input) {
      input.addEventListener("input", recalculate);
    }
  });

  window.addEventListener("load", postEmbedSize);
  window.addEventListener("resize", postEmbedSize);
  if (document.fonts && typeof document.fonts.ready === "object") {
    document.fonts.ready.then(postEmbedSize);
  }

  const resizeTarget = document.querySelector(".mini-card") || document.querySelector(".air-page");
  if (resizeTarget && typeof ResizeObserver !== "undefined") {
    const observer = new ResizeObserver(function () {
      postEmbedSize();
    });
    observer.observe(resizeTarget);
  }

  recalculate();
})();
