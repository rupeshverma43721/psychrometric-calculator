(function () {
  if (window.self !== window.top) {
    return;
  }

  const dockEnabled =
    (document.body && document.body.getAttribute("data-calculator-dock") === "enabled") ||
    Boolean(document.querySelector(".sheet-page"));

  if (!dockEnabled) {
    return;
  }

  const calculators = [
    {
      id: "relative-humidity",
      name: "Relative Humidity Calculator",
      description: "DB/WB based RH calculation",
      path: "relative-humidity-calculator.html",
      popupWidth: 760,
      popupHeight: 250,
      status: "active"
    },
    {
      id: "bypass-factor",
      name: "By Pass Factor Calculator",
      description: "Suction/Fresh/Recirculation based BF",
      path: "bypass-factor-calculator.html",
      popupWidth: 760,
      popupHeight: 292,
      status: "active"
    }
  ];

  const calculatorById = {};
  calculators.forEach(function (calc) {
    calculatorById[calc.id] = calc;
  });

  const host = document.createElement("div");
  host.innerHTML = [
    '<button type="button" class="calc-dock-toggle" aria-label="Open calculator list" aria-expanded="false">',
    '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">',
    '<rect x="4" y="2.5" width="16" height="19" rx="2.2" ry="2.2" fill="none" stroke="currentColor" stroke-width="1.8"/>',
    '<rect x="7" y="5.2" width="10" height="3.1" fill="currentColor"/>',
    '<rect x="7" y="10.5" width="2.8" height="2.8" fill="currentColor"/>',
    '<rect x="10.6" y="10.5" width="2.8" height="2.8" fill="currentColor"/>',
    '<rect x="14.2" y="10.5" width="2.8" height="2.8" fill="currentColor"/>',
    '<rect x="7" y="14.1" width="2.8" height="2.8" fill="currentColor"/>',
    '<rect x="10.6" y="14.1" width="2.8" height="2.8" fill="currentColor"/>',
    '<rect x="14.2" y="14.1" width="2.8" height="6.2" fill="currentColor"/>',
    '<rect x="7" y="17.7" width="6.4" height="2.6" fill="currentColor"/>',
    "</svg>",
    "</button>",
    '<section class="calc-dock-panel" aria-label="Calculator List">',
    '<div class="calc-dock-head">',
    "<h3>Calculator List</h3>",
    '<button type="button" class="calc-close-btn" aria-label="Close list">&times;</button>',
    "</div>",
    '<div class="calc-dock-search-wrap">',
    '<input type="search" class="calc-dock-search" placeholder="Search calculators..." aria-label="Search calculators">',
    "</div>",
    '<ul class="calc-dock-list"></ul>',
    "</section>"
  ].join("");

  document.body.appendChild(host);

  const toggleButton = host.querySelector(".calc-dock-toggle");
  const panel = host.querySelector(".calc-dock-panel");
  const panelCloseButton = panel.querySelector(".calc-close-btn");
  const searchInput = panel.querySelector(".calc-dock-search");
  const list = panel.querySelector(".calc-dock-list");

  let zIndexCounter = 3000;
  let windowSpawnCount = 0;
  let searchQuery = "";
  const openWindows = new Set();

  function clamp(value, min, max) {
    return Math.min(Math.max(value, min), max);
  }

  function toPixels(size, fallback) {
    if (typeof size === "number" && Number.isFinite(size)) {
      return Math.max(120, size) + "px";
    }
    if (typeof size === "string" && size.trim()) {
      return size;
    }
    return fallback;
  }

  function toNumber(size, fallback) {
    if (typeof size === "number" && Number.isFinite(size)) {
      return size;
    }

    if (typeof size === "string") {
      const parsed = Number.parseFloat(size);
      if (Number.isFinite(parsed)) {
        return parsed;
      }
    }

    return fallback;
  }

  function withEmbedMode(path) {
    const hashIndex = path.indexOf("#");
    const base = hashIndex >= 0 ? path.slice(0, hashIndex) : path;
    const hash = hashIndex >= 0 ? path.slice(hashIndex) : "";
    const joiner = base.includes("?") ? "&" : "?";
    return base + joiner + "embed=1" + hash;
  }

  function setPanelOpen(open) {
    panel.classList.toggle("open", open);
    toggleButton.setAttribute("aria-expanded", String(open));
  }

  function getTopWindow() {
    let topWindow = null;
    let topZ = -Infinity;

    openWindows.forEach(function (windowEl) {
      const z = Number(windowEl.style.zIndex || 0);
      if (z > topZ) {
        topZ = z;
        topWindow = windowEl;
      }
    });

    return topWindow;
  }

  function closeFloatingWindow(windowEl) {
    if (!openWindows.has(windowEl)) {
      return;
    }
    openWindows.delete(windowEl);
    windowEl.remove();
  }

  function bringToFront(windowEl) {
    zIndexCounter += 1;
    windowEl.style.zIndex = String(zIndexCounter);
  }

  function keepWindowInViewport(windowEl) {
    const rect = windowEl.getBoundingClientRect();
    const maxLeft = Math.max(0, window.innerWidth - rect.width);
    const maxTop = Math.max(0, window.innerHeight - rect.height);
    const nextLeft = clamp(rect.left, 0, maxLeft);
    const nextTop = clamp(rect.top, 0, maxTop);

    windowEl.style.left = nextLeft + "px";
    windowEl.style.top = nextTop + "px";
  }

  function resizeWindowToContent(windowEl, contentHeight) {
    if (!Number.isFinite(contentHeight) || contentHeight <= 0) {
      return;
    }

    const head = windowEl.querySelector(".calc-floating-head");
    if (!head) {
      return;
    }

    const minHeight = head.offsetHeight + 8;
    const maxHeight = Math.max(minHeight, window.innerHeight - 8);
    const nextHeight = clamp(Math.ceil(head.offsetHeight + contentHeight + 6), minHeight, maxHeight);

    windowEl.style.height = nextHeight + "px";
    keepWindowInViewport(windowEl);
  }

  function measureFrameAndResize(windowEl, frame) {
    try {
      const doc = frame.contentDocument;
      if (!doc) {
        return;
      }

      const body = doc.body;
      const root = doc.documentElement;
      const contentHeight = Math.max(
        body ? body.scrollHeight : 0,
        body ? body.offsetHeight : 0,
        root ? root.scrollHeight : 0,
        root ? root.offsetHeight : 0
      );
      resizeWindowToContent(windowEl, contentHeight);
    } catch (error) {
      // Ignore cross-context measurement failures.
    }
  }

  function createFloatingWindow(calculator) {
    const windowEl = document.createElement("article");
    windowEl.className = "calc-floating-window";
    windowEl.setAttribute("role", "dialog");
    windowEl.setAttribute("aria-modal", "false");
    windowEl.setAttribute("aria-label", calculator.name);

    windowEl.innerHTML = [
      '<div class="calc-floating-head" title="Drag to move">',
      '<h3 class="calc-floating-title"></h3>',
      '<button type="button" class="calc-close-btn" aria-label="Close calculator">&times;</button>',
      "</div>",
      '<iframe class="calc-floating-frame" src="about:blank" loading="lazy" scrolling="auto"></iframe>'
    ].join("");

    const titleEl = windowEl.querySelector(".calc-floating-title");
    const closeBtn = windowEl.querySelector(".calc-close-btn");
    const frame = windowEl.querySelector(".calc-floating-frame");
    const head = windowEl.querySelector(".calc-floating-head");

    titleEl.textContent = calculator.name;
    windowEl.style.width = toPixels(calculator.popupWidth, "760px");
    windowEl.style.height = toPixels(calculator.popupHeight, "260px");

    const widthNum = toNumber(calculator.popupWidth, 760);
    const heightNum = toNumber(calculator.popupHeight, 260);
    const offset = (windowSpawnCount % 8) * 24;

    const startLeft = clamp((window.innerWidth - widthNum) / 2 + offset, 8, Math.max(8, window.innerWidth - widthNum - 8));
    const startTop = clamp((window.innerHeight - heightNum) / 2 + offset, 8, Math.max(8, window.innerHeight - heightNum - 8));

    windowEl.style.left = startLeft + "px";
    windowEl.style.top = startTop + "px";
    windowEl.style.transform = "none";

    windowSpawnCount += 1;
    bringToFront(windowEl);

    frame.src = withEmbedMode(calculator.path);
    frame.addEventListener("load", function () {
      measureFrameAndResize(windowEl, frame);
    });

    closeBtn.addEventListener("click", function () {
      closeFloatingWindow(windowEl);
    });

    windowEl.addEventListener("mousedown", function () {
      bringToFront(windowEl);
    });

    windowEl.addEventListener("touchstart", function () {
      bringToFront(windowEl);
    }, { passive: true });

    let dragState = null;

    head.addEventListener("pointerdown", function (event) {
      if (event.button !== 0 || event.target.closest(".calc-close-btn")) {
        return;
      }

      bringToFront(windowEl);
      const rect = windowEl.getBoundingClientRect();
      dragState = {
        pointerId: event.pointerId,
        offsetX: event.clientX - rect.left,
        offsetY: event.clientY - rect.top
      };

      head.classList.add("dragging");
      head.setPointerCapture(event.pointerId);
      event.preventDefault();
    });

    head.addEventListener("pointermove", function (event) {
      if (!dragState || event.pointerId !== dragState.pointerId) {
        return;
      }

      const rect = windowEl.getBoundingClientRect();
      const maxLeft = Math.max(0, window.innerWidth - rect.width);
      const maxTop = Math.max(0, window.innerHeight - rect.height);

      const nextLeft = clamp(event.clientX - dragState.offsetX, 0, maxLeft);
      const nextTop = clamp(event.clientY - dragState.offsetY, 0, maxTop);

      windowEl.style.left = nextLeft + "px";
      windowEl.style.top = nextTop + "px";
    });

    function finishPointerDrag(event) {
      if (dragState && event.pointerId === dragState.pointerId) {
        dragState = null;
        head.classList.remove("dragging");
      }
    }

    head.addEventListener("pointerup", finishPointerDrag);
    head.addEventListener("pointercancel", finishPointerDrag);
    head.addEventListener("lostpointercapture", function () {
      dragState = null;
      head.classList.remove("dragging");
    });

    document.body.appendChild(windowEl);
    openWindows.add(windowEl);
  }

  function renderList() {
    const query = searchQuery.toLowerCase();
    const filteredCalculators = calculators.filter(function (calc) {
      if (!query) {
        return true;
      }
      return (
        calc.name.toLowerCase().includes(query) ||
        calc.description.toLowerCase().includes(query)
      );
    });

    if (!filteredCalculators.length) {
      list.innerHTML = '<li class="calc-dock-empty">No calculators found.</li>';
      return;
    }

    list.innerHTML = filteredCalculators
      .map(function (calc) {
        return [
          '<li class="calc-dock-item">',
          '<button type="button" data-calc-id="' + calc.id + '">',
          "<strong>" + calc.name + "</strong>",
          "<span>" + calc.description + "</span>",
          "</button>",
          "</li>"
        ].join("");
      })
      .join("");
  }

  toggleButton.addEventListener("click", function () {
    const next = !panel.classList.contains("open");
    setPanelOpen(next);
  });

  panelCloseButton.addEventListener("click", function () {
    setPanelOpen(false);
  });

  searchInput.addEventListener("input", function () {
    searchQuery = searchInput.value.trim();
    renderList();
  });

  list.addEventListener("click", function (event) {
    const button = event.target.closest("button[data-calc-id]");
    if (!button) {
      return;
    }

    const calculator = calculatorById[button.getAttribute("data-calc-id")];
    if (!calculator || calculator.status !== "active") {
      return;
    }

    createFloatingWindow(calculator);
    setPanelOpen(false);
  });

  document.addEventListener("click", function (event) {
    const trigger = event.target.closest("[data-import-calculator]");
    if (trigger) {
      const calculator = calculatorById[trigger.getAttribute("data-import-calculator")];
      if (calculator && calculator.status === "active") {
        event.preventDefault();
        createFloatingWindow(calculator);
      }
      return;
    }

    if (!host.contains(event.target)) {
      setPanelOpen(false);
    }
  });

  document.addEventListener("keydown", function (event) {
    if (event.key !== "Escape") {
      return;
    }

    if (panel.classList.contains("open")) {
      setPanelOpen(false);
      return;
    }

    const topWindow = getTopWindow();
    if (topWindow) {
      closeFloatingWindow(topWindow);
    }
  });

  window.addEventListener("message", function (event) {
    const payload = event.data;
    if (!payload || payload.type !== "calculator-embed-size") {
      return;
    }

    const contentHeight = Number(payload.height);
    openWindows.forEach(function (windowEl) {
      const frame = windowEl.querySelector(".calc-floating-frame");
      if (frame && frame.contentWindow === event.source) {
        resizeWindowToContent(windowEl, contentHeight);
      }
    });
  });

  window.addEventListener("resize", function () {
    openWindows.forEach(function (windowEl) {
      keepWindowInViewport(windowEl);
    });
  });

  renderList();
})();
