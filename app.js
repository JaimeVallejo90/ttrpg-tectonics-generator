(() => {
  const WIDTH = 297;
  const HEIGHT = 210;
  const GRID = 6;
  const INNER_START = 0;
  const INNER_COUNT = 6;
  const SECTOR_SIZE = 2;
  const SECTOR_COUNT = 3;
  const SUBGRID_SIZE = 6;
  const MICRO_PER_BIG = SUBGRID_SIZE / SECTOR_SIZE;
  const MICRO_COUNT = INNER_COUNT * MICRO_PER_BIG;
  const TOTAL_CELLS = MICRO_COUNT * MICRO_COUNT;
  const MAX_ATTEMPTS = 1000;

  const cellW = WIDTH / GRID;
  const cellH = HEIGHT / GRID;
  const innerLeft = cellW * INNER_START;
  const innerTop = cellH * INNER_START;
  const innerRight = innerLeft + INNER_COUNT * cellW;
  const innerBottom = innerTop + INNER_COUNT * cellH;
  const microW = cellW / MICRO_PER_BIG;
  const microH = cellH / MICRO_PER_BIG;
  const dotRadius = Math.min(microW, microH) * 0.12;
  const draftPointRadius = Math.min(microW, microH) * 0.14;
  const closeDistance = Math.min(microW, microH) * 0.6;
  const nodeRadius = Math.min(microW, microH) * 0.18;
  const nodeHitRadius = nodeRadius * 1.9;
  const continentFill = "rgba(194, 99, 42, 0.4)";
  const continentStroke = "rgba(27, 26, 23, 0.85)";
  const continentStrokeActive = "rgba(194, 99, 42, 0.95)";
  const draftStroke = "rgba(194, 99, 42, 0.8)";
  const draftPoint = "rgba(194, 99, 42, 0.9)";
  const draftPointActive = "rgba(79, 111, 106, 0.9)";

  const state = {
    selectedCells: new Set(),
    history: [],
    lastSector: null,
    lastCell: null,
    tool: "roll",
    isDrawing: false,
    currentLine: null,
    lines: [],
    continentPieces: [],
    continentDraft: [],
    continentCursor: null,
    continentDrag: null,
    selectedContinentId: null,
    continentHoverNode: null,
    nextContinentId: 1,
    clearConfirm: false,
  };

  const ui = {};

  function init() {
    ui.board = document.getElementById("board");
    ui.rollBtn = document.getElementById("roll-btn");
    ui.rollSixBtn = document.getElementById("roll-6-btn");
    ui.pencilBtn = document.getElementById("pencil-btn");
    ui.continentBtn = document.getElementById("continent-btn");
    ui.deleteContinentBtn = document.getElementById("delete-continent-btn");
    ui.clearBtn = document.getElementById("clear-btn");
    ui.toolMode = document.getElementById("tool-mode");
    ui.lastSector = document.getElementById("last-sector");
    ui.lastCell = document.getElementById("last-cell");
    ui.filled = document.getElementById("filled-count");
    ui.message = document.getElementById("message");
    ui.continentCanvas = document.getElementById("continent-canvas");
    ui.continentHint = document.getElementById("continent-hint");
    ui.continentCtx = ui.continentCanvas
      ? ui.continentCanvas.getContext("2d")
      : null;

    ui.board.setAttribute("viewBox", `0 0 ${WIDTH} ${HEIGHT}`);
    ui.board.setAttribute("preserveAspectRatio", "xMidYMid meet");
    if (ui.continentCanvas && ui.continentCtx) {
      updateOverlayCanvasSize();
    }

    ui.rollBtn.addEventListener("click", handleRoll);
    ui.rollSixBtn.addEventListener("click", handleRollSix);
    ui.pencilBtn.addEventListener("click", togglePencil);
    ui.continentBtn.addEventListener("click", toggleContinent);
    ui.deleteContinentBtn.addEventListener("click", handleDeleteContinent);
    ui.clearBtn.addEventListener("click", handleClearAll);
    ui.board.addEventListener("pointerdown", handlePointerDown);
    ui.board.addEventListener("pointermove", handlePointerMove);
    ui.board.addEventListener("pointerup", handlePointerUp);
    ui.board.addEventListener("pointerleave", handlePointerUp);
    ui.board.addEventListener("pointercancel", handlePointerUp);
    ui.board.addEventListener("contextmenu", handleContextMenu);
    if (ui.continentCanvas) {
      ui.continentCanvas.addEventListener("pointerdown", handleContinentPointerDown);
      ui.continentCanvas.addEventListener("pointermove", handleContinentPointerMove);
      ui.continentCanvas.addEventListener("pointerup", handleContinentPointerUp);
      ui.continentCanvas.addEventListener("pointerleave", handleContinentPointerUp);
      ui.continentCanvas.addEventListener("pointercancel", handleContinentPointerUp);
      ui.continentCanvas.addEventListener("contextmenu", handleContinentContextMenu);
    }
    window.addEventListener("resize", handleResize);
    window.addEventListener("keydown", handleKeydown);

    render();
  }

  function rollD6() {
    return Math.floor(Math.random() * 6) + 1;
  }

  function toGroupIndex(value) {
    return Math.floor((value - 1) / SECTOR_SIZE);
  }

  function groupLabel(index) {
    const start = index * SECTOR_SIZE + 1;
    return `${start}-${start + 1}`;
  }

  function handleRoll() {
    cancelClearConfirm();
    if (state.selectedCells.size >= TOTAL_CELLS) {
      setMessage("Grid is full.");
      render();
      return;
    }

    const result = rollRandomPoint();
    if (result) {
      setMessage(
        `Marked ${groupLabel(result.sectorCol)} x ${groupLabel(result.sectorRow)} (${result.cellCol}, ${result.cellRow}).`
      );
    } else {
      setMessage("No free cell found.");
    }
    render();
  }

  function handleRollSix() {
    cancelClearConfirm();
    if (state.selectedCells.size >= TOTAL_CELLS) {
      setMessage("Grid is full.");
      render();
      return;
    }

    let added = 0;
    let lastResult = null;

    for (let i = 0; i < 6; i += 1) {
      if (state.selectedCells.size >= TOTAL_CELLS) {
        break;
      }
      const result = rollRandomPoint();
      if (!result) {
        break;
      }
      added += 1;
      lastResult = result;
    }

    if (added === 0) {
      setMessage("No free cell found.");
    } else {
      const lastText = lastResult
        ? ` Last ${groupLabel(lastResult.sectorCol)} x ${groupLabel(
            lastResult.sectorRow
          )} (${lastResult.cellCol}, ${lastResult.cellRow}).`
        : "";
      setMessage(`Marked ${added} points.${lastText}`);
    }

    render();
  }

  function rollRandomPoint() {
    for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt += 1) {
      const sectorCol = toGroupIndex(rollD6());
      const sectorRow = toGroupIndex(rollD6());
      const cellCol = rollD6();
      const cellRow = rollD6();
      const microCol = sectorCol * SUBGRID_SIZE + (cellCol - 1);
      const microRow = sectorRow * SUBGRID_SIZE + (cellRow - 1);

      if (addCell(microCol, microRow, sectorCol, sectorRow, cellCol, cellRow)) {
        return { sectorCol, sectorRow, cellCol, cellRow };
      }
    }
    return null;
  }

  function setTool(nextTool) {
    cancelClearConfirm();
    const target = state.tool === nextTool ? "roll" : nextTool;
    state.tool = target;
    if (state.tool !== "pencil") {
      state.isDrawing = false;
      state.currentLine = null;
    }
    if (state.tool !== "continent") {
      state.continentDraft = [];
      state.continentCursor = null;
      state.continentDrag = null;
      state.continentHoverNode = null;
    }
    render();
  }

  function togglePencil() {
    setTool("pencil");
  }

  function toggleContinent() {
    setTool("continent");
  }

  function hasMarks() {
    return (
      state.selectedCells.size > 0 ||
      state.lines.length > 0 ||
      state.continentPieces.length > 0 ||
      state.continentDraft.length > 0
    );
  }

  function cancelClearConfirm() {
    if (state.clearConfirm) {
      state.clearConfirm = false;
    }
  }

  function handleClearAll() {
    if (!hasMarks()) {
      state.clearConfirm = false;
      setMessage("Nothing to clear.");
      render();
      return;
    }
    if (!state.clearConfirm) {
      state.clearConfirm = true;
      setMessage("Click Clear all again to confirm.");
      render();
      return;
    }
    const snapshot = createClearSnapshot();
    state.history.push({ type: "clear", snapshot });
    clearAllState();
    state.clearConfirm = false;
    setMessage("Cleared all marks.");
    render();
  }

  function clearAllState() {
    state.selectedCells.clear();
    state.lastSector = null;
    state.lastCell = null;
    state.currentLine = null;
    state.isDrawing = false;
    state.lines = [];
    state.continentPieces = [];
    state.continentDraft = [];
    state.continentCursor = null;
    state.continentDrag = null;
    state.selectedContinentId = null;
    state.continentHoverNode = null;
    state.nextContinentId = 1;
  }

  function createClearSnapshot() {
    return {
      selectedCells: Array.from(state.selectedCells),
      lines: state.lines.map((line) => ({ ...line })),
      lastSector: state.lastSector ? { ...state.lastSector } : null,
      lastCell: state.lastCell ? { ...state.lastCell } : null,
      continentPieces: state.continentPieces.map(cloneContinentPiece),
      selectedContinentId: state.selectedContinentId,
      nextContinentId: state.nextContinentId,
    };
  }

  function restoreClearSnapshot(snapshot) {
    state.selectedCells = new Set(snapshot.selectedCells);
    state.lines = snapshot.lines.map((line) => ({ ...line }));
    state.lastSector = snapshot.lastSector ? { ...snapshot.lastSector } : null;
    state.lastCell = snapshot.lastCell ? { ...snapshot.lastCell } : null;
    state.continentPieces = snapshot.continentPieces.map(cloneContinentPiece);
    state.selectedContinentId = snapshot.selectedContinentId;
    state.nextContinentId = snapshot.nextContinentId;
    state.currentLine = null;
    state.isDrawing = false;
    state.continentDraft = [];
    state.continentCursor = null;
    state.continentDrag = null;
    state.continentHoverNode = null;
    if (
      state.selectedContinentId &&
      !state.continentPieces.some((piece) => piece.id === state.selectedContinentId)
    ) {
      state.selectedContinentId = null;
    }
  }

  function handleDeleteContinent() {
    cancelClearConfirm();
    if (!deleteSelectedContinent()) {
      setMessage("Select a continent to delete.");
      render();
      return;
    }
    render();
  }

  function handlePointerDown(event) {
    if (state.tool !== "pencil" || event.button !== 0) {
      return;
    }
    cancelClearConfirm();
    const point = getSvgPointInPage(event, false);
    if (!point) {
      return;
    }
    state.isDrawing = true;
    state.currentLine = {
      x1: point.x,
      y1: point.y,
      x2: point.x,
      y2: point.y,
    };
    if (ui.board.setPointerCapture) {
      ui.board.setPointerCapture(event.pointerId);
    }
    render();
  }

  function handlePointerMove(event) {
    if (!state.isDrawing || state.tool !== "pencil") {
      return;
    }
    updateCurrentLine(event);
  }

  function handlePointerUp(event) {
    if (!state.isDrawing) {
      return;
    }
    state.isDrawing = false;
    finishCurrentLine();
    if (ui.board.releasePointerCapture) {
      try {
        ui.board.releasePointerCapture(event.pointerId);
      } catch (error) {
        // Ignore if pointer capture was not set.
      }
    }
  }

  function handleContextMenu(event) {
    event.preventDefault();
    cancelClearConfirm();
    if (state.isDrawing) {
      state.isDrawing = false;
      state.currentLine = null;
    }
    const point = getSvgPointRaw(event);
    if (!point) {
      return;
    }
    const removed = removeAtPoint(point.x, point.y);
    if (!removed) {
      setMessage("Nothing to delete.");
    }
    render();
  }

  function handleKeydown(event) {
    const key = event.key.toLowerCase();
    if (key === "delete") {
      if (isEditableTarget(event.target)) {
        return;
      }
      cancelClearConfirm();
      if (deleteSelectedContinent()) {
        event.preventDefault();
        render();
      }
      return;
    }
    if ((event.ctrlKey || event.metaKey) && key === "z") {
      if (isEditableTarget(event.target)) {
        return;
      }
      cancelClearConfirm();
      event.preventDefault();
      undoLast();
    }
  }

  function updateCurrentLine(event) {
    if (!state.currentLine) {
      return;
    }
    const point = getSvgPointInPage(event, true);
    if (!point) {
      return;
    }
    state.currentLine.x2 = point.x;
    state.currentLine.y2 = point.y;
    render();
  }

  function finishCurrentLine() {
    if (!state.currentLine) {
      return;
    }
    const { x1, y1, x2, y2 } = state.currentLine;
    const dx = x2 - x1;
    const dy = y2 - y1;
    const length = Math.hypot(dx, dy);
    const minLength = Math.min(microW, microH) * 0.2;
    if (length >= minLength) {
      state.lines.push({ x1, y1, x2, y2 });
      state.history.push({ type: "line" });
      setMessage("Line drawn.");
    }
    state.currentLine = null;
    render();
  }

  function getSvgPointInPage(event, clampToPage) {
    const rect = ui.board.getBoundingClientRect();
    if (!rect.width || !rect.height) {
      return null;
    }
    const x = ((event.clientX - rect.left) / rect.width) * WIDTH;
    const y = ((event.clientY - rect.top) / rect.height) * HEIGHT;
    if (!clampToPage) {
      if (x < 0 || x > WIDTH || y < 0 || y > HEIGHT) {
        return null;
      }
      return { x, y };
    }
    return {
      x: clamp(x, 0, WIDTH),
      y: clamp(y, 0, HEIGHT),
    };
  }

  function getSvgPointRaw(event) {
    return getSvgPointInPage(event, false);
  }

  function getSvgPointUnclamped(event) {
    const rect = ui.board.getBoundingClientRect();
    if (!rect.width || !rect.height) {
      return null;
    }
    return {
      x: ((event.clientX - rect.left) / rect.width) * WIDTH,
      y: ((event.clientY - rect.top) / rect.height) * HEIGHT,
    };
  }

  function removeAtPoint(x, y) {
    const pointHit = Math.min(microW, microH) * 0.35;
    const lineHit = Math.min(microW, microH) * 0.35;
    const pointCandidate = findNearestPoint(x, y, pointHit);
    const lineCandidate = findNearestLine(x, y, lineHit);

    if (pointCandidate && lineCandidate) {
      if (pointCandidate.distance <= lineCandidate.distance) {
        removePoint(pointCandidate.key, pointCandidate.col, pointCandidate.row);
        return true;
      }
      removeLine(lineCandidate.index);
      return true;
    }

    if (pointCandidate) {
      removePoint(pointCandidate.key, pointCandidate.col, pointCandidate.row);
      return true;
    }

    if (lineCandidate) {
      removeLine(lineCandidate.index);
      return true;
    }

    return false;
  }

  function findNearestPoint(x, y, radius) {
    let best = null;
    const radiusSq = radius * radius;
    state.selectedCells.forEach((key) => {
      const [col, row] = key.split(",").map(Number);
      const cx = innerLeft + col * microW + microW / 2;
      const cy = innerTop + row * microH + microH / 2;
      const dx = x - cx;
      const dy = y - cy;
      const distSq = dx * dx + dy * dy;
      if (distSq <= radiusSq) {
        if (!best || distSq < best.distanceSq) {
          best = {
            key,
            col,
            row,
            distanceSq: distSq,
          };
        }
      }
    });
    if (!best) {
      return null;
    }
    return {
      key: best.key,
      col: best.col,
      row: best.row,
      distance: Math.sqrt(best.distanceSq),
    };
  }

  function findNearestLine(x, y, radius) {
    let best = null;
    const radiusSq = radius * radius;
    state.lines.forEach((line, index) => {
      const distSq = distanceToSegmentSquared(x, y, line);
      if (distSq <= radiusSq) {
        if (!best || distSq < best.distanceSq) {
          best = { index, distanceSq: distSq };
        }
      }
    });
    if (!best) {
      return null;
    }
    return {
      index: best.index,
      distance: Math.sqrt(best.distanceSq),
    };
  }

  function distanceToSegmentSquared(px, py, line) {
    const { x1, y1, x2, y2 } = line;
    const dx = x2 - x1;
    const dy = y2 - y1;
    if (dx === 0 && dy === 0) {
      const nx = px - x1;
      const ny = py - y1;
      return nx * nx + ny * ny;
    }
    const t = ((px - x1) * dx + (py - y1) * dy) / (dx * dx + dy * dy);
    const clamped = Math.min(Math.max(t, 0), 1);
    const lx = x1 + clamped * dx;
    const ly = y1 + clamped * dy;
    const ox = px - lx;
    const oy = py - ly;
    return ox * ox + oy * oy;
  }

  function removePoint(key, microCol, microRow) {
    state.selectedCells.delete(key);
    const sectorCol = Math.floor(microCol / SUBGRID_SIZE);
    const sectorRow = Math.floor(microRow / SUBGRID_SIZE);
    const cellCol = (microCol % SUBGRID_SIZE) + 1;
    const cellRow = (microRow % SUBGRID_SIZE) + 1;
    state.history.push({
      type: "erase",
      itemType: "point",
      key,
      sector: { col: sectorCol, row: sectorRow },
      cell: { col: cellCol, row: cellRow },
    });
    syncLastPoint();
    setMessage("Point deleted.");
  }

  function removeLine(index) {
    const [line] = state.lines.splice(index, 1);
    if (!line) {
      return;
    }
    state.history.push({
      type: "erase",
      itemType: "line",
      line,
      index,
    });
    setMessage("Line deleted.");
  }

  function addCell(microCol, microRow, sectorCol, sectorRow, cellCol, cellRow) {
    const key = `${microCol},${microRow}`;
    if (state.selectedCells.has(key)) {
      return false;
    }
    state.selectedCells.add(key);
    const sector = { col: sectorCol, row: sectorRow };
    const cell = { col: cellCol, row: cellRow };
    state.history.push({ type: "point", key, sector, cell });
    state.lastSector = sector;
    state.lastCell = cell;
    return true;
  }

  function undoLast() {
    const last = state.history.pop();
    if (!last) {
      setMessage("Nothing to undo.");
      render();
      return;
    }
    state.clearConfirm = false;
    if (last.type === "point") {
      state.selectedCells.delete(last.key);
    }
    if (last.type === "line") {
      state.lines.pop();
    }
    if (last.type === "erase") {
      if (last.itemType === "point") {
        state.selectedCells.add(last.key);
      }
      if (last.itemType === "line") {
        const insertAt = clamp(last.index, 0, state.lines.length);
        state.lines.splice(insertAt, 0, last.line);
      }
    }
    if (last.type === "continent-add") {
      undoContinentAdd(last);
    }
    if (last.type === "continent-delete") {
      undoContinentDelete(last);
    }
    if (last.type === "continent-update") {
      undoContinentUpdate(last);
    }
    if (last.type === "clear") {
      restoreClearSnapshot(last.snapshot);
    }
    syncLastPoint();
    setMessage("Undo.");
    render();
  }

  function syncLastPoint() {
    for (let i = state.history.length - 1; i >= 0; i -= 1) {
      const entry = state.history[i];
      if (entry.type === "point" && state.selectedCells.has(entry.key)) {
        state.lastSector = entry.sector;
        state.lastCell = entry.cell;
        return;
      }
    }
    state.lastSector = null;
    state.lastCell = null;
  }

  function clamp(value, min, max) {
    return Math.min(Math.max(value, min), max);
  }

  function isEditableTarget(target) {
    if (!target) {
      return false;
    }
    const tag = target.tagName;
    return tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable;
  }

  function setMessage(text) {
    ui.message.textContent = text || "";
  }

  function render() {
    if (state.tool === "pencil") {
      ui.toolMode.textContent = "Paint plates";
    } else if (state.tool === "continent") {
      ui.toolMode.textContent = "Paint continent";
    } else {
      ui.toolMode.textContent = "Roll";
    }
    ui.lastSector.textContent = formatLastSector();
    ui.lastCell.textContent = formatLastCell();
    ui.filled.textContent = state.selectedCells.size;
    ui.rollBtn.disabled = state.selectedCells.size >= TOTAL_CELLS;
    ui.rollSixBtn.disabled = state.selectedCells.size >= TOTAL_CELLS;
    ui.clearBtn.disabled = !hasMarks();
    ui.clearBtn.textContent = state.clearConfirm ? "Confirm clear" : "Clear all";
    ui.pencilBtn.setAttribute("aria-pressed", state.tool === "pencil");
    ui.pencilBtn.classList.toggle("active", state.tool === "pencil");
    ui.continentBtn.setAttribute("aria-pressed", state.tool === "continent");
    ui.continentBtn.classList.toggle("active", state.tool === "continent");
    ui.deleteContinentBtn.disabled = !state.selectedContinentId;
    ui.board.classList.toggle("pencil-active", state.tool === "pencil");
    document.body.dataset.tool = state.tool;
    ui.board.innerHTML = buildSvg();
    drawContinents();
  }

  function formatLastSector() {
    if (!state.lastSector) {
      return "-";
    }
    return `${groupLabel(state.lastSector.col)} x ${groupLabel(state.lastSector.row)}`;
  }

  function formatLastCell() {
    if (!state.lastCell) {
      return "-";
    }
    return `${state.lastCell.col}, ${state.lastCell.row}`;
  }

  function buildSvg() {
    const parts = [];

    for (let i = 1; i < MICRO_COUNT; i += 1) {
      if (i % MICRO_PER_BIG === 0) {
        continue;
      }
      const x = innerLeft + i * microW;
      parts.push(
        `<line class="grid-line" x1="${x.toFixed(2)}" y1="${innerTop.toFixed(
          2
        )}" x2="${x.toFixed(2)}" y2="${innerBottom.toFixed(2)}" />`
      );
    }

    for (let i = 1; i < MICRO_COUNT; i += 1) {
      if (i % MICRO_PER_BIG === 0) {
        continue;
      }
      const y = innerTop + i * microH;
      parts.push(
        `<line class="grid-line" x1="${innerLeft.toFixed(2)}" y1="${y.toFixed(
          2
        )}" x2="${innerRight.toFixed(2)}" y2="${y.toFixed(2)}" />`
      );
    }

    for (let i = 0; i <= GRID; i += 1) {
      const x = (cellW * i).toFixed(2);
      const y = (cellH * i).toFixed(2);
      parts.push(`<line class="grid-line-strong" x1="${x}" y1="0" x2="${x}" y2="${HEIGHT}" />`);
      parts.push(`<line class="grid-line-strong" x1="0" y1="${y}" x2="${WIDTH}" y2="${y}" />`);
    }

    parts.push(
      `<rect class="grid-border" x="0.5" y="0.5" width="${WIDTH - 1}" height="${HEIGHT - 1}" />`
    );

    state.lines.forEach((line) => {
      parts.push(
        `<line class="pencil-line" x1="${line.x1.toFixed(2)}" y1="${line.y1.toFixed(
          2
        )}" x2="${line.x2.toFixed(2)}" y2="${line.y2.toFixed(2)}" />`
      );
    });

    if (state.currentLine) {
      parts.push(
        `<line class="pencil-line" x1="${state.currentLine.x1.toFixed(
          2
        )}" y1="${state.currentLine.y1.toFixed(2)}" x2="${state.currentLine.x2.toFixed(
          2
        )}" y2="${state.currentLine.y2.toFixed(2)}" />`
      );
    }

    state.selectedCells.forEach((key) => {
      const [col, row] = key.split(",").map(Number);
      const cx = innerLeft + col * microW + microW / 2;
      const cy = innerTop + row * microH + microH / 2;
      parts.push(
        `<circle class="cell-dot" cx="${cx.toFixed(2)}" cy="${cy.toFixed(
          2
        )}" r="${dotRadius.toFixed(2)}" />`
      );
    });

    return parts.join("");
  }

  function handleResize() {
    updateOverlayCanvasSize();
    drawContinents();
  }

  function updateOverlayCanvasSize() {
    if (!ui.continentCanvas || !ui.continentCtx) {
      return;
    }
    const rect = ui.continentCanvas.getBoundingClientRect();
    if (!rect.width || !rect.height) {
      return;
    }
    const ratio = Math.max(1, Math.min(2, window.devicePixelRatio || 1));
    const targetWidth = Math.round(rect.width * ratio);
    const targetHeight = Math.round(rect.height * ratio);
    if (
      ui.continentCanvas.width !== targetWidth ||
      ui.continentCanvas.height !== targetHeight
    ) {
      ui.continentCanvas.width = targetWidth;
      ui.continentCanvas.height = targetHeight;
    }
    const scaleX = targetWidth / WIDTH;
    const scaleY = targetHeight / HEIGHT;
    ui.continentCtx.setTransform(scaleX, 0, 0, scaleY, 0, 0);
    ui.continentCtx.lineJoin = "round";
    ui.continentCtx.lineCap = "round";
  }

  function handleContinentContextMenu(event) {
    if (state.tool === "continent") {
      event.preventDefault();
    }
  }

  function handleContinentPointerDown(event) {
    if (state.tool !== "continent") {
      return;
    }
    cancelClearConfirm();
    if (event.button !== 0 && event.button !== 2) {
      return;
    }
    const point = getSvgPointInPage(event, false);
    if (!point) {
      return;
    }

    if (event.button === 2) {
      const hit = findContinentAt(point);
      if (!hit) {
        return;
      }
      selectContinent(hit.id);
      state.continentDrag = {
        type: "rotate",
        id: hit.id,
        startAngle: Math.atan2(point.y - hit.y, point.x - hit.x),
        startRotation: hit.rotation,
        before: cloneContinentPiece(hit),
      };
      if (ui.continentCanvas.setPointerCapture) {
        ui.continentCanvas.setPointerCapture(event.pointerId);
      }
      ui.continentCanvas.style.cursor = "grabbing";
      event.preventDefault();
      render();
      return;
    }

    if (state.continentDraft.length > 0) {
      const first = state.continentDraft[0];
      const distance = Math.hypot(point.x - first.x, point.y - first.y);
      if (state.continentDraft.length >= 3 && distance <= closeDistance) {
        finalizeContinentDraft();
        render();
        return;
      }
      state.continentDraft.push(point);
      state.continentCursor = point;
      render();
      return;
    }

    const nodeHit = findContinentNodeAt(point);
    if (nodeHit) {
      selectContinent(nodeHit.piece.id);
      state.continentDrag = {
        type: "node",
        id: nodeHit.piece.id,
        index: nodeHit.index,
        offsetX: nodeHit.offsetX,
        offsetY: nodeHit.offsetY,
        before: cloneContinentPiece(nodeHit.piece),
      };
      if (ui.continentCanvas.setPointerCapture) {
        ui.continentCanvas.setPointerCapture(event.pointerId);
      }
      ui.continentCanvas.style.cursor = "grabbing";
      render();
      return;
    }

    const hit = findContinentAt(point);
    if (hit) {
      selectContinent(hit.id);
      state.continentDrag = {
        type: "move",
        id: hit.id,
        offsetX: point.x - hit.x,
        offsetY: point.y - hit.y,
        before: cloneContinentPiece(hit),
      };
      if (ui.continentCanvas.setPointerCapture) {
        ui.continentCanvas.setPointerCapture(event.pointerId);
      }
      ui.continentCanvas.style.cursor = "grabbing";
      render();
      return;
    }

    startContinentDraft(point);
    render();
  }

  function handleContinentPointerMove(event) {
    if (state.tool !== "continent") {
      return;
    }
    const point = state.continentDrag
      ? getSvgPointUnclamped(event)
      : getSvgPointInPage(event, true);
    if (!point) {
      return;
    }

    if (state.continentDrag) {
      const piece = getContinentById(state.continentDrag.id);
      if (!piece) {
        state.continentDrag = null;
        render();
        return;
      }
      if (state.continentDrag.type === "move") {
        piece.x = point.x - state.continentDrag.offsetX;
        piece.y = point.y - state.continentDrag.offsetY;
      } else if (state.continentDrag.type === "node") {
        const local = toLocalPoint(point, piece);
        piece.points[state.continentDrag.index] = [
          local[0] + state.continentDrag.offsetX,
          local[1] + state.continentDrag.offsetY,
        ];
        updateContinentRadius(piece);
      } else if (state.continentDrag.type === "rotate") {
        const angle = Math.atan2(point.y - piece.y, point.x - piece.x);
        piece.rotation = state.continentDrag.startRotation + (angle - state.continentDrag.startAngle);
      }
      render();
      return;
    }

    if (state.continentDraft.length > 0) {
      state.continentCursor = point;
      ui.continentCanvas.style.cursor = "crosshair";
      render();
      return;
    }

    const nodeHit = findContinentNodeAt(point);
    if (nodeHit) {
      const changed = updateHoverNode(nodeHit.piece.id, nodeHit.index);
      ui.continentCanvas.style.cursor = "pointer";
      if (changed) {
        render();
      }
      return;
    }

    const cleared = updateHoverNode(null, null);
    const hit = findContinentAt(point);
    ui.continentCanvas.style.cursor = hit ? "grab" : "crosshair";
    if (cleared) {
      render();
    }
  }

  function handleContinentPointerUp(event) {
    if (state.tool !== "continent") {
      return;
    }
    if (state.continentDrag) {
      commitContinentDrag();
      state.continentDrag = null;
      if (ui.continentCanvas.releasePointerCapture) {
        try {
          ui.continentCanvas.releasePointerCapture(event.pointerId);
        } catch (error) {
          // Ignore if pointer capture was not set.
        }
      }
    }
    if (state.tool === "continent") {
      ui.continentCanvas.style.cursor = "crosshair";
    }
    render();
  }

  function commitContinentDrag() {
    if (!state.continentDrag || !state.continentDrag.before) {
      return;
    }
    const piece = getContinentById(state.continentDrag.id);
    if (!piece) {
      return;
    }
    if (!hasContinentChanged(state.continentDrag.before, piece)) {
      return;
    }
    state.history.push({
      type: "continent-update",
      id: piece.id,
      before: state.continentDrag.before,
      after: cloneContinentPiece(piece),
    });
  }

  function startContinentDraft(point) {
    state.continentDraft = [point];
    state.continentCursor = point;
    state.selectedContinentId = null;
    state.continentHoverNode = null;
    setMessage("Click points, then click the first point again to close.");
  }

  function finalizeContinentDraft() {
    if (state.continentDraft.length < 3) {
      return;
    }
    const piece = createContinentPiece(state.continentDraft);
    state.continentPieces.push(piece);
    state.history.push({
      type: "continent-add",
      piece: cloneContinentPiece(piece),
    });
    selectContinent(piece.id);
    state.continentDraft = [];
    state.continentCursor = null;
    setMessage("Continent created.");
  }

  function createContinentPiece(points) {
    let sumX = 0;
    let sumY = 0;
    points.forEach((point) => {
      sumX += point.x;
      sumY += point.y;
    });
    const cx = sumX / points.length;
    const cy = sumY / points.length;
    const localPoints = points.map((point) => [point.x - cx, point.y - cy]);
    const piece = {
      id: `continent-${state.nextContinentId++}`,
      x: cx,
      y: cy,
      rotation: 0,
      points: localPoints,
    };
    updateContinentRadius(piece);
    return piece;
  }

  function cloneContinentPiece(piece) {
    const clone = {
      id: piece.id,
      x: piece.x,
      y: piece.y,
      rotation: piece.rotation,
      points: piece.points.map((point) => [point[0], point[1]]),
    };
    updateContinentRadius(clone);
    return clone;
  }

  function applyContinentSnapshot(target, snapshot) {
    target.x = snapshot.x;
    target.y = snapshot.y;
    target.rotation = snapshot.rotation;
    target.points = snapshot.points.map((point) => [point[0], point[1]]);
    updateContinentRadius(target);
  }

  function hasContinentChanged(before, piece) {
    const epsilon = 0.01;
    if (
      Math.abs(before.x - piece.x) > epsilon ||
      Math.abs(before.y - piece.y) > epsilon ||
      Math.abs(before.rotation - piece.rotation) > 0.001
    ) {
      return true;
    }
    if (before.points.length !== piece.points.length) {
      return true;
    }
    for (let i = 0; i < before.points.length; i += 1) {
      const [bx, by] = before.points[i];
      const [px, py] = piece.points[i];
      if (Math.abs(bx - px) > epsilon || Math.abs(by - py) > epsilon) {
        return true;
      }
    }
    return false;
  }

  function selectContinent(id) {
    state.selectedContinentId = id;
    bringContinentToFront(id);
  }

  function deleteSelectedContinent(recordHistory = true) {
    if (!state.selectedContinentId) {
      return false;
    }
    const index = state.continentPieces.findIndex(
      (piece) => piece.id === state.selectedContinentId
    );
    if (index < 0) {
      state.selectedContinentId = null;
      return false;
    }
    const [removed] = state.continentPieces.splice(index, 1);
    if (recordHistory && removed) {
      state.history.push({
        type: "continent-delete",
        piece: cloneContinentPiece(removed),
        index,
      });
    }
    state.selectedContinentId = null;
    state.continentHoverNode = null;
    setMessage("Continent deleted.");
    return true;
  }

  function undoContinentAdd(entry) {
    const index = state.continentPieces.findIndex(
      (piece) => piece.id === entry.piece.id
    );
    if (index >= 0) {
      state.continentPieces.splice(index, 1);
    }
    if (state.selectedContinentId === entry.piece.id) {
      state.selectedContinentId = null;
    }
    state.continentHoverNode = null;
  }

  function undoContinentDelete(entry) {
    const insertAt = clamp(entry.index, 0, state.continentPieces.length);
    const restored = cloneContinentPiece(entry.piece);
    state.continentPieces.splice(insertAt, 0, restored);
    state.selectedContinentId = restored.id;
    state.continentHoverNode = null;
  }

  function undoContinentUpdate(entry) {
    const piece = getContinentById(entry.id);
    if (!piece) {
      return;
    }
    applyContinentSnapshot(piece, entry.before);
  }

  function getContinentById(id) {
    return state.continentPieces.find((piece) => piece.id === id);
  }

  function bringContinentToFront(id) {
    const index = state.continentPieces.findIndex((piece) => piece.id === id);
    if (index < 0 || index === state.continentPieces.length - 1) {
      return;
    }
    const [piece] = state.continentPieces.splice(index, 1);
    state.continentPieces.push(piece);
  }

  function updateContinentRadius(piece) {
    let radius = 0;
    piece.points.forEach((point) => {
      const dist = Math.hypot(point[0], point[1]);
      if (dist > radius) {
        radius = dist;
      }
    });
    piece.radius = radius;
  }

  function findContinentAt(point) {
    for (let i = state.continentPieces.length - 1; i >= 0; i -= 1) {
      const piece = state.continentPieces[i];
      const dx = point.x - piece.x;
      const dy = point.y - piece.y;
      if (piece.radius && dx * dx + dy * dy > piece.radius * piece.radius) {
        continue;
      }
      const localPoint = toLocalPoint(point, piece);
      if (pointInPolygon(localPoint, piece.points)) {
        return piece;
      }
    }
    return null;
  }

  function findContinentNodeAt(point) {
    if (!state.selectedContinentId) {
      return null;
    }
    const piece = getContinentById(state.selectedContinentId);
    if (!piece) {
      return null;
    }
    for (let j = 0; j < piece.points.length; j += 1) {
      const world = toWorldPoint(piece, piece.points[j]);
      const dx = point.x - world.x;
      const dy = point.y - world.y;
      if (dx * dx + dy * dy <= nodeHitRadius * nodeHitRadius) {
        const local = toLocalPoint(point, piece);
        return {
          piece,
          index: j,
          offsetX: piece.points[j][0] - local[0],
          offsetY: piece.points[j][1] - local[1],
        };
      }
    }
    return null;
  }

  function updateHoverNode(id, index) {
    if (id === null) {
      if (state.continentHoverNode) {
        state.continentHoverNode = null;
        return true;
      }
      return false;
    }
    if (
      !state.continentHoverNode ||
      state.continentHoverNode.id !== id ||
      state.continentHoverNode.index !== index
    ) {
      state.continentHoverNode = { id, index };
      return true;
    }
    return false;
  }

  function toLocalPoint(point, piece) {
    const dx = point.x - piece.x;
    const dy = point.y - piece.y;
    const sin = Math.sin(-piece.rotation);
    const cos = Math.cos(-piece.rotation);
    return [dx * cos - dy * sin, dx * sin + dy * cos];
  }

  function toWorldPoint(piece, local) {
    const cos = Math.cos(piece.rotation);
    const sin = Math.sin(piece.rotation);
    return {
      x: piece.x + local[0] * cos - local[1] * sin,
      y: piece.y + local[0] * sin + local[1] * cos,
    };
  }

  function pointInPolygon(point, polygon) {
    const x = point[0];
    const y = point[1];
    let inside = false;
    for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i++) {
      const xi = polygon[i][0];
      const yi = polygon[i][1];
      const xj = polygon[j][0];
      const yj = polygon[j][1];
      const intersect =
        (yi > y) !== (yj > y) && x < ((xj - xi) * (y - yi)) / (yj - yi) + xi;
      if (intersect) {
        inside = !inside;
      }
    }
    return inside;
  }

  function drawContinents() {
    if (!ui.continentCtx) {
      return;
    }
    updateOverlayCanvasSize();
    ui.continentCtx.clearRect(0, 0, WIDTH, HEIGHT);

    state.continentPieces.forEach((piece) => {
      const isSelected = piece.id === state.selectedContinentId;
      drawContinentPiece(
        ui.continentCtx,
        piece,
        isSelected ? continentStrokeActive : continentStroke,
        isSelected
      );
      if (isSelected) {
        drawContinentNodes(ui.continentCtx, piece);
      }
    });

    drawContinentDraft(ui.continentCtx);
    updateContinentHint();
  }

  function drawContinentPiece(context, piece, strokeStyle, isSelected) {
    context.save();
    context.translate(piece.x, piece.y);
    context.rotate(piece.rotation);
    traceContinentPath(context, piece.points);
    context.fillStyle = continentFill;
    context.fill();
    context.lineWidth = isSelected ? 1.2 : 0.8;
    context.strokeStyle = strokeStyle;
    context.stroke();
    context.restore();
  }

  function drawContinentNodes(context, piece) {
    context.save();
    context.translate(piece.x, piece.y);
    context.rotate(piece.rotation);
    for (let i = 0; i < piece.points.length; i += 1) {
      const point = piece.points[i];
      const isHot =
        state.continentHoverNode &&
        state.continentHoverNode.id === piece.id &&
        state.continentHoverNode.index === i;
      context.beginPath();
      context.arc(point[0], point[1], nodeRadius, 0, Math.PI * 2);
      context.fillStyle = isHot ? draftPointActive : draftPoint;
      context.fill();
      context.lineWidth = 0.8;
      context.strokeStyle = isHot ? continentStrokeActive : continentStroke;
      context.stroke();
    }
    context.restore();
  }

  function drawContinentDraft(context) {
    if (state.continentDraft.length === 0) {
      return;
    }
    const points = state.continentDraft;
    const cursor = state.continentCursor;
    const first = points[0];
    const closeReady =
      cursor &&
      points.length >= 3 &&
      Math.hypot(cursor.x - first.x, cursor.y - first.y) <= closeDistance;

    context.save();
    context.strokeStyle = draftStroke;
    context.lineWidth = 0.7;
    context.beginPath();
    context.moveTo(points[0].x, points[0].y);
    for (let i = 1; i < points.length; i += 1) {
      context.lineTo(points[i].x, points[i].y);
    }
    if (cursor) {
      context.lineTo(cursor.x, cursor.y);
    }
    context.stroke();

    for (let i = 0; i < points.length; i += 1) {
      const point = points[i];
      context.beginPath();
      context.fillStyle =
        i === 0 && closeReady ? draftPointActive : draftPoint;
      context.arc(
        point.x,
        point.y,
        i === 0 ? draftPointRadius * 1.2 : draftPointRadius,
        0,
        Math.PI * 2
      );
      context.fill();
    }
    context.restore();
  }

  function updateContinentHint() {
    if (!ui.continentHint) {
      return;
    }
    if (state.tool !== "continent") {
      ui.continentHint.textContent = "Paint continent to add nodes";
      return;
    }
    if (state.continentDraft.length > 0) {
      ui.continentHint.textContent = "Click the first node again to close.";
      return;
    }
    if (state.selectedContinentId) {
      ui.continentHint.textContent =
        "Drag to move. Right-drag to rotate. Drag nodes to reshape.";
      return;
    }
    ui.continentHint.textContent =
      "Click to add nodes. Close by clicking the first node.";
  }

  function traceContinentPath(context, points) {
    if (!points.length) {
      return;
    }
    context.beginPath();
    context.moveTo(points[0][0], points[0][1]);
    for (let i = 1; i < points.length; i += 1) {
      context.lineTo(points[i][0], points[i][1]);
    }
    context.closePath();
  }

  init();
})();
