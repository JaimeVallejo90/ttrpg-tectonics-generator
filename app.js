(() => {
  const WIDTH = 297;
  const HEIGHT = 210;
  const GRID = 20;
  const TOTAL_CELLS = GRID * GRID;
  const MAX_ATTEMPTS = 1000;
  const MAX_PLATE_CONNECTIONS = 3;
  const CORNERS = [
    { id: "corner-tl-btn", col: 1, row: 1 },
    { id: "corner-tr-btn", col: GRID, row: 1 },
    { id: "corner-br-btn", col: GRID, row: GRID },
    { id: "corner-bl-btn", col: 1, row: GRID },
  ];

  const cellW = WIDTH / GRID;
  const cellH = HEIGHT / GRID;
  const innerLeft = 0;
  const innerTop = 0;
  const microW = cellW;
  const microH = cellH;
  const dotRadius = Math.min(microW, microH) * 0.12;
  const draftPointRadius = Math.min(microW, microH) * 0.14;
  const closeDistance = Math.min(microW, microH) * 0.6;
  const nodeRadius = Math.min(microW, microH) * 0.18;
  const nodeHitRadius = nodeRadius * 1.9;
  const plateHitRadius = Math.min(microW, microH) * 0.5;
  const continentFill = "rgba(194, 99, 42, 0.4)";
  const continentStroke = "rgba(27, 26, 23, 0.85)";
  const continentStrokeActive = "rgba(194, 99, 42, 0.95)";
  const draftStroke = "rgba(194, 99, 42, 0.8)";
  const draftPoint = "rgba(194, 99, 42, 0.9)";
  const draftPointActive = "rgba(79, 111, 106, 0.9)";
  const markerHitRadius = Math.min(microW, microH) * 0.45;
  const markerIconSize = Math.min(microW, microH) * 0.75;
  const markerNumberSize = Math.min(microW, microH) * 0.6;
  const volcanoIconSize = Math.min(microW, microH) * 0.55;
  const crustIcons = {
    ocean: "\u{1F30A}",
    continent: "\u26F0\uFE0F",
  };
  const volcanoIcon = "\u{1F30B}";
  const directionIcons = {
    1: "\u2B06\uFE0F",
    2: "\u2197\uFE0F",
    3: "\u2198\uFE0F",
    4: "\u2B07\uFE0F",
    5: "\u2199\uFE0F",
    6: "\u2196\uFE0F",
  };
  const directionLabels = {
    1: "N",
    2: "NE",
    3: "SE",
    4: "S",
    5: "SW",
    6: "NW",
  };

  const state = {
    selectedCells: new Set(),
    history: [],
    future: [],
    lastRoll: null,
    tool: "roll",
    isDrawing: false,
    currentLine: null,
    lines: [],
    plateDraft: null,
    wrapSide: null,
    markers: [],
    selectedMarkerId: null,
    markerDrag: null,
    continentPieces: [],
    continentDraft: [],
    continentCursor: null,
    continentDrag: null,
    selectedContinentId: null,
    continentHoverNode: null,
    nextContinentId: 1,
    nextMarkerId: 1,
    clearConfirm: false,
  };

  const ui = {};
  let edgeLayoutFrame = null;
  let lastEdgeLayout = null;

  function init() {
    ui.board = document.getElementById("board");
    ui.rollBtn = document.getElementById("roll-btn");
    ui.rollSixBtn = document.getElementById("roll-6-btn");
    ui.pencilBtn = document.getElementById("pencil-btn");
    ui.arrowBtn = document.getElementById("arrow-btn");
    ui.crustBtn = document.getElementById("crust-btn");
    ui.plateIdBtn = document.getElementById("plate-id-btn");
    ui.directionBtn = document.getElementById("direction-btn");
    ui.continentBtn = document.getElementById("continent-btn");
    ui.volcanoBtn = document.getElementById("volcano-btn");
    ui.deleteContinentBtn = document.getElementById("delete-continent-btn");
    ui.clearBtn = document.getElementById("clear-btn");
    ui.undoBtn = document.getElementById("undo-btn");
    ui.redoBtn = document.getElementById("redo-btn");
    ui.clearPointsBtn = document.getElementById("clear-points-btn");
    ui.clearArrowsBtn = document.getElementById("clear-arrows-btn");
    ui.clearCrustBtn = document.getElementById("clear-crust-btn");
    ui.clearDirectionsBtn = document.getElementById("clear-directions-btn");
    ui.clearPlatesBtn = document.getElementById("clear-plates-btn");
    ui.clearContinentsBtn = document.getElementById("clear-continents-btn");
    ui.clearVolcanoesBtn = document.getElementById("clear-volcanoes-btn");
    ui.toolMode = document.getElementById("tool-mode");
    ui.lastSector = document.getElementById("last-sector");
    ui.lastCell = document.getElementById("last-cell");
    ui.filled = document.getElementById("filled-count");
    ui.message = document.getElementById("message");
    ui.continentCanvas = document.getElementById("continent-canvas");
    ui.markerLayer = document.getElementById("marker-layer");
    ui.continentHint = document.getElementById("continent-hint");
    ui.header = document.querySelector(".app-header");
    ui.boardStack = document.querySelector(".board-stack");
    ui.edgeControls = document.getElementById("edge-controls");
    ui.wrapLeftBtn = document.getElementById("wrap-left-btn");
    ui.wrapRightBtn = document.getElementById("wrap-right-btn");
    ui.cornerButtons = CORNERS.map((corner) => ({
      ...corner,
      el: document.getElementById(corner.id),
    }));
    ui.continentCtx = ui.continentCanvas
      ? ui.continentCanvas.getContext("2d")
      : null;

    ui.board.setAttribute("viewBox", `0 0 ${WIDTH} ${HEIGHT}`);
    ui.board.setAttribute("preserveAspectRatio", "xMidYMid meet");
    if (ui.markerLayer) {
      ui.markerLayer.setAttribute("viewBox", `0 0 ${WIDTH} ${HEIGHT}`);
      ui.markerLayer.setAttribute("preserveAspectRatio", "xMidYMid meet");
    }
    if (ui.continentCanvas && ui.continentCtx) {
      updateOverlayCanvasSize();
    }

    ui.rollBtn.addEventListener("click", handleRoll);
    ui.rollSixBtn.addEventListener("click", handleRollSix);
    ui.pencilBtn.addEventListener("click", togglePencil);
    ui.arrowBtn.addEventListener("click", toggleArrow);
    ui.crustBtn.addEventListener("click", toggleCrust);
    ui.plateIdBtn.addEventListener("click", togglePlateId);
    ui.directionBtn.addEventListener("click", toggleDirection);
    ui.continentBtn.addEventListener("click", toggleContinent);
    ui.volcanoBtn.addEventListener("click", toggleVolcano);
    ui.deleteContinentBtn.addEventListener("click", handleDeleteContinent);
    ui.clearBtn.addEventListener("click", handleClearAll);
    ui.undoBtn.addEventListener("click", handleUndo);
    ui.redoBtn.addEventListener("click", handleRedo);
    ui.clearPointsBtn.addEventListener("click", handleClearPoints);
    ui.clearArrowsBtn.addEventListener("click", handleClearArrows);
    ui.clearCrustBtn.addEventListener("click", handleClearCrust);
    ui.clearDirectionsBtn.addEventListener("click", handleClearDirections);
    ui.clearPlatesBtn.addEventListener("click", handleClearPlates);
    ui.clearContinentsBtn.addEventListener("click", handleClearContinents);
    ui.clearVolcanoesBtn.addEventListener("click", handleClearVolcanoes);
    ui.board.addEventListener("pointerdown", handlePointerDown);
    ui.board.addEventListener("pointermove", handlePointerMove);
    ui.board.addEventListener("pointerup", handlePointerUp);
    ui.board.addEventListener("pointerleave", handlePointerUp);
    ui.board.addEventListener("pointercancel", handlePointerUp);
    ui.board.addEventListener("contextmenu", handleContextMenu);
    if (ui.wrapLeftBtn) {
      ui.wrapLeftBtn.addEventListener("click", () => toggleWrapSide("left"));
    }
    if (ui.wrapRightBtn) {
      ui.wrapRightBtn.addEventListener("click", () => toggleWrapSide("right"));
    }
    if (ui.cornerButtons) {
      ui.cornerButtons.forEach((corner) => {
        if (!corner.el) {
          return;
        }
        corner.el.addEventListener("click", () =>
          toggleCornerPoint(corner.col, corner.row)
        );
      });
    }
    if (ui.continentCanvas) {
      ui.continentCanvas.addEventListener("pointerdown", handleContinentPointerDown);
      ui.continentCanvas.addEventListener("pointermove", handleContinentPointerMove);
      ui.continentCanvas.addEventListener("pointerup", handleContinentPointerUp);
      ui.continentCanvas.addEventListener("pointerleave", handleContinentPointerUp);
      ui.continentCanvas.addEventListener("pointercancel", handleContinentPointerUp);
      ui.continentCanvas.addEventListener("contextmenu", handleContinentContextMenu);
    }
    if (ui.markerLayer) {
      ui.markerLayer.addEventListener("pointerdown", handleMarkerPointerDown);
      ui.markerLayer.addEventListener("pointermove", handleMarkerPointerMove);
      ui.markerLayer.addEventListener("pointerup", handleMarkerPointerUp);
      ui.markerLayer.addEventListener("pointerleave", handleMarkerPointerUp);
      ui.markerLayer.addEventListener("pointercancel", handleMarkerPointerUp);
      ui.markerLayer.addEventListener("contextmenu", handleMarkerContextMenu);
    }
    window.addEventListener("resize", handleResize);
    window.addEventListener("keydown", handleKeydown);

    updateHeaderOffset();
    render();
  }

  function recordHistory(entry) {
    state.history.push(entry);
    state.future = [];
  }

  function rollD6() {
    return Math.floor(Math.random() * 6) + 1;
  }

  function rollD20() {
    return Math.floor(Math.random() * 20) + 1;
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
      setMessage(`Marked (${result.col}, ${result.row}).`);
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
      const lastText = lastResult ? ` Last (${lastResult.col}, ${lastResult.row}).` : "";
      setMessage(`Marked ${added} points.${lastText}`);
    }

    render();
  }

  function rollRandomPoint() {
    for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt += 1) {
      const col = rollD20();
      const row = rollD20();
      const microCol = col - 1;
      const microRow = row - 1;

      if (addCell(microCol, microRow, col, row)) {
        return { col, row };
      }
    }
    return null;
  }

  function isLineTool(tool) {
    return tool === "pencil" || tool === "arrow";
  }

  function isMarkerTool(tool) {
    return tool === "crust" || tool === "plate-id" || tool === "direction" || tool === "volcano";
  }

  function setTool(nextTool) {
    cancelClearConfirm();
    const previousTool = state.tool;
    const target = state.tool === nextTool ? "roll" : nextTool;
    state.tool = target;
    if (previousTool !== state.tool) {
      state.isDrawing = false;
      state.currentLine = null;
      state.plateDraft = null;
      state.wrapSide = null;
    }
    if (!isMarkerTool(state.tool)) {
      state.markerDrag = null;
      if (isMarkerTool(previousTool)) {
        state.selectedMarkerId = null;
      }
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

  function toggleArrow() {
    setTool("arrow");
  }

  function toggleCrust() {
    setTool("crust");
  }

  function togglePlateId() {
    setTool("plate-id");
  }

  function toggleDirection() {
    setTool("direction");
  }

  function toggleContinent() {
    setTool("continent");
  }

  function toggleVolcano() {
    setTool("volcano");
  }

  function toggleWrapSide(side) {
    if (state.tool !== "pencil") {
      setMessage("Select Paint plates to use turn around.");
      render();
      return;
    }
    if (state.wrapSide === side) {
      state.wrapSide = null;
      setMessage("Turn around cleared.");
    } else {
      state.wrapSide = side;
      setMessage(`Next connection will turn around on the ${side}.`);
    }
    render();
  }

  function clearPlateDraft() {
    state.plateDraft = null;
    if (state.currentLine && (state.currentLine.lineType || "plate") === "plate") {
      state.currentLine = null;
    }
  }

  function getPointCenter(col, row) {
    return {
      x: innerLeft + col * microW + microW / 2,
      y: innerTop + row * microH + microH / 2,
    };
  }

  function buildPlatePoint(hit) {
    const center = getPointCenter(hit.col, hit.row);
    return {
      key: hit.key,
      col: hit.col,
      row: hit.row,
      x: center.x,
      y: center.y,
    };
  }

  function countPlateConnections(key) {
    let count = 0;
    state.lines.forEach((line) => {
      if ((line.lineType || "plate") !== "plate") {
        return;
      }
      if (line.startKey === key || line.endKey === key) {
        count += 1;
      }
    });
    return count;
  }

  function hasPlateConnection(keyA, keyB) {
    return state.lines.some((line) => {
      if ((line.lineType || "plate") !== "plate") {
        return false;
      }
      if (!line.startKey || !line.endKey) {
        return false;
      }
      return (
        (line.startKey === keyA && line.endKey === keyB) ||
        (line.startKey === keyB && line.endKey === keyA)
      );
    });
  }

  function addPlateConnection(start, end, wrapSide) {
    const line = {
      x1: start.x,
      y1: start.y,
      x2: end.x,
      y2: end.y,
      lineType: "plate",
      startKey: start.key,
      endKey: end.key,
      wrapSide: wrapSide || null,
    };
    state.lines.push(line);
    recordHistory({
      type: "line-add",
      line: { ...line },
      index: state.lines.length - 1,
    });
  }

  function getWrappedLineSegments(x1, y1, x2, y2, wrapSide) {
    if (!wrapSide) {
      return [{ x1, y1, x2, y2 }];
    }
    const wrapEdgeX = wrapSide === "left" ? 0 : WIDTH;
    const adjustedX2 = wrapSide === "left" ? x2 - WIDTH : x2 + WIDTH;
    const dx = adjustedX2 - x1;
    if (dx === 0) {
      return [{ x1, y1, x2, y2 }];
    }
    const t = (wrapEdgeX - x1) / dx;
    const yAtWrap = y1 + t * (y2 - y1);
    const otherX = wrapEdgeX === 0 ? WIDTH : 0;
    return [
      { x1, y1, x2: wrapEdgeX, y2: yAtWrap },
      { x1: otherX, y1: yAtWrap, x2, y2 },
    ];
  }

  function getLineSegments(line) {
    const lineType = line.lineType || "plate";
    if (lineType !== "plate") {
      return [line];
    }
    return getWrappedLineSegments(
      line.x1,
      line.y1,
      line.x2,
      line.y2,
      line.wrapSide
    );
  }

  function hasMarks() {
    return (
      state.selectedCells.size > 0 ||
      state.lines.length > 0 ||
      state.markers.length > 0 ||
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
      setMessage("Nothing to restart.");
      render();
      return;
    }
    if (!state.clearConfirm) {
      state.clearConfirm = true;
      setMessage("Click Restart again to confirm.");
      render();
      return;
    }
    const snapshot = createClearSnapshot();
    recordHistory({ type: "clear", snapshot });
    clearAllState();
    state.clearConfirm = false;
    setMessage("Restarted.");
    render();
  }

  function handleUndo() {
    cancelClearConfirm();
    undoLast();
  }

  function handleRedo() {
    cancelClearConfirm();
    redoLast();
  }

  function handleClearPoints() {
    cancelClearConfirm();
    clearPoints();
  }

  function handleClearArrows() {
    cancelClearConfirm();
    clearLinesByType("arrow", "Arrows");
  }

  function handleClearCrust() {
    cancelClearConfirm();
    clearMarkersByType("crust", "Crust markers");
  }

  function handleClearDirections() {
    cancelClearConfirm();
    clearMarkersByType("direction", "Directions");
  }

  function handleClearPlates() {
    cancelClearConfirm();
    clearMarkersByType("plate-id", "Plate numbers");
  }

  function handleClearContinents() {
    cancelClearConfirm();
    clearContinents();
  }

  function handleClearVolcanoes() {
    cancelClearConfirm();
    clearMarkersByType("volcano", "Volcanoes");
  }

  function clearPoints() {
    if (state.selectedCells.size === 0) {
      setMessage("No points to clear.");
      render();
      return;
    }
    clearPlateDraft();
    const points = Array.from(state.selectedCells);
    state.selectedCells.clear();
    state.lastRoll = null;
    recordHistory({ type: "bulk-clear", itemType: "points", points });
    setMessage("Points cleared.");
    render();
  }

  function clearLinesByType(lineType, label) {
    const removed = [];
    const remaining = [];
    state.lines.forEach((line, index) => {
      const currentType = line.lineType || "plate";
      if (currentType === lineType) {
        removed.push({ line: { ...line }, index });
      } else {
        remaining.push(line);
      }
    });
    if (removed.length === 0) {
      setMessage(`No ${label.toLowerCase()} to clear.`);
      render();
      return;
    }
    state.lines = remaining;
    if (state.currentLine && (state.currentLine.lineType || "plate") === lineType) {
      state.currentLine = null;
      state.isDrawing = false;
    }
    recordHistory({
      type: "bulk-clear",
      itemType: "lines",
      lineType,
      items: removed,
    });
    setMessage(`${label} cleared.`);
    render();
  }

  function clearMarkersByType(markerType, label) {
    const removed = [];
    const removedIds = new Set();
    state.markers.forEach((marker, index) => {
      if (marker.type === markerType) {
        removed.push({ marker: cloneMarker(marker), index });
        removedIds.add(marker.id);
      }
    });
    if (removed.length === 0) {
      setMessage(`No ${label.toLowerCase()} to clear.`);
      render();
      return;
    }
    const selectedMarkerId = state.selectedMarkerId;
    state.markers = state.markers.filter((marker) => !removedIds.has(marker.id));
    if (selectedMarkerId && removedIds.has(selectedMarkerId)) {
      state.selectedMarkerId = null;
    }
    if (state.markerDrag && removedIds.has(state.markerDrag.id)) {
      state.markerDrag = null;
    }
    recordHistory({
      type: "bulk-clear",
      itemType: "markers",
      markerType,
      items: removed,
      selectedMarkerId: selectedMarkerId && removedIds.has(selectedMarkerId)
        ? selectedMarkerId
        : null,
    });
    setMessage(`${label} cleared.`);
    render();
  }

  function clearContinents() {
    const hasDraft = state.continentDraft.length > 0;
    if (state.continentPieces.length === 0 && !hasDraft) {
      setMessage("No continents to clear.");
      render();
      return;
    }
    const removed = state.continentPieces.map((piece, index) => ({
      piece: cloneContinentPiece(piece),
      index,
    }));
    const selectedContinentId = state.selectedContinentId;
    state.continentPieces = [];
    state.continentDraft = [];
    state.continentCursor = null;
    state.continentDrag = null;
    state.continentHoverNode = null;
    state.selectedContinentId = null;
    if (removed.length > 0) {
      recordHistory({
        type: "bulk-clear",
        itemType: "continents",
        items: removed,
        selectedContinentId,
      });
    }
    setMessage("Continents cleared.");
    render();
  }

  function clearAllState() {
    state.selectedCells.clear();
    state.lastRoll = null;
    state.currentLine = null;
    state.isDrawing = false;
    state.lines = [];
    state.plateDraft = null;
    state.wrapSide = null;
    state.markers = [];
    state.selectedMarkerId = null;
    state.markerDrag = null;
    state.continentPieces = [];
    state.continentDraft = [];
    state.continentCursor = null;
    state.continentDrag = null;
    state.selectedContinentId = null;
    state.continentHoverNode = null;
    state.nextContinentId = 1;
    state.nextMarkerId = 1;
  }

  function createClearSnapshot() {
    return {
      selectedCells: Array.from(state.selectedCells),
      lines: state.lines.map((line) => ({ ...line })),
      lastRoll: state.lastRoll ? { ...state.lastRoll } : null,
      markers: state.markers.map(cloneMarker),
      selectedMarkerId: state.selectedMarkerId,
      continentPieces: state.continentPieces.map(cloneContinentPiece),
      selectedContinentId: state.selectedContinentId,
      nextContinentId: state.nextContinentId,
      nextMarkerId: state.nextMarkerId,
    };
  }

  function restoreClearSnapshot(snapshot) {
    state.selectedCells = new Set(snapshot.selectedCells);
    state.lines = snapshot.lines.map((line) => ({ ...line }));
    state.lastRoll = snapshot.lastRoll ? { ...snapshot.lastRoll } : null;
    state.markers = snapshot.markers.map(cloneMarker);
    state.selectedMarkerId = snapshot.selectedMarkerId;
    state.continentPieces = snapshot.continentPieces.map(cloneContinentPiece);
    state.selectedContinentId = snapshot.selectedContinentId;
    state.nextContinentId = snapshot.nextContinentId;
    state.nextMarkerId = snapshot.nextMarkerId;
    state.currentLine = null;
    state.isDrawing = false;
    state.plateDraft = null;
    state.wrapSide = null;
    state.markerDrag = null;
    state.continentDraft = [];
    state.continentCursor = null;
    state.continentDrag = null;
    state.continentHoverNode = null;
    if (
      state.selectedMarkerId &&
      !state.markers.some((marker) => marker.id === state.selectedMarkerId)
    ) {
      state.selectedMarkerId = null;
    }
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

  function handlePlatePointerDown(event) {
    if (event.button !== 0) {
      return;
    }
    cancelClearConfirm();
    if (state.selectedCells.size === 0) {
      setMessage("Roll points first.");
      render();
      return;
    }
    const point = getSvgPointInPage(event, false);
    if (!point) {
      return;
    }
    const hit = findNearestPoint(point.x, point.y, plateHitRadius);
    if (!hit) {
      if (state.plateDraft) {
        clearPlateDraft();
        setMessage("Chain ended.");
        render();
      }
      return;
    }

    const hitPoint = buildPlatePoint(hit);

    if (!state.plateDraft) {
      if (countPlateConnections(hitPoint.key) >= MAX_PLATE_CONNECTIONS) {
        setMessage(
          `This point already has ${MAX_PLATE_CONNECTIONS} connections.`
        );
        render();
        return;
      }
      state.plateDraft = hitPoint;
      setMessage("Select another point to connect.");
      render();
      return;
    }

    if (hitPoint.key === state.plateDraft.key) {
      clearPlateDraft();
      setMessage("Chain ended.");
      render();
      return;
    }

    if (countPlateConnections(state.plateDraft.key) >= MAX_PLATE_CONNECTIONS) {
      clearPlateDraft();
      setMessage(
        `This point already has ${MAX_PLATE_CONNECTIONS} connections.`
      );
      render();
      return;
    }

    if (countPlateConnections(hitPoint.key) >= MAX_PLATE_CONNECTIONS) {
      setMessage(
        `That point already has ${MAX_PLATE_CONNECTIONS} connections.`
      );
      render();
      return;
    }

    if (hasPlateConnection(state.plateDraft.key, hitPoint.key)) {
      setMessage("Points already connected.");
      render();
      return;
    }

    const wrapSide = state.wrapSide;
    addPlateConnection(state.plateDraft, hitPoint, wrapSide);
    if (wrapSide) {
      state.wrapSide = null;
    }
    if (countPlateConnections(hitPoint.key) >= MAX_PLATE_CONNECTIONS) {
      clearPlateDraft();
      setMessage("Line drawn. This point reached max connections.");
      render();
      return;
    }
    state.plateDraft = hitPoint;
    setMessage("Line drawn. Select the next point.");
    render();
  }

  function handlePointerDown(event) {
    if (state.tool === "pencil") {
      handlePlatePointerDown(event);
      return;
    }
    if (state.tool !== "arrow" || event.button !== 0) {
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
      lineType: "arrow",
    };
    if (ui.board.setPointerCapture) {
      ui.board.setPointerCapture(event.pointerId);
    }
    render();
  }

  function handlePointerMove(event) {
    if (state.tool !== "arrow" || !state.isDrawing) {
      return;
    }
    updateCurrentLine(event);
  }

  function handlePointerUp(event) {
    if (state.tool !== "arrow" || !state.isDrawing) {
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
    if (state.tool === "pencil" && state.plateDraft) {
      clearPlateDraft();
      setMessage("Chain ended.");
      render();
      return;
    }
    if (state.isDrawing) {
      state.isDrawing = false;
      state.currentLine = null;
    }
    const point = getSvgPointRaw(event);
    if (!point) {
      return;
    }
    const removed = removeAtPoint(point.x, point.y, { allowPoints: false });
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
      if (deleteSelectedMarker()) {
        event.preventDefault();
        render();
        return;
      }
      if (deleteSelectedContinent()) {
        event.preventDefault();
        render();
      }
      return;
    }
    const hasModifier = event.ctrlKey || event.metaKey;
    if (hasModifier && (key === "y" || (key === "z" && event.shiftKey))) {
      if (isEditableTarget(event.target)) {
        return;
      }
      cancelClearConfirm();
      event.preventDefault();
      redoLast();
      return;
    }
    if (hasModifier && key === "z") {
      if (isEditableTarget(event.target)) {
        return;
      }
      cancelClearConfirm();
      event.preventDefault();
      undoLast();
    }
  }

  function handleMarkerPointerDown(event) {
    if (!isMarkerTool(state.tool)) {
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
      const removed = removeMarkerAt(point);
      if (!removed) {
        setMessage("Nothing to delete.");
      }
      render();
      event.preventDefault();
      return;
    }

    const hit = findMarkerAt(point);
    if (hit) {
      selectMarker(hit.id);
      state.markerDrag = {
        id: hit.id,
        offsetX: point.x - hit.x,
        offsetY: point.y - hit.y,
        before: cloneMarker(hit),
      };
      if (ui.markerLayer && ui.markerLayer.setPointerCapture) {
        ui.markerLayer.setPointerCapture(event.pointerId);
      }
      if (ui.markerLayer) {
        ui.markerLayer.style.cursor = "grabbing";
      }
      render();
      return;
    }

    const marker = createMarkerFromTool(point);
    if (!marker) {
      return;
    }
    state.markers.push(marker);
    recordHistory({
      type: "marker-add",
      marker: cloneMarker(marker),
      index: state.markers.length - 1,
    });
    selectMarker(marker.id);
    render();
  }

  function handleMarkerPointerMove(event) {
    if (!isMarkerTool(state.tool)) {
      return;
    }
    const point = getSvgPointInPage(event, true);
    if (!point) {
      return;
    }

    if (state.markerDrag) {
      const marker = getMarkerById(state.markerDrag.id);
      if (!marker) {
        state.markerDrag = null;
        render();
        return;
      }
      marker.x = clamp(point.x - state.markerDrag.offsetX, 0, WIDTH);
      marker.y = clamp(point.y - state.markerDrag.offsetY, 0, HEIGHT);
      render();
      return;
    }

    const hit = findMarkerAt(point);
    if (ui.markerLayer) {
      ui.markerLayer.style.cursor = hit ? "grab" : "crosshair";
    }
  }

  function handleMarkerPointerUp(event) {
    if (!isMarkerTool(state.tool)) {
      return;
    }
    if (state.markerDrag) {
      commitMarkerDrag();
      state.markerDrag = null;
      if (ui.markerLayer && ui.markerLayer.releasePointerCapture) {
        try {
          ui.markerLayer.releasePointerCapture(event.pointerId);
        } catch (error) {
          // Ignore if pointer capture was not set.
        }
      }
    }
    if (ui.markerLayer) {
      ui.markerLayer.style.cursor = "crosshair";
    }
    render();
  }

  function handleMarkerContextMenu(event) {
    if (isMarkerTool(state.tool)) {
      event.preventDefault();
    }
  }

  function commitMarkerDrag() {
    if (!state.markerDrag || !state.markerDrag.before) {
      return;
    }
    const marker = getMarkerById(state.markerDrag.id);
    if (!marker) {
      return;
    }
    if (!hasMarkerChanged(state.markerDrag.before, marker)) {
      return;
    }
    recordHistory({
      type: "marker-update",
      id: marker.id,
      before: state.markerDrag.before,
      after: cloneMarker(marker),
    });
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
    const lineType = state.currentLine.lineType || "plate";
    const dx = x2 - x1;
    const dy = y2 - y1;
    const length = Math.hypot(dx, dy);
    const minLength =
      Math.min(microW, microH) * (lineType === "arrow" ? 0.3 : 0.2);
    if (length >= minLength) {
      const line = { x1, y1, x2, y2, lineType };
      state.lines.push(line);
      recordHistory({
        type: "line-add",
        line: { ...line },
        index: state.lines.length - 1,
      });
      setMessage(lineType === "arrow" ? "Arrow drawn." : "Line drawn.");
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

  function removeAtPoint(x, y, options = {}) {
    const { allowPoints = true } = options;
    const pointHit = Math.min(microW, microH) * 0.35;
    const lineHit = Math.min(microW, microH) * 0.35;
    const markerCandidate = findMarkerAt({ x, y });
    const pointCandidate = allowPoints ? findNearestPoint(x, y, pointHit) : null;
    const lineCandidate = findNearestLine(x, y, lineHit);

    if (markerCandidate) {
      deleteMarkerById(markerCandidate.id);
      return true;
    }

    if (lineCandidate) {
      removeLine(lineCandidate.index);
      return true;
    }

    if (pointCandidate) {
      removePoint(pointCandidate.key, pointCandidate.col, pointCandidate.row);
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
      const segments = getLineSegments(line);
      let distSq = Infinity;
      segments.forEach((segment) => {
        const candidate = distanceToSegmentSquared(x, y, segment);
        if (candidate < distSq) {
          distSq = candidate;
        }
      });
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
    const col = microCol + 1;
    const row = microRow + 1;
    recordHistory({
      type: "erase",
      itemType: "point",
      key,
      roll: { col, row },
    });
    syncLastPoint();
    setMessage("Point deleted.");
  }

  function removeLine(index) {
    const [line] = state.lines.splice(index, 1);
    if (!line) {
      return;
    }
    recordHistory({
      type: "erase",
      itemType: "line",
      line,
      index,
    });
    setMessage("Line deleted.");
  }

  function createMarkerFromTool(point) {
    if (state.tool === "crust") {
      const roll = rollD6();
      const crustType = roll <= 3 ? "ocean" : "continent";
      setMessage(
        crustType === "ocean" ? "Crust: oceanic." : "Crust: continental."
      );
      return createMarker("crust", point, crustType);
    }
    if (state.tool === "direction") {
      const roll = rollD6();
      const label = directionLabels[roll] || "?";
      setMessage(`Direction: ${label}.`);
      return createMarker("direction", point, roll);
    }
    if (state.tool === "plate-id") {
      const number = getNextPlateNumber();
      setMessage(`Plate ${number}.`);
      return createMarker("plate-id", point, number);
    }
    if (state.tool === "volcano") {
      setMessage("Volcano placed.");
      return createMarker("volcano", point, "volcano");
    }
    return null;
  }

  function createMarker(type, point, value) {
    return {
      id: `marker-${state.nextMarkerId++}`,
      type,
      x: point.x,
      y: point.y,
      value,
    };
  }

  function getNextPlateNumber() {
    const used = new Set();
    state.markers.forEach((marker) => {
      if (marker.type === "plate-id") {
        used.add(marker.value);
      }
    });
    let number = 1;
    while (used.has(number)) {
      number += 1;
    }
    return number;
  }

  function getMarkerLabel(marker) {
    if (marker.type === "plate-id") {
      return String(marker.value);
    }
    if (marker.type === "crust") {
      return crustIcons[marker.value] || "";
    }
    if (marker.type === "direction") {
      return directionIcons[marker.value] || "";
    }
    if (marker.type === "volcano") {
      return volcanoIcon;
    }
    return "";
  }

  function getMarkerFontSize(marker) {
    if (marker.type === "plate-id") {
      return markerNumberSize;
    }
    if (marker.type === "volcano") {
      return volcanoIconSize;
    }
    return markerIconSize;
  }

  function findMarkerAt(point) {
    for (let i = state.markers.length - 1; i >= 0; i -= 1) {
      const marker = state.markers[i];
      const dx = point.x - marker.x;
      const dy = point.y - marker.y;
      if (dx * dx + dy * dy <= markerHitRadius * markerHitRadius) {
        return marker;
      }
    }
    return null;
  }

  function removeMarkerAt(point) {
    const hit = findMarkerAt(point);
    if (!hit) {
      return false;
    }
    return deleteMarkerById(hit.id);
  }

  function deleteMarkerById(id, shouldRecordHistory = true) {
    const index = state.markers.findIndex((marker) => marker.id === id);
    if (index < 0) {
      return false;
    }
    const [removed] = state.markers.splice(index, 1);
    if (shouldRecordHistory && removed) {
      recordHistory({
        type: "marker-delete",
        marker: cloneMarker(removed),
        index,
      });
    }
    if (state.selectedMarkerId === id) {
      state.selectedMarkerId = null;
    }
    if (state.markerDrag && state.markerDrag.id === id) {
      state.markerDrag = null;
    }
    setMessage("Marker deleted.");
    return true;
  }

  function deleteSelectedMarker(shouldRecordHistory = true) {
    if (!state.selectedMarkerId) {
      return false;
    }
    return deleteMarkerById(state.selectedMarkerId, shouldRecordHistory);
  }

  function selectMarker(id) {
    state.selectedMarkerId = id;
    state.selectedContinentId = null;
    state.continentHoverNode = null;
    bringMarkerToFront(id);
  }

  function bringMarkerToFront(id) {
    const index = state.markers.findIndex((marker) => marker.id === id);
    if (index < 0 || index === state.markers.length - 1) {
      return;
    }
    const [marker] = state.markers.splice(index, 1);
    state.markers.push(marker);
  }

  function cloneMarker(marker) {
    return { ...marker };
  }

  function applyMarkerSnapshot(target, snapshot) {
    target.type = snapshot.type;
    target.x = snapshot.x;
    target.y = snapshot.y;
    target.value = snapshot.value;
  }

  function hasMarkerChanged(before, marker) {
    const epsilon = 0.01;
    if (before.type !== marker.type || before.value !== marker.value) {
      return true;
    }
    return (
      Math.abs(before.x - marker.x) > epsilon ||
      Math.abs(before.y - marker.y) > epsilon
    );
  }

  function getMarkerById(id) {
    return state.markers.find((marker) => marker.id === id);
  }

  function addCell(microCol, microRow, col, row) {
    const key = `${microCol},${microRow}`;
    if (state.selectedCells.has(key)) {
      return false;
    }
    state.selectedCells.add(key);
    const roll = { col, row };
    recordHistory({ type: "point", key, roll });
    state.lastRoll = roll;
    return true;
  }

  function toggleCornerPoint(col, row) {
    cancelClearConfirm();
    const microCol = col - 1;
    const microRow = row - 1;
    const key = `${microCol},${microRow}`;
    if (state.selectedCells.has(key)) {
      if (state.plateDraft && state.plateDraft.key === key) {
        clearPlateDraft();
      }
      removePoint(key, microCol, microRow);
      setMessage(`Corner (${col}, ${row}) removed.`);
    } else {
      addCell(microCol, microRow, col, row);
      setMessage(`Corner (${col}, ${row}) added.`);
    }
    render();
  }

  function undoLast() {
    const last = state.history.pop();
    if (!last) {
      setMessage("Nothing to undo.");
      render();
      return;
    }
    state.clearConfirm = false;
    applyUndoEntry(last);
    state.future.push(last);
    syncLastPoint();
    setMessage("Undo.");
    render();
  }

  function redoLast() {
    const next = state.future.pop();
    if (!next) {
      setMessage("Nothing to redo.");
      render();
      return;
    }
    state.clearConfirm = false;
    applyRedoEntry(next);
    state.history.push(next);
    syncLastPoint();
    setMessage("Redo.");
    render();
  }

  function applyUndoEntry(entry) {
    if (entry.type === "point") {
      state.selectedCells.delete(entry.key);
      return;
    }
    if (entry.type === "line-add") {
      removeLineBySnapshot(entry);
      return;
    }
    if (entry.type === "line") {
      state.lines.pop();
      return;
    }
    if (entry.type === "erase") {
      if (entry.itemType === "point") {
        state.selectedCells.add(entry.key);
        return;
      }
      if (entry.itemType === "line") {
        insertLineAt(entry.index, entry.line);
      }
      return;
    }
    if (entry.type === "marker-add") {
      undoMarkerAdd(entry);
      return;
    }
    if (entry.type === "marker-delete") {
      undoMarkerDelete(entry);
      return;
    }
    if (entry.type === "marker-update") {
      undoMarkerUpdate(entry);
      return;
    }
    if (entry.type === "continent-add") {
      undoContinentAdd(entry);
      return;
    }
    if (entry.type === "continent-delete") {
      undoContinentDelete(entry);
      return;
    }
    if (entry.type === "continent-update") {
      undoContinentUpdate(entry);
      return;
    }
    if (entry.type === "clear") {
      restoreClearSnapshot(entry.snapshot);
      return;
    }
    if (entry.type === "bulk-clear") {
      undoBulkClear(entry);
    }
  }

  function applyRedoEntry(entry) {
    if (entry.type === "point") {
      state.selectedCells.add(entry.key);
      return;
    }
    if (entry.type === "line-add") {
      insertLineAt(entry.index, entry.line);
      return;
    }
    if (entry.type === "line") {
      return;
    }
    if (entry.type === "erase") {
      if (entry.itemType === "point") {
        state.selectedCells.delete(entry.key);
        return;
      }
      if (entry.itemType === "line") {
        removeLineBySnapshot(entry);
      }
      return;
    }
    if (entry.type === "marker-add") {
      redoMarkerAdd(entry);
      return;
    }
    if (entry.type === "marker-delete") {
      redoMarkerDelete(entry);
      return;
    }
    if (entry.type === "marker-update") {
      redoMarkerUpdate(entry);
      return;
    }
    if (entry.type === "continent-add") {
      redoContinentAdd(entry);
      return;
    }
    if (entry.type === "continent-delete") {
      redoContinentDelete(entry);
      return;
    }
    if (entry.type === "continent-update") {
      redoContinentUpdate(entry);
      return;
    }
    if (entry.type === "clear") {
      clearAllState();
      return;
    }
    if (entry.type === "bulk-clear") {
      redoBulkClear(entry);
    }
  }

  function undoBulkClear(entry) {
    if (entry.itemType === "points") {
      entry.points.forEach((key) => state.selectedCells.add(key));
      return;
    }
    if (entry.itemType === "lines") {
      restoreLines(entry.items);
      return;
    }
    if (entry.itemType === "markers") {
      restoreMarkers(entry.items);
      if (entry.selectedMarkerId) {
        state.selectedMarkerId = entry.selectedMarkerId;
      }
      return;
    }
    if (entry.itemType === "continents") {
      restoreContinents(entry.items);
      if (entry.selectedContinentId) {
        state.selectedContinentId = entry.selectedContinentId;
      }
    }
  }

  function redoBulkClear(entry) {
    if (entry.itemType === "points") {
      entry.points.forEach((key) => state.selectedCells.delete(key));
      return;
    }
    if (entry.itemType === "lines") {
      removeLines(entry.items);
      return;
    }
    if (entry.itemType === "markers") {
      removeMarkers(entry.items);
      return;
    }
    if (entry.itemType === "continents") {
      removeContinents(entry.items);
    }
  }

  function restoreLines(items) {
    const sorted = [...items].sort((a, b) => a.index - b.index);
    sorted.forEach((item) => insertLineAt(item.index, item.line));
  }

  function removeLines(items) {
    const sorted = [...items].sort((a, b) => b.index - a.index);
    sorted.forEach((item) => removeLineBySnapshot(item));
  }

  function insertLineAt(index, line) {
    const insertAt = clamp(index, 0, state.lines.length);
    state.lines.splice(insertAt, 0, { ...line });
  }

  function removeLineBySnapshot(entry) {
    if (
      Number.isInteger(entry.index) &&
      entry.index >= 0 &&
      entry.index < state.lines.length
    ) {
      const candidate = state.lines[entry.index];
      if (candidate && linesMatch(candidate, entry.line)) {
        state.lines.splice(entry.index, 1);
        return;
      }
    }
    const matchIndex = state.lines.findIndex((line) => linesMatch(line, entry.line));
    if (matchIndex >= 0) {
      state.lines.splice(matchIndex, 1);
    }
  }

  function linesMatch(a, b) {
    if (!a || !b) {
      return false;
    }
    const epsilon = 0.0001;
    return (
      Math.abs(a.x1 - b.x1) <= epsilon &&
      Math.abs(a.y1 - b.y1) <= epsilon &&
      Math.abs(a.x2 - b.x2) <= epsilon &&
      Math.abs(a.y2 - b.y2) <= epsilon &&
      (a.lineType || "plate") === (b.lineType || "plate") &&
      (a.wrapSide || null) === (b.wrapSide || null)
    );
  }

  function restoreMarkers(items) {
    const sorted = [...items].sort((a, b) => a.index - b.index);
    sorted.forEach((item) => insertMarkerAt(item.index, item.marker));
  }

  function removeMarkers(items) {
    const ids = new Set(items.map((item) => item.marker.id));
    state.markers = state.markers.filter((marker) => !ids.has(marker.id));
    if (state.selectedMarkerId && ids.has(state.selectedMarkerId)) {
      state.selectedMarkerId = null;
    }
    if (state.markerDrag && ids.has(state.markerDrag.id)) {
      state.markerDrag = null;
    }
  }

  function restoreContinents(items) {
    const sorted = [...items].sort((a, b) => a.index - b.index);
    sorted.forEach((item) => insertContinentAt(item.index, item.piece));
  }

  function removeContinents(items) {
    const ids = new Set(items.map((item) => item.piece.id));
    state.continentPieces = state.continentPieces.filter(
      (piece) => !ids.has(piece.id)
    );
    if (state.selectedContinentId && ids.has(state.selectedContinentId)) {
      state.selectedContinentId = null;
    }
    if (state.continentDrag && ids.has(state.continentDrag.id)) {
      state.continentDrag = null;
    }
    state.continentHoverNode = null;
    state.continentDraft = [];
    state.continentCursor = null;
  }

  function insertMarkerAt(index, marker) {
    const insertAt = clamp(index, 0, state.markers.length);
    const restored = cloneMarker(marker);
    state.markers.splice(insertAt, 0, restored);
    return restored;
  }

  function redoMarkerAdd(entry) {
    const restored = insertMarkerAt(
      Number.isInteger(entry.index) ? entry.index : state.markers.length,
      entry.marker
    );
    selectMarker(restored.id);
  }

  function redoMarkerDelete(entry) {
    removeMarkers([{ marker: entry.marker, index: entry.index }]);
  }

  function redoMarkerUpdate(entry) {
    const marker = getMarkerById(entry.id);
    if (!marker) {
      return;
    }
    applyMarkerSnapshot(marker, entry.after);
  }

  function insertContinentAt(index, piece) {
    const insertAt = clamp(index, 0, state.continentPieces.length);
    const restored = cloneContinentPiece(piece);
    state.continentPieces.splice(insertAt, 0, restored);
    return restored;
  }

  function removeContinentById(id) {
    const index = state.continentPieces.findIndex((piece) => piece.id === id);
    if (index >= 0) {
      state.continentPieces.splice(index, 1);
    }
    if (state.selectedContinentId === id) {
      state.selectedContinentId = null;
    }
    state.continentHoverNode = null;
  }

  function redoContinentAdd(entry) {
    const restored = insertContinentAt(
      Number.isInteger(entry.index) ? entry.index : state.continentPieces.length,
      entry.piece
    );
    selectContinent(restored.id);
  }

  function redoContinentDelete(entry) {
    removeContinentById(entry.piece.id);
  }

  function redoContinentUpdate(entry) {
    const piece = getContinentById(entry.id);
    if (!piece) {
      return;
    }
    applyContinentSnapshot(piece, entry.after);
  }

  function syncLastPoint() {
    for (let i = state.history.length - 1; i >= 0; i -= 1) {
      const entry = state.history[i];
      if (entry.type === "point" && state.selectedCells.has(entry.key)) {
        state.lastRoll = entry.roll;
        return;
      }
    }
    state.lastRoll = null;
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

  function countMarkersByType() {
    const counts = {};
    state.markers.forEach((marker) => {
      counts[marker.type] = (counts[marker.type] || 0) + 1;
    });
    return counts;
  }

  function render() {
    if (state.tool === "pencil") {
      ui.toolMode.textContent = "Paint plates";
    } else if (state.tool === "arrow") {
      ui.toolMode.textContent = "Paint arrows";
    } else if (state.tool === "crust") {
      ui.toolMode.textContent = "Roll crust";
    } else if (state.tool === "plate-id") {
      ui.toolMode.textContent = "Identify plates";
    } else if (state.tool === "direction") {
      ui.toolMode.textContent = "Roll directions";
    } else if (state.tool === "continent") {
      ui.toolMode.textContent = "Paint continents";
    } else if (state.tool === "volcano") {
      ui.toolMode.textContent = "Add volcanoes";
    } else {
      ui.toolMode.textContent = "Roll";
    }
    ui.lastSector.textContent = formatLastX();
    ui.lastCell.textContent = formatLastY();
    ui.filled.textContent = state.selectedCells.size;
    const markerCounts = countMarkersByType();
    const arrowCount = state.lines.filter(
      (line) => (line.lineType || "plate") === "arrow"
    ).length;
    ui.rollBtn.disabled = state.selectedCells.size >= TOTAL_CELLS;
    ui.rollSixBtn.disabled = state.selectedCells.size >= TOTAL_CELLS;
    ui.clearBtn.disabled = !hasMarks();
    ui.clearBtn.textContent = state.clearConfirm ? "Confirm restart" : "Restart";
    ui.undoBtn.disabled = state.history.length === 0;
    ui.redoBtn.disabled = state.future.length === 0;
    ui.clearPointsBtn.disabled = state.selectedCells.size === 0;
    ui.clearArrowsBtn.disabled = arrowCount === 0;
    ui.clearCrustBtn.disabled = (markerCounts.crust || 0) === 0;
    ui.clearDirectionsBtn.disabled = (markerCounts.direction || 0) === 0;
    ui.clearPlatesBtn.disabled = (markerCounts["plate-id"] || 0) === 0;
    ui.clearContinentsBtn.disabled =
      state.continentPieces.length === 0 && state.continentDraft.length === 0;
    ui.clearVolcanoesBtn.disabled = (markerCounts.volcano || 0) === 0;
    ui.pencilBtn.setAttribute("aria-pressed", state.tool === "pencil");
    ui.pencilBtn.classList.toggle("active", state.tool === "pencil");
    ui.arrowBtn.setAttribute("aria-pressed", state.tool === "arrow");
    ui.arrowBtn.classList.toggle("active", state.tool === "arrow");
    ui.crustBtn.setAttribute("aria-pressed", state.tool === "crust");
    ui.crustBtn.classList.toggle("active", state.tool === "crust");
    ui.plateIdBtn.setAttribute("aria-pressed", state.tool === "plate-id");
    ui.plateIdBtn.classList.toggle("active", state.tool === "plate-id");
    ui.directionBtn.setAttribute("aria-pressed", state.tool === "direction");
    ui.directionBtn.classList.toggle("active", state.tool === "direction");
    ui.continentBtn.setAttribute("aria-pressed", state.tool === "continent");
    ui.continentBtn.classList.toggle("active", state.tool === "continent");
    ui.volcanoBtn.setAttribute("aria-pressed", state.tool === "volcano");
    ui.volcanoBtn.classList.toggle("active", state.tool === "volcano");
    const wrapEnabled = state.tool === "pencil";
    if (ui.wrapLeftBtn) {
      ui.wrapLeftBtn.disabled = !wrapEnabled;
      ui.wrapLeftBtn.setAttribute("aria-pressed", state.wrapSide === "left");
      ui.wrapLeftBtn.classList.toggle("active", state.wrapSide === "left");
    }
    if (ui.wrapRightBtn) {
      ui.wrapRightBtn.disabled = !wrapEnabled;
      ui.wrapRightBtn.setAttribute("aria-pressed", state.wrapSide === "right");
      ui.wrapRightBtn.classList.toggle("active", state.wrapSide === "right");
    }
    if (ui.cornerButtons) {
      ui.cornerButtons.forEach((corner) => {
        if (!corner.el) {
          return;
        }
        const key = `${corner.col - 1},${corner.row - 1}`;
        const isActive = state.selectedCells.has(key);
        corner.el.setAttribute("aria-pressed", isActive);
        corner.el.classList.toggle("active", isActive);
      });
    }
    ui.deleteContinentBtn.disabled = !state.selectedContinentId;
    ui.board.classList.toggle("pencil-active", isLineTool(state.tool));
    document.body.dataset.tool = state.tool;
    ui.board.innerHTML = buildSvg();
    drawContinents();
    drawMarkers();
    if (ui.markerLayer && !state.markerDrag) {
      ui.markerLayer.style.cursor = isMarkerTool(state.tool) ? "crosshair" : "default";
    }
    scheduleEdgeControlsLayout();
  }

  function formatLastX() {
    if (!state.lastRoll) {
      return "-";
    }
    return String(state.lastRoll.col);
  }

  function formatLastY() {
    if (!state.lastRoll) {
      return "-";
    }
    return String(state.lastRoll.row);
  }

  function buildSvg() {
    const parts = [];
    const arrowSize = Math.min(microW, microH) * 0.3;
    const arrowSizeText = arrowSize.toFixed(2);
    const arrowHalfText = (arrowSize / 2).toFixed(2);

    parts.push(
      `<defs><marker id="arrow-head" markerUnits="userSpaceOnUse" markerWidth="${arrowSizeText}" markerHeight="${arrowSizeText}" viewBox="0 0 ${arrowSizeText} ${arrowSizeText}" refX="${arrowSizeText}" refY="${arrowHalfText}" orient="auto"><path d="M 0 0 L ${arrowSizeText} ${arrowHalfText} L 0 ${arrowSizeText} z" fill="var(--ink)" /></marker></defs>`
    );

    for (let i = 1; i < GRID; i += 1) {
      const x = (cellW * i).toFixed(2);
      parts.push(
        `<line class="grid-line" x1="${x}" y1="0" x2="${x}" y2="${HEIGHT}" />`
      );
    }

    for (let i = 1; i < GRID; i += 1) {
      const y = (cellH * i).toFixed(2);
      parts.push(
        `<line class="grid-line" x1="0" y1="${y}" x2="${WIDTH}" y2="${y}" />`
      );
    }

    parts.push(
      `<rect class="grid-border" x="0.5" y="0.5" width="${WIDTH - 1}" height="${HEIGHT - 1}" />`
    );

    state.lines.forEach((line) => {
      const lineType = line.lineType || "plate";
      const lineClass = lineType === "arrow" ? "arrow-line" : "pencil-line";
      const markerEnd = lineType === "arrow" ? ' marker-end="url(#arrow-head)"' : "";
      const segments = getLineSegments(line);
      segments.forEach((segment) => {
        parts.push(
          `<line class="${lineClass}"${markerEnd} x1="${segment.x1.toFixed(
            2
          )}" y1="${segment.y1.toFixed(2)}" x2="${segment.x2.toFixed(
            2
          )}" y2="${segment.y2.toFixed(2)}" />`
        );
      });
    });

    if (state.currentLine) {
      const lineType = state.currentLine.lineType || "plate";
      const lineClass = lineType === "arrow" ? "arrow-line" : "pencil-line";
      const markerEnd = lineType === "arrow" ? ' marker-end="url(#arrow-head)"' : "";
      parts.push(
        `<line class="${lineClass}"${markerEnd} x1="${state.currentLine.x1.toFixed(
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

    if (state.plateDraft) {
      const ringRadius = dotRadius * 3.2;
      parts.push(
        `<circle class="plate-node-active" cx="${state.plateDraft.x.toFixed(
          2
        )}" cy="${state.plateDraft.y.toFixed(2)}" r="${ringRadius.toFixed(
          2
        )}" />`
      );
    }

    return parts.join("");
  }

  function buildMarkerSvg() {
    const parts = [];
    state.markers.forEach((marker) => {
      const label = escapeSvgText(getMarkerLabel(marker));
      const size = getMarkerFontSize(marker).toFixed(2);
      const selected = marker.id === state.selectedMarkerId ? " selected" : "";
      parts.push(
        `<text class="marker-text marker-${marker.type}${selected}" x="${marker.x.toFixed(
          2
        )}" y="${marker.y.toFixed(2)}" text-anchor="middle" dominant-baseline="central" font-size="${size}">${label}</text>`
      );
    });
    return parts.join("");
  }

  function drawMarkers() {
    if (!ui.markerLayer) {
      return;
    }
    ui.markerLayer.innerHTML = buildMarkerSvg();
  }

  function escapeSvgText(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function handleResize() {
    updateHeaderOffset();
    updateOverlayCanvasSize();
    drawContinents();
    scheduleEdgeControlsLayout();
  }

  function updateHeaderOffset() {
    if (!ui.header) {
      return;
    }
    const rect = ui.header.getBoundingClientRect();
    const offset = Math.max(0, rect.top + window.scrollY);
    document.documentElement.style.setProperty(
      "--header-offset",
      `${Math.round(offset)}px`
    );
  }

  function scheduleEdgeControlsLayout() {
    if (!ui.boardStack || !ui.edgeControls) {
      return;
    }
    if (edgeLayoutFrame) {
      return;
    }
    edgeLayoutFrame = window.requestAnimationFrame(() => {
      edgeLayoutFrame = null;
      updateEdgeControlsLayout();
    });
  }

  function updateEdgeControlsLayout() {
    if (!ui.boardStack || !ui.edgeControls) {
      return;
    }
    const rect = ui.boardStack.getBoundingClientRect();
    if (!rect.width || !rect.height) {
      return;
    }
    const left = Math.round(rect.left + window.scrollX);
    const top = Math.round(rect.top + window.scrollY);
    const width = Math.round(rect.width);
    const height = Math.round(rect.height);
    if (
      lastEdgeLayout &&
      lastEdgeLayout.left === left &&
      lastEdgeLayout.top === top &&
      lastEdgeLayout.width === width &&
      lastEdgeLayout.height === height
    ) {
      return;
    }
    lastEdgeLayout = {
      left,
      top,
      width,
      height,
    };
    document.documentElement.style.setProperty("--board-left", `${left}px`);
    document.documentElement.style.setProperty("--board-top", `${top}px`);
    document.documentElement.style.setProperty("--board-width", `${width}px`);
    document.documentElement.style.setProperty("--board-height", `${height}px`);
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

    if (event.button === 2 && state.continentDraft.length > 0) {
      state.continentDraft = [];
      state.continentCursor = null;
      setMessage("Continent draft cancelled.");
      render();
      return;
    }

    if (event.button === 2) {
      const hit = findContinentAt(point);
      if (!hit) {
        return;
      }
      selectContinent(hit.piece.id);
      state.continentDrag = {
        type: "rotate",
        id: hit.piece.id,
        shiftX: hit.shiftX,
        startAngle: Math.atan2(point.y - hit.y, point.x - hit.x),
        startRotation: hit.piece.rotation,
        before: cloneContinentPiece(hit.piece),
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
        shiftX: nodeHit.shiftX,
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
      selectContinent(hit.piece.id);
      state.continentDrag = {
        type: "move",
        id: hit.piece.id,
        shiftX: hit.shiftX,
        offsetX: point.x - hit.x,
        offsetY: point.y - hit.y,
        before: cloneContinentPiece(hit.piece),
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
        const local = toLocalPoint(
          point,
          piece,
          state.continentDrag.shiftX || 0
        );
        piece.points[state.continentDrag.index] = [
          local[0] + state.continentDrag.offsetX,
          local[1] + state.continentDrag.offsetY,
        ];
        updateContinentRadius(piece);
      } else if (state.continentDrag.type === "rotate") {
        const centerX = piece.x + (state.continentDrag.shiftX || 0);
        const angle = Math.atan2(point.y - piece.y, point.x - centerX);
        piece.rotation =
          state.continentDrag.startRotation +
          (angle - state.continentDrag.startAngle);
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
    recordHistory({
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
    setMessage("Click points, then click the first point again to close. Right click cancels.");
  }

  function finalizeContinentDraft() {
    if (state.continentDraft.length < 3) {
      return;
    }
    const piece = createContinentPiece(state.continentDraft);
    state.continentPieces.push(piece);
    recordHistory({
      type: "continent-add",
      piece: cloneContinentPiece(piece),
      index: state.continentPieces.length - 1,
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
    state.selectedMarkerId = null;
    state.markerDrag = null;
    bringContinentToFront(id);
  }

  function deleteSelectedContinent(shouldRecordHistory = true) {
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
    if (shouldRecordHistory && removed) {
      recordHistory({
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

  function undoMarkerAdd(entry) {
    const index = state.markers.findIndex((marker) => marker.id === entry.marker.id);
    if (index >= 0) {
      state.markers.splice(index, 1);
    }
    if (state.selectedMarkerId === entry.marker.id) {
      state.selectedMarkerId = null;
    }
  }

  function undoMarkerDelete(entry) {
    const insertAt = clamp(entry.index, 0, state.markers.length);
    const restored = cloneMarker(entry.marker);
    state.markers.splice(insertAt, 0, restored);
    state.selectedMarkerId = restored.id;
  }

  function undoMarkerUpdate(entry) {
    const marker = getMarkerById(entry.id);
    if (!marker) {
      return;
    }
    applyMarkerSnapshot(marker, entry.before);
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

  function getWrappedCenters(piece) {
    return [
      { x: piece.x, y: piece.y, shiftX: 0 },
      { x: piece.x - WIDTH, y: piece.y, shiftX: -WIDTH },
      { x: piece.x + WIDTH, y: piece.y, shiftX: WIDTH },
    ];
  }

  function getVisibleWrappedCenters(piece) {
    const radius = piece.radius || 0;
    return getWrappedCenters(piece).filter(
      (center) => center.x + radius >= 0 && center.x - radius <= WIDTH
    );
  }

  function findContinentAt(point) {
    for (let i = state.continentPieces.length - 1; i >= 0; i -= 1) {
      const piece = state.continentPieces[i];
      const centers = getVisibleWrappedCenters(piece);
      for (let j = 0; j < centers.length; j += 1) {
        const center = centers[j];
        const dx = point.x - center.x;
        const dy = point.y - center.y;
        if (piece.radius && dx * dx + dy * dy > piece.radius * piece.radius) {
          continue;
        }
        const localPoint = toLocalPoint(point, piece, center.shiftX);
        if (pointInPolygon(localPoint, piece.points)) {
          return { piece, x: center.x, y: center.y, shiftX: center.shiftX };
        }
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
    const centers = getVisibleWrappedCenters(piece);
    for (let i = 0; i < centers.length; i += 1) {
      const center = centers[i];
      for (let j = 0; j < piece.points.length; j += 1) {
        const world = toWorldPoint(piece, piece.points[j], center.shiftX);
        const dx = point.x - world.x;
        const dy = point.y - world.y;
        if (dx * dx + dy * dy <= nodeHitRadius * nodeHitRadius) {
          const local = toLocalPoint(point, piece, center.shiftX);
          return {
            piece,
            index: j,
            offsetX: piece.points[j][0] - local[0],
            offsetY: piece.points[j][1] - local[1],
            shiftX: center.shiftX,
          };
        }
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

  function toLocalPoint(point, piece, shiftX = 0) {
    const dx = point.x - (piece.x + shiftX);
    const dy = point.y - piece.y;
    const sin = Math.sin(-piece.rotation);
    const cos = Math.cos(-piece.rotation);
    return [dx * cos - dy * sin, dx * sin + dy * cos];
  }

  function toWorldPoint(piece, local, shiftX = 0) {
    const cos = Math.cos(piece.rotation);
    const sin = Math.sin(piece.rotation);
    return {
      x: piece.x + shiftX + local[0] * cos - local[1] * sin,
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
    const centers = getVisibleWrappedCenters(piece);
    centers.forEach((center) => {
      context.save();
      context.translate(center.x, center.y);
      context.rotate(piece.rotation);
      traceContinentPath(context, piece.points);
      context.fillStyle = continentFill;
      context.fill();
      context.lineWidth = isSelected ? 1.2 : 0.8;
      context.strokeStyle = strokeStyle;
      context.stroke();
      context.restore();
    });
  }

  function drawContinentNodes(context, piece) {
    const centers = getVisibleWrappedCenters(piece);
    centers.forEach((center) => {
      context.save();
      context.translate(center.x, center.y);
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
    });
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
    if (state.tool === "continent") {
      if (state.continentDraft.length > 0) {
        ui.continentHint.textContent =
          "Click the first node again to close. Right click cancels.";
        return;
      }
      if (state.selectedContinentId) {
        ui.continentHint.textContent =
          "Drag to move. Right-drag to rotate. Drag nodes to reshape.";
        return;
      }
      ui.continentHint.textContent =
        "Click to add nodes. Close by clicking the first node.";
      return;
    }
    if (state.tool === "crust") {
      ui.continentHint.textContent =
        "Click to roll crust and place an icon. Drag to move.";
      return;
    }
    if (state.tool === "plate-id") {
      ui.continentHint.textContent =
        "Click to place the next plate number. Drag to move.";
      return;
    }
    if (state.tool === "direction") {
      ui.continentHint.textContent =
        "Click to roll a direction arrow. Drag to move.";
      return;
    }
    if (state.tool === "volcano") {
      ui.continentHint.textContent = "Click to place a volcano. Drag to move.";
      return;
    }
    if (state.tool === "pencil") {
      ui.continentHint.textContent =
        "Click points to connect in a chain. Use the side turn-around buttons to wrap the next link. Right click or click empty space to stop. Max 3 connections per point.";
      return;
    }
    if (state.tool === "arrow") {
      ui.continentHint.textContent = "Click and drag to draw arrows.";
      return;
    }
    ui.continentHint.textContent = "Select a tool to begin.";
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
