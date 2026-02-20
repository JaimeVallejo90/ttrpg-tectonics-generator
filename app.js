(() => {
  const WIDTH = 297;
  const HEIGHT = 210;
  const GRID = 20;
  const TOTAL_CELLS = GRID * GRID;
  const MAX_ATTEMPTS = 1000;
  const MAX_PLATE_CONNECTIONS = 3;

  const cellW = WIDTH / GRID;
  const cellH = HEIGHT / GRID;
  const innerLeft = 0;
  const innerTop = 0;
  const microW = cellW;
  const microH = cellH;
  const dotRadius = Math.min(microW, microH) * 0.105;
  const draftPointRadius = Math.min(microW, microH) * 0.14;
  const closeDistance = Math.min(microW, microH) * 0.6;
  const nodeRadius = Math.min(microW, microH) * 0.18;
  const nodeHitRadius = nodeRadius * 1.9;
  const plateHitRadius = Math.min(microW, microH) * 0.5;
  const continentStroke = "rgba(20, 20, 20, 0.9)";
  const continentStrokeActive = "rgba(7, 7, 7, 0.98)";
  const draftPoint = "rgba(20, 20, 20, 0.85)";
  const draftPointActive = "rgba(7, 7, 7, 0.98)";
  const divideStroke = "rgba(14, 12, 10, 0.96)";
  const divideDashPattern = [6, 4];
  const divideLineWidth = 1;
  const surfacePaintColors = {
    continent: "rgba(20, 20, 20, 0.22)",
    ocean: "rgba(20, 20, 20, 0.12)",
  };
  const markerHitRadius = Math.min(microW, microH) * 0.45;
  const markerIconSize = Math.min(microW, microH) * 0.75;
  const markerNumberSize = Math.min(microW, microH) * 0.6;
  const volcanoIconSize = Math.min(microW, microH) * 0.55;
  const crustIcons = {
    ocean: "\u2248",
    mixed: "\u25D0",
    continent: "\u25CF",
  };
  const volcanoIcon = "\u25B2";
  const directionIcons = {
    1: "\u2191",
    2: "\u2197",
    3: "\u2198",
    4: "\u2193",
    5: "\u2199",
    6: "\u2196",
  };
  const directionLabels = {
    1: "N",
    2: "NE",
    3: "SE",
    4: "S",
    5: "SW",
    6: "NW",
  };
  const PLATE_STYLES = {
    boundary: "boundary",
    divergent: "divergent",
    convergent: "convergent",
    oblique: "oblique",
  };
  const DEFAULT_SURFACE_BRUSH_RADIUS = 4;
  const AUTO_ROLL_SAMPLE_COLS = GRID * 12;
  const AUTO_ROLL_SAMPLE_ROWS = Math.max(
    GRID,
    Math.round((AUTO_ROLL_SAMPLE_COLS * HEIGHT) / WIDTH)
  );

  const state = {
    selectedCells: new Set(),
    pointWeights: {},
    history: [],
    future: [],
    lastRoll: null,
    tool: "roll",
    isDrawing: false,
    currentLine: null,
    lines: [],
    plateDraft: null,
    plateStyle: PLATE_STYLES.boundary,
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
    surfacePaintStamps: [],
    surfacePaintDrag: null,
    surfaceBrushRadius: DEFAULT_SURFACE_BRUSH_RADIUS,
    nextContinentId: 1,
    nextMarkerId: 1,
    clearConfirm: false,
  };

  const ui = {};
  let edgeLayoutFrame = null;
  let lastEdgeLayout = null;
  let rightPanelOffsetFrame = null;
  let overlayRenderFrame = null;
  let overlayUnitsPerCssPixel = 1;
  let surfacePaintCacheCanvas = null;
  let surfacePaintCacheCtx = null;
  let surfacePaintCacheDirty = true;

  function init() {
    ui.board = document.getElementById("board");
    ui.rollBtn = document.getElementById("roll-btn");
    ui.rollSixBtn = document.getElementById("roll-6-btn");
    ui.rollAllCrustsBtn = document.getElementById("roll-all-crusts-btn");
    ui.rollAllDirectionsBtn = document.getElementById("roll-all-directions-btn");
    ui.connectPointsBtn = document.getElementById("connect-points-btn");
    ui.pencilBtn = document.getElementById("pencil-btn");
    ui.arrowBtn = document.getElementById("arrow-btn");
    ui.divideBtn = document.getElementById("divide-btn");
    ui.crustBtn = document.getElementById("crust-btn");
    ui.plateIdBtn = document.getElementById("plate-id-btn");
    ui.directionBtn = document.getElementById("direction-btn");
    ui.continentBtn = document.getElementById("continent-btn");
    ui.volcanoBtn = document.getElementById("volcano-btn");
    ui.moveIconBtn = document.getElementById("move-icon-btn");
    ui.deleteContinentBtn = document.getElementById("delete-continent-btn");
    ui.clearBtn = document.getElementById("clear-btn");
    ui.undoBtn = document.getElementById("undo-btn");
    ui.redoBtn = document.getElementById("redo-btn");
    ui.clearPointsBtn = document.getElementById("clear-points-btn");
    ui.clearArrowsBtn = document.getElementById("clear-arrows-btn");
    ui.clearDividesBtn = document.getElementById("clear-divides-btn");
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
    ui.workspacePanel = document.querySelector(".workspace-panel");
    ui.boardStack = document.querySelector(".board-stack");
    ui.edgeControls = document.getElementById("edge-controls");
    ui.wrapLeftBtn = document.getElementById("wrap-left-btn");
    ui.wrapRightBtn = document.getElementById("wrap-right-btn");
    ui.autoRollPanel = document.getElementById("auto-roll-panel");
    ui.plateModePanel = document.getElementById("plate-mode-panel");
    ui.plateBoundaryBtn = document.getElementById("plate-boundary-btn");
    ui.plateDivergentBtn = document.getElementById("plate-divergent-btn");
    ui.plateConvergentBtn = document.getElementById("plate-convergent-btn");
    ui.plateObliqueBtn = document.getElementById("plate-oblique-btn");
    ui.plateModeHint = document.getElementById("plate-mode-hint");
    ui.surfacePaintPanel = document.getElementById("surface-paint-panel");
    ui.paintContinentBtn = document.getElementById("paint-continent-btn");
    ui.paintOceanBtn = document.getElementById("paint-ocean-btn");
    ui.surfaceRadiusSlider = document.getElementById("surface-radius-slider");
    ui.surfaceRadiusValue = document.getElementById("surface-radius-value");
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
    if (ui.rollAllCrustsBtn) {
      ui.rollAllCrustsBtn.addEventListener("click", handleRollAllCrusts);
    }
    if (ui.rollAllDirectionsBtn) {
      ui.rollAllDirectionsBtn.addEventListener("click", handleRollAllDirections);
    }
    if (ui.connectPointsBtn) {
      ui.connectPointsBtn.addEventListener("click", handleConnectPoints);
    }
    ui.pencilBtn.addEventListener("click", togglePencil);
    ui.arrowBtn.addEventListener("click", toggleArrow);
    if (ui.divideBtn) {
      ui.divideBtn.addEventListener("click", toggleDivide);
    }
    ui.crustBtn.addEventListener("click", toggleCrust);
    ui.plateIdBtn.addEventListener("click", togglePlateId);
    ui.directionBtn.addEventListener("click", toggleDirection);
    ui.continentBtn.addEventListener("click", toggleContinent);
    ui.volcanoBtn.addEventListener("click", toggleVolcano);
    if (ui.moveIconBtn) {
      ui.moveIconBtn.addEventListener("click", toggleMoveIcon);
    }
    ui.deleteContinentBtn.addEventListener("click", handleDeleteContinent);
    ui.clearBtn.addEventListener("click", handleClearAll);
    ui.undoBtn.addEventListener("click", handleUndo);
    ui.redoBtn.addEventListener("click", handleRedo);
    ui.clearPointsBtn.addEventListener("click", handleClearPoints);
    ui.clearArrowsBtn.addEventListener("click", handleClearArrows);
    if (ui.clearDividesBtn) {
      ui.clearDividesBtn.addEventListener("click", handleClearDivides);
    }
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
    if (ui.plateBoundaryBtn) {
      ui.plateBoundaryBtn.addEventListener("click", () =>
        setPlateStyle(PLATE_STYLES.boundary)
      );
    }
    if (ui.plateDivergentBtn) {
      ui.plateDivergentBtn.addEventListener("click", () =>
        setPlateStyle(PLATE_STYLES.divergent)
      );
    }
    if (ui.plateConvergentBtn) {
      ui.plateConvergentBtn.addEventListener("click", () =>
        setPlateStyle(PLATE_STYLES.convergent)
      );
    }
    if (ui.plateObliqueBtn) {
      ui.plateObliqueBtn.addEventListener("click", () =>
        setPlateStyle(PLATE_STYLES.oblique)
      );
    }
    if (ui.paintContinentBtn) {
      ui.paintContinentBtn.addEventListener("click", togglePaintContinent);
    }
    if (ui.paintOceanBtn) {
      ui.paintOceanBtn.addEventListener("click", togglePaintOcean);
    }
    if (ui.surfaceRadiusSlider) {
      state.surfaceBrushRadius = clampSurfaceBrushRadius(
        Number(ui.surfaceRadiusSlider.value) || DEFAULT_SURFACE_BRUSH_RADIUS
      );
      ui.surfaceRadiusSlider.value = String(
        Math.round(state.surfaceBrushRadius)
      );
      ui.surfaceRadiusSlider.addEventListener("input", handleSurfaceRadiusInput);
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
    if (ui.workspacePanel) {
      ui.workspacePanel.addEventListener("animationend", scheduleEdgeControlsLayout);
    }
    if (document.fonts && document.fonts.ready) {
      document.fonts.ready.then(() => {
        scheduleEdgeControlsLayout();
      });
    }

    updateHeaderOffset();
    render();
    updateRightPanelOffset();
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
      if (result.hotspot) {
        setMessage(`Repeated (${result.col}, ${result.row}): point converted to volcano.`);
      } else {
        setMessage(`Marked (${result.col}, ${result.row}).`);
      }
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
    let hotspots = 0;
    let lastResult = null;

    for (let i = 0; i < 30; i += 1) {
      const result = rollRandomPoint();
      if (!result) {
        break;
      }
      if (result.hotspot) {
        hotspots += 1;
      } else {
        added += 1;
      }
      lastResult = result;
    }

    if (added === 0 && hotspots === 0) {
      setMessage("No free cell found.");
    } else {
      const lastText = lastResult ? ` Last (${lastResult.col}, ${lastResult.row}).` : "";
      setMessage(`Rolled 30 points. New: ${added}. Hotspots: ${hotspots}.${lastText}`);
    }

    render();
  }

  function handleRollAllCrusts() {
    cancelClearConfirm();
    rollAllPlateCenters("crust");
  }

  function handleRollAllDirections() {
    cancelClearConfirm();
    rollAllPlateCenters("direction");
  }

  function isPlateConnectablePointKey(key) {
    return getPointWeight(key) <= 1;
  }

  function getConnectablePlatePointCount() {
    let count = 0;
    state.selectedCells.forEach((key) => {
      if (isPlateConnectablePointKey(key)) {
        count += 1;
      }
    });
    return count;
  }

  function getSelectedPlatePointsForAutoConnect() {
    const points = [];
    state.selectedCells.forEach((key) => {
      if (!isPlateConnectablePointKey(key)) {
        return;
      }
      const [col, row] = key.split(",").map(Number);
      const center = getPointCenter(col, row);
      points.push({
        key,
        col,
        row,
        x: center.x,
        y: center.y,
      });
    });
    points.sort((a, b) => a.row - b.row || a.col - b.col);
    return points;
  }

  function getAutoConnectPairKey(keyA, keyB) {
    return keyA < keyB ? `${keyA}|${keyB}` : `${keyB}|${keyA}`;
  }

  function normalizeAngle(angle) {
    const fullTurn = Math.PI * 2;
    let normalized = angle % fullTurn;
    if (normalized > Math.PI) {
      normalized -= fullTurn;
    } else if (normalized < -Math.PI) {
      normalized += fullTurn;
    }
    return normalized;
  }

  function smallestAngleDifference(angleA, angleB) {
    const diff = Math.abs(normalizeAngle(angleA - angleB));
    return diff > Math.PI ? Math.PI * 2 - diff : diff;
  }

  function getAutoConnectWrapOptions(start, end) {
    const baseDx = end.x - start.x;
    const options = [
      { wrapSide: null, dx: baseDx },
      { wrapSide: "left", dx: baseDx - WIDTH },
      { wrapSide: "right", dx: baseDx + WIDTH },
    ];
    options.sort((a, b) => {
      const diff = Math.abs(a.dx) - Math.abs(b.dx);
      if (Math.abs(diff) > 0.0001) {
        return diff;
      }
      if (a.wrapSide && !b.wrapSide) {
        return -1;
      }
      if (!a.wrapSide && b.wrapSide) {
        return 1;
      }
      return 0;
    });
    return options.slice(0, 2);
  }

  function createAutoConnectCandidate(
    start,
    end,
    startIndex,
    endIndex,
    wrapSide = null
  ) {
    const baseDx = end.x - start.x;
    const dx =
      wrapSide === "left" ? baseDx - WIDTH : wrapSide === "right" ? baseDx + WIDTH : baseDx;
    const dy = end.y - start.y;
    const length = Math.hypot(dx, dy);
    const preferredLength = Math.min(WIDTH, HEIGHT) * 0.22;
    const longPenalty = Math.max(0, length - Math.min(WIDTH, HEIGHT) * 0.48) * 5.4;
    const extremePenalty = Math.max(0, length - Math.min(WIDTH, HEIGHT) * 0.62) * 22;
    const randomJitter = (Math.random() - 0.5) * Math.min(microW, microH) * 1.6;
    const fillScore =
      Math.abs(length - preferredLength) * 0.72 +
      length * 0.32 +
      longPenalty +
      extremePenalty +
      randomJitter +
      (wrapSide ? -1.6 : 0);
    const treeScore =
      length +
      longPenalty * 1.45 +
      extremePenalty * 2.7 +
      Math.random() * Math.min(microW, microH) * 1.2 +
      (wrapSide ? -1.1 : 0);
    const angleFromStart = Math.atan2(dy, dx);
    const angleFromEnd = normalizeAngle(angleFromStart + Math.PI);

    return {
      pairKey: getAutoConnectPairKey(start.key, end.key),
      start,
      end,
      startKey: start.key,
      endKey: end.key,
      startIndex,
      endIndex,
      wrapSide,
      length,
      fillScore,
      treeScore,
      angleFromStart,
      angleFromEnd,
      curve: null,
      hitSegments: null,
    };
  }

  function buildAutoConnectCandidates(points) {
    const candidates = [];
    for (let i = 0; i < points.length; i += 1) {
      for (let j = i + 1; j < points.length; j += 1) {
        const wrapOptions = getAutoConnectWrapOptions(points[i], points[j]);
        wrapOptions.forEach((option) => {
          candidates.push(
            createAutoConnectCandidate(points[i], points[j], i, j, option.wrapSide)
          );
        });
      }
    }
    return candidates;
  }

  function createUnionFind(size) {
    const parent = Array.from({ length: size }, (_, index) => index);
    const rank = new Array(size).fill(0);

    function find(index) {
      let root = index;
      while (parent[root] !== root) {
        root = parent[root];
      }
      let current = index;
      while (parent[current] !== current) {
        const next = parent[current];
        parent[current] = root;
        current = next;
      }
      return root;
    }

    function union(a, b) {
      const rootA = find(a);
      const rootB = find(b);
      if (rootA === rootB) {
        return false;
      }
      if (rank[rootA] < rank[rootB]) {
        parent[rootA] = rootB;
      } else if (rank[rootA] > rank[rootB]) {
        parent[rootB] = rootA;
      } else {
        parent[rootB] = rootA;
        rank[rootA] += 1;
      }
      return true;
    }

    function connected(a, b) {
      return find(a) === find(b);
    }

    return { find, union, connected };
  }

  function getAutoConnectAnglePenalty(nodeAngles, key, candidateAngle) {
    const angles = nodeAngles.get(key);
    if (!angles || angles.length === 0) {
      return 0;
    }
    let minDiff = Math.PI;
    for (let i = 0; i < angles.length; i += 1) {
      const diff = smallestAngleDifference(candidateAngle, angles[i]);
      if (diff < minDiff) {
        minDiff = diff;
      }
    }
    const threshold = Math.PI / 4.3;
    if (minDiff >= threshold) {
      return 0;
    }
    return ((threshold - minDiff) / threshold) * 3.6;
  }

  function autoConnectPointOnSegment(
    px,
    py,
    x1,
    y1,
    x2,
    y2,
    epsilon = 0.0001
  ) {
    return (
      px >= Math.min(x1, x2) - epsilon &&
      px <= Math.max(x1, x2) + epsilon &&
      py >= Math.min(y1, y2) - epsilon &&
      py <= Math.max(y1, y2) + epsilon
    );
  }

  function autoConnectSegmentCross(segmentA, segmentB) {
    const epsilon = 0.0001;
    if (
      Math.max(segmentA.x1, segmentA.x2) + epsilon <
        Math.min(segmentB.x1, segmentB.x2) ||
      Math.max(segmentB.x1, segmentB.x2) + epsilon <
        Math.min(segmentA.x1, segmentA.x2) ||
      Math.max(segmentA.y1, segmentA.y2) + epsilon <
        Math.min(segmentB.y1, segmentB.y2) ||
      Math.max(segmentB.y1, segmentB.y2) + epsilon <
        Math.min(segmentA.y1, segmentA.y2)
    ) {
      return false;
    }

    const cross = (ax, ay, bx, by, cx, cy) =>
      (bx - ax) * (cy - ay) - (by - ay) * (cx - ax);

    const d1 = cross(
      segmentA.x1,
      segmentA.y1,
      segmentA.x2,
      segmentA.y2,
      segmentB.x1,
      segmentB.y1
    );
    const d2 = cross(
      segmentA.x1,
      segmentA.y1,
      segmentA.x2,
      segmentA.y2,
      segmentB.x2,
      segmentB.y2
    );
    const d3 = cross(
      segmentB.x1,
      segmentB.y1,
      segmentB.x2,
      segmentB.y2,
      segmentA.x1,
      segmentA.y1
    );
    const d4 = cross(
      segmentB.x1,
      segmentB.y1,
      segmentB.x2,
      segmentB.y2,
      segmentA.x2,
      segmentA.y2
    );

    const hasProperCrossing =
      ((d1 > epsilon && d2 < -epsilon) || (d1 < -epsilon && d2 > epsilon)) &&
      ((d3 > epsilon && d4 < -epsilon) || (d3 < -epsilon && d4 > epsilon));
    if (hasProperCrossing) {
      return true;
    }

    if (
      Math.abs(d1) <= epsilon &&
      autoConnectPointOnSegment(
        segmentB.x1,
        segmentB.y1,
        segmentA.x1,
        segmentA.y1,
        segmentA.x2,
        segmentA.y2,
        epsilon
      )
    ) {
      return true;
    }
    if (
      Math.abs(d2) <= epsilon &&
      autoConnectPointOnSegment(
        segmentB.x2,
        segmentB.y2,
        segmentA.x1,
        segmentA.y1,
        segmentA.x2,
        segmentA.y2,
        epsilon
      )
    ) {
      return true;
    }
    if (
      Math.abs(d3) <= epsilon &&
      autoConnectPointOnSegment(
        segmentA.x1,
        segmentA.y1,
        segmentB.x1,
        segmentB.y1,
        segmentB.x2,
        segmentB.y2,
        epsilon
      )
    ) {
      return true;
    }
    if (
      Math.abs(d4) <= epsilon &&
      autoConnectPointOnSegment(
        segmentA.x2,
        segmentA.y2,
        segmentB.x1,
        segmentB.y1,
        segmentB.x2,
        segmentB.y2,
        epsilon
      )
    ) {
      return true;
    }
    return false;
  }

  function createAutoConnectCurve() {
    const curve = createRandomPlateCurve();
    if (curve && curve.mode === "tectonic") {
      curve.roughness = clamp(curve.roughness * 0.65, 0.018, 0.075);
      curve.broadBend = clamp(curve.broadBend * 0.6, -0.16, 0.16);
      curve.alongScale = clamp(curve.alongScale * 0.55, 0.003, 0.03);
    }
    return curve;
  }

  function getCandidateCurveSegments(candidate) {
    if (candidate.hitSegments) {
      return candidate.hitSegments;
    }
    if (!candidate.curve) {
      candidate.curve = createAutoConnectCurve();
    }
    const wrappedSegments = getWrappedLineSegments(
      candidate.start.x,
      candidate.start.y,
      candidate.end.x,
      candidate.end.y,
      candidate.wrapSide
    );
    candidate.hitSegments = wrappedSegments.flatMap((segment) =>
      buildPlateCurveHitSegments(segment, candidate.curve)
    );
    return candidate.hitSegments;
  }

  function candidateCrossesAcceptedEdges(candidate, acceptedEdges) {
    const candidateSegments = getCandidateCurveSegments(candidate);
    for (let i = 0; i < acceptedEdges.length; i += 1) {
      const existing = acceptedEdges[i];
      const sharesEndpoint =
        candidate.startKey === existing.startKey ||
        candidate.startKey === existing.endKey ||
        candidate.endKey === existing.startKey ||
        candidate.endKey === existing.endKey;
      if (sharesEndpoint) {
        continue;
      }
      const existingSegments = getCandidateCurveSegments(existing);
      for (let a = 0; a < candidateSegments.length; a += 1) {
        for (let b = 0; b < existingSegments.length; b += 1) {
          if (autoConnectSegmentCross(candidateSegments[a], existingSegments[b])) {
            return true;
          }
        }
      }
    }
    return false;
  }

  function acceptAutoConnectCandidate(candidate, context) {
    getCandidateCurveSegments(candidate);
    context.acceptedEdges.push(candidate);
    context.connectedPairs.add(candidate.pairKey);
    context.degrees.set(
      candidate.startKey,
      (context.degrees.get(candidate.startKey) || 0) + 1
    );
    context.degrees.set(
      candidate.endKey,
      (context.degrees.get(candidate.endKey) || 0) + 1
    );
    const startAngles = context.nodeAngles.get(candidate.startKey);
    const endAngles = context.nodeAngles.get(candidate.endKey);
    if (startAngles) {
      startAngles.push(candidate.angleFromStart);
    }
    if (endAngles) {
      endAngles.push(candidate.angleFromEnd);
    }
    const startNeighbors = context.adjacency.get(candidate.startKey);
    const endNeighbors = context.adjacency.get(candidate.endKey);
    if (startNeighbors && endNeighbors) {
      let triangleDelta = 0;
      startNeighbors.forEach((neighborKey) => {
        if (endNeighbors.has(neighborKey)) {
          triangleDelta += 1;
        }
      });
      context.triangleCount += triangleDelta;
    }
    if (startNeighbors) {
      startNeighbors.add(candidate.endKey);
    }
    if (endNeighbors) {
      endNeighbors.add(candidate.startKey);
    }
    context.uf.union(candidate.startIndex, candidate.endIndex);
  }

  function pickAutoConnectCandidateForNode(
    nodeKey,
    candidatesByNode,
    context,
    targetDegree,
    strictCrossing
  ) {
    const pool = candidatesByNode.get(nodeKey) || [];
    let best = null;
    let bestScore = Infinity;

    for (let i = 0; i < pool.length; i += 1) {
      const candidate = pool[i];
      if (context.connectedPairs.has(candidate.pairKey)) {
        continue;
      }

      const startDegree = context.degrees.get(candidate.startKey) || 0;
      const endDegree = context.degrees.get(candidate.endKey) || 0;
      if (startDegree >= MAX_PLATE_CONNECTIONS || endDegree >= MAX_PLATE_CONNECTIONS) {
        continue;
      }

      const otherKey =
        candidate.startKey === nodeKey ? candidate.endKey : candidate.startKey;
      const otherDegree = context.degrees.get(otherKey) || 0;
      let score = candidate.fillScore;

      if (otherDegree < targetDegree) {
        score -= 18;
      } else {
        score += 6;
      }
      if (otherDegree === 0) {
        score -= 8;
      }

      const nodeAngle =
        candidate.startKey === nodeKey
          ? candidate.angleFromStart
          : candidate.angleFromEnd;
      const otherAngle =
        candidate.startKey === nodeKey
          ? candidate.angleFromEnd
          : candidate.angleFromStart;
      score += getAutoConnectAnglePenalty(context.nodeAngles, nodeKey, nodeAngle) * 8;
      score += getAutoConnectAnglePenalty(context.nodeAngles, otherKey, otherAngle) * 5.6;

      if (!context.uf.connected(candidate.startIndex, candidate.endIndex)) {
        score -= 10;
      }

      const crosses = candidateCrossesAcceptedEdges(candidate, context.acceptedEdges);
      if (strictCrossing && crosses) {
        continue;
      }
      if (!strictCrossing && crosses) {
        score += 220;
      }

      const startNeighbors = context.adjacency.get(candidate.startKey);
      const endNeighbors = context.adjacency.get(candidate.endKey);
      let triangleDelta = 0;
      if (startNeighbors && endNeighbors) {
        startNeighbors.forEach((neighborKey) => {
          if (endNeighbors.has(neighborKey)) {
            triangleDelta += 1;
          }
        });
      }
      if (triangleDelta > 0) {
        const baselinePenalty = triangleDelta * 95;
        const overflowTriangles = Math.max(
          0,
          context.triangleCount + triangleDelta - 2
        );
        const overflowPenalty = overflowTriangles * 620;
        score += baselinePenalty + overflowPenalty;
      }

      score += Math.random() * 0.8;
      if (score < bestScore) {
        best = candidate;
        bestScore = score;
      }
    }

    return best;
  }

  function fillAutoConnectDegreeTarget(
    targetDegree,
    strictCrossing,
    points,
    context,
    candidatesByNode
  ) {
    let changed = true;
    let safety = points.length * MAX_PLATE_CONNECTIONS * 6;

    while (changed && safety > 0) {
      changed = false;
      safety -= 1;
      const ordered = [...points].sort((a, b) => {
        const degreeA = context.degrees.get(a.key) || 0;
        const degreeB = context.degrees.get(b.key) || 0;
        if (degreeA !== degreeB) {
          return degreeA - degreeB;
        }
        return Math.random() - 0.5;
      });

      for (let i = 0; i < ordered.length; i += 1) {
        const node = ordered[i];
        const degree = context.degrees.get(node.key) || 0;
        if (degree >= targetDegree || degree >= MAX_PLATE_CONNECTIONS) {
          continue;
        }
        const candidate = pickAutoConnectCandidateForNode(
          node.key,
          candidatesByNode,
          context,
          targetDegree,
          strictCrossing
        );
        if (!candidate) {
          continue;
        }
        acceptAutoConnectCandidate(candidate, context);
        changed = true;
      }
    }
  }

  function countUnionFindComponents(uf, size) {
    const roots = new Set();
    for (let i = 0; i < size; i += 1) {
      roots.add(uf.find(i));
    }
    return roots.size;
  }

  function buildAutoConnectNetwork(points) {
    if (!points || points.length < 2) {
      return {
        edges: [],
        degrees: new Map(),
        components: points && points.length ? points.length : 0,
      };
    }

    const candidates = buildAutoConnectCandidates(points);
    const degrees = new Map();
    const nodeAngles = new Map();
    const candidatesByNode = new Map();
    const adjacency = new Map();
    points.forEach((point) => {
      degrees.set(point.key, 0);
      nodeAngles.set(point.key, []);
      candidatesByNode.set(point.key, []);
      adjacency.set(point.key, new Set());
    });
    candidates.forEach((candidate) => {
      const startCandidates = candidatesByNode.get(candidate.startKey);
      const endCandidates = candidatesByNode.get(candidate.endKey);
      if (startCandidates) {
        startCandidates.push(candidate);
      }
      if (endCandidates) {
        endCandidates.push(candidate);
      }
    });

    const context = {
      acceptedEdges: [],
      connectedPairs: new Set(),
      degrees,
      nodeAngles,
      adjacency,
      triangleCount: 0,
      uf: createUnionFind(points.length),
    };

    const treeCandidates = [...candidates].sort((a, b) => a.treeScore - b.treeScore);
    for (let i = 0; i < treeCandidates.length; i += 1) {
      const candidate = treeCandidates[i];
      if (context.connectedPairs.has(candidate.pairKey)) {
        continue;
      }
      if (context.uf.connected(candidate.startIndex, candidate.endIndex)) {
        continue;
      }
      const startDegree = context.degrees.get(candidate.startKey) || 0;
      const endDegree = context.degrees.get(candidate.endKey) || 0;
      if (startDegree >= MAX_PLATE_CONNECTIONS || endDegree >= MAX_PLATE_CONNECTIONS) {
        continue;
      }
      if (candidateCrossesAcceptedEdges(candidate, context.acceptedEdges)) {
        continue;
      }
      acceptAutoConnectCandidate(candidate, context);
    }

    for (let i = 0; i < treeCandidates.length; i += 1) {
      const candidate = treeCandidates[i];
      if (context.connectedPairs.has(candidate.pairKey)) {
        continue;
      }
      if (context.uf.connected(candidate.startIndex, candidate.endIndex)) {
        continue;
      }
      const startDegree = context.degrees.get(candidate.startKey) || 0;
      const endDegree = context.degrees.get(candidate.endKey) || 0;
      if (startDegree >= MAX_PLATE_CONNECTIONS || endDegree >= MAX_PLATE_CONNECTIONS) {
        continue;
      }
      acceptAutoConnectCandidate(candidate, context);
    }

    fillAutoConnectDegreeTarget(2, true, points, context, candidatesByNode);
    fillAutoConnectDegreeTarget(3, true, points, context, candidatesByNode);

    const hasUnderTwo = points.some((point) => (context.degrees.get(point.key) || 0) < 2);
    if (hasUnderTwo) {
      fillAutoConnectDegreeTarget(2, false, points, context, candidatesByNode);
    }
    const hasUnderThree = points.some((point) => (context.degrees.get(point.key) || 0) < 3);
    if (hasUnderThree) {
      fillAutoConnectDegreeTarget(3, false, points, context, candidatesByNode);
    }

    return {
      edges: context.acceptedEdges,
      degrees: context.degrees,
      components: countUnionFindComponents(context.uf, points.length),
    };
  }

  function countAutoConnectCrossings(edges) {
    if (!edges || edges.length <= 1) {
      return 0;
    }
    let count = 0;
    for (let i = 0; i < edges.length; i += 1) {
      const edgeA = edges[i];
      const segmentsA = getCandidateCurveSegments(edgeA);
      for (let j = i + 1; j < edges.length; j += 1) {
        const edgeB = edges[j];
        if (
          edgeA.startKey === edgeB.startKey ||
          edgeA.startKey === edgeB.endKey ||
          edgeA.endKey === edgeB.startKey ||
          edgeA.endKey === edgeB.endKey
        ) {
          continue;
        }
        const segmentsB = getCandidateCurveSegments(edgeB);
        let crossed = false;
        for (let a = 0; a < segmentsA.length && !crossed; a += 1) {
          for (let b = 0; b < segmentsB.length; b += 1) {
            if (autoConnectSegmentCross(segmentsA[a], segmentsB[b])) {
              crossed = true;
              break;
            }
          }
        }
        if (crossed) {
          count += 1;
        }
      }
    }
    return count;
  }

  function countAutoConnectTriangles(edges) {
    if (!edges || edges.length < 3) {
      return 0;
    }
    const adjacency = new Map();
    edges.forEach((edge) => {
      if (!adjacency.has(edge.startKey)) {
        adjacency.set(edge.startKey, new Set());
      }
      if (!adjacency.has(edge.endKey)) {
        adjacency.set(edge.endKey, new Set());
      }
      adjacency.get(edge.startKey).add(edge.endKey);
      adjacency.get(edge.endKey).add(edge.startKey);
    });

    let triangles = 0;
    const keys = [...adjacency.keys()].sort();
    for (let i = 0; i < keys.length; i += 1) {
      const a = keys[i];
      const neighborsA = [...adjacency.get(a)].filter((key) => key > a).sort();
      for (let j = 0; j < neighborsA.length; j += 1) {
        const b = neighborsA[j];
        const neighborsB = adjacency.get(b);
        if (!neighborsB) {
          continue;
        }
        for (let k = j + 1; k < neighborsA.length; k += 1) {
          const c = neighborsA[k];
          if (neighborsB.has(c)) {
            triangles += 1;
          }
        }
      }
    }
    return triangles;
  }

  function getAutoConnectNetworkStats(points, network) {
    const degrees = (network && network.degrees) || new Map();
    const edges = (network && network.edges) || [];
    let isolated = 0;
    let underTwo = 0;
    let underThree = 0;
    points.forEach((point) => {
      const degree = degrees.get(point.key) || 0;
      if (degree <= 0) {
        isolated += 1;
      }
      if (degree < 2) {
        underTwo += 1;
      }
      if (degree < 3) {
        underThree += 1;
      }
    });

    const longThreshold = Math.min(WIDTH, HEIGHT) * 0.56;
    let longPenalty = 0;
    let totalLength = 0;
    edges.forEach((edge) => {
      const length = Number(edge.length) || 0;
      totalLength += length;
      const overflow = Math.max(0, length - longThreshold);
      longPenalty += overflow * overflow;
    });

    return {
      components: Number(network && network.components) || points.length,
      isolated,
      underTwo,
      underThree,
      crossings: countAutoConnectCrossings(edges),
      triangles: countAutoConnectTriangles(edges),
      longPenalty,
      totalLength,
    };
  }

  function scoreAutoConnectNetworkStats(stats) {
    const triangleOverflow = Math.max(0, stats.triangles - 2);
    return (
      stats.components * 1e9 +
      stats.isolated * 2.6e8 +
      stats.underTwo * 7.2e7 +
      stats.crossings * 1.8e7 +
      stats.triangles * 4.8e5 +
      triangleOverflow * 1.4e7 +
      stats.underThree * 2.2e6 +
      stats.longPenalty * 32 +
      stats.totalLength
    );
  }

  function buildBestAutoConnectNetwork(points) {
    const attemptCount = clamp(Math.round(10 + points.length * 0.35), 10, 28);
    let best = null;

    for (let attempt = 0; attempt < attemptCount; attempt += 1) {
      const network = buildAutoConnectNetwork(points);
      const stats = getAutoConnectNetworkStats(points, network);
      const score = scoreAutoConnectNetworkStats(stats);
      if (!best || score < best.score) {
        best = { network, stats, score };
      }
      if (
        stats.components === 1 &&
        stats.isolated === 0 &&
        stats.underTwo === 0 &&
        stats.crossings === 0 &&
        stats.triangles <= 2 &&
        stats.underThree <= 1
      ) {
        break;
      }
    }

    return best;
  }

  function handleConnectPoints() {
    cancelClearConfirm();
    if (state.selectedCells.size < 2) {
      setMessage("Roll at least two points first.");
      render();
      return;
    }
    if (getPlateLineCount() > 0) {
      setMessage("Connect points requires an empty boundary layer. Clear boundaries first.");
      render();
      return;
    }

    clearPlateDraft();
    state.wrapSide = null;
    const points = getSelectedPlatePointsForAutoConnect();
    if (points.length < 2) {
      setMessage("Need at least 2 points. Volcano hotspots are skipped.");
      render();
      return;
    }
    const bestResult = buildBestAutoConnectNetwork(points);
    const network = bestResult ? bestResult.network : buildAutoConnectNetwork(points);
    if (!network.edges.length) {
      setMessage("No valid connections were generated.");
      render();
      return;
    }

    network.edges.forEach((edge) => {
      addPlateConnection(
        edge.start,
        edge.end,
        edge.wrapSide,
        PLATE_STYLES.boundary,
        edge.curve
      );
    });

    const degreeValues = points.map((point) => network.degrees.get(point.key) || 0);
    const stats = bestResult
      ? bestResult.stats
      : getAutoConnectNetworkStats(points, network);
    const atLeastTwo = degreeValues.filter((degree) => degree >= 2).length;
    const atLeastThree = degreeValues.filter((degree) => degree >= 3).length;
    const isolated = degreeValues.filter((degree) => degree === 0).length;
    const wrapLinks = network.edges.filter((edge) => Boolean(edge.wrapSide)).length;
    const componentText = stats.components > 1 ? ` Components: ${stats.components}.` : "";
    const crossingText = stats.crossings > 0 ? ` crossings: ${stats.crossings},` : "";
    const triangleText = stats.triangles > 0 ? ` triangles: ${stats.triangles},` : "";
    setMessage(
      `Connected ${network.edges.length} boundaries. >=3: ${atLeastThree}/${points.length}, >=2: ${atLeastTwo}/${points.length}, isolated: ${isolated},${crossingText}${triangleText} wrap: ${wrapLinks}.${componentText}`
    );
    render();
  }

  function rollRandomPoint() {
    const col = rollD20();
    const row = rollD20();
    const microCol = col - 1;
    const microRow = row - 1;
    return addCell(microCol, microRow, col, row);
  }

  function getPlateLineCount() {
    let count = 0;
    state.lines.forEach((line) => {
      if ((line.lineType || "plate") === "plate") {
        count += 1;
      }
    });
    return count;
  }

  function collectPlateBarrierSegments() {
    const segments = [];
    state.lines.forEach((line) => {
      if ((line.lineType || "plate") !== "plate") {
        return;
      }
      getLineHitSegments(line).forEach((segment) => {
        segments.push({
          x1: segment.x1,
          y1: segment.y1,
          x2: segment.x2,
          y2: segment.y2,
        });
      });
    });
    return segments;
  }

  function markBarrierDisk(mask, cols, rows, x, y, radiusX, radiusY) {
    const minCol = clamp(Math.floor((x - radiusX) / WIDTH * cols), 0, cols - 1);
    const maxCol = clamp(Math.floor((x + radiusX) / WIDTH * cols), 0, cols - 1);
    const minRow = clamp(Math.floor((y - radiusY) / HEIGHT * rows), 0, rows - 1);
    const maxRow = clamp(Math.floor((y + radiusY) / HEIGHT * rows), 0, rows - 1);
    const rxSq = Math.max(0.0001, radiusX * radiusX);
    const rySq = Math.max(0.0001, radiusY * radiusY);
    const cellW = WIDTH / cols;
    const cellH = HEIGHT / rows;
    for (let row = minRow; row <= maxRow; row += 1) {
      const cy = (row + 0.5) * cellH;
      const dy = cy - y;
      for (let col = minCol; col <= maxCol; col += 1) {
        const cx = (col + 0.5) * cellW;
        const dx = cx - x;
        if ((dx * dx) / rxSq + (dy * dy) / rySq <= 1) {
          mask[row * cols + col] = 1;
        }
      }
    }
  }

  function buildBarrierMaskForAutoRoll(segments, cols, rows) {
    const mask = new Uint8Array(cols * rows);
    const seamBlocked = new Uint8Array(rows);
    if (!segments.length) {
      return { mask, seamBlocked };
    }
    const sampleStep = Math.max(0.5, Math.min(WIDTH / cols, HEIGHT / rows) * 0.36);
    const brushRadius = Math.max(0.65, Math.min(WIDTH / cols, HEIGHT / rows) * 0.9);
    const seamMargin = Math.max(brushRadius * 1.05, WIDTH / cols * 0.9);
    const seamRadiusRows = Math.max(1, Math.ceil(brushRadius / (HEIGHT / rows)));

    segments.forEach((segment) => {
      const dx = segment.x2 - segment.x1;
      const dy = segment.y2 - segment.y1;
      const length = Math.hypot(dx, dy);
      const steps = Math.max(1, Math.ceil(length / sampleStep));
      for (let i = 0; i <= steps; i += 1) {
        const t = i / steps;
        const x = segment.x1 + dx * t;
        const y = segment.y1 + dy * t;
        markBarrierDisk(mask, cols, rows, x, y, brushRadius, brushRadius);
        if (x <= seamMargin || x >= WIDTH - seamMargin) {
          const centerRow = clamp(Math.floor((y / HEIGHT) * rows), 0, rows - 1);
          const startRow = Math.max(0, centerRow - seamRadiusRows);
          const endRow = Math.min(rows - 1, centerRow + seamRadiusRows);
          for (let row = startRow; row <= endRow; row += 1) {
            seamBlocked[row] = 1;
          }
        }
      }
    });

    return { mask, seamBlocked };
  }

  function buildDistanceMapFromBarrierMask(mask, cols, rows) {
    const maxDistance = 65535;
    const distances = new Uint16Array(cols * rows);
    distances.fill(maxDistance);
    const queue = [];

    for (let i = 0; i < mask.length; i += 1) {
      if (mask[i]) {
        distances[i] = 0;
        queue.push(i);
      }
    }
    if (queue.length === 0) {
      return distances;
    }

    for (let head = 0; head < queue.length; head += 1) {
      const index = queue[head];
      const currentDistance = distances[index];
      const row = Math.floor(index / cols);
      const col = index - row * cols;
      const nextDistance = currentDistance + 1;

      const neighborCols = [(col + 1) % cols, (col - 1 + cols) % cols];
      for (let i = 0; i < neighborCols.length; i += 1) {
        const neighborIndex = row * cols + neighborCols[i];
        if (nextDistance < distances[neighborIndex]) {
          distances[neighborIndex] = nextDistance;
          queue.push(neighborIndex);
        }
      }

      if (row + 1 < rows) {
        const downIndex = (row + 1) * cols + col;
        if (nextDistance < distances[downIndex]) {
          distances[downIndex] = nextDistance;
          queue.push(downIndex);
        }
      }
      if (row - 1 >= 0) {
        const upIndex = (row - 1) * cols + col;
        if (nextDistance < distances[upIndex]) {
          distances[upIndex] = nextDistance;
          queue.push(upIndex);
        }
      }
    }

    return distances;
  }

  function collectFloodRegionFromSampleGrid(
    startCol,
    startRow,
    cols,
    rows,
    cellW,
    cellH,
    mask,
    seamBlocked,
    visited
  ) {
    const queue = [{ col: startCol, row: startRow }];
    let readIndex = 0;
    visited[startRow * cols + startCol] = 1;
    const twoPi = Math.PI * 2;
    let count = 0;
    let sumY = 0;
    let sumSin = 0;
    let sumCos = 0;
    const cells = [];

    while (readIndex < queue.length) {
      const current = queue[readIndex];
      readIndex += 1;
      const currentIndex = current.row * cols + current.col;
      cells.push(currentIndex);
      count += 1;
      const x = (current.col + 0.5) * cellW;
      const y = (current.row + 0.5) * cellH;
      const angle = (x / WIDTH) * twoPi;
      sumSin += Math.sin(angle);
      sumCos += Math.cos(angle);
      sumY += y;

      const neighbors = [
        { col: current.col + 1, row: current.row, seamRow: current.row },
        { col: current.col - 1, row: current.row, seamRow: current.row },
        { col: current.col, row: current.row + 1, seamRow: null },
        { col: current.col, row: current.row - 1, seamRow: null },
      ];

      neighbors.forEach((neighbor) => {
        const nextRow = neighbor.row;
        if (nextRow < 0 || nextRow >= rows) {
          return;
        }
        let nextCol = neighbor.col;
        if (nextCol < 0 || nextCol >= cols) {
          if (neighbor.seamRow !== null && seamBlocked[neighbor.seamRow]) {
            return;
          }
          nextCol = (nextCol + cols) % cols;
        }
        const nextIndex = nextRow * cols + nextCol;
        if (visited[nextIndex] || mask[nextIndex]) {
          return;
        }
        visited[nextIndex] = 1;
        queue.push({ col: nextCol, row: nextRow });
      });
    }

    if (count <= 0) {
      return null;
    }
    let centerX = WIDTH / 2;
    if (Math.abs(sumSin) > 0.0001 || Math.abs(sumCos) > 0.0001) {
      const angle = Math.atan2(sumSin, sumCos);
      centerX = ((angle < 0 ? angle + twoPi : angle) / twoPi) * WIDTH;
    }
    return {
      size: count,
      cells,
      meanX: centerX,
      meanY: sumY / count,
    };
  }

  function chooseAutoRollRegionCenter(region, cols, cellW, cellH, distanceMap) {
    if (!region || !region.cells || region.cells.length === 0) {
      return { x: WIDTH / 2, y: HEIGHT / 2 };
    }
    let bestIndex = region.cells[0];
    let bestDistance = -1;
    let bestTieDistance = Infinity;
    region.cells.forEach((index) => {
      const row = Math.floor(index / cols);
      const col = index - row * cols;
      const x = (col + 0.5) * cellW;
      const y = (row + 0.5) * cellH;
      const distance = distanceMap[index];
      const rawDx = Math.abs(x - region.meanX);
      const dx = Math.min(rawDx, WIDTH - rawDx);
      const dy = y - region.meanY;
      const tieDistance = dx * dx + dy * dy;
      if (distance > bestDistance) {
        bestDistance = distance;
        bestTieDistance = tieDistance;
        bestIndex = index;
        return;
      }
      if (distance === bestDistance && tieDistance < bestTieDistance) {
        bestTieDistance = tieDistance;
        bestIndex = index;
      }
    });
    const row = Math.floor(bestIndex / cols);
    const col = bestIndex - row * cols;
    return {
      x: (col + 0.5) * cellW,
      y: (row + 0.5) * cellH,
    };
  }

  function detectPlateRegions() {
    const segments = collectPlateBarrierSegments();
    if (!segments.length) {
      return [];
    }

    const cols = AUTO_ROLL_SAMPLE_COLS;
    const rows = AUTO_ROLL_SAMPLE_ROWS;
    const cellW = WIDTH / cols;
    const cellH = HEIGHT / rows;
    const { mask, seamBlocked } = buildBarrierMaskForAutoRoll(segments, cols, rows);
    const distanceMap = buildDistanceMapFromBarrierMask(mask, cols, rows);
    const visited = new Uint8Array(cols * rows);
    const regions = [];
    const minArea = Math.max(10, Math.round(cols * rows * 0.0003));

    for (let row = 0; row < rows; row += 1) {
      for (let col = 0; col < cols; col += 1) {
        const index = row * cols + col;
        if (visited[index] || mask[index]) {
          continue;
        }
        const region = collectFloodRegionFromSampleGrid(
          col,
          row,
          cols,
          rows,
          cellW,
          cellH,
          mask,
          seamBlocked,
          visited
        );
        if (!region || region.size < minArea) {
          continue;
        }
        regions.push({
          size: region.size,
          center: chooseAutoRollRegionCenter(region, cols, cellW, cellH, distanceMap),
        });
      }
    }

    regions.sort((a, b) => b.size - a.size);
    return regions;
  }

  function dedupeAutoRollRegions(regions) {
    if (!regions || regions.length <= 1) {
      return regions || [];
    }
    const minDistance = Math.max(2.2, Math.min(microW, microH) * 1.08);
    const minDistanceSq = minDistance * minDistance;
    const deduped = [];

    regions.forEach((region) => {
      const duplicate = deduped.some((candidate) => {
        const rawDx = Math.abs(region.center.x - candidate.center.x);
        const dx = Math.min(rawDx, WIDTH - rawDx);
        const dy = region.center.y - candidate.center.y;
        return dx * dx + dy * dy <= minDistanceSq;
      });
      if (!duplicate) {
        deduped.push(region);
      }
    });

    return deduped;
  }

  function getAutoRollMarkerPadding() {
    return Math.max(2, markerIconSize * 0.9);
  }

  function clampAutoRollPoint(point) {
    const padding = getAutoRollMarkerPadding();
    return {
      x: clamp(point.x, padding, WIDTH - padding),
      y: clamp(point.y, padding, HEIGHT - padding),
    };
  }

  function createAutoMarkerFromType(markerType, point) {
    const safePoint = clampAutoRollPoint(point);
    if (markerType === "crust") {
      const roll = rollD6();
      let crustType = "mixed";
      if (roll <= 2) {
        crustType = "ocean";
      } else if (roll === 6) {
        crustType = "continent";
      }
      return createMarker("crust", safePoint, crustType);
    }
    if (markerType === "direction") {
      const padding = getAutoRollMarkerPadding();
      const offset = Math.min(microW, microH) * 0.86;
      let offsetX = safePoint.x + offset;
      if (offsetX > WIDTH - padding) {
        offsetX = safePoint.x - offset;
      }
      if (offsetX < padding) {
        offsetX = safePoint.x + offset;
      }
      const offsetPoint = {
        x: clamp(offsetX, padding, WIDTH - padding),
        y: safePoint.y,
      };
      return createMarker("direction", offsetPoint, rollD6());
    }
    return null;
  }

  function rollAllPlateCenters(markerType) {
    if (getPlateLineCount() === 0) {
      setMessage("Paint boundaries first to auto roll by plate.");
      render();
      return;
    }

    const regions = dedupeAutoRollRegions(detectPlateRegions());
    if (!regions.length) {
      setMessage("No plate regions detected.");
      render();
      return;
    }

    let created = 0;
    let oceanic = 0;
    let mixed = 0;
    let continental = 0;
    const directionCount = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 };

    regions.forEach((region) => {
      const marker = createAutoMarkerFromType(markerType, region.center);
      if (!marker) {
        return;
      }
      state.markers.push(marker);
      recordHistory({
        type: "marker-add",
        marker: cloneMarker(marker),
        index: state.markers.length - 1,
      });
      created += 1;
      if (markerType === "crust") {
        if (marker.value === "ocean") {
          oceanic += 1;
        } else if (marker.value === "continent") {
          continental += 1;
        } else {
          mixed += 1;
        }
      } else if (markerType === "direction") {
        directionCount[marker.value] = (directionCount[marker.value] || 0) + 1;
      }
    });

    if (created === 0) {
      setMessage("No markers were created.");
      render();
      return;
    }

    if (markerType === "crust") {
      setMessage(
        `Rolled crust for ${created} plates (oceanic: ${oceanic}, mixed: ${mixed}, continental: ${continental}).`
      );
    } else if (markerType === "direction") {
      const summary = [1, 2, 3, 4, 5, 6]
        .map((value) => `${directionLabels[value]}:${directionCount[value] || 0}`)
        .join(", ");
      setMessage(`Rolled directions for ${created} plates (${summary}).`);
    } else {
      setMessage(`Auto roll created ${created} markers.`);
    }

    render();
  }

  function isLineTool(tool) {
    return tool === "pencil" || tool === "arrow" || tool === "divide";
  }

  function isSurfacePaintTool(tool) {
    return tool === "paint-continent" || tool === "paint-ocean";
  }

  function getSurfacePaintTypeFromTool(tool, button = 0) {
    if (button === 2) {
      return "erase";
    }
    if (tool === "paint-continent") {
      return "continent";
    }
    if (tool === "paint-ocean") {
      return "ocean";
    }
    return null;
  }

  function clampSurfaceBrushRadius(value) {
    return clamp(value, 1, 15);
  }

  function handleSurfaceRadiusInput(event) {
    const next = clampSurfaceBrushRadius(
      Number(event.target.value) || state.surfaceBrushRadius
    );
    state.surfaceBrushRadius = next;
    render();
  }

  function isMarkerTool(tool) {
    return (
      tool === "crust" ||
      tool === "plate-id" ||
      tool === "direction" ||
      tool === "volcano" ||
      tool === "move-icon"
    );
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
    if (state.tool === "pencil" && previousTool !== "pencil") {
      state.plateStyle = PLATE_STYLES.boundary;
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
    if (!isSurfacePaintTool(state.tool)) {
      state.surfacePaintDrag = null;
    }
    render();
  }

  function togglePencil() {
    setTool("pencil");
  }

  function toggleArrow() {
    setTool("arrow");
  }

  function toggleDivide() {
    setTool("divide");
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

  function togglePaintContinent() {
    setTool("paint-continent");
  }

  function togglePaintOcean() {
    setTool("paint-ocean");
  }

  function toggleVolcano() {
    setTool("volcano");
  }

  function toggleMoveIcon() {
    setTool("move-icon");
  }

  function getPlateStyleLabel(style) {
    if (style === PLATE_STYLES.divergent) {
      return "Divergent";
    }
    if (style === PLATE_STYLES.convergent) {
      return "Convergent";
    }
    if (style === PLATE_STYLES.oblique) {
      return "Oblique";
    }
    return "Boundary";
  }

  function setPlateStyle(style) {
    if (!Object.values(PLATE_STYLES).includes(style)) {
      return;
    }
    if (state.tool !== "pencil") {
      setTool("pencil");
    }
    state.plateStyle = style;
    setMessage(`Plate mode: ${getPlateStyleLabel(style)}.`);
    render();
  }

  function toggleWrapSide(side) {
    if (state.tool !== "pencil") {
      setMessage("Select Paint Boundaries to use turn around.");
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

  function getPlateLineStyle(line) {
    if (!line || (line.lineType || "plate") !== "plate") {
      return PLATE_STYLES.boundary;
    }
    if (Object.values(PLATE_STYLES).includes(line.plateStyle)) {
      return line.plateStyle;
    }
    return PLATE_STYLES.boundary;
  }

  function getDivergentOffsetDirection(line) {
    if (line && (line.divergentSide === -1 || line.divergentSide === 1)) {
      return line.divergentSide;
    }
    const startKey = String((line && line.startKey) || "");
    const endKey = String((line && line.endKey) || "");
    const key = startKey <= endKey
      ? `${startKey}|${endKey}`
      : `${endKey}|${startKey}`;
    let hash = 0;
    for (let i = 0; i < key.length; i += 1) {
      hash = (hash * 31 + key.charCodeAt(i)) & 0x7fffffff;
    }
    return hash % 2 === 0 ? 1 : -1;
  }

  function findPlateConnection(keyA, keyB) {
    const index = state.lines.findIndex((line) => {
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
    if (index < 0) {
      return null;
    }
    return {
      index,
      line: state.lines[index],
    };
  }

  function setPlateConnectionStyle(connection, plateStyle) {
    if (!connection || !connection.line) {
      return false;
    }
    const targetStyle = Object.values(PLATE_STYLES).includes(plateStyle)
      ? plateStyle
      : PLATE_STYLES.boundary;
    const before = { ...connection.line };
    if (getPlateLineStyle(connection.line) === targetStyle) {
      return false;
    }
    connection.line.plateStyle = targetStyle;
    if (
      targetStyle === PLATE_STYLES.divergent &&
      connection.line.divergentSide !== -1 &&
      connection.line.divergentSide !== 1
    ) {
      connection.line.divergentSide = Math.random() < 0.5 ? -1 : 1;
    }
    recordHistory({
      type: "line-update",
      index: connection.index,
      startKey: connection.line.startKey || null,
      endKey: connection.line.endKey || null,
      before,
      after: { ...connection.line },
    });
    return true;
  }

  function updatePlateLineStyleAtPoint(point, plateStyle) {
    if (!point) {
      return { handled: false, updated: false };
    }
    const hitRadius = Math.min(microW, microH) * 0.32;
    const lineHit = findNearestPlateLine(point.x, point.y, hitRadius);
    if (!lineHit) {
      return { handled: false, updated: false };
    }
    const connection = {
      index: lineHit.index,
      line: state.lines[lineHit.index],
    };
    const targetStyle = Object.values(PLATE_STYLES).includes(plateStyle)
      ? plateStyle
      : PLATE_STYLES.boundary;
    const updated = setPlateConnectionStyle(connection, targetStyle);
    const label = getPlateStyleLabel(targetStyle).toLowerCase();
    setMessage(updated ? `Line set to ${label}.` : `Line already ${label}.`);
    return { handled: true, updated };
  }

  function deletePlateLineAtPoint(point) {
    if (!point) {
      return { handled: false };
    }
    const hitRadius = Math.min(microW, microH) * 0.32;
    const lineHit = findNearestPlateLine(point.x, point.y, hitRadius);
    if (!lineHit) {
      return { handled: false };
    }
    removeLine(lineHit.index);
    return { handled: true };
  }

  function randomBetween(min, max) {
    return min + Math.random() * (max - min);
  }

  function createRandomPlateCurve() {
    const knotCount = Math.floor(randomBetween(2, 6));
    const persistence = randomBetween(0.36, 0.72);
    const nodes = [];
    let previousNormal = randomBetween(-1, 1);
    for (let i = 0; i < knotCount; i += 1) {
      const keepDirection = i > 0 && Math.random() < persistence;
      let normal = keepDirection
        ? previousNormal + randomBetween(-0.34, 0.34)
        : randomBetween(-1, 1);
      normal = clamp(normal, -1, 1);
      previousNormal = normal;
      nodes.push({
        t: randomBetween(-0.22, 0.22),
        n: normal,
        a: randomBetween(-1, 1),
      });
    }
    return {
      mode: "tectonic",
      knotCount,
      roughness: randomBetween(0.028, 0.1),
      broadBend: randomBetween(-0.2, 0.2),
      alongScale: randomBetween(0.005, 0.055),
      nodes,
    };
  }

  function buildTectonicCurvePoints(segment, curve) {
    const dx = segment.x2 - segment.x1;
    const dy = segment.y2 - segment.y1;
    const length = Math.hypot(dx, dy);
    if (length <= 0.0001) {
      return [
        { x: segment.x1, y: segment.y1 },
        { x: segment.x2, y: segment.y2 },
      ];
    }
    const tx = dx / length;
    const ty = dy / length;
    const nx = -ty;
    const ny = tx;
    const knotCount = clamp(Math.round(Number(curve.knotCount) || 3), 2, 6);
    const roughness = clamp(Number(curve.roughness) || 0.06, 0.012, 0.15);
    const broadBend = clamp(Number(curve.broadBend) || 0, -0.34, 0.34);
    const alongScale = clamp(Number(curve.alongScale) || 0.03, 0, 0.09);
    const nodes = Array.isArray(curve.nodes) ? curve.nodes : [];
    const points = [{ x: segment.x1, y: segment.y1 }];
    let previousT = 0;

    for (let i = 1; i <= knotCount; i += 1) {
      const node = nodes[i - 1] || {};
      const baseT = i / (knotCount + 1);
      const jitter = clamp(Number(node.t) || 0, -0.32, 0.32);
      let t = baseT + jitter / (knotCount + 1);
      const minT = previousT + 0.06;
      const remaining = knotCount - i + 1;
      const maxT = 1 - remaining * 0.05;
      t = clamp(t, minT, maxT);
      previousT = t;

      const broadWave = Math.sin(t * Math.PI);
      const localNormal = clamp(Number(node.n) || 0, -1, 1);
      const localAlong = clamp(Number(node.a) || 0, -1, 1);
      const normalOffset = length * (broadBend * broadWave + localNormal * roughness);
      const alongOffset = length * localAlong * alongScale;
      points.push({
        x: segment.x1 + dx * t + nx * normalOffset + tx * alongOffset,
        y: segment.y1 + dy * t + ny * normalOffset + ty * alongOffset,
      });
    }

    points.push({ x: segment.x2, y: segment.y2 });
    return points;
  }

  function buildSegmentsFromPoints(points) {
    if (!points || points.length < 2) {
      return [];
    }
    const result = [];
    for (let i = 1; i < points.length; i += 1) {
      result.push({
        x1: points[i - 1].x,
        y1: points[i - 1].y,
        x2: points[i].x,
        y2: points[i].y,
      });
    }
    return result;
  }

  function addPlateConnection(
    start,
    end,
    wrapSide,
    plateStyle = PLATE_STYLES.boundary,
    curve = null
  ) {
    const line = {
      x1: start.x,
      y1: start.y,
      x2: end.x,
      y2: end.y,
      lineType: "plate",
      startKey: start.key,
      endKey: end.key,
      wrapSide: wrapSide || null,
      plateStyle: Object.values(PLATE_STYLES).includes(plateStyle)
        ? plateStyle
        : PLATE_STYLES.boundary,
      curve: curve ? JSON.parse(JSON.stringify(curve)) : createRandomPlateCurve(),
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

  function getPlateCurveDefinition(segment, curve) {
    if (!curve) {
      return null;
    }
    const dx = segment.x2 - segment.x1;
    const dy = segment.y2 - segment.y1;
    const length = Math.hypot(dx, dy);
    if (length <= 0.0001) {
      return null;
    }
    const nx = -dy / length;
    const ny = dx / length;
    const side = curve.side === -1 ? -1 : 1;
    const strength = clamp(Number(curve.strength) || 0.18, 0.05, 0.42);
    const amplitude = length * strength;
    const mode = curve.mode === "s" ? "s" : "arc";
    if (mode === "s") {
      const split = clamp(Number(curve.split) || 0.35, 0.2, 0.45);
      const skew = clamp(Number(curve.skew) || 1, 0.6, 1.4);
      const t1 = split;
      const t2 = 1 - split;
      const base1X = segment.x1 + dx * t1;
      const base1Y = segment.y1 + dy * t1;
      const base2X = segment.x1 + dx * t2;
      const base2Y = segment.y1 + dy * t2;
      return {
        mode,
        c1: {
          x: base1X + nx * amplitude * side,
          y: base1Y + ny * amplitude * side,
        },
        c2: {
          x: base2X - nx * amplitude * side * skew,
          y: base2Y - ny * amplitude * side * skew,
        },
      };
    }
    const bias = clamp(Number(curve.bias) || 0, -0.22, 0.22);
    const t = clamp(0.5 + bias, 0.2, 0.8);
    const baseX = segment.x1 + dx * t;
    const baseY = segment.y1 + dy * t;
    return {
      mode,
      c1: {
        x: baseX + nx * amplitude * side,
        y: baseY + ny * amplitude * side,
      },
    };
  }

  function evaluatePlateCurvePoint(segment, definition, t) {
    const oneMinus = 1 - t;
    if (!definition || definition.mode === "line") {
      return {
        x: segment.x1 + (segment.x2 - segment.x1) * t,
        y: segment.y1 + (segment.y2 - segment.y1) * t,
      };
    }
    if (definition.mode === "s") {
      const c1 = definition.c1;
      const c2 = definition.c2;
      const x =
        oneMinus * oneMinus * oneMinus * segment.x1 +
        3 * oneMinus * oneMinus * t * c1.x +
        3 * oneMinus * t * t * c2.x +
        t * t * t * segment.x2;
      const y =
        oneMinus * oneMinus * oneMinus * segment.y1 +
        3 * oneMinus * oneMinus * t * c1.y +
        3 * oneMinus * t * t * c2.y +
        t * t * t * segment.y2;
      return { x, y };
    }
    const c1 = definition.c1;
    const x =
      oneMinus * oneMinus * segment.x1 +
      2 * oneMinus * t * c1.x +
      t * t * segment.x2;
    const y =
      oneMinus * oneMinus * segment.y1 +
      2 * oneMinus * t * c1.y +
      t * t * segment.y2;
    return { x, y };
  }

  function buildPlateCurvePathData(segment, curve) {
    if (curve && curve.mode === "tectonic") {
      return buildPolylinePathData(buildTectonicCurvePoints(segment, curve));
    }
    const definition = getPlateCurveDefinition(segment, curve);
    if (!definition) {
      return `M ${segment.x1.toFixed(2)} ${segment.y1.toFixed(2)} L ${segment.x2.toFixed(
        2
      )} ${segment.y2.toFixed(2)}`;
    }
    if (definition.mode === "s") {
      return `M ${segment.x1.toFixed(2)} ${segment.y1.toFixed(2)} C ${definition.c1.x.toFixed(
        2
      )} ${definition.c1.y.toFixed(2)} ${definition.c2.x.toFixed(2)} ${definition.c2.y.toFixed(
        2
      )} ${segment.x2.toFixed(2)} ${segment.y2.toFixed(2)}`;
    }
    return `M ${segment.x1.toFixed(2)} ${segment.y1.toFixed(2)} Q ${definition.c1.x.toFixed(
      2
    )} ${definition.c1.y.toFixed(2)} ${segment.x2.toFixed(2)} ${segment.y2.toFixed(2)}`;
  }

  function buildPlateCurveHitSegments(segment, curve) {
    if (curve && curve.mode === "tectonic") {
      return buildSegmentsFromPoints(buildTectonicCurvePoints(segment, curve));
    }
    const definition = getPlateCurveDefinition(segment, curve);
    if (!definition) {
      return [segment];
    }
    const steps = definition.mode === "s" ? 18 : 12;
    const result = [];
    let previous = { x: segment.x1, y: segment.y1 };
    for (let i = 1; i <= steps; i += 1) {
      const t = i / steps;
      const current = evaluatePlateCurvePoint(segment, definition, t);
      result.push({
        x1: previous.x,
        y1: previous.y,
        x2: current.x,
        y2: current.y,
      });
      previous = current;
    }
    return result;
  }

  function buildPlateCurvePoints(segment, curve) {
    if (curve && curve.mode === "tectonic") {
      return buildTectonicCurvePoints(segment, curve);
    }
    const definition = getPlateCurveDefinition(segment, curve);
    if (!definition) {
      return [
        { x: segment.x1, y: segment.y1 },
        { x: segment.x2, y: segment.y2 },
      ];
    }
    const steps = definition.mode === "s" ? 26 : 18;
    const points = [{ x: segment.x1, y: segment.y1 }];
    for (let i = 1; i < steps; i += 1) {
      points.push(evaluatePlateCurvePoint(segment, definition, i / steps));
    }
    points.push({ x: segment.x2, y: segment.y2 });
    return points;
  }

  function buildPolylinePathData(points) {
    if (!points || points.length === 0) {
      return "";
    }
    const commands = points.map(
      (point, index) =>
        `${index === 0 ? "M" : "L"} ${point.x.toFixed(2)} ${point.y.toFixed(2)}`
    );
    return commands.join(" ");
  }

  function buildPolylineMetrics(points) {
    const cumulative = [0];
    let total = 0;
    for (let i = 1; i < points.length; i += 1) {
      total += Math.hypot(
        points[i].x - points[i - 1].x,
        points[i].y - points[i - 1].y
      );
      cumulative.push(total);
    }
    return { cumulative, total };
  }

  function samplePolylineAtDistance(points, metrics, distance) {
    if (!points || points.length === 0) {
      return null;
    }
    if (points.length === 1 || metrics.total <= 0.0001) {
      return { x: points[0].x, y: points[0].y, nx: 0, ny: -1 };
    }
    const target = clamp(distance, 0, metrics.total);
    let index = 0;
    while (
      index < metrics.cumulative.length - 2 &&
      metrics.cumulative[index + 1] < target
    ) {
      index += 1;
    }
    const start = points[index];
    const end = points[index + 1] || start;
    const startDistance = metrics.cumulative[index];
    const endDistance = metrics.cumulative[index + 1] ?? startDistance;
    const segmentLength = endDistance - startDistance;
    const t =
      segmentLength > 0.0001 ? (target - startDistance) / segmentLength : 0;
    const x = start.x + (end.x - start.x) * t;
    const y = start.y + (end.y - start.y) * t;
    let dx = end.x - start.x;
    let dy = end.y - start.y;
    let length = Math.hypot(dx, dy);
    if (length <= 0.0001) {
      const previous = points[Math.max(0, index - 1)];
      const next = points[Math.min(points.length - 1, index + 2)];
      dx = next.x - previous.x;
      dy = next.y - previous.y;
      length = Math.hypot(dx, dy);
    }
    if (length <= 0.0001) {
      return { x, y, nx: 0, ny: -1 };
    }
    return { x, y, nx: -dy / length, ny: dx / length };
  }

  function offsetPolylinePoints(points, offset) {
    if (!points || points.length === 0) {
      return [];
    }
    if (points.length === 1 || Math.abs(offset) <= 0.0001) {
      return points.map((point) => ({ x: point.x, y: point.y }));
    }
    const result = [];
    for (let i = 0; i < points.length; i += 1) {
      const previous = points[Math.max(0, i - 1)];
      const next = points[Math.min(points.length - 1, i + 1)];
      let dx = next.x - previous.x;
      let dy = next.y - previous.y;
      let length = Math.hypot(dx, dy);
      if (length <= 0.0001) {
        result.push({ x: points[i].x, y: points[i].y });
        continue;
      }
      const nx = -dy / length;
      const ny = dx / length;
      result.push({
        x: points[i].x + nx * offset,
        y: points[i].y + ny * offset,
      });
    }
    return result;
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

  function getLineHitSegments(line) {
    const segments = getLineSegments(line);
    if ((line.lineType || "plate") !== "plate") {
      return segments;
    }
    if (!line.curve) {
      return segments;
    }
    return segments.flatMap((segment) => buildPlateCurveHitSegments(segment, line.curve));
  }

  function buildConvergentPathData(segment, curve = null) {
    const basePoints = buildPlateCurvePoints(segment, curve);
    const metrics = buildPolylineMetrics(basePoints);
    const length = metrics.total;
    if (length <= 0.0001) {
      return buildPolylinePathData(basePoints);
    }
    const amplitude = Math.min(microW, microH) * 0.22;
    const step = Math.min(microW, microH) * 0.4;
    const waveCount = Math.max(2, Math.floor(length / step));
    const points = [{ x: basePoints[0].x, y: basePoints[0].y }];
    for (let i = 1; i < waveCount; i += 1) {
      const sample = samplePolylineAtDistance(
        basePoints,
        metrics,
        (length * i) / waveCount
      );
      if (!sample) {
        continue;
      }
      const direction = i % 2 === 0 ? -1 : 1;
      points.push({
        x: sample.x + sample.nx * amplitude * direction,
        y: sample.y + sample.ny * amplitude * direction,
      });
    }
    points.push({
      x: basePoints[basePoints.length - 1].x,
      y: basePoints[basePoints.length - 1].y,
    });
    return buildPolylinePathData(points);
  }

  function buildObliquePathData(segment, curve = null) {
    const basePoints = buildPlateCurvePoints(segment, curve);
    const metrics = buildPolylineMetrics(basePoints);
    const length = metrics.total;
    if (length <= 0.0001) {
      return buildPolylinePathData(basePoints);
    }
    const amplitude = Math.min(microW, microH) * 0.16;
    const step = Math.min(microW, microH) * 0.48;
    const waveCount = Math.max(2, Math.floor(length / step));
    const start = basePoints[0];
    const commands = [`M ${start.x.toFixed(2)} ${start.y.toFixed(2)}`];
    for (let i = 1; i <= waveCount; i += 1) {
      const end = samplePolylineAtDistance(
        basePoints,
        metrics,
        (length * i) / waveCount
      );
      const mid = samplePolylineAtDistance(
        basePoints,
        metrics,
        (length * (i - 0.5)) / waveCount
      );
      if (!end || !mid) {
        continue;
      }
      const direction = i % 2 === 0 ? -1 : 1;
      const ctrlX = mid.x + mid.nx * amplitude * direction;
      const ctrlY = mid.y + mid.ny * amplitude * direction;
      commands.push(
        `Q ${ctrlX.toFixed(2)} ${ctrlY.toFixed(2)} ${end.x.toFixed(2)} ${end.y.toFixed(2)}`
      );
    }
    return commands.join(" ");
  }

  function hasMarks() {
    return (
      state.selectedCells.size > 0 ||
      state.lines.length > 0 ||
      state.markers.length > 0 ||
      state.surfacePaintStamps.length > 0 ||
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

  function handleClearDivides() {
    cancelClearConfirm();
    clearLinesByType("divide", "Divides");
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
    const points = Array.from(state.selectedCells).map((key) => ({
      key,
      weight: getPointWeight(key),
    }));
    state.selectedCells.clear();
    state.pointWeights = {};
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
      setMessage("No cratons to clear.");
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
    setMessage("Cratons cleared.");
    render();
  }

  function clearAllState() {
    state.selectedCells.clear();
    state.pointWeights = {};
    state.lastRoll = null;
    state.currentLine = null;
    state.isDrawing = false;
    state.lines = [];
    state.plateDraft = null;
    state.plateStyle = PLATE_STYLES.boundary;
    state.wrapSide = null;
    state.markers = [];
    state.selectedMarkerId = null;
    state.markerDrag = null;
    state.surfacePaintStamps = [];
    invalidateSurfacePaintCache();
    state.surfacePaintDrag = null;
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
      pointWeights: { ...state.pointWeights },
      lines: state.lines.map((line) => ({ ...line })),
      lastRoll: state.lastRoll ? { ...state.lastRoll } : null,
      plateStyle: state.plateStyle,
      markers: state.markers.map(cloneMarker),
      surfacePaintStamps: state.surfacePaintStamps.map(cloneSurfaceStamp),
      selectedMarkerId: state.selectedMarkerId,
      continentPieces: state.continentPieces.map(cloneContinentPiece),
      selectedContinentId: state.selectedContinentId,
      nextContinentId: state.nextContinentId,
      nextMarkerId: state.nextMarkerId,
    };
  }

  function restoreClearSnapshot(snapshot) {
    state.selectedCells = new Set(snapshot.selectedCells);
    if (snapshot.pointWeights) {
      state.pointWeights = { ...snapshot.pointWeights };
    } else {
      state.pointWeights = {};
      state.selectedCells.forEach((key) => {
        state.pointWeights[key] = 1;
      });
    }
    state.lines = snapshot.lines.map((line) => ({ ...line }));
    state.lastRoll = snapshot.lastRoll ? { ...snapshot.lastRoll } : null;
    state.plateStyle =
      snapshot.plateStyle && Object.values(PLATE_STYLES).includes(snapshot.plateStyle)
        ? snapshot.plateStyle
        : PLATE_STYLES.boundary;
    state.markers = snapshot.markers.map(cloneMarker);
    state.surfacePaintStamps = snapshot.surfacePaintStamps
      ? snapshot.surfacePaintStamps.map(cloneSurfaceStamp)
      : [];
    invalidateSurfacePaintCache();
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
    state.surfacePaintDrag = null;
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
      setMessage("Select a craton to delete.");
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
    const point = getSvgPointInPage(event, false);
    if (!point) {
      return;
    }
    if (state.selectedCells.size === 0) {
      setMessage("Roll points first.");
      render();
      return;
    }
    if (state.plateStyle !== PLATE_STYLES.boundary) {
      const styleResult = updatePlateLineStyleAtPoint(point, state.plateStyle);
      if (styleResult.handled) {
        clearPlateDraft();
      } else {
        clearPlateDraft();
        setMessage(
          `${getPlateStyleLabel(
            state.plateStyle
          )} mode changes existing boundary lines.`
        );
      }
      render();
      return;
    }
    const hit = findNearestPoint(point.x, point.y, plateHitRadius);
    if (!hit) {
      const styleResult = updatePlateLineStyleAtPoint(point, state.plateStyle);
      if (styleResult.handled) {
        clearPlateDraft();
        render();
        return;
      }
      if (state.plateDraft) {
        clearPlateDraft();
        setMessage("Chain ended.");
        render();
      }
      return;
    }

    const hitPoint = buildPlatePoint(hit);
    if (!isPlateConnectablePointKey(hitPoint.key)) {
      setMessage("Hotspot point: this point is reserved and cannot be connected.");
      render();
      return;
    }

    if (
      state.plateDraft &&
      !isPlateConnectablePointKey(state.plateDraft.key)
    ) {
      clearPlateDraft();
      setMessage("Chain ended: hotspot points cannot be connected.");
      render();
      return;
    }

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

    const existingConnection = findPlateConnection(
      state.plateDraft.key,
      hitPoint.key
    );
    if (existingConnection) {
      setMessage("Points already connected. Left click the line to convert it.");
      render();
      return;
    }

    const wrapSide = state.wrapSide;
    addPlateConnection(state.plateDraft, hitPoint, wrapSide, state.plateStyle);
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

  function isStrokeTool(tool) {
    return tool === "arrow" || tool === "divide";
  }

  function handlePointerDown(event) {
    if (state.tool === "pencil") {
      handlePlatePointerDown(event);
      return;
    }
    if (!isStrokeTool(state.tool) || event.button !== 0) {
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
      lineType: state.tool === "divide" ? "divide" : "arrow",
    };
    if (ui.board.setPointerCapture) {
      ui.board.setPointerCapture(event.pointerId);
    }
    render();
  }

  function handlePointerMove(event) {
    if (!isStrokeTool(state.tool) || !state.isDrawing) {
      return;
    }
    updateCurrentLine(event);
  }

  function handlePointerUp(event) {
    if (!isStrokeTool(state.tool) || !state.isDrawing) {
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
    const point = getSvgPointRaw(event);
    if (state.tool === "pencil") {
      if (state.plateStyle === PLATE_STYLES.boundary) {
        const deleteResult = deletePlateLineAtPoint(point);
        if (deleteResult.handled) {
          clearPlateDraft();
          render();
          return;
        }
      } else {
        const resetResult = updatePlateLineStyleAtPoint(point, PLATE_STYLES.boundary);
        if (resetResult.handled) {
          clearPlateDraft();
          render();
          return;
        }
      }
      if (state.plateDraft) {
        clearPlateDraft();
        setMessage("Chain ended.");
      } else {
        if (state.plateStyle === PLATE_STYLES.boundary) {
          setMessage("Right click a plate line to delete it.");
        } else {
          setMessage("Right click a plate line to reset it to boundary.");
        }
      }
      render();
      return;
    }
    if (state.isDrawing) {
      state.isDrawing = false;
      state.currentLine = null;
    }
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
      if (state.tool === "move-icon") {
        event.preventDefault();
        return;
      }
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
      Math.min(microW, microH) *
      (lineType === "arrow" || lineType === "divide" ? 0.3 : 0.2);
    if (length >= minLength) {
      const line = { x1, y1, x2, y2, lineType };
      state.lines.push(line);
      recordHistory({
        type: "line-add",
        line: { ...line },
        index: state.lines.length - 1,
      });
      if (lineType === "arrow") {
        setMessage("Arrow drawn.");
      } else if (lineType === "divide") {
        setMessage("Divide line drawn.");
      } else {
        setMessage("Line drawn.");
      }
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

  function findNearestLine(x, y, radius, filterFn = null) {
    let best = null;
    const radiusSq = radius * radius;
    state.lines.forEach((line, index) => {
      if (filterFn && !filterFn(line, index)) {
        return;
      }
      const segments = getLineHitSegments(line);
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

  function findNearestPlateLine(x, y, radius) {
    return findNearestLine(
      x,
      y,
      radius,
      (line) => (line.lineType || "plate") === "plate"
    );
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
    const weight = getPointWeight(key);
    state.selectedCells.delete(key);
    delete state.pointWeights[key];
    const col = microCol + 1;
    const row = microRow + 1;
    recordHistory({
      type: "erase",
      itemType: "point",
      key,
      weight,
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
      let crustType = "mixed";
      if (roll <= 2) {
        crustType = "ocean";
      } else if (roll === 6) {
        crustType = "continent";
      }
      if (crustType === "ocean") {
        setMessage("Crust: oceanic.");
      } else if (crustType === "continent") {
        setMessage("Crust: continental.");
      } else {
        setMessage("Crust: mixed plate.");
      }
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
    if (marker.type === "crust" && marker.value === "continent") {
      return markerIconSize * 1.5;
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

  function cloneSurfaceStamp(stamp) {
    return {
      x: stamp.x,
      y: stamp.y,
      radius: stamp.radius,
      type: stamp.type,
    };
  }

  function invalidateSurfacePaintCache() {
    surfacePaintCacheDirty = true;
  }

  function ensureSurfacePaintCacheContext() {
    if (!ui.continentCanvas) {
      return null;
    }
    const targetWidth = Math.max(1, ui.continentCanvas.width || WIDTH);
    const targetHeight = Math.max(1, ui.continentCanvas.height || HEIGHT);
    if (!surfacePaintCacheCanvas) {
      surfacePaintCacheCanvas = document.createElement("canvas");
      surfacePaintCacheCtx = surfacePaintCacheCanvas.getContext("2d");
      surfacePaintCacheDirty = true;
    }
    if (!surfacePaintCacheCtx) {
      return null;
    }
    if (
      surfacePaintCacheCanvas.width !== targetWidth ||
      surfacePaintCacheCanvas.height !== targetHeight
    ) {
      surfacePaintCacheCanvas.width = targetWidth;
      surfacePaintCacheCanvas.height = targetHeight;
      surfacePaintCacheDirty = true;
    }
    const scaleX = targetWidth / WIDTH;
    const scaleY = targetHeight / HEIGHT;
    surfacePaintCacheCtx.setTransform(scaleX, 0, 0, scaleY, 0, 0);
    surfacePaintCacheCtx.lineJoin = "round";
    surfacePaintCacheCtx.lineCap = "round";
    return surfacePaintCacheCtx;
  }

  function clearSurfacePaintCache() {
    if (!surfacePaintCacheCanvas || !surfacePaintCacheCtx) {
      return;
    }
    surfacePaintCacheCtx.save();
    surfacePaintCacheCtx.setTransform(1, 0, 0, 1, 0, 0);
    surfacePaintCacheCtx.clearRect(
      0,
      0,
      surfacePaintCacheCanvas.width,
      surfacePaintCacheCanvas.height
    );
    surfacePaintCacheCtx.restore();
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

  function getPointWeight(key) {
    return state.pointWeights[key] || 1;
  }

  function getPointDotRadius(key) {
    return getPointWeight(key) > 1 ? dotRadius * 1.85 : dotRadius;
  }

  function findVolcanoMarkerAtPoint(point) {
    const epsilon = 0.0001;
    return (
      state.markers.find(
        (marker) =>
          marker.type === "volcano" &&
          Math.abs(marker.x - point.x) <= epsilon &&
          Math.abs(marker.y - point.y) <= epsilon
      ) || null
    );
  }

  function convertPointToVolcano(key, microCol, microRow, roll) {
    const beforeWeight = getPointWeight(key);
    if (state.plateDraft && state.plateDraft.key === key) {
      clearPlateDraft();
    }
    state.selectedCells.delete(key);
    delete state.pointWeights[key];
    const point = getPointCenter(microCol, microRow);
    let addedMarker = null;
    if (!findVolcanoMarkerAtPoint(point)) {
      const marker = createMarker("volcano", point, "volcano");
      state.markers.push(marker);
      addedMarker = cloneMarker(marker);
    }
    recordHistory({
      type: "point-hotspot",
      key,
      roll,
      beforeWeight,
      marker: addedMarker,
      markerIndex: addedMarker ? state.markers.length - 1 : null,
    });
    state.lastRoll = roll;
    return {
      key,
      col: roll.col,
      row: roll.row,
      repeated: true,
      hotspot: true,
    };
  }

  function addCell(microCol, microRow, col, row) {
    const key = `${microCol},${microRow}`;
    const roll = { col, row };
    if (state.selectedCells.has(key)) {
      return convertPointToVolcano(key, microCol, microRow, roll);
    }
    state.selectedCells.add(key);
    state.pointWeights[key] = 1;
    recordHistory({ type: "point", key, roll, weight: 1 });
    state.lastRoll = roll;
    return {
      key,
      col,
      row,
      repeated: false,
      hotspot: false,
    };
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
      delete state.pointWeights[entry.key];
      return;
    }
    if (entry.type === "point-hotspot") {
      state.selectedCells.add(entry.key);
      state.pointWeights[entry.key] = entry.beforeWeight || 1;
      if (entry.marker) {
        removeMarkers([{ marker: entry.marker, index: entry.markerIndex ?? -1 }]);
      }
      return;
    }
    if (entry.type === "point-weight") {
      if (state.selectedCells.has(entry.key)) {
        state.pointWeights[entry.key] = entry.beforeWeight || 1;
      }
      return;
    }
    if (entry.type === "line-add") {
      removeLineBySnapshot(entry);
      return;
    }
    if (entry.type === "line-update") {
      undoLineUpdate(entry);
      return;
    }
    if (entry.type === "line") {
      state.lines.pop();
      return;
    }
    if (entry.type === "erase") {
      if (entry.itemType === "point") {
        state.selectedCells.add(entry.key);
        state.pointWeights[entry.key] = entry.weight || 1;
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
      state.pointWeights[entry.key] = entry.weight || 1;
      return;
    }
    if (entry.type === "point-hotspot") {
      state.selectedCells.delete(entry.key);
      delete state.pointWeights[entry.key];
      if (entry.marker) {
        removeMarkers([{ marker: entry.marker, index: entry.markerIndex ?? -1 }]);
        insertMarkerAt(
          Number.isInteger(entry.markerIndex) ? entry.markerIndex : state.markers.length,
          entry.marker
        );
      }
      return;
    }
    if (entry.type === "point-weight") {
      if (state.selectedCells.has(entry.key)) {
        state.pointWeights[entry.key] = entry.afterWeight || 1;
      }
      return;
    }
    if (entry.type === "line-add") {
      insertLineAt(entry.index, entry.line);
      return;
    }
    if (entry.type === "line-update") {
      redoLineUpdate(entry);
      return;
    }
    if (entry.type === "line") {
      return;
    }
    if (entry.type === "erase") {
      if (entry.itemType === "point") {
        state.selectedCells.delete(entry.key);
        delete state.pointWeights[entry.key];
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
      entry.points.forEach((point) => {
        const key = typeof point === "string" ? point : point.key;
        const weight =
          typeof point === "string" ? 1 : point.weight || 1;
        state.selectedCells.add(key);
        state.pointWeights[key] = weight;
      });
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
      entry.points.forEach((point) => {
        const key = typeof point === "string" ? point : point.key;
        state.selectedCells.delete(key);
        delete state.pointWeights[key];
      });
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

  function linesMatchGeometry(a, b) {
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
      (a.wrapSide || null) === (b.wrapSide || null) &&
      (a.startKey || null) === (b.startKey || null) &&
      (a.endKey || null) === (b.endKey || null)
    );
  }

  function linesMatch(a, b) {
    return (
      linesMatchGeometry(a, b) &&
      getPlateLineStyle(a) === getPlateLineStyle(b)
    );
  }

  function findLineUpdateTarget(entry) {
    if (entry.startKey && entry.endKey) {
      const match = findPlateConnection(entry.startKey, entry.endKey);
      if (match) {
        return match;
      }
    }
    if (
      Number.isInteger(entry.index) &&
      entry.index >= 0 &&
      entry.index < state.lines.length
    ) {
      return { index: entry.index, line: state.lines[entry.index] };
    }
    const snapshot = entry.after || entry.before;
    if (!snapshot) {
      return null;
    }
    const matchIndex = state.lines.findIndex((line) =>
      linesMatchGeometry(line, snapshot)
    );
    if (matchIndex < 0) {
      return null;
    }
    return { index: matchIndex, line: state.lines[matchIndex] };
  }

  function undoLineUpdate(entry) {
    const target = findLineUpdateTarget(entry);
    if (!target || !entry.before) {
      return;
    }
    state.lines[target.index] = { ...target.line, ...entry.before };
  }

  function redoLineUpdate(entry) {
    const target = findLineUpdateTarget(entry);
    if (!target || !entry.after) {
      return;
    }
    state.lines[target.index] = { ...target.line, ...entry.after };
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
      if (
        (entry.type === "point" || entry.type === "point-weight") &&
        state.selectedCells.has(entry.key)
      ) {
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
    if (overlayRenderFrame) {
      window.cancelAnimationFrame(overlayRenderFrame);
      overlayRenderFrame = null;
    }
    let toolModeLabel = "Roll";
    if (state.tool === "pencil") {
      toolModeLabel = "Paint Boundaries";
    } else if (state.tool === "arrow") {
      toolModeLabel = "Paint arrows";
    } else if (state.tool === "divide") {
      toolModeLabel = "Paint Divides";
    } else if (state.tool === "crust") {
      toolModeLabel = "Roll crust";
    } else if (state.tool === "plate-id") {
      toolModeLabel = "Identify plates";
    } else if (state.tool === "direction") {
      toolModeLabel = "Roll directions";
    } else if (state.tool === "continent") {
      toolModeLabel = "Paint Cratons";
    } else if (state.tool === "paint-continent") {
      toolModeLabel = "Paint Continents";
    } else if (state.tool === "paint-ocean") {
      toolModeLabel = "Paint Oceans";
    } else if (state.tool === "volcano") {
      toolModeLabel = "Add volcanoes";
    } else if (state.tool === "move-icon") {
      toolModeLabel = "Move icon";
    }
    if (ui.toolMode) {
      ui.toolMode.textContent = toolModeLabel;
    }
    if (ui.lastSector) {
      ui.lastSector.textContent = formatLastX();
    }
    if (ui.lastCell) {
      ui.lastCell.textContent = formatLastY();
    }
    if (ui.filled) {
      ui.filled.textContent = state.selectedCells.size;
    }
    const markerCounts = countMarkersByType();
    const arrowCount = state.lines.filter(
      (line) => (line.lineType || "plate") === "arrow"
    ).length;
    const divideCount = state.lines.filter(
      (line) => (line.lineType || "plate") === "divide"
    ).length;
    const plateBoundaryCount = getPlateLineCount();
    ui.rollBtn.disabled = state.selectedCells.size >= TOTAL_CELLS;
    ui.rollSixBtn.disabled = state.selectedCells.size >= TOTAL_CELLS;
    if (ui.rollAllCrustsBtn) {
      ui.rollAllCrustsBtn.disabled = plateBoundaryCount === 0;
    }
    if (ui.rollAllDirectionsBtn) {
      ui.rollAllDirectionsBtn.disabled = plateBoundaryCount === 0;
    }
    if (ui.connectPointsBtn) {
      ui.connectPointsBtn.disabled =
        getConnectablePlatePointCount() < 2 || plateBoundaryCount > 0;
    }
    ui.clearBtn.disabled = !hasMarks();
    ui.clearBtn.textContent = state.clearConfirm ? "Confirm restart" : "Restart";
    ui.undoBtn.disabled = state.history.length === 0;
    ui.redoBtn.disabled = state.future.length === 0;
    ui.clearPointsBtn.disabled = state.selectedCells.size === 0;
    ui.clearArrowsBtn.disabled = arrowCount === 0;
    if (ui.clearDividesBtn) {
      ui.clearDividesBtn.disabled = divideCount === 0;
    }
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
    if (ui.divideBtn) {
      ui.divideBtn.setAttribute("aria-pressed", state.tool === "divide");
      ui.divideBtn.classList.toggle("active", state.tool === "divide");
    }
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
    if (ui.moveIconBtn) {
      ui.moveIconBtn.setAttribute("aria-pressed", state.tool === "move-icon");
      ui.moveIconBtn.classList.toggle("active", state.tool === "move-icon");
    }
    if (ui.paintContinentBtn) {
      ui.paintContinentBtn.setAttribute(
        "aria-pressed",
        state.tool === "paint-continent"
      );
      ui.paintContinentBtn.classList.toggle(
        "active",
        state.tool === "paint-continent"
      );
    }
    if (ui.paintOceanBtn) {
      ui.paintOceanBtn.setAttribute("aria-pressed", state.tool === "paint-ocean");
      ui.paintOceanBtn.classList.toggle("active", state.tool === "paint-ocean");
    }
    if (ui.surfaceRadiusSlider) {
      const roundedRadius = Math.round(clampSurfaceBrushRadius(state.surfaceBrushRadius));
      ui.surfaceRadiusSlider.value = String(roundedRadius);
      if (ui.surfaceRadiusValue) {
        ui.surfaceRadiusValue.textContent = String(roundedRadius);
      }
    }
    const plateModeVisible = state.tool === "pencil";
    if (ui.plateModePanel) {
      ui.plateModePanel.hidden = !plateModeVisible;
    }
    if (ui.surfacePaintPanel) {
      ui.surfacePaintPanel.hidden = false;
    }
    if (ui.plateBoundaryBtn) {
      const selected = state.plateStyle === PLATE_STYLES.boundary;
      ui.plateBoundaryBtn.setAttribute("aria-pressed", selected);
      ui.plateBoundaryBtn.classList.toggle("active", selected);
    }
    if (ui.plateDivergentBtn) {
      const selected = state.plateStyle === PLATE_STYLES.divergent;
      ui.plateDivergentBtn.setAttribute("aria-pressed", selected);
      ui.plateDivergentBtn.classList.toggle("active", selected);
    }
    if (ui.plateConvergentBtn) {
      const selected = state.plateStyle === PLATE_STYLES.convergent;
      ui.plateConvergentBtn.setAttribute("aria-pressed", selected);
      ui.plateConvergentBtn.classList.toggle("active", selected);
    }
    if (ui.plateObliqueBtn) {
      const selected = state.plateStyle === PLATE_STYLES.oblique;
      ui.plateObliqueBtn.setAttribute("aria-pressed", selected);
      ui.plateObliqueBtn.classList.toggle("active", selected);
    }
    if (ui.plateModeHint) {
      if (state.plateStyle === PLATE_STYLES.divergent) {
        ui.plateModeHint.textContent =
          "Left click a boundary line to keep the main line and add one parallel line on one side. Right click resets a plate line to boundary.";
      } else if (state.plateStyle === PLATE_STYLES.convergent) {
        ui.plateModeHint.textContent =
          "Left click a boundary line to add a convergent zigzag line. Right click resets a plate line to boundary.";
      } else if (state.plateStyle === PLATE_STYLES.oblique) {
        ui.plateModeHint.textContent =
          "Left click a boundary line to add an oblique curved zigzag. Right click resets a plate line to boundary.";
      } else {
        ui.plateModeHint.textContent =
          "Boundary mode: click points to connect in a chain. Volcano hotspots are markers and cannot be connected. Right click a boundary line deletes it.";
      }
    }
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
    ui.deleteContinentBtn.disabled = !state.selectedContinentId;
    ui.board.classList.toggle("pencil-active", isLineTool(state.tool));
    document.body.dataset.tool = state.tool;
    ui.board.innerHTML = buildSvg();
    drawContinents();
    drawMarkers();
    scheduleRightPanelOffsetUpdate();
    if (ui.continentCanvas && !state.continentDrag && !state.surfacePaintDrag) {
      if (state.tool === "continent" || isSurfacePaintTool(state.tool)) {
        ui.continentCanvas.style.cursor = "crosshair";
      } else {
        ui.continentCanvas.style.cursor = "default";
      }
    }
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
    const arrowTailXText = Math.max(0.7, arrowSize * 0.16).toFixed(2);
    const arrowTopText = Math.max(0.7, arrowSize * 0.16).toFixed(2);
    const arrowBottomText = (arrowSize - Math.max(0.7, arrowSize * 0.16)).toFixed(2);
    const arrowTipXText = (arrowSize - Math.max(0.6, arrowSize * 0.1)).toFixed(2);
    const arrowStrokeText = Math.max(0.75, arrowSize * 0.17).toFixed(2);

    parts.push(
      `<defs><marker id="arrow-head" markerUnits="userSpaceOnUse" markerWidth="${arrowSizeText}" markerHeight="${arrowSizeText}" viewBox="0 0 ${arrowSizeText} ${arrowSizeText}" refX="${arrowSizeText}" refY="${arrowHalfText}" orient="auto"><path d="M ${arrowTailXText} ${arrowTopText} L ${arrowTipXText} ${arrowHalfText} L ${arrowTailXText} ${arrowBottomText}" fill="none" stroke="var(--ink)" stroke-width="${arrowStrokeText}" stroke-linecap="round" stroke-linejoin="round" /></marker></defs>`
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
      const segments = getLineSegments(line);
      if (lineType === "arrow") {
        segments.forEach((segment) => {
          parts.push(
            `<line class="arrow-line" marker-end="url(#arrow-head)" x1="${segment.x1.toFixed(
              2
            )}" y1="${segment.y1.toFixed(2)}" x2="${segment.x2.toFixed(
              2
            )}" y2="${segment.y2.toFixed(2)}" />`
          );
        });
        return;
      }
      if (lineType === "divide") {
        segments.forEach((segment) => {
          parts.push(
            `<line class="divide-line" x1="${segment.x1.toFixed(2)}" y1="${segment.y1.toFixed(
              2
            )}" x2="${segment.x2.toFixed(2)}" y2="${segment.y2.toFixed(2)}" />`
          );
        });
        return;
      }
      const plateStyle = getPlateLineStyle(line);
      segments.forEach((segment) => {
        const curvePoints = buildPlateCurvePoints(segment, line.curve);
        const basePathData = buildPlateCurvePathData(segment, line.curve);
        if (plateStyle === PLATE_STYLES.divergent) {
          const offset = Math.min(microW, microH) * 0.19;
          const side = getDivergentOffsetDirection(line);
          const parallelPath = buildPolylinePathData(
            offsetPolylinePoints(curvePoints, offset * side)
          );
          parts.push(
            `<path class="pencil-line plate-boundary" fill="none" d="${basePathData}" />`
          );
          parts.push(
            `<path class="pencil-line plate-divergent" fill="none" d="${parallelPath}" />`
          );
          return;
        }
        parts.push(
          `<path class="${
            plateStyle === PLATE_STYLES.convergent ||
            plateStyle === PLATE_STYLES.oblique
              ? "pencil-line plate-convergent-base"
              : "pencil-line plate-boundary"
          }" fill="none" d="${basePathData}" />`
        );
        if (plateStyle === PLATE_STYLES.convergent) {
          parts.push(
            `<path class="pencil-line plate-convergent-zigzag" fill="none" d="${buildConvergentPathData(
              segment,
              line.curve
            )}" />`
          );
        } else if (plateStyle === PLATE_STYLES.oblique) {
          parts.push(
            `<path class="pencil-line plate-oblique-curve" fill="none" d="${buildObliquePathData(
              segment,
              line.curve
            )}" />`
          );
        }
      });
    });

    if (state.currentLine) {
      const lineType = state.currentLine.lineType || "plate";
      const lineClass =
        lineType === "arrow"
          ? "arrow-line"
          : lineType === "divide"
            ? "divide-line"
            : "pencil-line";
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
      const radius = getPointDotRadius(key);
      parts.push(
        `<circle class="cell-dot" cx="${cx.toFixed(2)}" cy="${cy.toFixed(
          2
        )}" r="${radius.toFixed(2)}" />`
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
    scheduleRightPanelOffsetUpdate();
    updateOverlayCanvasSize();
    drawContinents();
    scheduleEdgeControlsLayout();
  }

  function scheduleOverlayRender() {
    if (overlayRenderFrame) {
      return;
    }
    overlayRenderFrame = window.requestAnimationFrame(() => {
      overlayRenderFrame = null;
      drawContinents();
    });
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

  function scheduleRightPanelOffsetUpdate() {
    if (rightPanelOffsetFrame) {
      return;
    }
    rightPanelOffsetFrame = window.requestAnimationFrame(() => {
      rightPanelOffsetFrame = null;
      updateRightPanelOffset();
    });
  }

  function measurePanelHeight(panel) {
    if (!panel) {
      return 0;
    }
    if (!panel.hidden) {
      return panel.getBoundingClientRect().height;
    }
    const previousHidden = panel.hidden;
    const previousVisibility = panel.style.visibility;
    const previousPointerEvents = panel.style.pointerEvents;
    panel.hidden = false;
    panel.style.visibility = "hidden";
    panel.style.pointerEvents = "none";
    const height = panel.getBoundingClientRect().height;
    panel.hidden = previousHidden;
    panel.style.visibility = previousVisibility;
    panel.style.pointerEvents = previousPointerEvents;
    return height;
  }

  function updateRightPanelOffset() {
    const panelGap = 14;
    const autoHeight = ui.autoRollPanel ? measurePanelHeight(ui.autoRollPanel) : 0;
    const plateOffset = autoHeight > 0 ? Math.round(autoHeight + panelGap) : 0;
    let surfaceOffset = plateOffset;
    if (ui.plateModePanel) {
      const plateHeight = measurePanelHeight(ui.plateModePanel);
      surfaceOffset += Math.round(plateHeight + panelGap);
    }
    document.documentElement.style.setProperty(
      "--right-tertiary-offset",
      `${plateOffset}px`
    );
    document.documentElement.style.setProperty(
      "--right-secondary-offset",
      `${surfaceOffset}px`
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
    overlayUnitsPerCssPixel = WIDTH / rect.width;
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

  function applyDivideCanvasStrokeStyle(context) {
    const unitsPerCssPx = Math.max(0.0001, overlayUnitsPerCssPixel || 1);
    context.lineWidth = divideLineWidth * unitsPerCssPx;
    context.strokeStyle = divideStroke;
    context.setLineDash(divideDashPattern.map((value) => value * unitsPerCssPx));
    context.lineCap = "round";
    context.lineJoin = "round";
  }

  function handleContinentContextMenu(event) {
    if (isSurfacePaintTool(state.tool)) {
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    if (state.tool === "continent") {
      event.stopPropagation();
      event.preventDefault();
    }
  }

  function handleContinentPointerDown(event) {
    if (isSurfacePaintTool(state.tool)) {
      handleSurfacePaintPointerDown(event);
      return;
    }
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
      setMessage("Craton draft cancelled.");
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
    if (isSurfacePaintTool(state.tool)) {
      handleSurfacePaintPointerMove(event);
      return;
    }
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
    if (isSurfacePaintTool(state.tool)) {
      handleSurfacePaintPointerUp(event);
      return;
    }
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

  function handleSurfacePaintPointerDown(event) {
    cancelClearConfirm();
    if (event.button !== 0 && event.button !== 2) {
      return;
    }
    if (event.button === 2) {
      event.preventDefault();
      event.stopPropagation();
    }
    const point = getSvgPointInPage(event, true);
    if (!point) {
      return;
    }
    const type = getSurfacePaintTypeFromTool(state.tool, event.button);
    if (!type) {
      return;
    }
    const drag = {
      paintType: type,
      lastPoint: point,
      stamps: [],
    };
    state.surfacePaintDrag = drag;
    addSurfacePaintStamp(point, type, state.surfaceBrushRadius, drag.stamps);
    if (ui.continentCanvas && ui.continentCanvas.setPointerCapture) {
      ui.continentCanvas.setPointerCapture(event.pointerId);
    }
    if (ui.continentCanvas) {
      ui.continentCanvas.style.cursor = "crosshair";
    }
    scheduleOverlayRender();
  }

  function handleSurfacePaintPointerMove(event) {
    if (!state.surfacePaintDrag) {
      return;
    }
    const point = getSvgPointInPage(event, true);
    if (!point) {
      return;
    }
    const type = state.surfacePaintDrag.paintType;
    if (!type) {
      return;
    }
    addSurfacePaintSegment(
      state.surfacePaintDrag.lastPoint,
      point,
      type,
      state.surfaceBrushRadius,
      state.surfacePaintDrag.stamps
    );
    state.surfacePaintDrag.lastPoint = point;
    scheduleOverlayRender();
  }

  function handleSurfacePaintPointerUp(event) {
    if (!state.surfacePaintDrag) {
      return;
    }
    if (ui.continentCanvas && ui.continentCanvas.releasePointerCapture) {
      try {
        ui.continentCanvas.releasePointerCapture(event.pointerId);
      } catch (error) {
        // Ignore if pointer capture was not set.
      }
    }
    const stamps = state.surfacePaintDrag.stamps || [];
    if (stamps.length > 0) {
      if (state.surfacePaintDrag.paintType === "erase") {
        setMessage("Surface erased.");
      } else if (state.tool === "paint-continent") {
        setMessage("Continent paint applied.");
      } else if (state.tool === "paint-ocean") {
        setMessage("Ocean paint applied.");
      }
    }
    state.surfacePaintDrag = null;
    if (ui.continentCanvas) {
      ui.continentCanvas.style.cursor = "crosshair";
    }
    render();
  }

  function addSurfacePaintSegment(from, to, type, radius, output) {
    const dx = to.x - from.x;
    const dy = to.y - from.y;
    const length = Math.hypot(dx, dy);
    if (length <= 0.001) {
      return;
    }
    const step = Math.max(0.9, radius * 0.38);
    const steps = Math.max(1, Math.ceil(length / step));
    for (let i = 1; i <= steps; i += 1) {
      const t = i / steps;
      addSurfacePaintStamp(
        {
          x: from.x + dx * t,
          y: from.y + dy * t,
        },
        type,
        radius,
        output
      );
    }
  }

  function addSurfacePaintStamp(point, type, radius, output) {
    const stamp = {
      x: clamp(point.x, 0, WIDTH),
      y: clamp(point.y, 0, HEIGHT),
      radius: clampSurfaceBrushRadius(radius),
      type,
    };
    state.surfacePaintStamps.push(stamp);
    applySurfaceStampToCache(stamp);
    if (output) {
      output.push(cloneSurfaceStamp(stamp));
    }
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
    setMessage("Cratons: click points, then click the first point again to close. Right click cancels.");
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
    setMessage("Craton created.");
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
    setMessage("Craton deleted.");
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
    drawSurfacePaint(ui.continentCtx);

    state.continentPieces.forEach((piece) => {
      const isSelected = piece.id === state.selectedContinentId;
      drawContinentPiece(ui.continentCtx, piece);
      if (isSelected && state.tool === "continent") {
        drawContinentNodes(ui.continentCtx, piece);
      }
    });

    drawContinentDraft(ui.continentCtx);
    updateContinentHint();
  }

  function drawSingleSurfacePaintStamp(context, stamp) {
    const fill = surfacePaintColors[stamp.type];
    const radius = clampSurfaceBrushRadius(stamp.radius || state.surfaceBrushRadius);
    const left = clamp(stamp.x - radius, 0, WIDTH);
    const top = clamp(stamp.y - radius, 0, HEIGHT);
    const right = clamp(stamp.x + radius, 0, WIDTH);
    const bottom = clamp(stamp.y + radius, 0, HEIGHT);

    context.save();
    context.beginPath();
    context.arc(stamp.x, stamp.y, radius, 0, Math.PI * 2);
    context.clip();
    context.clearRect(left, top, Math.max(0, right - left), Math.max(0, bottom - top));
    context.restore();

    if (!fill) {
      return;
    }
    context.save();
    context.beginPath();
    context.arc(stamp.x, stamp.y, radius, 0, Math.PI * 2);
    context.fillStyle = fill;
    context.fill();
    context.restore();
  }

  function rebuildSurfacePaintCache() {
    const cacheCtx = ensureSurfacePaintCacheContext();
    if (!cacheCtx) {
      surfacePaintCacheDirty = true;
      return false;
    }
    clearSurfacePaintCache();
    state.surfacePaintStamps.forEach((stamp) => {
      drawSingleSurfacePaintStamp(cacheCtx, stamp);
    });
    surfacePaintCacheDirty = false;
    return true;
  }

  function applySurfaceStampToCache(stamp) {
    const cacheCtx = ensureSurfacePaintCacheContext();
    if (!cacheCtx) {
      surfacePaintCacheDirty = true;
      return;
    }
    if (surfacePaintCacheDirty) {
      rebuildSurfacePaintCache();
      return;
    }
    drawSingleSurfacePaintStamp(cacheCtx, stamp);
  }

  function drawSurfacePaintFromCache(context) {
    if (context !== ui.continentCtx || !state.surfacePaintStamps.length) {
      return false;
    }
    if (surfacePaintCacheDirty && !rebuildSurfacePaintCache()) {
      return false;
    }
    if (!surfacePaintCacheCanvas) {
      return false;
    }
    context.drawImage(surfacePaintCacheCanvas, 0, 0, WIDTH, HEIGHT);
    return true;
  }

  function drawSurfacePaint(context) {
    if (!state.surfacePaintStamps.length) {
      return;
    }
    if (drawSurfacePaintFromCache(context)) {
      return;
    }
    state.surfacePaintStamps.forEach((stamp) => {
      drawSingleSurfacePaintStamp(context, stamp);
    });
  }

  function drawContinentPiece(context, piece) {
    const centers = getVisibleWrappedCenters(piece);
    centers.forEach((center) => {
      context.save();
      context.translate(center.x, center.y);
      context.rotate(piece.rotation);
      traceContinentPath(context, piece.points);
      applyDivideCanvasStrokeStyle(context);
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
    applyDivideCanvasStrokeStyle(context);
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
        "Cratons: click to add nodes. Close by clicking the first node. Closed cratons draw dashed outlines (no fill).";
      return;
    }
    if (state.tool === "paint-continent") {
      ui.continentHint.textContent =
        "Paint Continents: drag to paint soft brown zones. Right click draws eraser brush.";
      return;
    }
    if (state.tool === "paint-ocean") {
      ui.continentHint.textContent =
        "Paint Oceans: drag to paint soft blue zones. Right click draws eraser brush.";
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
    if (state.tool === "move-icon") {
      ui.continentHint.textContent =
        "Move icon: drag any marker to reposition it. Right click does not delete in this mode.";
      return;
    }
    if (state.tool === "pencil") {
      if (state.plateStyle === PLATE_STYLES.divergent) {
        ui.continentHint.textContent =
          "Divergent mode: left click a plate line to keep the main line and add one parallel line. Right click resets to boundary.";
      } else if (state.plateStyle === PLATE_STYLES.convergent) {
        ui.continentHint.textContent =
          "Convergent mode: left click a plate line to add zigzag. Right click a plate line to reset boundary.";
      } else if (state.plateStyle === PLATE_STYLES.oblique) {
        ui.continentHint.textContent =
          "Oblique mode: left click a plate line to add curved zigzag hills. Right click resets to boundary.";
      } else {
        ui.continentHint.textContent =
          "Boundary mode: click points to connect in a chain. Volcano hotspots are markers and cannot be connected. Right click a boundary line deletes it.";
      }
      return;
    }
    if (state.tool === "arrow") {
      ui.continentHint.textContent = "Click and drag to draw arrows.";
      return;
    }
    if (state.tool === "divide") {
      ui.continentHint.textContent =
        "Click and drag to draw dashed divide lines.";
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

