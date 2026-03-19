let pdfDoc = null;
let pageNum = 1;
const BASE_RENDER_SCALE = Math.max(2, window.devicePixelRatio || 1);
let scale = BASE_RENDER_SCALE;
let canvas = null;
let ctx = null;
let selection = null;
let isSelecting = false;
let startX, startY;
let currentMode = null; // 'image' or 'text'
let pdfUrl = null;
let sourcePdfOverride = null;

// editing state
let placedItems = [];   // array of {type:'text'|'image', x,y,w,h, ...}
let editingItem = null; // reference to item being edited

// keep track of all applied edits so they can be re‑opened and redrawn
let edits = [];
let editingIndex = null; // index of entry being edited, null for new
let previewEdit = null;  // temporary preview drawn above committed edits
let selectedImageDataUrl = null;
let selectedImageObj = null;
let selectionBgColor = 'rgb(255, 255, 255)';
let imageMouseTransformState = null;
let redrawRequestId = 0;
let redrawRafId = null;
let pageBaseCanvas = null;
let pageBaseCtx = null;
let renderedBasePageNum = null;
let renderedBaseScale = null;
let renderGeneration = 0;
let hideSelectedAppliedImageDuringEdit = false;
let textLinesModeEnabled = false;
let viewZoom = 1;
let pendingZoomFocus = null;
let basePageWidth = 0;
let basePageHeight = 0;
let fitZoomBase = 1;
const AUTO_FIT_MIN_FONT_SIZE = 12;
const DEFAULT_IMAGE_TRANSFORM = {
    scale: 1,
    offsetX: 0,
    offsetY: 0,
    cropLeft: 0,
    cropRight: 0,
    cropTop: 0,
    cropBottom: 0
};
let currentImageTransform = { ...DEFAULT_IMAGE_TRANSFORM };
const DEFAULT_IMAGE_FILTER = {
    type: 'none',
    intensity: 100
};
const PDF_IMAGE_EXPORT_UPSCALE = 3;
const PDF_IMAGE_EXPORT_MAX_PIXELS = 22000000;
let currentImageFilter = { ...DEFAULT_IMAGE_FILTER };

function normalizeImageFilter(filter = {}) {
    const merged = { ...DEFAULT_IMAGE_FILTER, ...(filter || {}) };
    const allowedTypes = ['none', 'grayscale', 'sepia', 'sepia-magenta'];
    return {
        type: allowedTypes.includes(merged.type) ? merged.type : 'none',
        intensity: clamp(Number(merged.intensity) || 0, 0, 100)
    };
}

function buildCanvasFilter(filter = {}) {
    const normalized = normalizeImageFilter(filter);
    if (normalized.type === 'none' || normalized.intensity <= 0) {
        return 'none';
    }
    const ratio = normalized.intensity / 100;
    switch (normalized.type) {
        case 'grayscale':
            return `grayscale(${normalized.intensity}%)`;
        case 'sepia':
            return `sepia(${normalized.intensity}%)`;
        case 'sepia-magenta': {
            const ratio = normalized.intensity / 100;
            const hue = -35 * ratio;
            const sat = 100 + (150 * ratio);
            return `sepia(${normalized.intensity}%) hue-rotate(${hue}deg) saturate(${sat}%)`;
        }
        default:
            return 'none';
    }
}

function normalizeImageTransform(transform = {}) {
    const merged = { ...DEFAULT_IMAGE_TRANSFORM, ...(transform || {}) };
    const normalized = {
        scale: clamp(Number(merged.scale) || 1, 0.2, 4),
        offsetX: clamp(Number(merged.offsetX) || 0, -5000, 5000),
        offsetY: clamp(Number(merged.offsetY) || 0, -5000, 5000),
        cropLeft: clamp(Number(merged.cropLeft) || 0, 0, 0.9),
        cropRight: clamp(Number(merged.cropRight) || 0, 0, 0.9),
        cropTop: clamp(Number(merged.cropTop) || 0, 0, 0.9),
        cropBottom: clamp(Number(merged.cropBottom) || 0, 0, 0.9)
    };

    const sumX = normalized.cropLeft + normalized.cropRight;
    if (sumX > 0.9) {
        const factor = 0.9 / sumX;
        normalized.cropLeft *= factor;
        normalized.cropRight *= factor;
    }

    const sumY = normalized.cropTop + normalized.cropBottom;
    if (sumY > 0.9) {
        const factor = 0.9 / sumY;
        normalized.cropTop *= factor;
        normalized.cropBottom *= factor;
    }

    return normalized;
}

// helper to paint the current edits back onto the canvas
function clamp(value, min, max) {
    return Math.max(min, Math.min(max, value));
}

function getScaleToBase() {
    return BASE_RENDER_SCALE / Math.max(0.0001, scale);
}

function getScaleFromBase() {
    return Math.max(0.0001, scale) / BASE_RENDER_SCALE;
}

function canvasRectToBaseRect(rect) {
    const f = getScaleToBase();
    return {
        x: rect.x * f,
        y: rect.y * f,
        w: rect.w * f,
        h: rect.h * f
    };
}

function baseRectToCanvasRect(rect) {
    const f = getScaleFromBase();
    return {
        x: rect.x * f,
        y: rect.y * f,
        w: rect.w * f,
        h: rect.h * f
    };
}

function getEntryBaseRect(entry) {
    if (!entry) return null;
    if (entry.coordSpace === 'base') {
        return { x: entry.x, y: entry.y, w: entry.w, h: entry.h };
    }
    return canvasRectToBaseRect(entry);
}

function fontSizeCanvasToBase(fontSize) {
    const size = Number(fontSize) || 16;
    return size * getScaleToBase();
}

function fontSizeBaseToCanvas(fontSize) {
    const size = Number(fontSize) || 16;
    return size * getScaleFromBase();
}

function mapTextLinesCanvasToBase(lines = []) {
    return (Array.isArray(lines) ? lines : []).map((line) => ({
        ...line,
        fontSize: fontSizeCanvasToBase(line && line.fontSize)
    }));
}

function mapTextLinesBaseToCanvas(lines = []) {
    return (Array.isArray(lines) ? lines : []).map((line) => ({
        ...line,
        fontSize: fontSizeBaseToCanvas(line && line.fontSize)
    }));
}

function getEntryCanvasRect(entry) {
    if (!entry) return null;
    if (entry.coordSpace === 'base') {
        return baseRectToCanvasRect(entry);
    }
    return { x: entry.x, y: entry.y, w: entry.w, h: entry.h };
}

function transformBaseToCanvas(transform = {}) {
    const normalized = normalizeImageTransform(transform || {});
    const f = getScaleFromBase();
    return normalizeImageTransform({
        ...normalized,
        offsetX: normalized.offsetX * f,
        offsetY: normalized.offsetY * f
    });
}

function transformCanvasToBase(transform = {}) {
    const normalized = normalizeImageTransform(transform || {});
    const f = getScaleToBase();
    return normalizeImageTransform({
        ...normalized,
        offsetX: normalized.offsetX * f,
        offsetY: normalized.offsetY * f
    });
}

function getEntryCanvasImageTransform(entry) {
    if (!entry) return normalizeImageTransform({});
    if (entry.coordSpace === 'base') {
        return transformBaseToCanvas(entry.imageTransform || {});
    }
    return normalizeImageTransform(entry.imageTransform || {});
}

function getEntryBaseImageTransform(entry) {
    if (!entry) return normalizeImageTransform({});
    if (entry.coordSpace === 'base') {
        return normalizeImageTransform(entry.imageTransform || {});
    }
    return transformCanvasToBase(entry.imageTransform || {});
}

function migrateLegacyEditsToBase() {
    if (!Array.isArray(edits)) return;
    edits.forEach((entry) => {
        if (!entry || entry.coordSpace === 'base') return;
        const baseRect = canvasRectToBaseRect(entry);
        entry.x = baseRect.x;
        entry.y = baseRect.y;
        entry.w = baseRect.w;
        entry.h = baseRect.h;
        if (entry.type === 'text') {
            if (Array.isArray(entry.lines) && entry.lines.length > 0) {
                entry.lines = mapTextLinesCanvasToBase(entry.lines);
            }
            if (entry.fontSize) {
                entry.fontSize = fontSizeCanvasToBase(entry.fontSize);
            }
        }
        if (entry.type === 'image' && entry.imageTransform) {
            entry.imageTransform = transformCanvasToBase(entry.imageTransform);
        }
        entry.coordSpace = 'base';
    });
}

function getImageRenderedRect(selectionRect, transform = {}, img) {
    if (!selectionRect || !img) return null;
    const dw = Math.max(1, selectionRect.w);
    const dh = Math.max(1, selectionRect.h);
    const sw = img.width;
    const sh = img.height;
    if (!sw || !sh) return null;

    const normalized = normalizeImageTransform(transform);
    const totalCropX = clamp(normalized.cropLeft + normalized.cropRight, 0, 0.9);
    const totalCropY = clamp(normalized.cropTop + normalized.cropBottom, 0, 0.9);

    const baseContainScale = Math.min(dw / sw, dh / sh);
    const baseContainW = sw * baseContainScale;
    const baseContainH = sh * baseContainScale;

    const drawW = baseContainW * normalized.scale;
    const drawH = baseContainH * normalized.scale;
    const drawX = selectionRect.x + (dw - drawW) / 2 + normalized.offsetX;
    const drawY = selectionRect.y + (dh - drawH) / 2 + normalized.offsetY;

    const dstX = drawX + drawW * normalized.cropLeft;
    const dstY = drawY + drawH * normalized.cropTop;
    const dstW = clamp(drawW * (1 - totalCropX), 1, Math.abs(drawW));
    const dstH = clamp(drawH * (1 - totalCropY), 1, Math.abs(drawH));

    return { x: dstX, y: dstY, w: dstW, h: dstH };
}

function drawImageCover(img, x, y, w, h, transform = {}, filter = {}) {
    const sw = img.width;
    const sh = img.height;
    if (!sw || !sh) return;

    const normalized = normalizeImageTransform(transform);
    const totalCropX = clamp(normalized.cropLeft + normalized.cropRight, 0, 0.9);
    const totalCropY = clamp(normalized.cropTop + normalized.cropBottom, 0, 0.9);
    const srcX = sw * normalized.cropLeft;
    const srcY = sh * normalized.cropTop;
    const srcW = clamp(sw * (1 - totalCropX), 1, sw);
    const srcH = clamp(sh * (1 - totalCropY), 1, sh);
    const rendered = getImageRenderedRect({ x, y, w, h }, normalized, img);
    if (!rendered) return;
    const dstX = rendered.x;
    const dstY = rendered.y;
    const dstW = rendered.w;
    const dstH = rendered.h;
    const canvasFilter = buildCanvasFilter(filter);

    ctx.save();
    // always paint a white base under the selected image region
    // so transparent/partially covered areas keep a clean background.
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(x, y, Math.max(1, w), Math.max(1, h));
    ctx.beginPath();
    ctx.rect(x, y, Math.max(1, w), Math.max(1, h));
    ctx.clip();

    ctx.filter = canvasFilter;
    ctx.drawImage(img, srcX, srcY, srcW, srcH, dstX, dstY, dstW, dstH);
    ctx.restore();
}

function drawTextEntry(entry) {
    const rect = getEntryCanvasRect(entry);
    if (!rect) return false;
    const x = rect.x;
    const y = rect.y;
    const w = rect.w;
    const h = rect.h;

    ctx.fillStyle = entry.bgColor || 'white';
    ctx.fillRect(x, y, w, h);
    const rawLines = Array.isArray(entry.lines) && entry.lines.length > 0
        ? entry.lines
        : [{
            text: entry.text || '',
            fontFamily: entry.fontFamily || 'Arial',
            fontSize: entry.coordSpace === 'base'
                ? fontSizeBaseToCanvas(entry.fontSize || 16)
                : (entry.fontSize || 16),
            textColor: entry.textColor || '#000000'
        }];
    const lines = entry.coordSpace === 'base' ? mapTextLinesBaseToCanvas(rawLines) : rawLines;

    const zoomFactor = entry.coordSpace === 'base' ? getScaleFromBase() : 1;
    const padX = Math.max(2, 10 * zoomFactor);
    const padY = Math.max(2, 10 * zoomFactor);
    const lineSpacing = Math.max(0.6, Number(entry.lineSpacing) || 1);
    const baseLineGap = Math.max(1, 4 * zoomFactor);
    const lineGap = baseLineGap * lineSpacing;
    const bottomPad = Math.max(1, 6 * zoomFactor);
    const clipInset = Math.max(1, 2 * zoomFactor);
    const isCentered = entry.textAlign === 'center' || !!entry.centerText;
    const autoFitText = !!(entry.autoFitText || entry.autoFitSingleLine);
    const defaultTextColor = entry.textColor || resolveTextColorForBackground(entry.bgPreset || 'white');

    const maxWidth = Math.max(1, w - (2 * padX));
    let drawY = y + padY;
    const maxY = y + h - bottomPad;

    ctx.save();
    ctx.beginPath();
    ctx.rect(x + clipInset, y + clipInset, Math.max(1, w - (2 * clipInset)), Math.max(1, h - (2 * clipInset)));
    ctx.clip();

    let hasOverflow = false;

    const splitLongToken = (token) => {
        const chunks = [];
        let chunk = '';
        for (const char of token) {
            const attempt = chunk + char;
            if (ctx.measureText(attempt).width <= maxWidth || chunk.length === 0) {
                chunk = attempt;
            } else {
                chunks.push(chunk);
                chunk = char;
            }
        }
        if (chunk) chunks.push(chunk);
        return chunks;
    };

    const buildRenderedLines = (scaleFactor = 1) => {
        const rendered = [];
        for (const ln of lines) {
            const text = (ln.text || '').trim();
            const fontFamily = ln.fontFamily || 'Arial';
            const fontSize = Math.max(6, (parseInt(ln.fontSize, 10) || 16) * scaleFactor);

            if (!text) continue;

            ctx.font = buildCanvasFont(fontSize, fontFamily);
            const words = text.split(' ');
            let currentLine = '';
            const wrappedLines = [];

            for (const word of words) {
                if (!word) continue;
                if (ctx.measureText(word).width > maxWidth) {
                    if (currentLine) {
                        wrappedLines.push(currentLine.trimEnd());
                        currentLine = '';
                    }
                    const pieces = splitLongToken(word);
                    pieces.forEach((piece, idx) => {
                        if (idx < pieces.length - 1) {
                            wrappedLines.push(piece);
                        } else {
                            currentLine = piece + ' ';
                        }
                    });
                    continue;
                }

                const testLine = currentLine + word + ' ';
                if (ctx.measureText(testLine).width > maxWidth && currentLine !== '') {
                    wrappedLines.push(currentLine.trimEnd());
                    currentLine = word + ' ';
                } else {
                    currentLine = testLine;
                }
            }

            if (currentLine) wrappedLines.push(currentLine.trimEnd());

            wrappedLines.forEach((wrappedText) => {
                ctx.font = buildCanvasFont(fontSize, fontFamily);
                rendered.push({
                    text: wrappedText,
                    fontFamily,
                    fontSize,
                    textColor: defaultTextColor,
                    width: ctx.measureText(wrappedText).width
                });
            });
        }
        return rendered;
    };

    const measureRenderedHeight = (rendered, gap) => rendered.reduce((total, line, index) => {
        return total + line.fontSize + (index < rendered.length - 1 ? gap : 0);
    }, 0);

    let renderedLines = buildRenderedLines(1);
    let renderLineGap = lineGap;
    if (autoFitText) {
        let scaleFactor = 1;
        let fitted = renderedLines;
        let fittedGap = lineGap;
        let foundFit = false;
        while (scaleFactor >= 0.35) {
            const candidate = buildRenderedLines(scaleFactor);
            const candidateGap = lineGap * scaleFactor;
            if (measureRenderedHeight(candidate, candidateGap) <= Math.max(1, maxY - (y + padY))) {
                fitted = candidate;
                fittedGap = candidateGap;
                foundFit = true;
                break;
            }
            fitted = candidate;
            fittedGap = candidateGap;
            scaleFactor -= scaleFactor > 0.6 ? 0.05 : 0.02;
        }
        renderedLines = fitted;
        renderLineGap = fittedGap;
        hasOverflow = !foundFit && measureRenderedHeight(renderedLines, renderLineGap) > Math.max(1, maxY - (y + padY));
    }

    const fitTextWithEllipsis = (rawText, availableWidth) => {
        const textValue = `${rawText || ''}`;
        if (ctx.measureText(textValue).width <= availableWidth) return textValue;
        const ellipsis = '...';
        let cut = textValue;
        while (cut.length > 0 && ctx.measureText(cut + ellipsis).width > availableWidth) {
            cut = cut.slice(0, -1);
        }
        return cut ? (cut + ellipsis) : ellipsis;
    };

    ctx.textBaseline = 'top';
    if (autoFitText && renderedLines.length === 1) {
        const single = renderedLines[0];
        ctx.font = buildCanvasFont(single.fontSize, single.fontFamily);
        ctx.fillStyle = single.textColor;
        const autoY = y + Math.max(padY, (h - single.fontSize) / 2);
        const autoX = x + padX + (isCentered ? Math.max(0, (maxWidth - single.width) / 2) : 0);
        ctx.fillText(single.text, autoX, autoY);
    } else {
        for (let i = 0; i < renderedLines.length; i++) {
            const rendered = renderedLines[i];
            if (drawY + rendered.fontSize > maxY) {
                hasOverflow = true;
                ctx.restore();
                return hasOverflow;
            }
            ctx.font = buildCanvasFont(rendered.fontSize, rendered.fontFamily);
            ctx.fillStyle = rendered.textColor;

            let textToDraw = rendered.text;
            let textWidth = rendered.width;
            const hasMore = i < renderedLines.length - 1;
            const nextY = drawY + rendered.fontSize + renderLineGap;
            if (hasMore && nextY + renderedLines[i + 1].fontSize > maxY) {
                textToDraw = fitTextWithEllipsis(`${rendered.text} `, maxWidth);
                textWidth = ctx.measureText(textToDraw).width;
                hasOverflow = true;
            }

            const drawX = x + padX + (isCentered ? Math.max(0, (maxWidth - textWidth) / 2) : 0);
            ctx.fillText(textToDraw, drawX, drawY);
            drawY += rendered.fontSize + renderLineGap;
            if (hasMore && nextY + renderedLines[i + 1].fontSize > maxY) {
                ctx.restore();
                return hasOverflow;
            }
        }
    }

    ctx.restore();
    return hasOverflow;
}

function setTextOverflowIndicator(isVisible) {
    const indicator = document.getElementById('text-overflow-indicator');
    if (!indicator) return;
    indicator.style.display = isVisible ? 'block' : 'none';
}

function applyEdits() {
    edits.forEach((entry, index) => {
        if (hideSelectedAppliedImageDuringEdit && editingIndex !== null && index === editingIndex) {
            return;
        }
        if (entry.type === 'image') {
            const rect = getEntryCanvasRect(entry);
            if (!rect) return;
            const drawTransform = getEntryCanvasImageTransform(entry);
            if (entry.imgObj && entry.imgObj.complete) {
                drawImageCover(entry.imgObj, rect.x, rect.y, rect.w, rect.h, drawTransform, entry.imageFilter);
            } else {
                const img = entry.imgObj || new Image();
                entry.imgObj = img;
                img.onload = () => {
                    if (redrawRequestId > 0) redrawCanvas();
                };
                if (!img.src) img.src = entry.dataUrl;
            }
        } else if (entry.type === 'text') {
            drawTextEntry(entry);
        }
    });
}

function drawPreviewEdit() {
    if (!previewEdit) {
        setTextOverflowIndicator(false);
        return;
    }
    if (previewEdit.type === 'text') {
        const overflow = !!drawTextEntry(previewEdit);
        setTextOverflowIndicator(overflow);
    } else if (previewEdit.type === 'image' && previewEdit.dataUrl) {
        setTextOverflowIndicator(false);
        if (previewEdit.imgObj && previewEdit.imgObj.complete) {
            drawImageCover(previewEdit.imgObj, previewEdit.x, previewEdit.y, previewEdit.w, previewEdit.h, previewEdit.imageTransform, previewEdit.imageFilter);
            return;
        }
        const img = previewEdit.imgObj || new Image();
        previewEdit.imgObj = img;
        img.onload = () => {
            if (previewEdit && previewEdit.imgObj === img) {
                redrawCanvas();
            }
        };
        if (!img.src) img.src = previewEdit.dataUrl;
    }
}

function redrawCanvas() {
    if (!pdfDoc) return;
    const requestId = ++redrawRequestId;
    if (drawBasePageToCanvas()) {
        if (requestId !== redrawRequestId) return;
        applyEdits();
        drawPreviewEdit();
        applyPendingZoomFocus();
        return;
    }
    const generation = ++renderGeneration;
    renderPage(pageNum, generation).then((ok) => {
        if (!ok) return;
        if (requestId !== redrawRequestId) return;
        if (!drawBasePageToCanvas()) return;
        applyEdits();
        drawPreviewEdit();
        applyPendingZoomFocus();
    });
}

function scheduleRedrawCanvas() {
    if (redrawRafId !== null) return;
    redrawRafId = window.requestAnimationFrame(() => {
        redrawRafId = null;
        redrawCanvas();
    });
}

function serializeEditableEdits() {
    return edits.map((entry) => {
        if (!entry || !entry.type) return null;
        if (entry.type === 'image') {
            return {
                type: 'image',
                x: entry.x,
                y: entry.y,
                w: entry.w,
                h: entry.h,
                coordSpace: entry.coordSpace || 'base',
                dataUrl: entry.dataUrl || '',
                imageTransform: normalizeImageTransform(entry.imageTransform || {}),
                imageFilter: normalizeImageFilter(entry.imageFilter || {})
            };
        }
        if (entry.type === 'text') {
            return {
                type: 'text',
                x: entry.x,
                y: entry.y,
                w: entry.w,
                h: entry.h,
                coordSpace: entry.coordSpace || 'base',
                text: entry.text || '',
                lines: Array.isArray(entry.lines) ? entry.lines.map((ln) => ({
                    text: ln.text || '',
                    fontFamily: ln.fontFamily || 'Arial',
                    fontSize: Number(ln.fontSize) || 16,
                    textColor: ln.textColor || '#000000'
                })) : [],
                fontFamily: entry.fontFamily || 'Arial',
                fontSize: Number(entry.fontSize) || 16,
                textColor: entry.textColor || '#000000',
                bgColor: entry.bgColor || '#ffffff',
                bgPreset: entry.bgPreset || 'white',
                autoFitSingleLine: !!entry.autoFitSingleLine,
                autoFitText: !!(entry.autoFitText || entry.autoFitSingleLine),
                lineSpacing: Number(entry.lineSpacing) || 1,
                textAlign: entry.textAlign === 'center' ? 'center' : 'left',
                centerText: !!entry.centerText
            };
        }
        return null;
    }).filter(Boolean);
}

async function loadEditableProject(projectUrl) {
    if (!projectUrl) return;
    try {
        const res = await fetch(projectUrl, { cache: 'no-store' });
        if (!res.ok) return;
        const project = await res.json();
        applyEditableProjectData(project);
    } catch (err) {
        console.warn('Could not load editable project:', err);
    }
}

function applyEditableProjectData(project) {
    const loaded = Array.isArray(project && project.edits) ? project.edits : [];
    sourcePdfOverride = (project && project.sourcePdf) ? project.sourcePdf : sourcePdfOverride;

    edits = loaded.map((entry) => {
        if (!entry || !entry.type) return null;
        if (entry.type === 'image') {
            const normalized = {
                type: 'image',
                x: Number(entry.x) || 0,
                y: Number(entry.y) || 0,
                w: Number(entry.w) || 0,
                h: Number(entry.h) || 0,
                coordSpace: entry.coordSpace || 'base',
                dataUrl: entry.dataUrl || '',
                imageTransform: normalizeImageTransform(entry.imageTransform || {}),
                imageFilter: normalizeImageFilter(entry.imageFilter || {})
            };
            if (normalized.dataUrl) {
                const img = new Image();
                img.src = normalized.dataUrl;
                normalized.imgObj = img;
            }
            return normalized;
        }
        if (entry.type === 'text') {
            return {
                type: 'text',
                x: Number(entry.x) || 0,
                y: Number(entry.y) || 0,
                w: Number(entry.w) || 0,
                h: Number(entry.h) || 0,
                coordSpace: entry.coordSpace || 'base',
                text: entry.text || '',
                lines: Array.isArray(entry.lines) ? entry.lines : [],
                fontFamily: entry.fontFamily || 'Arial',
                fontSize: Number(entry.fontSize) || 16,
                textColor: entry.textColor || '#000000',
                bgColor: entry.bgColor || '#ffffff',
                bgPreset: entry.bgPreset || 'white',
                autoFitSingleLine: !!entry.autoFitSingleLine,
                autoFitText: !!(entry.autoFitText || entry.autoFitSingleLine),
                lineSpacing: Number(entry.lineSpacing) || 1,
                textAlign: entry.textAlign === 'center' ? 'center' : 'left',
                centerText: !!entry.centerText
            };
        }
        return null;
    }).filter(Boolean);

    previewEdit = null;
    editingIndex = null;
    selection = null;
    hideEditOverlay();
    scheduleRedrawCanvas();
}

function applyViewZoom() {
    if (!canvas) return;
    const zoomValue = document.getElementById('zoom-value');
    if (zoomValue) {
        zoomValue.textContent = `${Math.round(viewZoom * 100)}%`;
    }
    if (selection) {
        updateOverlayBoundsFromSelection();
    }
}

function updatePanControls() {
    const editorArea = document.getElementById('editor-area');
    const panX = document.getElementById('pan-x');
    const panY = document.getElementById('pan-y');
    if (!editorArea || !panX || !panY) return;

    const maxX = Math.max(0, Math.round(editorArea.scrollWidth - editorArea.clientWidth));
    const maxY = Math.max(0, Math.round(editorArea.scrollHeight - editorArea.clientHeight));

    panX.max = String(maxX);
    panY.max = String(maxY);
    panX.value = String(Math.round(editorArea.scrollLeft));
    panY.value = String(Math.round(editorArea.scrollTop));
    panX.disabled = maxX <= 0;
    panY.disabled = maxY <= 0;
}

function updateFitZoomBase() {
    const editorArea = document.getElementById('editor-area');
    if (!editorArea || basePageWidth <= 0 || basePageHeight <= 0) {
        fitZoomBase = 1;
        return;
    }

    const availableW = Math.max(100, editorArea.clientWidth - 32);
    const availableH = Math.max(100, editorArea.clientHeight - 100);
    const fitW = availableW / basePageWidth;
    const fitH = availableH / basePageHeight;
    fitZoomBase = clamp(Math.min(fitW, fitH), 0.1, 5);
}

function applyPendingZoomFocus() {
    if (!pendingZoomFocus) return;
    const editorArea = document.getElementById('editor-area');
    if (!editorArea || !canvas) {
        pendingZoomFocus = null;
        return;
    }

    const editorRect = editorArea.getBoundingClientRect();
    const canvasRect = canvas.getBoundingClientRect();
    const focus = pendingZoomFocus;

    const anchorX = focus.anchorClientX;
    const anchorY = focus.anchorClientY;
    const rx = clamp(focus.rx, 0, 1);
    const ry = clamp(focus.ry, 0, 1);

    const canvasStartX = editorArea.scrollLeft + (canvasRect.left - editorRect.left);
    const canvasStartY = editorArea.scrollTop + (canvasRect.top - editorRect.top);
    const pointX = canvasStartX + rx * canvasRect.width;
    const pointY = canvasStartY + ry * canvasRect.height;

    editorArea.scrollLeft = pointX - (anchorX - editorRect.left);
    editorArea.scrollTop = pointY - (anchorY - editorRect.top);

    pendingZoomFocus = null;
    updatePanControls();
}

function rescaleRectInPlace(rect, ratio) {
    if (!rect || !Number.isFinite(ratio) || ratio <= 0) return;
    rect.x *= ratio;
    rect.y *= ratio;
    rect.w *= ratio;
    rect.h *= ratio;
}

function rescaleLiveStateForZoom(prevScale, nextScale) {
    if (!Number.isFinite(prevScale) || !Number.isFinite(nextScale) || prevScale <= 0 || nextScale <= 0) {
        return;
    }
    const ratio = nextScale / prevScale;
    if (Math.abs(ratio - 1) < 0.000001) return;

    if (selection) {
        rescaleRectInPlace(selection, ratio);
    }

    if (Number.isFinite(startX)) startX *= ratio;
    if (Number.isFinite(startY)) startY *= ratio;

    if (imageMouseTransformState) {
        if (imageMouseTransformState.startSelection) {
            rescaleRectInPlace(imageMouseTransformState.startSelection, ratio);
        }
        if (imageMouseTransformState.startTransform) {
            imageMouseTransformState.startTransform.offsetX = (imageMouseTransformState.startTransform.offsetX || 0) * ratio;
            imageMouseTransformState.startTransform.offsetY = (imageMouseTransformState.startTransform.offsetY || 0) * ratio;
        }
    }

    if (currentImageTransform) {
        currentImageTransform = normalizeImageTransform({
            ...currentImageTransform,
            offsetX: (currentImageTransform.offsetX || 0) * ratio,
            offsetY: (currentImageTransform.offsetY || 0) * ratio
        });
    }

    if (previewEdit && previewEdit.type === 'image' && previewEdit.coordSpace !== 'base') {
        previewEdit.x *= ratio;
        previewEdit.y *= ratio;
        previewEdit.w *= ratio;
        previewEdit.h *= ratio;
        if (previewEdit.imageTransform) {
            previewEdit.imageTransform = normalizeImageTransform({
                ...previewEdit.imageTransform,
                offsetX: (previewEdit.imageTransform.offsetX || 0) * ratio,
                offsetY: (previewEdit.imageTransform.offsetY || 0) * ratio
            });
        }
    }
}

function setViewZoom(newZoom, options = {}) {
    const editorArea = document.getElementById('editor-area');
    if (canvas && editorArea) {
        const canvasRect = canvas.getBoundingClientRect();
        const editorRect = editorArea.getBoundingClientRect();
        const fallbackX = editorRect.left + editorRect.width / 2;
        const fallbackY = editorRect.top + editorRect.height / 2;
        const anchorClientX = options.anchorClientX ?? fallbackX;
        const anchorClientY = options.anchorClientY ?? fallbackY;
        const rx = canvasRect.width > 0 ? (anchorClientX - canvasRect.left) / canvasRect.width : 0.5;
        const ry = canvasRect.height > 0 ? (anchorClientY - canvasRect.top) / canvasRect.height : 0.5;
        pendingZoomFocus = { anchorClientX, anchorClientY, rx, ry };
    }

    migrateLegacyEditsToBase();
    viewZoom = clamp(newZoom, 0.5, 3);
    updateFitZoomBase();
    // 100% in UI means "fit to view". Higher/lower values zoom from that baseline.
    const previousScale = scale;
    const effectiveZoom = viewZoom * fitZoomBase;
    const nextScale = BASE_RENDER_SCALE * effectiveZoom;
    rescaleLiveStateForZoom(previousScale, nextScale);
    scale = nextScale;
    // force rerender in the new scale (otherwise redraw may reuse old base bitmap)
    renderedBaseScale = null;
    const zoomRange = document.getElementById('zoom-range');
    if (zoomRange) {
        zoomRange.value = String(Math.round(viewZoom * 100));
    }
    scheduleRedrawCanvas();
    setTimeout(updatePanControls, 0);
}

// Wait for DOM to be ready
function initDOM() {
    canvas = document.getElementById('pdf-canvas');
    console.log('initDOM: canvas =', canvas);
    if (!canvas) {
        console.error('Canvas element not found');
        return false;
    }
    ctx = canvas.getContext('2d');
    pageBaseCanvas = document.createElement('canvas');
    pageBaseCtx = pageBaseCanvas.getContext('2d');
    console.log('initDOM: ctx =', ctx);
    applyViewZoom();
    return true;
}

function initPdf(url, projectUrl = null, sourcePdfName = null) {
    console.log('initPdf called with url:', url);
    if (!initDOM()) {
        console.error('initDOM failed');
        return;
    }
    // ensure PDF.js worker is set
    try {
        if (window.pdfjsLib && !pdfjsLib.GlobalWorkerOptions.workerSrc) {
            // use local worker to avoid CDN/tracking prevention issues
            pdfjsLib.GlobalWorkerOptions.workerSrc = '/static/js/pdf.worker.min.js';
            console.log('pdf.worker configured (local):', pdfjsLib.GlobalWorkerOptions.workerSrc);
        }
    } catch (e) {
        console.warn('Could not set pdf.worker:', e);
    }
    pdfUrl = url;
    sourcePdfOverride = sourcePdfName || null;
    const loadingTask = pdfjsLib.getDocument(url);
    loadingTask.promise.then(function(pdf) {
        console.log('PDF loaded successfully');
        pdfDoc = pdf;
        renderAndRedraw();
        console.log('First page rendered, setting up event listeners');
        setupEventListeners();
        if (projectUrl) {
            loadEditableProject(projectUrl);
        }
    }).catch(err => {
        console.error('Error loading PDF:', err);
    });
}

function renderPage(num, generation = null) {
    return pdfDoc.getPage(num).then(function(page) {
        const baseViewport = page.getViewport({ scale: BASE_RENDER_SCALE });
        basePageWidth = baseViewport.width;
        basePageHeight = baseViewport.height;
        let viewport = page.getViewport({ scale: scale });
        const tempCanvas = document.createElement('canvas');
        tempCanvas.width = viewport.width;
        tempCanvas.height = viewport.height;
        const tempCtx = tempCanvas.getContext('2d');

        let renderContext = {
            canvasContext: tempCtx,
            viewport: viewport
        };
        return page.render(renderContext).promise.then(() => {
            if (generation !== null && generation !== renderGeneration) {
                return false;
            }
            pageBaseCanvas.width = tempCanvas.width;
            pageBaseCanvas.height = tempCanvas.height;
            pageBaseCtx.clearRect(0, 0, pageBaseCanvas.width, pageBaseCanvas.height);
            pageBaseCtx.drawImage(tempCanvas, 0, 0);
            renderedBasePageNum = num;
            renderedBaseScale = scale;
            return true;
        });
    });
}

function drawBasePageToCanvas() {
    if (!pageBaseCanvas || renderedBasePageNum !== pageNum || renderedBaseScale !== scale || !pageBaseCanvas.width || !pageBaseCanvas.height) {
        return false;
    }
    if (canvas.width !== pageBaseCanvas.width || canvas.height !== pageBaseCanvas.height) {
        canvas.width = pageBaseCanvas.width;
        canvas.height = pageBaseCanvas.height;
    }
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(pageBaseCanvas, 0, 0);
    applyViewZoom();
    return true;
}

// helper to render and then draw all placed items
function renderAndRedraw() {
    if (!pdfDoc) return Promise.resolve();
    const generation = ++renderGeneration;
    return renderPage(pageNum, generation).then((ok) => {
        if (!ok) return;
        if (!drawBasePageToCanvas()) return;
        drawAllItems();
    });
}

// draw every item in placedItems onto canvas
function drawAllItems() {
    placedItems.forEach(item => {
        if (item.type === 'text') {
            // draw background
            ctx.fillStyle = item.bgColor || 'white';
            ctx.fillRect(item.x, item.y, item.w, item.h);
            // draw text
            ctx.font = buildCanvasFont(item.fontSize, item.fontFamily);
            ctx.fillStyle = 'black';
            ctx.textBaseline = 'top';
            const words = item.text.split(' ');
            let line = '';
            let y = item.y + 10;
            const maxWidth = item.w - 20;
            for (let word of words) {
                const testLine = line + word + ' ';
                const metrics = ctx.measureText(testLine);
                if (metrics.width > maxWidth && line !== '') {
                    ctx.fillText(line, item.x + 10, y);
                    line = word + ' ';
                    y += item.fontSize + 5;
                } else {
                    line = testLine;
                }
            }
            if (line) {
                ctx.fillText(line, item.x + 10, y);
            }
        } else if (item.type === 'image') {
            if (item.img) {
                ctx.drawImage(item.img, item.x, item.y, item.w, item.h);
            } else if (item.src) {
                const imgObj = new Image();
                imgObj.onload = () => {
                    item.img = imgObj;
                    ctx.drawImage(imgObj, item.x, item.y, item.w, item.h);
                };
                imgObj.src = item.src;
            }
        }
    });
}

function getSidebarTextConfig() {
    const textInput = document.getElementById('sidebar-text-input');
    const fontFamilyInput = document.getElementById('sidebar-font-family');
    const fontSizeInput = document.getElementById('sidebar-font-size');
    const bgColorInput = document.getElementById('sidebar-bg-color');
    const autoFitInput = document.getElementById('sidebar-text-autofit');
    const lineSpacingInput = document.getElementById('sidebar-line-spacing');
    const centerTextBtn = document.getElementById('sidebar-text-center');
    const lineRows = document.querySelectorAll('.text-line-row');

    const rowLines = Array.from(lineRows).map(row => {
        const lineTextInput = row.querySelector('.line-text');
        const lineFamilyInput = row.querySelector('.line-font-family');
        const lineSizeInput = row.querySelector('.line-font-size');
        return {
            text: lineTextInput ? lineTextInput.value : '',
            fontFamily: lineFamilyInput ? lineFamilyInput.value : (fontFamilyInput ? fontFamilyInput.value : 'Arial'),
            fontSize: lineSizeInput ? parseInt(lineSizeInput.value, 10) || 16 : 16
        };
    }).filter(line => (line.text || '').trim().length > 0);

    const fallbackRawText = textInput ? textInput.value : '';
    const fallbackLineEntries = fallbackRawText.split(/\r?\n/).map(lineText => ({
        text: lineText,
        fontFamily: fontFamilyInput ? fontFamilyInput.value : 'Arial',
        fontSize: fontSizeInput ? parseInt(fontSizeInput.value, 10) || 16 : 16
    }));
    const hasFallbackContent = fallbackLineEntries.some(line => (line.text || '').trim().length > 0);

    const effectiveLines = (textLinesModeEnabled && rowLines.length > 0)
        ? rowLines
        : (hasFallbackContent ? fallbackLineEntries : []);

    const bgPreset = bgColorInput ? bgColorInput.value : 'white';
    const effectiveTextColor = resolveTextColorForBackground(bgPreset);
    const lineSpacing = Math.max(0.6, parseFloat(lineSpacingInput ? lineSpacingInput.value : '1') || 1);
    const isCentered = !!(centerTextBtn && centerTextBtn.classList.contains('mode-active'));

    const adjustedLines = effectiveLines.map(line => ({
        ...line,
        textColor: effectiveTextColor
    }));

    return {
        text: adjustedLines.map(line => line.text).join('\n'),
        fontFamily: fontFamilyInput ? fontFamilyInput.value : 'Arial',
        fontSize: fontSizeInput ? parseInt(fontSizeInput.value, 10) || 16 : 16,
        textColor: effectiveTextColor,
        lines: adjustedLines,
        bgPreset,
        autoFitSingleLine: !!(autoFitInput && autoFitInput.checked),
        autoFitText: !!(autoFitInput && autoFitInput.checked),
        lineSpacing,
        textAlign: isCentered ? 'center' : 'left',
        centerText: isCentered
    };
}

function resetTextSidebarForNewEntry() {
    const textInput = document.getElementById('sidebar-text-input');
    const fontFamilyInput = document.getElementById('sidebar-font-family');
    const fontSizeInput = document.getElementById('sidebar-font-size');
    const bgColorInput = document.getElementById('sidebar-bg-color');
    const autoFitInput = document.getElementById('sidebar-text-autofit');
    const lineSpacingInput = document.getElementById('sidebar-line-spacing');
    const centerTextBtn = document.getElementById('sidebar-text-center');

    if (textInput) textInput.value = '';
    if (fontFamilyInput) fontFamilyInput.value = 'Arial';
    if (fontSizeInput) fontSizeInput.value = 16;
    if (bgColorInput) bgColorInput.value = 'white';
    if (autoFitInput) autoFitInput.checked = false;
    if (lineSpacingInput) lineSpacingInput.value = '1';
    if (centerTextBtn) centerTextBtn.classList.remove('mode-active');

    textLinesModeEnabled = false;
    populateTextLineEditor([], {
        text: '',
        fontFamily: 'Arial',
        fontSize: 16,
        textColor: resolveTextColorForBackground('white'),
        lineSpacing: 1,
        textAlign: 'left'
    });
    setTextOverflowIndicator(false);
}

function resolveTextColorForBackground(bgPreset) {
    // readable defaults by background choice
    switch (bgPreset) {
        case 'blue':
        case 'red':
            return '#ffffff';
        case 'yellow':
        case 'white':
        default:
            return '#000000';
    }
}

function getBackgroundColorFromPreset(preset) {
    switch (preset) {
        case 'blue':
            return '#3b82f6';
        case 'yellow':
            return '#facc15';
        case 'red':
            return '#ef4444';
        case 'white':
        default:
            return '#ffffff';
    }
}

function resolveCanvasFontParts(fontFamily) {
    const raw = (fontFamily || 'Arial').toString().trim();
    const normalized = raw.toLowerCase();
    const style = (normalized.includes('italic') || normalized.includes('oblique')) ? 'italic' : 'normal';
    const weight = normalized.includes('bold') ? 'bold' : 'normal';

    let family = 'Arial, Helvetica, sans-serif';
    if (normalized.includes('times') || normalized.includes('georgia')) {
        family = '"Times New Roman", Times, serif';
    } else if (normalized.includes('helvetica')) {
        family = 'Helvetica, Arial, sans-serif';
    } else if (normalized.includes('courier')) {
        family = '"Courier New", Courier, monospace';
    } else if (normalized.includes('verdana')) {
        family = 'Verdana, Geneva, sans-serif';
    } else if (normalized.includes('tahoma')) {
        family = 'Tahoma, Verdana, sans-serif';
    } else if (normalized.includes('trebuchet')) {
        family = '"Trebuchet MS", Helvetica, sans-serif';
    } else if (normalized.includes('calibri')) {
        family = 'Calibri, Arial, sans-serif';
    }

    return { style, weight, family };
}

function buildCanvasFont(fontSize, fontFamily) {
    const size = Number(fontSize) || 16;
    const parts = resolveCanvasFontParts(fontFamily);
    return `${parts.style} ${parts.weight} ${size}px ${parts.family}`;
}

function fitSingleLineFontSizeHybrid(text, fontFamily, boxWidth, boxHeight) {
    const padX = 10;
    const padY = 10;
    const heightBasedSize = clamp(Math.round(Math.max(8, (Number(boxHeight) || 0) - (padY + 2))), 8, 200);
    const minSize = Math.min(heightBasedSize, AUTO_FIT_MIN_FONT_SIZE);
    const maxWidth = Math.max(1, (Number(boxWidth) || 0) - (2 * padX));
    const content = `${text || ''}`.trim();

    if (!content) {
        return heightBasedSize;
    }

    let fitted = heightBasedSize;
    while (fitted > minSize) {
        ctx.font = buildCanvasFont(fitted, fontFamily || 'Arial');
        if (ctx.measureText(content).width <= maxWidth) {
            return fitted;
        }
        fitted -= 1;
    }
    return minSize;
}

function getSimpleFontOptionsHtml(selectedFont = 'Arial') {
    const fonts = [
        'Arial',
        'Arial Bold',
        'Verdana',
        'Tahoma',
        'Helvetica Oblique',
        'Helvetica Bold Oblique',
        'Trebuchet MS',
        'Calibri',
        'Times New Roman',
        'Times Bold',
        'Times Italic',
        'Times Bold Italic',
        'Georgia',
        'Courier New'
    ];
    return fonts
        .map(font => `<option value="${font}" ${font === selectedFont ? 'selected' : ''}>${font}</option>`)
        .join('');
}

function buildTextLineRow(line = {}, index = 1) {
    const text = line.text || '';
    const fontFamily = line.fontFamily || 'Arial';
    const fontSize = parseInt(line.fontSize, 10) || 16;
    return `
        <div class="text-line-row">
            <input type="text" class="line-text" value="${text.replace(/"/g, '&quot;')}" placeholder="Linha ${index}">
            <div class="text-line-row-controls">
                <select class="line-font-family">${getSimpleFontOptionsHtml(fontFamily)}</select>
                <input type="number" class="line-font-size" min="8" max="72" value="${fontSize}">
            </div>
        </div>
    `;
}

function populateTextLineEditor(lines = [], fallbackCfg = null) {
    const container = document.getElementById('text-lines-editor');
    if (!container) return;

    let sourceLines = Array.isArray(lines) ? lines : [];
    if (sourceLines.length === 0 && fallbackCfg && fallbackCfg.text) {
        sourceLines = fallbackCfg.text.split(/\r?\n/).map(text => ({
            text,
            fontFamily: fallbackCfg.fontFamily || 'Arial',
            fontSize: fallbackCfg.fontSize || 16
        }));
    }

    if (sourceLines.length === 0) {
        sourceLines = [{
            text: '',
            fontFamily: (fallbackCfg && fallbackCfg.fontFamily) || 'Arial',
            fontSize: (fallbackCfg && fallbackCfg.fontSize) || 16
        }];
    }

    container.innerHTML = sourceLines
        .map((line, index) => buildTextLineRow(line, index + 1))
        .join('');
}

function splitRawTextIntoEditableLines(rawText) {
    const raw = (rawText || '').trim();
    if (!raw) return [];

    let parts = raw.split(/\r?\n+/).map(part => part.trim()).filter(Boolean);
    if (parts.length <= 1 && raw.includes('|')) {
        parts = raw.split('|').map(part => part.trim()).filter(Boolean);
    }
    if (parts.length <= 1 && raw.includes(';')) {
        parts = raw.split(';').map(part => part.trim()).filter(Boolean);
    }

    return parts;
}

function getRangeValue(id, fallback = 0) {
    const input = document.getElementById(id);
    if (!input) return fallback;
    const value = parseFloat(input.value);
    return Number.isFinite(value) ? value : fallback;
}

function setRangeValue(id, value) {
    const input = document.getElementById(id);
    if (input) input.value = String(value);
}

function getImageTransformConfig() {
    const base = currentImageTransform || DEFAULT_IMAGE_TRANSFORM;
    return normalizeImageTransform({
        ...base,
        scale: getRangeValue('image-scale', 100) / 100
    });
}

function getImageFilterConfig() {
    const base = currentImageFilter || DEFAULT_IMAGE_FILTER;
    const typeInput = document.getElementById('image-filter-type');
    return normalizeImageFilter({
        ...base,
        type: typeInput ? typeInput.value : base.type,
        intensity: getRangeValue('image-filter-intensity', base.intensity)
    });
}

function setImageFilterControls(filter = {}) {
    currentImageFilter = normalizeImageFilter(filter);
    const typeInput = document.getElementById('image-filter-type');
    if (typeInput) typeInput.value = currentImageFilter.type;
    setRangeValue('image-filter-intensity', Math.round(currentImageFilter.intensity));
    updateImageFilterLabels();
}

function updateImageFilterLabels() {
    const input = document.getElementById('image-filter-intensity');
    const output = document.getElementById('image-filter-intensity-val');
    if (input && output) output.textContent = `${input.value}%`;
}

function setImageTransformControls(transform = {}) {
    currentImageTransform = normalizeImageTransform(transform);
    setRangeValue('image-scale', Math.round((currentImageTransform.scale ?? 1) * 100));
    updateImageTransformLabels();
}

function applyImageTransformAndPreview(transform) {
    setImageTransformControls(normalizeImageTransform(transform));
    if (currentMode === 'image' && selection && selectedImageDataUrl) {
        updateImagePreview();
    }
}

function applyImageFilterAndPreview(filter) {
    setImageFilterControls(normalizeImageFilter(filter));
    if (currentMode === 'image' && selection && selectedImageDataUrl) {
        updateImagePreview();
    }
}

function actionHasHorizontal(action) {
    return action.includes('e') || action.includes('w');
}

function actionHasVertical(action) {
    return action.includes('n') || action.includes('s');
}

function resizeSelectionFromHandle(baseSel, action, dxCanvas, dyCanvas) {
    const minSize = 10;
    let left = baseSel.x;
    let top = baseSel.y;
    let right = baseSel.x + baseSel.w;
    let bottom = baseSel.y + baseSel.h;

    if (action.includes('e')) {
        right = clamp(baseSel.x + baseSel.w + dxCanvas, baseSel.x + minSize, canvas.width);
    }
    if (action.includes('w')) {
        left = clamp(baseSel.x + dxCanvas, 0, baseSel.x + baseSel.w - minSize);
    }
    if (action.includes('s')) {
        bottom = clamp(baseSel.y + baseSel.h + dyCanvas, baseSel.y + minSize, canvas.height);
    }
    if (action.includes('n')) {
        top = clamp(baseSel.y + dyCanvas, 0, baseSel.y + baseSel.h - minSize);
    }

    return {
        x: left,
        y: top,
        w: right - left,
        h: bottom - top
    };
}

function resizeSelectionFromHandleWithAspect(baseSel, action, dxCanvas, dyCanvas) {
    const rawSel = resizeSelectionFromHandle(baseSel, action, dxCanvas, dyCanvas);
    const aspect = baseSel.w / Math.max(1, baseSel.h);
    if (!Number.isFinite(aspect) || aspect <= 0) return rawSel;

    const minSize = 10;
    let w = rawSel.w;
    let h = rawSel.h;

    if (actionHasHorizontal(action) && !actionHasVertical(action)) {
        h = Math.max(minSize, w / aspect);
    } else if (!actionHasHorizontal(action) && actionHasVertical(action)) {
        w = Math.max(minSize, h * aspect);
    } else {
        const scaleX = rawSel.w / Math.max(1, baseSel.w);
        const scaleY = rawSel.h / Math.max(1, baseSel.h);
        const scale = Math.min(scaleX, scaleY);
        w = Math.max(minSize, baseSel.w * scale);
        h = Math.max(minSize, baseSel.h * scale);
    }

    const baseRight = baseSel.x + baseSel.w;
    const baseBottom = baseSel.y + baseSel.h;
    let x = baseSel.x;
    let y = baseSel.y;

    if (action === 'e') {
        x = baseSel.x;
        y = baseSel.y + (baseSel.h - h) / 2;
    } else if (action === 'w') {
        x = baseRight - w;
        y = baseSel.y + (baseSel.h - h) / 2;
    } else if (action === 's') {
        x = baseSel.x + (baseSel.w - w) / 2;
        y = baseSel.y;
    } else if (action === 'n') {
        x = baseSel.x + (baseSel.w - w) / 2;
        y = baseBottom - h;
    } else if (action === 'se') {
        x = baseSel.x;
        y = baseSel.y;
    } else if (action === 'sw') {
        x = baseRight - w;
        y = baseSel.y;
    } else if (action === 'ne') {
        x = baseSel.x;
        y = baseBottom - h;
    } else if (action === 'nw') {
        x = baseRight - w;
        y = baseBottom - h;
    }

    x = clamp(x, 0, Math.max(0, canvas.width - w));
    y = clamp(y, 0, Math.max(0, canvas.height - h));

    return { x, y, w, h };
}

function cropAndResizeFromHandle(baseTransform, baseSel, action, dxCanvas, dyCanvas) {
    const nextSel = { ...baseSel };
    const next = { ...baseTransform };

    const rightInward = -dxCanvas;
    const leftInward = dxCanvas;
    const bottomInward = -dyCanvas;
    const topInward = dyCanvas;

    if (action.includes('e')) {
        next.cropRight = clamp(baseTransform.cropRight + (rightInward / Math.max(1, baseSel.w)), 0, 0.9);
    }
    if (action.includes('w')) {
        next.cropLeft = clamp(baseTransform.cropLeft + (leftInward / Math.max(1, baseSel.w)), 0, 0.9);
    }
    if (action.includes('s')) {
        next.cropBottom = clamp(baseTransform.cropBottom + (bottomInward / Math.max(1, baseSel.h)), 0, 0.9);
    }
    if (action.includes('n')) {
        next.cropTop = clamp(baseTransform.cropTop + (topInward / Math.max(1, baseSel.h)), 0, 0.9);
    }

    if ((next.cropLeft + next.cropRight) > 0.95) {
        if (action.includes('w')) {
            next.cropLeft = clamp(0.95 - next.cropRight, 0, 0.95);
        } else {
            next.cropRight = clamp(0.95 - next.cropLeft, 0, 0.95);
        }
    }
    if ((next.cropTop + next.cropBottom) > 0.95) {
        if (action.includes('n')) {
            next.cropTop = clamp(0.95 - next.cropBottom, 0, 0.95);
        } else {
            next.cropBottom = clamp(0.95 - next.cropTop, 0, 0.95);
        }
    }

    return { nextTransform: next, nextSelection: nextSel };
}

function resizeImageTransformFromDiagonalHandle(baseTransform, baseSel, action, dxCanvas, dyCanvas) {
    const next = normalizeImageTransform(baseTransform || {});
    const signX = action.includes('e') ? 1 : -1;
    const signY = action.includes('s') ? 1 : -1;
    const deltaX = (dxCanvas * signX) / Math.max(1, baseSel.w);
    const deltaY = (dyCanvas * signY) / Math.max(1, baseSel.h);
    const delta = (deltaX + deltaY) / 2;
    next.scale = clamp(next.scale + delta, 0.2, 4);
    return next;
}

function isDiagonalHandleAction(action) {
    return action === 'ne' || action === 'nw' || action === 'se' || action === 'sw';
}

function updateImageTransformLabels() {
    const labels = [
        ['image-scale', 'image-scale-val', true]
    ];
    labels.forEach(([inputId, outputId]) => {
        const input = document.getElementById(inputId);
        const output = document.getElementById(outputId);
        if (input && output) output.textContent = `${input.value}%`;
    });
}

function isSideHandleAction(action) {
    return action === 'n' || action === 's' || action === 'e' || action === 'w';
}

function buildTextEntryFromSelection() {
    if (!selection) return null;
    const cfg = getSidebarTextConfig();
    if (!cfg.text) return null;
    const resolvedBgColor = getBackgroundColorFromPreset(cfg.bgPreset);
    let resolvedLines = (cfg.lines || []).map(line => ({ ...line }));

    const baseRect = canvasRectToBaseRect(selection);

    if (cfg.autoFitText && resolvedLines.length === 1 && resolvedLines[0] && !(resolvedLines[0].text || '').includes('\n')) {
        const line = resolvedLines[0];
        const fittedSize = fitSingleLineFontSizeHybrid(
            line.text || '',
            line.fontFamily || cfg.fontFamily,
            baseRect.w,
            baseRect.h
        );
        line.fontSize = fontSizeBaseToCanvas(fittedSize);
        resolvedLines = [line];
    }

    const baseLines = mapTextLinesCanvasToBase(resolvedLines);

    return {
        type: 'text',
        x: baseRect.x,
        y: baseRect.y,
        w: baseRect.w,
        h: baseRect.h,
        coordSpace: 'base',
        text: cfg.text,
        lines: baseLines,
        fontFamily: cfg.fontFamily,
        fontSize: fontSizeCanvasToBase(cfg.fontSize),
        textColor: cfg.textColor,
        bgColor: resolvedBgColor,
        bgPreset: cfg.bgPreset,
        autoFitSingleLine: cfg.autoFitSingleLine,
        autoFitText: cfg.autoFitText,
        lineSpacing: cfg.lineSpacing,
        textAlign: cfg.textAlign,
        centerText: cfg.centerText
    };
}

function updateTextPreview() {
    const entry = buildTextEntryFromSelection();
    previewEdit = entry;
    if (!entry) {
        setTextOverflowIndicator(false);
    }
    scheduleRedrawCanvas();
}

function triggerLiveTextPreview() {
    if (selection && (currentMode === 'text' || (editingIndex !== null && edits[editingIndex] && edits[editingIndex].type === 'text'))) {
        updateTextPreview();
    }
}

function updateImagePreview() {
    if (!selection || !selectedImageDataUrl) {
        previewEdit = null;
        setTextOverflowIndicator(false);
        scheduleRedrawCanvas();
        return;
    }
    setTextOverflowIndicator(false);
    const transform = getImageTransformConfig();
    previewEdit = {
        type: 'image',
        x: selection.x,
        y: selection.y,
        w: selection.w,
        h: selection.h,
        dataUrl: selectedImageDataUrl,
        imgObj: selectedImageObj,
        imageTransform: transform,
        imageFilter: getImageFilterConfig()
    };
    scheduleRedrawCanvas();
}

function loadImageFromDataUrl(dataUrl) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = reject;
        img.src = dataUrl;
    });
}

async function renderImageEntryForPdf(entry) {
    const srcImg = (entry.imgObj && entry.imgObj.complete) ? entry.imgObj : await loadImageFromDataUrl(entry.dataUrl);
    const baseW = Math.max(1, Math.round(entry.w));
    const baseH = Math.max(1, Math.round(entry.h));

    // Export image edits at a higher resolution for better PDF print quality.
    let exportScale = PDF_IMAGE_EXPORT_UPSCALE;
    const pixelBudget = Math.max(1, baseW * baseH);
    if (pixelBudget * exportScale * exportScale > PDF_IMAGE_EXPORT_MAX_PIXELS) {
        exportScale = Math.sqrt(PDF_IMAGE_EXPORT_MAX_PIXELS / pixelBudget);
    }
    exportScale = clamp(exportScale, 1, PDF_IMAGE_EXPORT_UPSCALE);

    const outW = Math.max(1, Math.round(baseW * exportScale));
    const outH = Math.max(1, Math.round(baseH * exportScale));

    const offscreen = document.createElement('canvas');
    offscreen.width = outW;
    offscreen.height = outH;
    const octx = offscreen.getContext('2d');

    octx.fillStyle = '#ffffff';
    octx.fillRect(0, 0, outW, outH);

    const normalized = normalizeImageTransform(entry.imageTransform || {});
    const exportTransform = normalizeImageTransform({
        ...normalized,
        offsetX: normalized.offsetX * exportScale,
        offsetY: normalized.offsetY * exportScale
    });
    const sw = srcImg.width;
    const sh = srcImg.height;
    const totalCropX = clamp(exportTransform.cropLeft + exportTransform.cropRight, 0, 0.9);
    const totalCropY = clamp(exportTransform.cropTop + exportTransform.cropBottom, 0, 0.9);
    const srcX = sw * exportTransform.cropLeft;
    const srcY = sh * exportTransform.cropTop;
    const srcW = clamp(sw * (1 - totalCropX), 1, sw);
    const srcH = clamp(sh * (1 - totalCropY), 1, sh);
    const rendered = getImageRenderedRect({ x: 0, y: 0, w: outW, h: outH }, exportTransform, srcImg);

    octx.save();
    octx.beginPath();
    octx.rect(0, 0, outW, outH);
    octx.clip();
    octx.filter = buildCanvasFilter(entry.imageFilter || {});
    octx.drawImage(srcImg, srcX, srcY, srcW, srcH, rendered.x, rendered.y, rendered.w, rendered.h);
    octx.restore();

    return offscreen.toDataURL('image/png');
}

async function prepareEditsForPdfSave() {
    const prepared = [];
    for (const entry of edits) {
        const rect = getEntryCanvasRect(entry);
        if (!rect) continue;
        if (entry.type === 'image') {
            const canvasTransform = getEntryCanvasImageTransform(entry);
            const imageData = await renderImageEntryForPdf({
                ...entry,
                x: rect.x,
                y: rect.y,
                w: rect.w,
                h: rect.h,
                imageTransform: canvasTransform
            });
            prepared.push({
                type: 'image',
                x: rect.x,
                y: rect.y,
                w: rect.w,
                h: rect.h,
                imageData
            });
        } else if (entry.type === 'text') {
            const exportLines = entry.coordSpace === 'base'
                ? mapTextLinesBaseToCanvas(entry.lines || [])
                : (entry.lines || []);
            prepared.push({
                type: 'text',
                x: rect.x,
                y: rect.y,
                w: rect.w,
                h: rect.h,
                lines: exportLines,
                text: entry.text || '',
                bgColor: entry.bgColor || '#ffffff',
                autoFitSingleLine: !!entry.autoFitSingleLine,
                autoFitText: !!(entry.autoFitText || entry.autoFitSingleLine),
                lineSpacing: Number(entry.lineSpacing) || 1,
                textAlign: entry.textAlign === 'center' ? 'center' : 'left',
                centerText: !!entry.centerText
            });
        }
    }
    return prepared;
}

// hit-test returns item under given canvas coords or null
function findItemAt(x, y) {
    for (let i = placedItems.length - 1; i >= 0; i--) {
        const it = placedItems[i];
        if (x >= it.x && y >= it.y && x <= it.x + it.w && y <= it.y + it.h) {
            return it;
        }
    }
    return null;
}

// Extract dominant color from selected area
function getDominantColor(x, y, w, h) {
    // sample border pixels around the selection area. this usually represents
    // the true background better than center sampling (which can hit text).
    if (w <= 0 || h <= 0) return 'rgb(255, 255, 255)';

    const left = Math.max(0, Math.floor(x));
    const top = Math.max(0, Math.floor(y));
    const right = Math.min(canvas.width - 1, Math.floor(x + w - 1));
    const bottom = Math.min(canvas.height - 1, Math.floor(y + h - 1));

    if (right <= left || bottom <= top) return 'rgb(255, 255, 255)';

    const inset = Math.max(1, Math.floor(Math.min(w, h) * 0.08));
    const sampleStep = Math.max(1, Math.floor(Math.min(w, h) / 30));

    let r = 0, g = 0, b = 0, count = 0;

    function sampleAt(sx, sy) {
        if (sx < 0 || sy < 0 || sx >= canvas.width || sy >= canvas.height) return;
        const pixel = ctx.getImageData(sx, sy, 1, 1).data;
        r += pixel[0];
        g += pixel[1];
        b += pixel[2];
        count++;
    }

    for (let sx = left + inset; sx <= right - inset; sx += sampleStep) {
        sampleAt(sx, top + inset);
        sampleAt(sx, bottom - inset);
    }

    for (let sy = top + inset; sy <= bottom - inset; sy += sampleStep) {
        sampleAt(left + inset, sy);
        sampleAt(right - inset, sy);
    }

    if (count === 0) {
        const cx = Math.floor(x + w / 2);
        const cy = Math.floor(y + h / 2);
        const pixel = ctx.getImageData(cx, cy, 1, 1).data;
        return `rgb(${pixel[0]}, ${pixel[1]}, ${pixel[2]})`;
    }

    r = Math.round(r / count);
    g = Math.round(g / count);
    b = Math.round(b / count);
    return `rgb(${r}, ${g}, ${b})`;
}

// Show edit overlay with detected color
function showEditOverlay() {
    if (!selection) return;
    hideSelectedAppliedImageDuringEdit = false;
    // when opening overlay because user clicked an existing edit,
    // editingIndex will point to the entry and we should pre-populate
    const previewImg = document.getElementById('overlay-preview');
    if (editingIndex !== null) {
        const entry = edits[editingIndex];
        if (entry) {
            if (entry.type === 'text') {
                const ti = document.getElementById('overlay-text-input');
                if (ti) ti.value = entry.text;
                document.getElementById('overlay-font-family').value = entry.fontFamily;
                document.getElementById('overlay-font-size').value = entry.fontSize;
            } else if (entry.type === 'image' && previewImg) {
                previewImg.src = entry.dataUrl;
                previewImg.style.display = 'block';
                selectedImageDataUrl = entry.dataUrl;
                selectedImageObj = entry.imgObj && entry.imgObj.complete ? entry.imgObj : null;
                if (!selectedImageObj) {
                    selectedImageObj = new Image();
                    selectedImageObj.onload = () => {
                        updateImagePreview();
                    };
                    selectedImageObj.src = entry.dataUrl;
                }
                const canvasTransform = getEntryCanvasImageTransform(entry);
                setImageTransformControls(canvasTransform);
                setImageFilterControls(entry.imageFilter || {});
                hideSelectedAppliedImageDuringEdit = true;
                previewEdit = {
                    type: 'image',
                    x: selection.x,
                    y: selection.y,
                    w: selection.w,
                    h: selection.h,
                    dataUrl: selectedImageDataUrl,
                    imgObj: selectedImageObj,
                    imageTransform: canvasTransform,
                    imageFilter: normalizeImageFilter(entry.imageFilter || {})
                };
                scheduleRedrawCanvas();
            }
        }
    } else {
        if (previewImg) previewImg.style.display = 'none';
    }
    console.log('showEditOverlay()', currentMode, selection);
    const overlay = document.getElementById('edit-overlay');
    const editorArea = document.getElementById('editor-area');
    const editorRect = editorArea.getBoundingClientRect();
    const canvasRect = canvas.getBoundingClientRect();
    const scaleX = canvasRect.width / canvas.width;
    const scaleY = canvasRect.height / canvas.height;

    // Position and size converted to CSS pixels using scale factors
    overlay.style.left = (editorArea.scrollLeft + (canvasRect.left - editorRect.left) + selection.x * scaleX) + 'px';
    overlay.style.top = (editorArea.scrollTop + (canvasRect.top - editorRect.top) + selection.y * scaleY) + 'px';
    overlay.style.width = selection.w * scaleX + 'px';
    overlay.style.height = selection.h * scaleY + 'px';

    overlay.style.backgroundColor = 'transparent';

    // Overlay is now mainly a visual frame. Main controls live in sidebar tabs.
    const imageControl = document.getElementById('overlay-image-btn');
    const textControlGroup = document.querySelector('.text-controls');
    const overlayDeleteBtn = document.getElementById('overlay-delete-btn');
    const imageMouseLayer = document.getElementById('image-mouse-layer');
    const overlayControls = document.getElementById('overlay-controls');
    if (imageControl) imageControl.style.display = 'none';
    if (textControlGroup) textControlGroup.style.display = 'none';
    if (overlayDeleteBtn) overlayDeleteBtn.style.display = editingIndex !== null ? 'block' : 'none';
    if (imageMouseLayer) imageMouseLayer.style.display = currentMode === 'image' ? 'block' : 'none';
    if (overlayControls) overlayControls.style.display = currentMode === 'image' ? 'none' : 'flex';
    overlay.classList.toggle('image-mode', currentMode === 'image');

    // Sync sidebar controls when editing an existing text item
    if (editingIndex !== null) {
        const entry = edits[editingIndex];
        if (entry && entry.type === 'text') {
            const uiLines = entry.coordSpace === 'base'
                ? mapTextLinesBaseToCanvas(entry.lines || [])
                : (entry.lines || []);
            const uiFontSize = entry.coordSpace === 'base'
                ? Math.round(fontSizeBaseToCanvas(entry.fontSize || 16))
                : (entry.fontSize || 16);
            const textInput = document.getElementById('sidebar-text-input');
            const fontFamilyInput = document.getElementById('sidebar-font-family');
            const fontSizeInput = document.getElementById('sidebar-font-size');
            const bgColorInput = document.getElementById('sidebar-bg-color');
            const autoFitInput = document.getElementById('sidebar-text-autofit');
            const lineSpacingInput = document.getElementById('sidebar-line-spacing');
            const centerTextBtn = document.getElementById('sidebar-text-center');
            if (textInput) textInput.value = entry.text || '';
            if (fontFamilyInput) fontFamilyInput.value = entry.fontFamily || 'Arial';
            if (fontSizeInput) fontSizeInput.value = uiFontSize;
            if (bgColorInput) bgColorInput.value = entry.bgPreset || 'white';
            if (autoFitInput) autoFitInput.checked = !!(entry.autoFitText || entry.autoFitSingleLine);
            if (lineSpacingInput) lineSpacingInput.value = String(Number(entry.lineSpacing) || 1);
            if (centerTextBtn) centerTextBtn.classList.toggle('mode-active', entry.textAlign === 'center' || !!entry.centerText);
            textLinesModeEnabled = !!(uiLines && uiLines.length > 1);
            populateTextLineEditor(uiLines || [], {
                text: entry.text || '',
                fontFamily: entry.fontFamily || 'Arial',
                fontSize: uiFontSize
            });
        }
    }

    overlay.style.display = 'flex';

    document.getElementById('sel-x').value = selection.x;
    document.getElementById('sel-y').value = selection.y;
    document.getElementById('sel-w').value = selection.w;
    document.getElementById('sel-h').value = selection.h;
}

function updateOverlayBoundsFromSelection() {
    if (!selection || !canvas) return;
    const overlay = document.getElementById('edit-overlay');
    const editorArea = document.getElementById('editor-area');
    if (!overlay || !editorArea) return;
    const editorRect = editorArea.getBoundingClientRect();
    const canvasRect = canvas.getBoundingClientRect();
    const scaleX = canvasRect.width / canvas.width;
    const scaleY = canvasRect.height / canvas.height;
    overlay.style.left = (editorArea.scrollLeft + (canvasRect.left - editorRect.left) + selection.x * scaleX) + 'px';
    overlay.style.top = (editorArea.scrollTop + (canvasRect.top - editorRect.top) + selection.y * scaleY) + 'px';
    overlay.style.width = selection.w * scaleX + 'px';
    overlay.style.height = selection.h * scaleY + 'px';

    document.getElementById('sel-x').value = selection.x;
    document.getElementById('sel-y').value = selection.y;
    document.getElementById('sel-w').value = selection.w;
    document.getElementById('sel-h').value = selection.h;
}

function hideEditOverlay() {
    // simply hide the overlay and restore cursor; do not touch currentMode here.
    // callers who want to stop selecting should clear `currentMode` explicitly.
    const overlay = document.getElementById('edit-overlay');
    if (overlay) {
        overlay.style.display = 'none';
        overlay.classList.remove('image-mode');
    }
    if (canvas) canvas.style.cursor = '';
    imageMouseTransformState = null;
    hideSelectedAppliedImageDuringEdit = false;
    setTextOverflowIndicator(false);
}

// Setup all event listeners after DOM is ready
function setupEventListeners() {
    console.log('setupEventListeners called');
    try {
        const selectImageBtn = document.getElementById('select-image');
        const editAppliedImageBtn = document.getElementById('edit-applied-image');
        const cancelEditAppliedImageBtn = document.getElementById('cancel-edit-applied-image');
        const selectTextBtn = document.getElementById('select-text');
        const backToSelectionTextBtn = document.getElementById('back-to-selection-text');
        const backToSelectionImageBtn = document.getElementById('back-to-selection-image');
        // set up sidebar tabs if present
        setupTabSwitching();
        const imageUploadInput = document.getElementById('image-upload');
        const overlayImageBtn = document.getElementById('overlay-image-btn');
        const overlayTextBtn = document.getElementById('overlay-text-btn');
        const overlayCancelBtn = document.getElementById('overlay-cancel-btn');
        const overlayDeleteBtn = document.getElementById('overlay-delete-btn');
        const imageMouseLayer = document.getElementById('image-mouse-layer');
        const sidebarTextInput = document.getElementById('sidebar-text-input');
        const sidebarFontFamily = document.getElementById('sidebar-font-family');
        const sidebarFontSize = document.getElementById('sidebar-font-size');
        const sidebarBgColor = document.getElementById('sidebar-bg-color');
        const sidebarTextAutofit = document.getElementById('sidebar-text-autofit');
        const sidebarLineSpacing = document.getElementById('sidebar-line-spacing');
        const sidebarTextCenter = document.getElementById('sidebar-text-center');
        const splitTextLinesBtn = document.getElementById('split-text-lines');
        const textLinesEditor = document.getElementById('text-lines-editor');
        const previewTextBtn = document.getElementById('preview-text');
        const applyTextBtn = document.getElementById('apply-text');
        const chooseImageBtn = document.getElementById('choose-image');
        const previewImageBtn = document.getElementById('preview-image');
        const applyImageBtn = document.getElementById('apply-image');
        const outputPdfNameInput = document.getElementById('pdf-output-name');
        const editorArea = document.getElementById('editor-area');
        const editableDownloadLink = document.getElementById('editable-download-link');
        const importEditableBtn = document.getElementById('import-editable-btn');
        const importEditableFileInput = document.getElementById('import-editable-file');
        const zoomOutBtn = document.getElementById('zoom-out');
        const zoomInBtn = document.getElementById('zoom-in');
        const zoomResetBtn = document.getElementById('zoom-reset');
        const zoomRangeInput = document.getElementById('zoom-range');
        const panXInput = document.getElementById('pan-x');
        const panYInput = document.getElementById('pan-y');
        const imageFilterTypeInput = document.getElementById('image-filter-type');
        const imageFilterIntensityInput = document.getElementById('image-filter-intensity');
        const imageTransformInputs = [
            document.getElementById('image-scale')
        ];
        const saveBtn = document.getElementById('save-canvas');
        const editLink = document.getElementById('edit-link');
        let pickAppliedImageMode = false;

        const baseInputName = (() => {
            const raw = (pdfUrl || '').split('/').pop() || 'documento.pdf';
            return raw.replace(/\.pdf$/i, '') + '-edited';
        })();
        if (outputPdfNameInput && !outputPdfNameInput.value.trim()) {
            outputPdfNameInput.value = baseInputName;
        }

        const setPickAppliedImageMode = (enabled) => {
            pickAppliedImageMode = enabled;
            if (editAppliedImageBtn) {
                editAppliedImageBtn.classList.toggle('mode-active', enabled);
                editAppliedImageBtn.textContent = enabled ? 'Clique numa imagem para editar' : 'Editar imagem aplicada';
            }
            if (cancelEditAppliedImageBtn) {
                cancelEditAppliedImageBtn.style.display = enabled ? 'block' : 'none';
            }
            if (canvas) {
                if (enabled) {
                    canvas.style.cursor = 'pointer';
                } else if (currentMode === 'image' || currentMode === 'text') {
                    canvas.style.cursor = 'crosshair';
                } else {
                    canvas.style.cursor = '';
                }
            }
        };

        const returnToSelectionMenu = () => {
            currentMode = null;
            editingIndex = null;
            selection = null;
            previewEdit = null;
            hideEditOverlay();
            setPickAppliedImageMode(false);
            if (canvas) canvas.style.cursor = '';
            redrawCanvas();
        };
        
        console.log('selectImageBtn:', selectImageBtn);
        console.log('selectTextBtn:', selectTextBtn);
        console.log('imageUploadInput:', imageUploadInput);
        
        // Sidebar buttons
        if (selectImageBtn) {
            selectImageBtn.addEventListener('click', () => {
                console.log('select-image clicked');
                setPickAppliedImageMode(false);
                currentMode = 'image';
                editingIndex = null; // new edit
                previewEdit = null;
                hideEditOverlay();
                canvas.style.cursor = 'crosshair';
                setImageTransformControls(DEFAULT_IMAGE_TRANSFORM);
                setImageFilterControls(DEFAULT_IMAGE_FILTER);
                redrawCanvas();
            });
        }
        
        if (selectTextBtn) {
            selectTextBtn.addEventListener('click', () => {
                console.log('select-text clicked');
                setPickAppliedImageMode(false);
                currentMode = 'text';
                editingIndex = null;
                previewEdit = null;
                hideEditOverlay();
                canvas.style.cursor = 'crosshair';
                resetTextSidebarForNewEntry();
                redrawCanvas();
            });
        }

        if (backToSelectionTextBtn) {
            backToSelectionTextBtn.addEventListener('click', () => {
                returnToSelectionMenu();
            });
        }

        if (backToSelectionImageBtn) {
            backToSelectionImageBtn.addEventListener('click', () => {
                returnToSelectionMenu();
            });
        }

        if (editAppliedImageBtn) {
            editAppliedImageBtn.addEventListener('click', () => {
                const nextEnabled = !pickAppliedImageMode;
                currentMode = null;
                editingIndex = null;
                selection = null;
                previewEdit = null;
                hideEditOverlay();
                setPickAppliedImageMode(nextEnabled);
                redrawCanvas();
            });
        }

        if (cancelEditAppliedImageBtn) {
            cancelEditAppliedImageBtn.addEventListener('click', () => {
                currentMode = null;
                editingIndex = null;
                selection = null;
                previewEdit = null;
                hideEditOverlay();
                setPickAppliedImageMode(false);
                redrawCanvas();
            });
        }

        const liveTextPreviewHandler = () => {
            triggerLiveTextPreview();
        };
        if (sidebarTextInput) sidebarTextInput.addEventListener('input', liveTextPreviewHandler);
        if (sidebarFontFamily) sidebarFontFamily.addEventListener('change', liveTextPreviewHandler);
        if (sidebarFontSize) sidebarFontSize.addEventListener('input', liveTextPreviewHandler);
        if (sidebarBgColor) sidebarBgColor.addEventListener('change', liveTextPreviewHandler);
        if (sidebarTextAutofit) sidebarTextAutofit.addEventListener('change', liveTextPreviewHandler);
        if (sidebarLineSpacing) sidebarLineSpacing.addEventListener('input', liveTextPreviewHandler);
        if (sidebarTextCenter) {
            sidebarTextCenter.addEventListener('click', () => {
                sidebarTextCenter.classList.toggle('mode-active');
                liveTextPreviewHandler();
            });
        }

        if (splitTextLinesBtn) {
            splitTextLinesBtn.addEventListener('click', () => {
                const cfg = getSidebarTextConfig();
                const rawText = sidebarTextInput ? sidebarTextInput.value : '';
                let splitLines = splitRawTextIntoEditableLines(rawText);

                if (splitLines.length === 0 && cfg.lines && cfg.lines.length > 0) {
                    splitLines = cfg.lines.map(line => (line.text || '').trim()).filter(Boolean);
                }

                if (splitLines.length <= 1) {
                    const single = splitLines[0] || (rawText || '').trim();
                    splitLines = [single, '', ''];
                }

                const lineEntries = splitLines.map(text => ({
                    text,
                    fontFamily: cfg.fontFamily || 'Arial',
                    fontSize: cfg.fontSize || 16
                }));

                textLinesModeEnabled = true;
                populateTextLineEditor(lineEntries, cfg);
                liveTextPreviewHandler();
            });
        }

        if (textLinesEditor) {
            textLinesEditor.addEventListener('input', liveTextPreviewHandler);
            textLinesEditor.addEventListener('change', liveTextPreviewHandler);
        }

        if (previewTextBtn) {
            previewTextBtn.addEventListener('click', () => {
                if (currentMode === 'text' && selection) updateTextPreview();
            });
        }

        if (applyTextBtn) {
            applyTextBtn.addEventListener('click', () => {
                if (currentMode !== 'text' || !selection) return;
                const entry = buildTextEntryFromSelection();
                if (!entry) return;
                if (editingIndex !== null) {
                    edits[editingIndex] = entry;
                } else {
                    edits.push(entry);
                }
                previewEdit = null;
                hideEditOverlay();
                selection = null;
                editingIndex = null;
                currentMode = null;
                resetTextSidebarForNewEntry();
                redrawCanvas();
            });
        }

        if (chooseImageBtn) {
            chooseImageBtn.addEventListener('click', () => {
                if (currentMode !== 'image') return;
                // ensure selecting the same file still triggers change event
                imageUploadInput.value = '';
                imageUploadInput.click();
            });
        }

        imageTransformInputs.forEach(input => {
            if (!input) return;
            input.addEventListener('input', () => {
                updateImageTransformLabels();
                if (currentMode === 'image' && selection && selectedImageDataUrl) {
                    updateImagePreview();
                }
            });
        });
        updateImageTransformLabels();

        if (imageFilterTypeInput) {
            imageFilterTypeInput.addEventListener('change', () => {
                const cfg = getImageFilterConfig();
                if (cfg.type !== 'none' && cfg.intensity <= 0) {
                    setRangeValue('image-filter-intensity', 100);
                    updateImageFilterLabels();
                }
                applyImageFilterAndPreview(getImageFilterConfig());
            });
        }
        if (imageFilterIntensityInput) {
            imageFilterIntensityInput.addEventListener('input', () => {
                updateImageFilterLabels();
                applyImageFilterAndPreview(getImageFilterConfig());
            });
        }
        updateImageFilterLabels();

        if (zoomRangeInput) {
            zoomRangeInput.addEventListener('input', () => {
                const value = parseFloat(zoomRangeInput.value);
                if (Number.isFinite(value)) {
                    setViewZoom(value / 100);
                }
            });
        }
        if (zoomOutBtn) {
            zoomOutBtn.addEventListener('click', () => {
                setViewZoom(viewZoom - 0.1);
            });
        }
        if (zoomInBtn) {
            zoomInBtn.addEventListener('click', () => {
                setViewZoom(viewZoom + 0.1);
            });
        }
        if (zoomResetBtn) {
            zoomResetBtn.addEventListener('click', () => {
                setViewZoom(1);
            });
        }

        if (panXInput && editorArea) {
            panXInput.addEventListener('input', () => {
                editorArea.scrollLeft = parseFloat(panXInput.value) || 0;
            });
        }
        if (panYInput && editorArea) {
            panYInput.addEventListener('input', () => {
                editorArea.scrollTop = parseFloat(panYInput.value) || 0;
            });
        }
        if (editorArea) {
            editorArea.addEventListener('scroll', () => {
                updatePanControls();
                const overlay = document.getElementById('edit-overlay');
                if (selection && overlay && overlay.style.display !== 'none') {
                    updateOverlayBoundsFromSelection();
                }
            });
        }

        window.addEventListener('resize', () => {
            setViewZoom(viewZoom);
        });

        // prevent browser zoom on Ctrl+wheel inside editor and use app zoom instead
        if (editorArea) {
            editorArea.addEventListener('wheel', (e) => {
                if (!e.ctrlKey) return;
                e.preventDefault();
                const delta = e.deltaY > 0 ? -0.1 : 0.1;
                setViewZoom(viewZoom + delta, { anchorClientX: e.clientX, anchorClientY: e.clientY });
            }, { passive: false });
        }

        setViewZoom(viewZoom);
        updatePanControls();

        if (previewImageBtn) {
            previewImageBtn.addEventListener('click', () => {
                if (currentMode === 'image' && selection) updateImagePreview();
            });
        }

        if (imageMouseLayer) {
            imageMouseLayer.addEventListener('pointerdown', (e) => {
                if (currentMode !== 'image' || !selection || !selectedImageDataUrl) return;
                const action = e.target.dataset.action || 'move';
                if (action !== 'move' && !isDiagonalHandleAction(action) && !isSideHandleAction(action)) return;
                const canvasRect = canvas.getBoundingClientRect();
                imageMouseTransformState = {
                    action,
                    startClientX: e.clientX,
                    startClientY: e.clientY,
                    scaleX: canvas.width / canvasRect.width,
                    scaleY: canvas.height / canvasRect.height,
                    startSelection: { ...selection },
                    startTransform: getImageTransformConfig()
                };
                imageMouseLayer.setPointerCapture(e.pointerId);
                e.preventDefault();
                e.stopPropagation();
            });

            imageMouseLayer.addEventListener('pointermove', (e) => {
                if (!imageMouseTransformState) return;
                const dx = e.clientX - imageMouseTransformState.startClientX;
                const dy = e.clientY - imageMouseTransformState.startClientY;
                const dxCanvas = dx * imageMouseTransformState.scaleX;
                const dyCanvas = dy * imageMouseTransformState.scaleY;
                const baseSel = imageMouseTransformState.startSelection;
                if (imageMouseTransformState.action === 'move') {
                    const baseTransform = imageMouseTransformState.startTransform || getImageTransformConfig();
                    const movedTransform = {
                        ...baseTransform,
                        offsetX: (baseTransform.offsetX || 0) + dxCanvas,
                        offsetY: (baseTransform.offsetY || 0) + dyCanvas
                    };
                    applyImageTransformAndPreview(movedTransform);
                    e.preventDefault();
                    return;
                }

                if (isSideHandleAction(imageMouseTransformState.action)) {
                    const baseTransform = imageMouseTransformState.startTransform || getImageTransformConfig();
                    const cropResult = cropAndResizeFromHandle(
                        baseTransform,
                        baseSel,
                        imageMouseTransformState.action,
                        dxCanvas,
                        dyCanvas
                    );
                    applyImageTransformAndPreview(cropResult.nextTransform);
                    e.preventDefault();
                    return;
                }

                if (!isDiagonalHandleAction(imageMouseTransformState.action)) {
                    return;
                }

                const baseTransform = imageMouseTransformState.startTransform || getImageTransformConfig();
                const resizedTransform = resizeImageTransformFromDiagonalHandle(
                    baseTransform,
                    baseSel,
                    imageMouseTransformState.action,
                    dxCanvas,
                    dyCanvas
                );
                applyImageTransformAndPreview(resizedTransform);
                e.preventDefault();
            });

            imageMouseLayer.addEventListener('pointerup', (e) => {
                if (!imageMouseTransformState) return;
                imageMouseTransformState = null;
                if (imageMouseLayer.hasPointerCapture(e.pointerId)) {
                    imageMouseLayer.releasePointerCapture(e.pointerId);
                }
            });

            imageMouseLayer.addEventListener('pointercancel', () => {
                imageMouseTransformState = null;
            });
        }

        if (applyImageBtn) {
            applyImageBtn.addEventListener('click', () => {
                if (currentMode !== 'image' || !selection || !selectedImageDataUrl) return;
                const existingEntry = editingIndex !== null ? edits[editingIndex] : null;
                let reusableImgObj = (() => {
                    if (selectedImageObj && selectedImageObj.complete) {
                        return selectedImageObj;
                    }
                    if (previewEdit && previewEdit.type === 'image' && previewEdit.imgObj && previewEdit.imgObj.complete) {
                        return previewEdit.imgObj;
                    }
                    if (existingEntry && existingEntry.type === 'image' && existingEntry.imgObj && existingEntry.imgObj.complete && existingEntry.dataUrl === selectedImageDataUrl) {
                        return existingEntry.imgObj;
                    }
                    return null;
                })();

                if (!reusableImgObj) {
                    reusableImgObj = new Image();
                    reusableImgObj.onload = () => scheduleRedrawCanvas();
                    reusableImgObj.src = selectedImageDataUrl;
                } else if (!reusableImgObj.src) {
                    reusableImgObj.src = selectedImageDataUrl;
                }

                const entry = {
                    ...canvasRectToBaseRect(selection),
                    type: 'image',
                    coordSpace: 'base',
                    dataUrl: selectedImageDataUrl,
                    imageTransform: transformCanvasToBase(getImageTransformConfig()),
                    imageFilter: normalizeImageFilter(getImageFilterConfig()),
                    imgObj: reusableImgObj
                };
                if (editingIndex !== null) {
                    edits[editingIndex] = entry;
                } else {
                    edits.push(entry);
                }
                previewEdit = null;
                hideEditOverlay();
                setPickAppliedImageMode(false);
                selection = null;
                editingIndex = null;
                currentMode = 'image';
                canvas.style.cursor = 'crosshair';
                redrawCanvas();
                window.requestAnimationFrame(() => redrawCanvas());
            });
        }
        
        // Use global pointer handlers as a fallback so selection works even
        // if some element overlays the canvas. Map pointer coordinates to canvas
        // taking into account any CSS scaling of the <canvas> element.
        function toCanvasCoords(clientX, clientY) {
            const rect = canvas.getBoundingClientRect();
            const scaleX = canvas.width / rect.width;
            const scaleY = canvas.height / rect.height;
            return {
                x: (clientX - rect.left) * scaleX,
                y: (clientY - rect.top) * scaleY,
                rect,
                scaleX,
                scaleY
            };
        }

        document.addEventListener('pointerdown', (e) => {
            // if the click started inside the edit overlay, let it handle the event
            const overlayElt = document.getElementById('edit-overlay');
            if (overlayElt && overlayElt.contains(e.target)) {
                return;
            }

            // Only interact with selections/edits when pointer starts on the canvas.
            // This avoids hijacking footer sliders and other UI controls.
            if (e.target !== canvas) {
                return;
            }

            // if we're not in selection mode, maybe the user clicked an existing edit
            if (!currentMode) {
                const coords = toCanvasCoords(e.clientX, e.clientY);
                for (let i = 0; i < edits.length; i++) {
                    const ed = edits[i];
                    if (pickAppliedImageMode && ed.type !== 'image') {
                        continue;
                    }
                    const rect = getEntryCanvasRect(ed);
                    if (!rect) continue;
                    if (coords.x >= rect.x && coords.y >= rect.y && coords.x <= rect.x + rect.w && coords.y <= rect.y + rect.h) {
                        console.log('clicked existing edit', i, ed);
                        selection = { x: rect.x, y: rect.y, w: rect.w, h: rect.h };
                        selectionBgColor = ed.bgColor || getDominantColor(rect.x, rect.y, rect.w, rect.h);
                        editingIndex = i;
                        currentMode = ed.type;
                        setPickAppliedImageMode(false);
                        showEditOverlay();
                        return;
                    }
                }
                return;
            }

            const coords = toCanvasCoords(e.clientX, e.clientY);
            // ignore if outside canvas
            if (coords.x < 0 || coords.y < 0 || coords.x > canvas.width || coords.y > canvas.height) return;
            startX = coords.x;
            startY = coords.y;
            isSelecting = true;
            console.log('pointerdown -> selection started at', startX, startY);
            const selRect = document.getElementById('sel-rect');
            if (selRect) {
                selRect.style.display = 'block';
                const editorArea = document.getElementById('editor-area');
                const editorRect = editorArea.getBoundingClientRect();
                // convert canvas coords back to CSS pixels
                const cssX = startX / coords.scaleX;
                const cssY = startY / coords.scaleY;
                selRect.style.left = (editorArea.scrollLeft + (coords.rect.left - editorRect.left) + cssX) + 'px';
                selRect.style.top = (editorArea.scrollTop + (coords.rect.top - editorRect.top) + cssY) + 'px';
                selRect.style.width = '0px';
                selRect.style.height = '0px';
            }
            e.preventDefault();
        });

        document.addEventListener('pointermove', (e) => {
            if (!isSelecting) return;
            const coords = toCanvasCoords(e.clientX, e.clientY);
            // clamp coords inside canvas pixel dimensions
            const x = Math.max(0, Math.min(coords.x, canvas.width));
            const y = Math.max(0, Math.min(coords.y, canvas.height));
            selection = { x: Math.min(startX, x), y: Math.min(startY, y), w: Math.abs(x - startX), h: Math.abs(y - startY) };
            const selRect = document.getElementById('sel-rect');
            if (selRect && selection) {
                const editorArea = document.getElementById('editor-area');
                const editorRect = editorArea.getBoundingClientRect();
                const cssX = selection.x / coords.scaleX;
                const cssY = selection.y / coords.scaleY;
                const cssW = selection.w / coords.scaleX;
                const cssH = selection.h / coords.scaleY;
                selRect.style.left = (editorArea.scrollLeft + (coords.rect.left - editorRect.left) + cssX) + 'px';
                selRect.style.top = (editorArea.scrollTop + (coords.rect.top - editorRect.top) + cssY) + 'px';
                selRect.style.width = cssW + 'px';
                selRect.style.height = cssH + 'px';
            }
            // only prevent default while actively selecting
            e.preventDefault();
        });

        document.addEventListener('pointerup', (e) => {
            if (!isSelecting) return;
            isSelecting = false;
            console.log('pointerup, selection size:', selection && selection.w, 'x', selection && selection.h);
            const selRect = document.getElementById('sel-rect');
            if (selRect) selRect.style.display = 'none';
            if (selection && selection.w > 10 && selection.h > 10) {
                selectionBgColor = getDominantColor(selection.x, selection.y, selection.w, selection.h);
                showEditOverlay();
                triggerLiveTextPreview();
            } else {
                console.log('Selection too small');
            }
            // preventDefault to stop unwanted drags after selection
            e.preventDefault();
        });
        
        // Image upload
        if (imageUploadInput) {
            imageUploadInput.addEventListener('change', (e) => {
                console.log('Image upload change event');
                const file = e.target.files[0];
                if (!file || !selection) return;

                // for a fresh image insertion, start from neutral transform/filter
                if (editingIndex === null) {
                    setImageTransformControls(DEFAULT_IMAGE_TRANSFORM);
                    setImageFilterControls(DEFAULT_IMAGE_FILTER);
                }

                const reader = new FileReader();
                reader.onload = () => {
                    const dataUrl = reader.result;
                    if (typeof dataUrl !== 'string') return;
                    const img = new Image();
                    img.onload = function() {
                        console.log('Image loaded, ready to update entry');
                        selectedImageDataUrl = dataUrl;
                        selectedImageObj = img;
                        updateImagePreview();
                    };
                    img.src = dataUrl;
                };
                reader.readAsDataURL(file);
            });
        }
        
        // Overlay buttons
        if (overlayImageBtn) {
            overlayImageBtn.addEventListener('click', () => {
                console.log('overlay-image-btn clicked');
                if (currentMode === 'image' && selection) {
                    imageUploadInput.value = '';
                    imageUploadInput.click();
                }
            });
        }
        
        if (overlayTextBtn) {
            overlayTextBtn.addEventListener('click', () => {
                console.log('overlay-text-btn clicked');
                // keep compatibility: trigger sidebar apply flow
                if (applyTextBtn) applyTextBtn.click();
            });
        }
        
        if (overlayDeleteBtn) {
            overlayDeleteBtn.addEventListener('click', () => {
                console.log('overlay-delete-btn clicked');
                if (editingIndex !== null) {
                    edits.splice(editingIndex, 1);
                    redrawCanvas();
                }
                editingIndex = null;
                previewEdit = null;
                hideEditOverlay();
                currentMode = null;
            });
        }

        if (overlayCancelBtn) {
            overlayCancelBtn.addEventListener('click', () => {
                console.log('overlay-cancel-btn clicked');
                hideEditOverlay();
                // keep currentMode so user can immediately start a new selection
                editingIndex = null;
                previewEdit = null;
                redrawCanvas();
            });
        }

        if (saveBtn) {
            saveBtn.addEventListener('click', async () => {
                console.log('save-canvas clicked');
                const desiredPdfName = outputPdfNameInput && outputPdfNameInput.value.trim()
                    ? outputPdfNameInput.value.trim()
                    : baseInputName;
                const sourcePdfName = sourcePdfOverride || ((pdfUrl || '').split('/').pop() || '');
                const currentPdfName = ((pdfUrl || '').split('/').pop() || '');
                const preparedEdits = await prepareEditsForPdfSave();
                fetch('/save', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ 
                        sourcePdf: sourcePdfName,
                        currentPdf: currentPdfName,
                        pdfFilename: desiredPdfName,
                        canvasWidth: canvas.width,
                        canvasHeight: canvas.height,
                        edits: preparedEdits,
                        editableEdits: serializeEditableEdits()
                    })
                }).then(async res => {
                    const body = await res.json();
                    if (!res.ok) {
                        throw new Error(body.message || body.error || 'Falha ao salvar');
                    }
                    return body;
                })
                  .then(res => {
                      if (res.status === 'ok') {
                          const link = document.getElementById('download-link');
                          link.href = res.pdf || res.png || '#';
                          link.style.display = 'inline-block';
                          link.textContent = res.pdf ? 'Download PDF' : 'Download PNG';
                          if (editableDownloadLink && res.editable) {
                              editableDownloadLink.href = res.editable;
                              editableDownloadLink.style.display = 'inline-block';
                              editableDownloadLink.setAttribute('download', 'projeto-editavel.json');
                          }
                          if (editLink && res.editUrl) {
                              editLink.href = res.editUrl;
                              editLink.style.display = 'inline-block';
                          }
                      } else {
                          alert(res.message || 'Falha ao salvar');
                      }
                  }).catch(err => { alert('Erro: ' + (err && err.message ? err.message : err)); });
            });
        }

        if (importEditableBtn && importEditableFileInput) {
            importEditableBtn.addEventListener('click', () => {
                importEditableFileInput.value = '';
                importEditableFileInput.click();
            });

            importEditableFileInput.addEventListener('change', async (e) => {
                const file = e.target.files && e.target.files[0];
                if (!file) return;
                try {
                    const raw = await file.text();
                    const project = JSON.parse(raw);
                    if (!project || !Array.isArray(project.edits)) {
                        alert('Ficheiro inválido de cópia editável.');
                        return;
                    }
                    applyEditableProjectData(project);
                    alert('Cópia editável carregada com sucesso.');
                } catch (err) {
                    alert('Erro ao carregar cópia editável: ' + err);
                }
            });
        }
        
        populateTextLineEditor([], getSidebarTextConfig());
        console.log('All event listeners setup completed successfully');
        
    } catch (err) {
        console.error('Error in setupEventListeners:', err);
    }
}

// tab-switching logic for sidebar
function setupTabSwitching() {
    const tabs = document.querySelectorAll('.tab-btn');
    const contents = document.querySelectorAll('.tab-content');
    tabs.forEach(btn => {
        btn.addEventListener('click', () => {
            tabs.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            contents.forEach(c => c.style.display = 'none');
            const id = btn.dataset.tab + '-tab';
            const target = document.getElementById(id);
            if (target) target.style.display = 'block';
        });
    });
}
