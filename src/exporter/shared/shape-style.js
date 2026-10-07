/**
 * DrawingML shape-properties fragments (`a:prstGeom`/fill/`a:ln`/
 * `a:effectLst`) shared between PPTX's `<p:spPr>` (pptx/shapes.js) and
 * DOCX's `<wps:spPr>` (docx/document.js) — both use the exact same
 * `a:CT_ShapeProperties` content type (ECMA-376 §20.1.2.2.35), so the same
 * fragment slots into either host element unchanged; only the surrounding
 * tag name differs per format.
 */
import { XmlBuilder } from '../../utils/xml-builder.js';

// CSS's gradient angle is measured clockwise from pointing up (0deg = up,
// 90deg = right); DrawingML's `a:lin ang` is measured clockwise from
// pointing right (0 = right). Rotating by -90deg (mod 360) converts one
// convention to the other; `a:lin ang` itself is in 60,000ths of a degree.
export function gradientAngleEmu(angleDeg) {
  const drawingMlDeg = ((angleDeg - 90) % 360 + 360) % 360;
  return Math.round(drawingMlDeg * 60000);
}

// CT_ShapeProperties requires gsLst before lin (ECMA-376 §20.1.8.33).
export function gradFillXml(gradient) {
  const b = new XmlBuilder();
  b.el('a:gradFill', { rotWithShape: 1 });
  b.el('a:gsLst');
  for (const stop of gradient.stops) {
    b.el('a:gs', { pos: Math.round(stop.offset * 1000) });
    b.self('a:srgbClr', { val: stop.color });
    b.up();
  }
  b.up();
  b.self('a:lin', { ang: gradientAngleEmu(gradient.angleDeg), scaled: 1 });
  b.up();
  return b.end();
}

// A shape's own fill — the real multi-stop gradient when the source CSS had
// one (see extractBackgroundGradient), else the flat backgroundColor
// approximation every other exporter already uses, else no fill at all
// (a border-only card with no background).
export function shapeFillXml(node) {
  if (node.backgroundGradient) return gradFillXml(node.backgroundGradient);
  if (node.backgroundColor) {
    const b = new XmlBuilder();
    b.el('a:solidFill');
    b.self('a:srgbClr', { val: node.backgroundColor });
    b.up();
    return b.end();
  }
  return '<a:noFill/>';
}

// `border` is the `{ color, sizeEighthPt, style }` shape extractTableBorder/
// resolveBorder produce — sizeEighthPt matches DOCX's own w:sz unit (eighths
// of a point), converted here straight to EMU.
const EMU_PER_PT = 12700;

export function shapeLineXml(border) {
  if (!border) return '';
  const b = new XmlBuilder();
  b.el('a:ln', { w: Math.round((border.sizeEighthPt / 8) * EMU_PER_PT) });
  b.el('a:solidFill');
  b.self('a:srgbClr', { val: border.color });
  b.up();
  b.up();
  return b.end();
}

// DrawingML can represent a real drop shadow (`a:outerShdw`), but nothing
// in the IDM carries a CSS box-shadow's own offset/blur/spread/color — see
// hasBoxShadow — so this is one fixed, native-looking shadow (the same
// preset PowerPoint's own "Offset: Bottom Right" shadow style uses) applied
// whenever the source declared *any* real box-shadow, rather than an
// attempt to reproduce its exact parameters.
export function shadowEffectXml() {
  const b = new XmlBuilder();
  b.el('a:effectLst');
  b.el('a:outerShdw', { blurRad: 90000, dist: 38100, dir: 5400000, rotWithShape: 0 });
  b.el('a:srgbClr', { val: '000000' });
  b.self('a:alpha', { val: 20000 });
  b.up();
  b.up();
  b.up();
  return b.end();
}

// Whether `node` needs a real geometry/fill/effect that plain paragraph
// borders/shading (OOXML's w:pBdr/w:shd, or a table's own borders/tcPr
// shading) simply cannot represent at all: a rounded corner, a circle/oval,
// a drop shadow, or an explicit fixed height (w:pBdr/w:shd has no height
// concept at all — a plain paragraph is always exactly as tall as its own
// text). A plain flat background/gradient/border with square corners and no
// declared height is left alone — every exporter's existing flat-border/
// flat-fill rendering already handles that case fine (see
// extractBackgroundColor's own gradient fallback), so only promoting the
// genuinely unrepresentable cases avoids turning every bordered box into a
// heavier DrawingML shape unnecessarily. `circular` implies `rounded` is
// already set too (see dom-to-idm.js's isCircularBorderRadius/
// hasBorderRadius), so this is mostly belt-and-suspenders, not a new trigger
// in practice.
export function needsDrawingShape(node) {
  return Boolean(node.rounded || node.shadow || node.circular || node.heightPx != null);
}

// The `a:prstGeom` preset to draw a card/box shape as, given what its own
// CSS declared. `circular` (a `border-radius:50%`-or-more box — see
// isCircularBorderRadius) always wins over a plain `rounded`: a box meant to
// be a full circle/oval should never fall back to a merely-rounded
// rectangle. Shared between PPTX's renderCardShape and DOCX's
// renderCardShapeDrawing, which both make this exact three-way choice.
// Resolves one SVG vector shape (see dom-to-idm.js's extractSvgShapes /
// idm/index.js's svgShapes) into the DrawingML geometry both DOCX
// (document.js's wp:anchor-based renderer) and PPTX (shapes.js's
// renderSvgShapesShape) need to draw it — `x`/`y`/`cx`/`cy` are all in the
// same unit as `scale` produces (EMU, in both current callers), with the
// shape's own top-left already resolved from whichever kind-specific
// geometry it carries. PPTX's own renderSvgShapesShape predates this
// extraction and still inlines the identical math itself (not worth the
// regression risk of refactoring already-shipped, tested code just for
// DRY) — this is the canonical version new callers (DOCX) build on.
export function resolveSvgShapeGeometry(shape, scale) {
  switch (shape.kind) {
    case 'rect':
      return { prst: 'rect', x: shape.x * scale, y: shape.y * scale, cx: shape.width * scale, cy: shape.height * scale };
    case 'circle':
      return { prst: 'ellipse', x: (shape.cx - shape.r) * scale, y: (shape.cy - shape.r) * scale, cx: shape.r * 2 * scale, cy: shape.r * 2 * scale };
    case 'ellipse':
      return { prst: 'ellipse', x: (shape.cx - shape.rx) * scale, y: (shape.cy - shape.ry) * scale, cx: shape.rx * 2 * scale, cy: shape.ry * 2 * scale };
    case 'line': {
      const x = Math.min(shape.x1, shape.x2) * scale;
      const y = Math.min(shape.y1, shape.y2) * scale;
      const cx = Math.abs(shape.x2 - shape.x1) * scale;
      const cy = Math.abs(shape.y2 - shape.y1) * scale;
      // DrawingML's `line` preset always draws its diagonal from top-left to
      // bottom-right of its own bounding box — flipH mirrors it to instead
      // run top-right-to-bottom-left when the real endpoints go that way
      // (an XOR: either axis reversing alone flips it, both together don't).
      const flipH = (shape.x2 < shape.x1) !== (shape.y2 < shape.y1);
      return { prst: 'line', x, y, cx, cy, flipH };
    }
    case 'polygon':
    case 'polyline': {
      const scaled = shape.points.map(([x, y]) => [x * scale, y * scale]);
      const minX = Math.min(...scaled.map((p) => p[0]));
      const minY = Math.min(...scaled.map((p) => p[1]));
      const maxX = Math.max(...scaled.map((p) => p[0]));
      const maxY = Math.max(...scaled.map((p) => p[1]));
      return {
        prst: 'custom',
        x: minX, y: minY, cx: Math.max(1, maxX - minX), cy: Math.max(1, maxY - minY),
        points: scaled.map(([x, y]) => [Math.round(x - minX), Math.round(y - minY)]),
        closed: shape.kind === 'polygon',
      };
    }
    default:
      return undefined;
  }
}

export function cardPresetGeometry(node) {
  if (node.circular) return 'ellipse';
  return node.border || node.rounded ? 'roundRect' : 'rect';
}
