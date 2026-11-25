import { flatText } from '../model.ts'
import type {
  Atom,
  AtomKind,
  Block,
  BorderEdge,
  Cell,
  CellBorders,
  HighlightName,
  LineSpacing,
  Marks,
  MediaKey,
  Paragraph,
  ParaId,
  ParaLook,
  Pos,
  Pt,
  Row,
  RunLook,
  Section,
  Table,
  TableBorders,
} from '../model.ts'

/** View-only decorations. Rendered, never stored, never measured into the cache. */
export type Decoration = { readonly kind: 'preedit'; readonly pos: Pos; readonly text: string }

export interface RenderContext {
  /** Resolved look for a run: rdocx's (paragraph style, character style) look overlaid with direct marks. Cached. */
  look(para: Paragraph, marks: Marks): RunLook
  paraLook(para: Paragraph): ParaLook
  /** CSS font-family stack for a family name, from fonts.ts's alias table. */
  fontStack(family: string): string
  /** Word line height of one line of this font at this size, from the font file's OS/2 and hhea metrics (backend). */
  lineHeightPx(family: string, sizePt: Pt, rule: LineSpacing): number
  /** List marker text for a paragraph, derived from document order. Null when not in a list. */
  marker(id: ParaId): string | null
  imageUrl(key: MediaKey): string
  /** Section grid line pitch in px, or null. Exact spacing stays absolute; other rules snap up to a whole number of rows. */
  gridPitchPx: number | null
  /**
   * Whether the resolved face for this family has CJK glyphs. Undefined until metrics have loaded.
   * An east-asian segment whose face has none is painted with 宋体.
   */
  cjkGlyphs?: (family: string) => boolean | undefined
  /** Text inside page-number fields, set per page by pages.ts for header and footer stories. */
  pageField?: { page: number; pages: number }
}

export const PT_TO_PX = 4 / 3

const HIGHLIGHT: Record<HighlightName, string> = {
  yellow: '#ffff00',
  green: '#00ff00',
  cyan: '#00ffff',
  magenta: '#ff00ff',
  blue: '#0000ff',
  red: '#ff0000',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#c0c0c0',
  black: '#000000',
}

function isEastAsian(cp: number): boolean {
  return (
    (cp >= 0x1100 && cp <= 0x11ff) ||
    (cp >= 0x2e80 && cp <= 0x9fff) ||
    (cp >= 0xa960 && cp <= 0xa97f) ||
    (cp >= 0xac00 && cp <= 0xd7ff) ||
    (cp >= 0xf900 && cp <= 0xfaff) ||
    (cp >= 0xfe10 && cp <= 0xfe1f) ||
    (cp >= 0xfe30 && cp <= 0xfe4f) ||
    (cp >= 0xff00 && cp <= 0xffef) ||
    (cp >= 0x1b000 && cp <= 0x1b16f) ||
    (cp >= 0x1f200 && cp <= 0x1f2ff) ||
    (cp >= 0x20000 && cp <= 0x3fffd)
  )
}

function isComplex(cp: number): boolean {
  return (
    (cp >= 0x0590 && cp <= 0x05ff) ||
    (cp >= 0x0600 && cp <= 0x06ff) ||
    (cp >= 0x0900 && cp <= 0x097f) ||
    (cp >= 0x0e00 && cp <= 0x0e7f)
  )
}

export type ScriptSlot = 'eastAsia' | 'ascii' | 'hAnsi' | 'cs'

function slotFor(cp: number): ScriptSlot {
  if (isEastAsian(cp)) return 'eastAsia'
  if (isComplex(cp)) return 'cs'
  if (cp <= 0x7f) return 'ascii'
  return 'hAnsi'
}

function familyFor(cp: number, look: RunLook): string {
  const slot = slotFor(cp)
  if (slot === 'eastAsia') return look.fonts.eastAsia
  if (slot === 'cs') return look.fonts.cs
  if (slot === 'ascii') return look.fonts.ascii
  return look.fonts.hAnsi
}

export interface ScriptSegment {
  readonly text: string
  readonly family: string
  readonly slot: ScriptSlot
}

/**
 * Split text into segments that each use one font slot, the way Word and rdocx classify characters: CJK ideographs,
 * kana, hangul and full-width forms use the East Asian slot; the rest use ASCII, high ANSI or complex script slots.
 */
export function scriptSegments(text: string, look: RunLook): readonly ScriptSegment[] {
  const out: ScriptSegment[] = []
  let buf = ''
  let family = ''
  let slot: ScriptSlot = 'ascii'
  for (const ch of text) {
    const cp = ch.codePointAt(0) ?? 0
    const nextSlot = slotFor(cp)
    const next = familyFor(cp, look)
    if (buf.length > 0 && (next !== family || nextSlot !== slot)) {
      out.push({ text: buf, family, slot })
      buf = ''
    }
    family = next
    slot = nextSlot
    buf += ch
  }
  if (buf.length > 0) out.push({ text: buf, family, slot })
  return out
}

/**
 * The family that paints one segment. Han in a face with no CJK glyphs takes 宋体; Latin in the same run keeps its own face.
 * `cjk` undefined means the metrics have not loaded, so the segment stays on the run's family.
 */
export function paintFamily(family: string, slot: ScriptSlot, cjk: boolean | undefined): string {
  if (slot === 'eastAsia' && cjk === false) return '宋体'
  return family
}

/** Grid snap from `restore_word_line_heights`: exact spacing is the author's height, and other rules take whole rows. */
export function snapLinePx(height: number, pitch: number | null, rule: LineSpacing): number {
  if (rule.rule === 'exact' || pitch === null || !(pitch > 0)) return height
  return Math.max(1, Math.ceil(height / pitch - 1e-9)) * pitch
}

/**
 * Space before a cell paragraph adds below the previous paragraph. The first paragraph of a cell, and one
 * after a nested table, keep the full space before. Between two paragraphs the gap is the larger of the two.
 */
export function collapsedSpaceBefore(beforePx: number, previousAfterPx: number | null): number {
  if (previousAfterPx === null) return Math.max(0, beforePx)
  return Math.max(0, beforePx - previousAfterPx)
}

function px(pt: number): number {
  return pt * PT_TO_PX
}

function readO(el: Element): number {
  const n = Number(el.getAttribute('data-o'))
  return Number.isFinite(n) ? n : 0
}

function blockHasPara(block: Block, id: ParaId): boolean {
  if (block.kind === 'p') return block.id === id
  if (block.kind === 'opaque') return block.preview.some((para) => para.id === id)
  for (const row of block.rows) {
    for (const cell of row.cells) {
      for (const child of cell.blocks) {
        if (blockHasPara(child, id)) return true
      }
    }
  }
  return false
}

function decoFor(block: Block, deco: Decoration | undefined): Decoration | undefined {
  if (!deco) return undefined
  return blockHasPara(block, deco.pos.p) ? deco : undefined
}

function lineHeight(ctx: RenderContext, para: Paragraph, family: string, sizePt: Pt): number {
  const rule = ctx.paraLook(para).spacing.line
  return snapLinePx(ctx.lineHeightPx(family, sizePt, rule), ctx.gridPitchPx, rule)
}

function paintedFamily(ctx: RenderContext, family: string, slot: ScriptSlot): string {
  return paintFamily(family, slot, ctx.cjkGlyphs?.(family))
}

function applyRun(
  span: HTMLElement,
  ctx: RenderContext,
  para: Paragraph,
  look: RunLook,
  family: string,
  marks: Marks,
): void {
  const sizePx = px(look.sizePt)
  span.style.fontFamily = ctx.fontStack(family)
  span.style.fontSize = `${marks.vertAlign ? sizePx * 0.7 : sizePx}px`
  span.style.lineHeight = `${lineHeight(ctx, para, family, look.sizePt)}px`
  span.style.fontWeight = look.bold ? '700' : '400'
  span.style.fontStyle = look.italic ? 'italic' : 'normal'
  const deco = [look.underline || marks.link ? 'underline' : '', look.strike ? 'line-through' : ''].filter((part) => part.length > 0)
  if (deco.length > 0) span.style.textDecoration = deco.join(' ')
  if (marks.link) span.style.color = '#0563c1'
  else if (look.color !== 'auto') span.style.color = `#${look.color}`
  if (look.highlight) span.style.backgroundColor = HIGHLIGHT[look.highlight]
  if (look.caps) span.style.textTransform = 'uppercase'
  if (look.charSpacingPt) span.style.letterSpacing = `${px(look.charSpacingPt)}px`
  if (marks.vertAlign === 'super' || marks.vertAlign === 'sub') span.style.verticalAlign = marks.vertAlign
}

function firstLinePx(ctx: RenderContext, para: Paragraph, look: ParaLook): number {
  if (look.indent.firstLineChars) {
    const run = ctx.look(para, {})
    return (look.indent.firstLineChars / 100) * px(run.sizePt)
  }
  return px(look.indent.firstLine)
}

/**
 * Marker box and first-line content start in px. Clamping the marker to >= 0 keeps it inside `.page-body`'s clip;
 * setting `paddingLeft` to `markerLeft + markerWidth` guarantees text and the caret never overlap the marker, even when
 * a paragraph carries an explicit `left: 0` or `hanging: 0` indent.
 */
export function markerBoxPx(
  leftPx: number,
  hangingPx: number,
): { readonly markerLeft: number; readonly markerWidth: number; readonly paddingLeft: number } {
  const width = hangingPx > 0 ? hangingPx : px(18)
  const markerLeft = hangingPx > 0 ? Math.max(0, leftPx - hangingPx) : Math.max(0, leftPx)
  return {
    markerLeft,
    markerWidth: width,
    paddingLeft: markerLeft + width,
  }
}

function applyParaBox(el: HTMLElement, ctx: RenderContext, para: Paragraph, beforePx?: number): ParaLook {
  const look = ctx.paraLook(para)
  const before = beforePx ?? px(look.spacing.before)
  const after = px(look.spacing.after)
  const left = px(look.indent.left)
  const hanging = px(look.indent.hanging)
  const right = px(look.indent.right)
  const marker = ctx.marker(para.id)
  el.style.paddingTop = `${before}px`
  el.style.paddingBottom = `${after}px`
  el.style.paddingRight = `${right}px`
  if (marker) {
    const box = markerBoxPx(left, hanging)
    el.style.paddingLeft = `${box.paddingLeft}px`
    el.style.textIndent = '0'
  } else {
    el.style.paddingLeft = `${left + hanging}px`
    el.style.textIndent = `${firstLinePx(ctx, para, look) - hanging}px`
  }
  if (look.align === 'distribute') {
    el.style.textAlign = 'justify'
    el.style.textAlignLast = 'justify'
  } else {
    el.style.textAlign = look.align
  }
  el.dataset.keepNext = look.keepNext ? '1' : '0'
  el.dataset.keepLines = look.keepLines ? '1' : '0'
  el.dataset.pageBreakBefore = look.pageBreakBefore ? '1' : '0'
  el.dataset.widow = look.widowControl ? '1' : '0'
  el.dataset.spaceBefore = String(before)
  el.dataset.spaceAfter = String(after)
  return look
}

function fieldLabel(atom: Extract<AtomKind, { type: 'field' }>, ctx: RenderContext): string {
  if (atom.field === 'page' && ctx.pageField) return String(ctx.pageField.page)
  if (atom.field === 'numPages' && ctx.pageField) return String(ctx.pageField.pages)
  return atom.display
}

function renderAtom(atom: Atom, offset: number, ctx: RenderContext, para: Paragraph): HTMLElement {
  const kind = atom.atom
  if (kind.type === 'image') {
    const img = document.createElement('img')
    img.className = 'atom'
    img.dataset.o = String(offset)
    img.dataset.atom = 'image'
    img.alt = ''
    img.draggable = false
    img.src = ctx.imageUrl(kind.media)
    img.style.width = `${px(kind.widthPt)}px`
    img.style.height = `${px(kind.heightPt)}px`
    return img
  }
  const span = document.createElement('span')
  span.className = 'atom'
  span.dataset.o = String(offset)
  span.dataset.atom = kind.type
  span.contentEditable = 'false'
  if (kind.type === 'field') {
    span.classList.add('field')
    span.textContent = fieldLabel(kind, ctx)
  } else if (kind.type === 'opaque') {
    span.classList.add('opaque-atom')
    span.textContent = kind.display.length > 0 ? kind.display : kind.label
  } else {
    span.classList.add('break')
    span.textContent = '\u200b'
  }
  if (kind.type !== 'pageBreak' && kind.type !== 'columnBreak') {
    const look = ctx.look(para, {})
    applyRun(span, ctx, para, look, look.fonts.eastAsia, {})
  }
  return span
}

interface TextPiece {
  readonly kind: 'text'
  readonly text: string
  readonly o: number
  readonly look: RunLook
  readonly family: string
  readonly slot: ScriptSlot
  readonly marks: Marks
}

interface AtomPiece {
  readonly kind: 'atom'
  readonly o: number
  readonly atom: Atom
}

type Piece = TextPiece | AtomPiece

function piecesOf(para: Paragraph, ctx: RenderContext): Piece[] {
  const pieces: Piece[] = []
  let o = 0
  for (const inline of para.inlines) {
    if (inline.kind === 'text') {
      const look = ctx.look(para, inline.marks)
      for (const seg of scriptSegments(inline.text, look)) {
        pieces.push({ kind: 'text', text: seg.text, o, look, family: seg.family, slot: seg.slot, marks: inline.marks })
        o += seg.text.length
      }
    } else {
      pieces.push({ kind: 'atom', o, atom: inline })
      o += 1
    }
  }
  return pieces
}

function appendText(parent: HTMLElement, ctx: RenderContext, para: Paragraph, piece: TextPiece): void {
  const span = document.createElement('span')
  span.dataset.o = String(piece.o)
  span.textContent = piece.text
  applyRun(span, ctx, para, piece.look, paintedFamily(ctx, piece.family, piece.slot), piece.marks)
  parent.append(span)
}

export interface PreeditRunStyle {
  readonly slot: ScriptSlot
  readonly fontFamily: string
  readonly fontSizePx: number
  readonly fontWeight: '400' | '700'
  readonly fontStyle: 'normal' | 'italic'
  readonly letterSpacingPx: number
  readonly color: string
  readonly vertAlign: 'super' | 'sub' | undefined
}

export function preeditRunStyle(ctx: RenderContext, para: Paragraph, marks: Marks, text: string): PreeditRunStyle {
  const look = ctx.look(para, marks)
  const slot: ScriptSlot = isMostlyEastAsian(text) ? 'eastAsia' : 'ascii'
  const rawFamily = slot === 'eastAsia' ? look.fonts.eastAsia : look.fonts.ascii
  const family = paintedFamily(ctx, rawFamily, slot)
  const sizePx = px(look.sizePt)
  return {
    slot,
    fontFamily: ctx.fontStack(family),
    fontSizePx: marks.vertAlign ? sizePx * 0.7 : sizePx,
    fontWeight: look.bold ? '700' : '400',
    fontStyle: look.italic ? 'italic' : 'normal',
    letterSpacingPx: look.charSpacingPt ? px(look.charSpacingPt) : 0,
    color: marks.link ? '#0563c1' : look.color !== 'auto' ? `#${look.color}` : '#222',
    vertAlign: marks.vertAlign === 'super' || marks.vertAlign === 'sub' ? marks.vertAlign : undefined,
  }
}

function appendPreedit(parent: HTMLElement, ctx: RenderContext, para: Paragraph, deco: Decoration): void {
  const span = document.createElement('span')
  span.className = 'preedit'
  span.dataset.o = String(deco.pos.o)
  span.textContent = deco.text
  const look = ctx.look(para, {})
  const slot: ScriptSlot = isMostlyEastAsian(deco.text) ? 'eastAsia' : 'ascii'
  const family = slot === 'eastAsia' ? look.fonts.eastAsia : look.fonts.ascii
  applyRun(span, ctx, para, look, paintedFamily(ctx, family, slot), {})
  parent.append(span)
}

function isMostlyEastAsian(text: string): boolean {
  for (const ch of text) {
    if (isEastAsian(ch.codePointAt(0) ?? 0)) return true
  }
  return false
}

function renderParagraph(para: Paragraph, ctx: RenderContext, deco: Decoration | undefined, beforePx?: number, splitBreaks = true): HTMLElement {
  const el = document.createElement('div')
  el.className = 'blk p'
  el.dataset.id = para.id
  const look = applyParaBox(el, ctx, para, beforePx)
  const before = beforePx ?? px(look.spacing.before)
  const marker = ctx.marker(para.id)
  if (marker) {
    const span = document.createElement('span')
    span.className = 'marker'
    span.contentEditable = 'false'
    span.textContent = marker
    const box = markerBoxPx(px(look.indent.left), px(look.indent.hanging))
    span.style.left = `${box.markerLeft}px`
    span.style.top = `${before}px`
    span.style.width = `${box.markerWidth}px`
    const run = ctx.look(para, {})
    applyRun(span, ctx, para, run, run.fonts.eastAsia, {})
    el.append(span)
  }
  const active = deco && deco.pos.p === para.id ? deco : undefined
  const pieces = piecesOf(para, ctx)
  if (pieces.length === 0) {
    const span = document.createElement('span')
    span.dataset.o = '0'
    span.textContent = '\u200b'
    const run = ctx.look(para, {})
    applyRun(span, ctx, para, run, run.fonts.eastAsia, {})
    if (active && active.pos.o === 0) appendPreedit(el, ctx, para, active)
    el.append(span)
    return finishParagraph(el, para)
  }
  let preeditDone = false
  const emitPreedit = (at: number): void => {
    if (!active || preeditDone || active.pos.o !== at) return
    appendPreedit(el, ctx, para, active)
    preeditDone = true
  }
  for (const piece of pieces) {
    if (piece.kind === 'atom') {
      emitPreedit(piece.o)
      const atom = renderAtom(piece.atom, piece.o, ctx, para)
      if (splitBreaks && piece.atom.atom.type === 'pageBreak') stylePageBreak(atom, el)
      el.append(atom)
      continue
    }
    const end = piece.o + piece.text.length
    if (active && !preeditDone && active.pos.o > piece.o && active.pos.o < end) {
      const cut = active.pos.o - piece.o
      appendText(el, ctx, para, { ...piece, text: piece.text.slice(0, cut) })
      appendPreedit(el, ctx, para, active)
      preeditDone = true
      appendText(el, ctx, para, { ...piece, text: piece.text.slice(cut), o: active.pos.o })
      continue
    }
    emitPreedit(piece.o)
    appendText(el, ctx, para, piece)
  }
  if (active && !preeditDone) appendPreedit(el, ctx, para, active)
  return finishParagraph(el, para)
}

function lineHasInk(parent: HTMLElement): boolean {
  const children = parent.childNodes
  for (let i = children.length - 1; i >= 0; i--) {
    const node = children[i]
    if (node === undefined) continue
    if (node.nodeType === Node.TEXT_NODE) {
      const text = (node.textContent ?? '').replaceAll('\u200b', '')
      if (text.length > 0) return true
      continue
    }
    if (!(node instanceof HTMLElement)) continue
    if (node.classList.contains('marker')) continue
    if (node.dataset.atom === 'pageBreak') return false
    return true
  }
  return false
}

function stylePageBreak(el: HTMLElement, parent: HTMLElement): void {
  el.style.display = 'block'
  el.style.width = '100%'
  el.style.overflow = 'hidden'
  // A break after text closes that line. A break with nothing on its line takes one:
  // the line breaker flushes a line before the forced page break, and an empty line has height.
  el.style.height = lineHasInk(parent) ? '0' : '1em'
}

function finishParagraph(el: HTMLElement, para: Paragraph): HTMLElement {
  // One accessible name for the whole paragraph. Hidden children stop a span per run from being read as its own line.
  el.setAttribute('aria-label', flatText(para))
  for (const child of el.children) child.setAttribute('aria-hidden', 'true')
  return el
}

function rowspanAt(table: Table, rowIndex: number, col: number): number {
  let span = 1
  for (let r = rowIndex + 1; r < table.rows.length; r++) {
    const row = table.rows[r]
    if (!row) break
    let cursor = 0
    let found = false
    for (const cell of row.cells) {
      if (cursor === col) {
        if (cell.vMerge !== 'continue') return span
        found = true
        span += 1
        break
      }
      cursor += cell.span
    }
    if (!found) break
  }
  return span
}

export interface CellEdges {
  readonly top: BorderEdge
  readonly left: BorderEdge
  readonly bottom: BorderEdge
  readonly right: BorderEdge
}

const NO_EDGES: CellEdges = { top: null, left: null, bottom: null, right: null }

/**
 * The four edges a cell shows. Outer sides use the table's outer edges, inner
 * sides use insideH and insideV, and an edge the cell names replaces the table.
 */
export function cellEdges(table: Table, rowIndex: number, cellIndex: number): CellEdges {
  const row = table.rows[rowIndex]
  const cell = row?.cells[cellIndex]
  if (!row || !cell) return NO_EDGES
  const col = columnIndex(table, rowIndex, cellIndex)
  const columns = columnCount(table, rowIndex)
  const covered = cell.vMerge === 'restart' ? rowspanAt(table, rowIndex, col) : 1
  const lastRow = rowIndex + covered - 1
  return {
    top: resolveEdge(cell.borders, 'top', table.props.borders, rowIndex === 0 ? 'top' : 'insideH'),
    left: resolveEdge(cell.borders, 'left', table.props.borders, col === 0 ? 'left' : 'insideV'),
    bottom: resolveEdge(cell.borders, 'bottom', table.props.borders, lastRow >= table.rows.length - 1 ? 'bottom' : 'insideH'),
    right: resolveEdge(cell.borders, 'right', table.props.borders, col + cell.span >= columns ? 'right' : 'insideV'),
  }
}

/**
 * The edges this cell paints. A line between two cells is painted once: the
 * left cell owns a vertical line and the upper cell owns a horizontal line.
 * rdocx paints each cell's resolved edge on its own, so a side with no line
 * drops out. When both sides are lines, the later stroke covers the earlier
 * one, and that later stroke is the one kept: the right neighbour's left, or
 * the lower neighbour's top.
 */
export function paintedCellEdges(table: Table, rowIndex: number, cellIndex: number): CellEdges {
  const edges = cellEdges(table, rowIndex, cellIndex)
  const row = table.rows[rowIndex]
  const cell = row?.cells[cellIndex]
  if (!row || !cell) return edges
  const col = columnIndex(table, rowIndex, cellIndex)
  const covered = cell.vMerge === 'restart' ? rowspanAt(table, rowIndex, col) : 1
  const below = facingEdge(table, rowIndex + covered, 1, col, cell.span, 'top')
  const beside = facingEdge(table, rowIndex, covered, col + cell.span, 1, 'left')
  return {
    top: rowIndex === 0 ? edges.top : null,
    left: col === 0 ? edges.left : null,
    bottom: sharedEdge(edges.bottom, below),
    right: sharedEdge(edges.right, beside),
  }
}

/** Inset shadow for the edges this cell paints. `none` when it paints no line. */
export function cellBoxShadow(edges: CellEdges): string {
  const parts: string[] = []
  // box-shadow is one solid band, so dashed, dotted and double use that band.
  addBand(parts, edges.top, 0, 1)
  addBand(parts, edges.left, 1, 0)
  addBand(parts, edges.bottom, 0, -1)
  addBand(parts, edges.right, -1, 0)
  return parts.length === 0 ? 'none' : parts.join(', ')
}

function resolveEdge(
  own: CellBorders | undefined,
  side: keyof CellBorders,
  tableBorders: TableBorders | undefined,
  tableSide: keyof TableBorders,
): BorderEdge {
  if (own !== undefined && Object.hasOwn(own, side)) {
    const value = own[side]
    return value === undefined ? null : value
  }
  if (tableBorders === undefined) return null
  return tableBorders[tableSide]
}

function columnIndex(table: Table, rowIndex: number, cellIndex: number): number {
  const cells = table.rows[rowIndex]?.cells ?? []
  let col = 0
  for (let i = 0; i < cellIndex; i++) col += cells[i]?.span ?? 0
  return col
}

function columnCount(table: Table, rowIndex: number): number {
  const spanned = table.rows[rowIndex]?.cells.reduce((sum, cell) => sum + cell.span, 0) ?? 0
  return Math.max(table.grid.length, spanned)
}

/** A line with a positive width, or null when this side draws nothing. */
function drawnEdge(edge: BorderEdge): BorderEdge {
  if (edge !== null && edge.widthPt > 0) return edge
  return null
}

/**
 * One stroke for a shared boundary. The neighbour is painted later, so its
 * line covers the owner's where both draw. A side that draws nothing leaves
 * the other side's line in place.
 */
function sharedEdge(owner: BorderEdge, facing: BorderEdge): BorderEdge {
  return drawnEdge(facing) ?? drawnEdge(owner)
}

/**
 * The edge of the first cell that faces this span, in row-major order.
 * A vertical merge paints from its restart cell. When several cells meet one
 * edge of a span and they differ, the first is the one that counts.
 */
function facingEdge(
  table: Table,
  rowIndex: number,
  rowCount: number,
  col: number,
  colCount: number,
  side: 'top' | 'left',
): BorderEdge {
  const rowEnd = Math.min(table.rows.length, rowIndex + rowCount)
  const colEnd = col + colCount
  for (let row = rowIndex; row < rowEnd; row++) {
    for (let cursor = col; cursor < colEnd; cursor++) {
      const at = cellCovering(table, row, cursor)
      if (!at) continue
      return cellEdges(table, at.row, at.cell)[side]
    }
  }
  return null
}

/** The cell that paints this grid position. A vertical merge paints from its restart. */
function cellCovering(table: Table, rowIndex: number, col: number): { readonly row: number; readonly cell: number } | null {
  let row = rowIndex
  let cell = cellIndexAt(table, row, col)
  if (cell === null) return null
  while (row > 0 && table.rows[row]?.cells?.[cell]?.vMerge === 'continue') {
    const above = cellIndexAt(table, row - 1, col)
    if (above === null) break
    row -= 1
    cell = above
  }
  return { row, cell }
}

function cellIndexAt(table: Table, rowIndex: number, col: number): number | null {
  const cells = table.rows[rowIndex]?.cells
  if (!cells) return null
  let cursor = 0
  for (let index = 0; index < cells.length; index++) {
    const span = cells[index]?.span ?? 0
    if (col >= cursor && col < cursor + span) return index
    cursor += span
  }
  return null
}

function addBand(parts: string[], edge: BorderEdge, xSign: number, ySign: number): void {
  if (edge === null || !(edge.widthPt > 0)) return
  const color = /^[0-9A-Fa-f]{6}$/.test(edge.color) ? `#${edge.color}` : '#000000'
  const width = edge.widthPt * PT_TO_PX
  parts.push(`inset ${cssPx(xSign * width)} ${cssPx(ySign * width)} 0 ${color}`)
}

function cssPx(value: number): string {
  return `${Math.round(value * 1000) / 1000}px`
}

function appendCellBlocks(td: HTMLElement, blocks: readonly Block[], ctx: RenderContext, deco: Decoration | undefined): void {
  let previousAfter: number | null = null
  for (const child of blocks) {
    if (child.kind !== 'p') {
      td.append(renderBlock(child, ctx, decoFor(child, deco)))
      previousAfter = null
      continue
    }
    const before = collapsedSpaceBefore(px(ctx.paraLook(child).spacing.before), previousAfter)
    td.append(renderParagraph(child, ctx, decoFor(child, deco), before, false))
    previousAfter = px(ctx.paraLook(child).spacing.after)
  }
}

const WORD_CELL_PAD_PT = { top: 0, right: 5.4, bottom: 0, left: 5.4 }

export interface CellPad {
  readonly top: number
  readonly right: number
  readonly bottom: number
  readonly left: number
}

function padEdge(cell: number | undefined, table: number | undefined, fallback: number): number {
  if (cell !== undefined) return cell
  if (table !== undefined) return table
  return fallback
}

/** Cell, then table, then Word's 0, 5.4 pt, 0, 5.4 pt. An explicit 0 wins. */
export function cellPadPt(table: Table, cell: Cell): CellPad {
  const tableMargins = table.props.margins
  const cellMargins = cell.margins
  return {
    top: padEdge(cellMargins?.top, tableMargins?.top, WORD_CELL_PAD_PT.top),
    right: padEdge(cellMargins?.right, tableMargins?.right, WORD_CELL_PAD_PT.right),
    bottom: padEdge(cellMargins?.bottom, tableMargins?.bottom, WORD_CELL_PAD_PT.bottom),
    left: padEdge(cellMargins?.left, tableMargins?.left, WORD_CELL_PAD_PT.left),
  }
}

export interface RowFrame {
  readonly height: string | null
  readonly minHeight: string | null
  readonly clip: boolean
}

/**
 * Exact clips. atLeast, and a height with no rule, is a minimum.
 * A table row honors `height` as that minimum and ignores `min-height`, so the
 * minimum is written to both: `height` for the layout, `minHeight` for the paginator.
 */
export function rowFrame(row: Row): RowFrame {
  const height = row.props.heightPt
  if (height === undefined) return { height: null, minHeight: null, clip: false }
  const pxHeight = `${px(height)}px`
  if (row.props.heightRule === 'exact') return { height: pxHeight, minHeight: null, clip: true }
  return { height: pxHeight, minHeight: pxHeight, clip: false }
}

function renderTable(table: Table, ctx: RenderContext, deco: Decoration | undefined): HTMLElement {
  const el = document.createElement('div')
  el.className = 'blk t'
  el.dataset.id = table.id
  el.dataset.keepNext = '0'
  el.dataset.keepLines = '0'
  el.dataset.pageBreakBefore = '0'
  el.dataset.widow = '0'
  el.dataset.spaceBefore = '0'
  el.dataset.spaceAfter = '0'
  const tableEl = document.createElement('table')
  const cols = document.createElement('colgroup')
  let sum = 0
  for (const width of table.grid) {
    const col = document.createElement('col')
    const widthPx = px(width)
    sum += widthPx
    col.style.width = `${widthPx}px`
    cols.append(col)
  }
  tableEl.style.width = `${sum}px`
  tableEl.append(cols)
  const body = document.createElement('tbody')
  let headerRows = 0
  let headerOpen = true
  for (let r = 0; r < table.rows.length; r++) {
    const row = table.rows[r]
    if (!row) continue
    if (headerOpen && row.props.header) headerRows += 1
    else headerOpen = false
    const tr = document.createElement('tr')
    if (row.props.header) tr.dataset.header = '1'
    tr.dataset.cantSplit = row.props.cantSplit === true ? '1' : '0'
    const frame = rowFrame(row)
    if (frame.clip) tr.dataset.exact = '1'
    if (frame.height !== null) tr.style.height = frame.height
    if (frame.minHeight !== null) tr.style.minHeight = frame.minHeight
    let col = 0
    for (let c = 0; c < row.cells.length; c++) {
      const cell = row.cells[c]
      if (!cell) continue
      if (cell.vMerge === 'continue') {
        col += cell.span
        continue
      }
      const td = document.createElement('td')
      if (cell.span > 1) td.colSpan = cell.span
      const rows = cell.vMerge === 'restart' ? rowspanAt(table, r, col) : 1
      if (rows > 1) td.rowSpan = rows
      const shadow = cellBoxShadow(paintedCellEdges(table, r, c))
      if (shadow !== 'none') td.style.boxShadow = shadow
      const pad = cellPadPt(table, cell)
      td.style.padding = `${px(pad.top)}px ${px(pad.right)}px ${px(pad.bottom)}px ${px(pad.left)}px`
      if (frame.clip && frame.height !== null) {
        td.style.height = frame.height
        td.style.maxHeight = frame.height
        td.style.overflow = 'hidden'
      }
      if (cell.shading) td.style.backgroundColor = `#${cell.shading}`
      if (cell.vAlign === 'center') td.style.verticalAlign = 'middle'
      else if (cell.vAlign) td.style.verticalAlign = cell.vAlign
      appendCellBlocks(td, cell.blocks, ctx, deco)
      tr.append(td)
      col += cell.span
    }
    body.append(tr)
  }
  tableEl.append(body)
  el.dataset.headerRows = String(headerRows)
  el.append(tableEl)
  return el
}

function renderOpaque(block: Extract<Block, { kind: 'opaque' }>, ctx: RenderContext, deco: Decoration | undefined): HTMLElement {
  const el = document.createElement('div')
  el.className = 'blk o'
  el.dataset.id = block.id
  el.dataset.keepNext = '0'
  el.dataset.keepLines = '0'
  el.dataset.pageBreakBefore = '0'
  el.dataset.widow = '0'
  el.dataset.spaceBefore = '0'
  el.dataset.spaceAfter = '0'
  const label = document.createElement('div')
  label.className = 'opaque-label'
  label.textContent = block.label
  el.append(label)
  for (const para of block.preview) el.append(renderBlock(para, ctx, decoFor(para, deco)))
  return el
}

/** Render one block. Pure in its inputs; pages.ts caches the result per block object and clones it to mount. */
export function renderBlock(block: Block, ctx: RenderContext, deco?: Decoration): HTMLElement {
  if (block.kind === 'p') return renderParagraph(block, ctx, deco)
  if (block.kind === 'table') return renderTable(block, ctx, deco)
  return renderOpaque(block, ctx, deco)
}

function slots(blockEl: HTMLElement): HTMLElement[] {
  const out: HTMLElement[] = []
  for (const child of blockEl.children) {
    if (!(child instanceof HTMLElement)) continue
    if (!child.hasAttribute('data-o')) continue
    if (child.classList.contains('preedit') || child.classList.contains('marker')) continue
    out.push(child)
  }
  return out
}

function pointBefore(el: Element): { node: Node; offset: number } {
  const parent = el.parentNode
  if (!parent) return { node: el, offset: 0 }
  return { node: parent, offset: childIndex(parent, el) }
}

function pointAfter(el: Element): { node: Node; offset: number } {
  const parent = el.parentNode
  if (!parent) return { node: el, offset: 0 }
  return { node: parent, offset: childIndex(parent, el) + 1 }
}

function childIndex(parent: Node, el: Element): number {
  for (let i = 0; i < parent.childNodes.length; i++) {
    if (parent.childNodes[i] === el) return i
  }
  return 0
}

/** The DOM point inside a rendered block element (one mounted copy) for a flat-text offset. */
export function domPoint(blockEl: HTMLElement, o: number): { node: Node; offset: number } {
  const spanList = slots(blockEl)
  if (spanList.length === 0) return { node: blockEl, offset: 0 }
  for (let i = 0; i < spanList.length; i++) {
    const span = spanList[i]
    if (!span) continue
    const start = readO(span)
    if (span.classList.contains('atom')) {
      if (o <= start) return pointBefore(span)
      if (o === start + 1) {
        const next = spanList[i + 1]
        if (next && readO(next) === o) continue
        return pointAfter(span)
      }
      continue
    }
    const text = span.firstChild
    if (!(text instanceof Text)) continue
    if (text.data === '\u200b') {
      if (o <= start) return { node: text, offset: 0 }
      continue
    }
    if (o < start + text.length) return { node: text, offset: Math.max(0, o - start) }
    if (o === start + text.length) {
      const next = spanList[i + 1]
      if (next && readO(next) <= o) continue
      return { node: text, offset: text.length }
    }
  }
  const last = spanList[spanList.length - 1]
  if (!last) return { node: blockEl, offset: 0 }
  if (last.classList.contains('atom')) return pointAfter(last)
  const text = last.firstChild
  if (text instanceof Text && text.data !== '\u200b') return { node: text, offset: text.length }
  if (text instanceof Text) return { node: text, offset: 0 }
  return pointAfter(last)
}

let shared: Range | null = null

/**
 * The one Range every measurement reuses. acdert updates each live Range on every DOM mutation and collects them
 * late, so a Range per measurement makes each keystroke slower with every document opened. Read it before the next
 * call takes it.
 */
export function scratchRange(): Range {
  shared ??= document.createRange()
  return shared
}

/**
 * The shared Range set over [from, to) in one rendered paragraph element; collapsed when from === to. Read it before
 * the next call takes it.
 */
export function rangeIn(blockEl: HTMLElement, from: number, to: number): Range {
  const start = domPoint(blockEl, Math.min(from, to))
  const end = domPoint(blockEl, Math.max(from, to))
  const range = scratchRange()
  range.setStart(start.node, start.offset)
  range.setEnd(end.node, end.offset)
  return range
}

/**
 * Map a DOM point (from caretPositionFromPoint) back to a model position. Null when the point is outside any editable
 * paragraph (a list marker resolves to offset 0, an atom to the nearer side, a U+200B placeholder to 0).
 */
export function posFromDom(node: Node, offset: number): Pos | null {
  const start = node instanceof Element ? node : node.parentElement
  if (!start) return null
  const para = start.closest('.blk.p')
  if (!(para instanceof HTMLElement)) return null
  const id = para.dataset.id
  if (!id) return null
  const p = id as ParaId
  if (start.closest('.marker')) return { p, o: 0 }
  const preedit = start.closest('.preedit')
  if (preedit instanceof HTMLElement && para.contains(preedit)) return { p, o: readO(preedit) }
  const atom = start.closest('.atom')
  if (atom instanceof HTMLElement && para.contains(atom)) return atomSide(atom, node, offset, p)

  if (node instanceof Text) {
    const host = node.parentElement
    if (!(host instanceof HTMLElement) || !para.contains(host)) return { p, o: 0 }
    if (host.classList.contains('marker')) return { p, o: 0 }
    if (node.data === '\u200b') return { p, o: readO(host) }
    const base = readO(host)
    const within = Math.max(0, Math.min(offset, node.data.length))
    return { p, o: base + within }
  }

  if (node instanceof Element && para.contains(node)) {
    const child = node.childNodes[offset]
    if (child instanceof Element && child.classList.contains('atom')) return { p, o: readO(child) }
    if (child instanceof Text) return posFromDom(child, 0)
    if (offset > 0) {
      const prev = node.childNodes[offset - 1]
      if (prev instanceof Element && prev.classList.contains('atom')) return { p, o: readO(prev) + 1 }
      if (prev instanceof Text) return posFromDom(prev, prev.data.length)
    }
    return { p, o: 0 }
  }
  return { p, o: 0 }
}

function atomSide(atom: HTMLElement, node: Node, offset: number, p: ParaId): Pos {
  const base = readO(atom)
  const parent = atom.parentNode
  if (node === parent && parent) {
    const index = childIndex(parent, atom)
    return { p, o: offset > index ? base + 1 : base }
  }
  return { p, o: base }
}

/** The element of block `id` inside a mounted container. Pages may hold two copies; pages.ts picks by geometry. */
export function blockElements(root: ParentNode, id: string): readonly HTMLElement[] {
  return [...root.querySelectorAll<HTMLElement>(`.blk[data-id="${CSS.escape(id)}"]`)]
}

/** CSS for a section's page box and content box, in px at 100 percent zoom (1pt = 4/3 px). */
export function pageBoxStyle(section: Section): { page: Partial<CSSStyleDeclaration>; content: Partial<CSSStyleDeclaration> } {
  const page = section.page
  const margin = page.margin
  const width = px(page.widthPt)
  const height = px(page.heightPt)
  const left = px(margin.left + margin.gutter)
  const right = px(margin.right)
  const top = px(margin.top)
  const bottom = px(margin.bottom)
  return {
    page: {
      boxSizing: 'border-box',
      position: 'relative',
      width: `${width}px`,
      height: `${height}px`,
    },
    content: {
      boxSizing: 'border-box',
      position: 'absolute',
      top: `${top}px`,
      right: `${right}px`,
      bottom: `${bottom}px`,
      left: `${left}px`,
    },
  }
}
