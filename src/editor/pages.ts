import { t } from '../i18n.ts'
import { flatText, marksAt } from '../model.ts'
import type {
  Block,
  BlockId,
  Doc,
  LineSpacing,
  Marks,
  Paragraph,
  ParaId,
  Pos,
  Section,
  Selection,
  StoryKey,
  Table,
} from '../model.ts'
import { documentGridPitch } from './derive.ts'
import { imageChrome, selectedImage } from './edit.ts'
import type { TextMatch } from './find.ts'
import { caretAtPreeditEnd, hostPreeditBox, isHostPreeditStyle } from './input.ts'
import {
  HANDLE_HIT_SCREEN_PX,
  HANDLE_SCREEN_PX,
  RESIZE_HANDLES,
  handleOrigin,
} from './image-geom.ts'
import type { ImageGesture, PageBox } from './image-geom.ts'
import {
  blockElements,
  domPoint,
  pageBoxStyle,
  posFromDom,
  preeditRunStyle,
  PT_TO_PX,
  rangeIn,
  renderBlock,
  scratchRange,
} from './render.ts'
import type { Decoration, PreeditRunStyle, RenderContext } from './render.ts'

const EPS = 0.51
const STACK_PAD = 16
const PAGE_GAP = 16

export interface Measure {
  readonly widthPx: number
  readonly heightPx: number
  /** Bottom of each line, relative to the block's top, padding included. Paragraphs only. */
  readonly lineBottoms: Float32Array
  /** Bottom of each row. Tables only. */
  readonly rowBottoms?: Float32Array
  /** Y of each allowed cut inside a row, relative to the block top, ascending. Tables only. Row bottoms are not in it. */
  readonly rowCuts?: Float32Array
  /** 1 where the kit would split a row that still fits on a page. 0 is used only when the row is taller than the page. Absent means every cut is legal. */
  readonly rowCutSplit?: Uint8Array
  /** Minimum height of the first fragment of each row. Absent when every row has none. */
  readonly rowMinPx?: Float32Array
  readonly headerRows: number
  /** 1 when the kit keeps the row whole if it fits on a page. A row taller than the page is still cut. */
  readonly cantSplitRows?: Uint8Array
  readonly spaceBeforePx: number
  readonly spaceAfterPx: number
  readonly keepNext: boolean
  readonly keepLines: boolean
  readonly pageBreakBefore: boolean
  readonly widowControl: boolean
  /** Y of each inline page break, relative to the block top. Paragraphs only. */
  readonly inlineBreaks?: Float32Array
}

interface LineRect {
  top: number
  bottom: number
  left: number
  right: number
  anchor: number
}

function groupLineRects(rects: readonly DOMRect[]): LineRect[] {
  const usable = rects.filter((rect) => rect.width > 0.4 && rect.height > 0.4)
  usable.sort((a, b) => a.top - b.top || a.left - b.left)
  const lines: LineRect[] = []
  for (const rect of usable) {
    const last = lines[lines.length - 1]
    // Classify by the fragment's top. Growing last.bottom would let one tall rect swallow every line below it.
    if (last && rect.top < last.anchor + 8) {
      last.top = Math.min(last.top, rect.top)
      last.bottom = Math.max(last.bottom, rect.bottom)
      last.left = Math.min(last.left, rect.left)
      last.right = Math.max(last.right, rect.right)
    } else {
      lines.push({ top: rect.top, bottom: rect.bottom, left: rect.left, right: rect.right, anchor: rect.top })
    }
  }
  return lines
}

function contentLineRects(el: HTMLElement): LineRect[] {
  const rects: DOMRect[] = []
  let saw = false
  // One range around the children makes acdert return each child's whole box, so each child is measured alone.
  for (const child of el.children) {
    if (!(child instanceof HTMLElement) || child.classList.contains('marker')) continue
    saw = true
    const boxes = childContentRects(child)
    if (boxes.length > 0) rects.push(...boxes)
    else if (child.classList.contains('atom')) rects.push(child.getBoundingClientRect())
  }
  if (!saw) {
    const range = scratchRange()
    range.selectNodeContents(el)
    return groupLineRects([...range.getClientRects()])
  }
  return groupLineRects(rects)
}

function childContentRects(child: HTMLElement): DOMRect[] {
  const range = scratchRange()
  range.selectNodeContents(child)
  return [...range.getClientRects()]
}

export interface CutCell {
  readonly top: number
  readonly bottom: number
  readonly contentEnd: number
  readonly breaks: readonly { readonly y: number; readonly tight: boolean }[]
}

/** Allowed cuts inside rows. A y stays when every cell whose box contains it has a break there or has already ended. */
export function tableRowCuts(
  rowBottoms: Float32Array,
  cells: readonly CutCell[],
): { cuts: Float32Array; split: Uint8Array } {
  const ys: number[] = []
  for (const cell of cells) {
    for (const item of cell.breaks) {
      if (Number.isFinite(item.y)) ys.push(item.y)
    }
  }
  ys.sort((a, b) => a - b)
  const kept: number[] = []
  const bits: number[] = []
  let last = Number.NEGATIVE_INFINITY
  for (const y of ys) {
    if (y <= last + EPS) continue
    if (!yInsideRow(rowBottoms, y)) continue
    let tight = true
    let covered = false
    let blocked = false
    for (const cell of cells) {
      if (!cellCovers(cell, y)) continue
      covered = true
      if (!cellAccepts(cell, y, false)) {
        blocked = true
        break
      }
      if (!cellAccepts(cell, y, true)) tight = false
    }
    if (!covered || blocked) continue
    last = y
    kept.push(y)
    bits.push(tight ? 1 : 0)
  }
  return { cuts: Float32Array.from(kept), split: Uint8Array.from(bits) }
}

function yInsideRow(bottoms: Float32Array, y: number): boolean {
  let top = 0
  for (let i = 0; i < bottoms.length; i++) {
    const bottom = bottoms[i] ?? top
    if (y > top + EPS && y < bottom - EPS) return true
    if (y <= bottom + EPS) return false
    top = bottom
  }
  return false
}

function cellCovers(cell: CutCell, y: number): boolean {
  return y > cell.top + 0.01 && y <= cell.bottom + EPS
}

function cellAccepts(cell: CutCell, y: number, tight: boolean): boolean {
  if (y >= cell.contentEnd - EPS) return true
  for (const item of cell.breaks) {
    if (Math.abs(item.y - y) > EPS) continue
    if (tight && !item.tight) continue
    return true
  }
  return false
}

interface ContentBreak {
  readonly y: number
  readonly tight: boolean
}

interface TableReading {
  readonly rowBottoms: Float32Array
  readonly cantSplitRows: Uint8Array
  readonly rowCuts?: Float32Array
  readonly rowCutSplit?: Uint8Array
  readonly rowMinPx?: Float32Array
}

function readTable(table: HTMLTableElement, boxTop: number, zoom: number): TableReading {
  const z = zoom > 0 ? zoom : 1
  const n = table.rows.length
  const rowBottoms = new Float32Array(n)
  const cantSplitRows = new Uint8Array(n)
  const mins = new Float32Array(n)
  let anyMin = false
  const cells: CutCell[] = []
  for (let i = 0; i < n; i++) {
    const row = table.rows[i]
    if (!row) continue
    let bottom = 0
    for (const cell of row.cells) {
      const rect = cell.getBoundingClientRect()
      bottom = Math.max(bottom, rect.bottom - boxTop)
      const top = (rect.top - boxTop) / z
      const content = cellBreaks(cell, boxTop, z, top)
      cells.push({ top, bottom: (rect.bottom - boxTop) / z, contentEnd: content.contentEnd, breaks: content.breaks })
    }
    rowBottoms[i] = bottom / z
    if (rowRefusesSplit(row)) cantSplitRows[i] = 1
    const min = parseFloat(row.style.minHeight)
    if (Number.isFinite(min) && min > 0) {
      mins[i] = min
      anyMin = true
    }
  }
  for (let i = 1; i < n; i++) {
    const prev = rowBottoms[i - 1] ?? 0
    if ((rowBottoms[i] ?? 0) < prev + 0.5) rowBottoms[i] = prev + 0.5
  }
  const found = tableRowCuts(rowBottoms, cells)
  let loose = false
  for (let i = 0; i < found.split.length; i++) if ((found.split[i] ?? 1) === 0) loose = true
  return {
    rowBottoms,
    cantSplitRows,
    ...(found.cuts.length > 0 ? { rowCuts: found.cuts } : {}),
    ...(loose ? { rowCutSplit: found.split } : {}),
    ...(anyMin ? { rowMinPx: mins } : {}),
  }
}

function rowRefusesSplit(row: HTMLTableRowElement): boolean {
  // The kit refuses the whole row for cantSplit, an exact height, or any direct paragraph with keep-next or a page break.
  if (row.dataset.cantSplit === '1' || row.dataset.exact === '1') return true
  for (const cell of row.cells) {
    for (const child of cell.children) {
      if (!(child instanceof HTMLElement) || !child.classList.contains('p')) continue
      if (child.dataset.keepNext === '1' || child.dataset.pageBreakBefore === '1') return true
    }
  }
  return false
}

function cellBreaks(
  cell: HTMLElement,
  boxTop: number,
  zoom: number,
  top: number,
): { breaks: ContentBreak[]; contentEnd: number } {
  const breaks: ContentBreak[] = []
  let contentEnd = top
  for (const child of cell.children) {
    if (!(child instanceof HTMLElement)) continue
    for (const item of blockBreaks(child, boxTop, zoom)) {
      breaks.push(item)
      if (item.y > contentEnd) contentEnd = item.y
    }
  }
  return { breaks, contentEnd }
}

function blockBreaks(el: HTMLElement, boxTop: number, zoom: number): ContentBreak[] {
  if (el.classList.contains('p')) return paragraphBreaks(el, boxTop, zoom)
  if (el.classList.contains('t')) return nestedBreaks(el, boxTop, zoom)
  const bottom = (el.getBoundingClientRect().bottom - boxTop) / zoom
  return Number.isFinite(bottom) ? [{ y: bottom, tight: true }] : []
}

function paragraphBreaks(el: HTMLElement, boxTop: number, zoom: number): ContentBreak[] {
  const lines = contentLineRects(el)
  if (lines.length === 0) {
    const bottom = (el.getBoundingClientRect().bottom - boxTop) / zoom
    return Number.isFinite(bottom) ? [{ y: bottom, tight: true }] : []
  }
  // A line split needs four lines and the widow minimum on both sides. keep-lines, keep-next, and a page break do not.
  const blocked = el.dataset.keepLines === '1' || el.dataset.keepNext === '1' || el.dataset.pageBreakBefore === '1'
  const canSplit = !blocked && lines.length >= 4
  const minimum = el.dataset.widow === '1' ? 2 : 1
  const breaks: ContentBreak[] = []
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]
    if (!line) continue
    const y = (line.bottom - boxTop) / zoom
    if (!Number.isFinite(y)) continue
    const kept = i + 1
    const after = lines.length - kept
    const tight = i === lines.length - 1 || (canSplit && kept >= minimum && after >= minimum)
    breaks.push({ y, tight })
  }
  return breaks
}

function nestedBreaks(el: HTMLElement, boxTop: number, zoom: number): ContentBreak[] {
  const table = el.querySelector(':scope > table')
  if (!(table instanceof HTMLTableElement) || table.rows.length === 0) {
    const bottom = (el.getBoundingClientRect().bottom - boxTop) / zoom
    return Number.isFinite(bottom) ? [{ y: bottom, tight: true }] : []
  }
  const measured = readTable(table, boxTop, zoom)
  const breaks: ContentBreak[] = []
  for (let i = 0; i < measured.rowBottoms.length; i++) {
    breaks.push({ y: measured.rowBottoms[i] ?? 0, tight: true })
  }
  const cuts = measured.rowCuts
  const split = measured.rowCutSplit
  if (cuts) {
    for (let i = 0; i < cuts.length; i++) {
      const tight = !split || (split[i] ?? 1) !== 0
      breaks.push({ y: cuts[i] ?? 0, tight })
    }
  }
  return breaks
}

/** Reads a mounted or hidden block element: offsetHeight, plus line bottoms grouped from Range.getClientRects(). */
export function measureElement(el: HTMLElement, zoom: number): Measure {
  const z = zoom > 0 ? zoom : 1
  const box = el.getBoundingClientRect()
  const style = getComputedStyle(el)
  const spaceBeforePx = parseFloat(style.paddingTop) || 0
  const spaceAfterPx = parseFloat(style.paddingBottom) || 0
  const lines = el.classList.contains('t') ? [] : contentLineRects(el)
  const lineBottoms = new Float32Array(lines.length)
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]
    if (!line) continue
    lineBottoms[i] = (line.bottom - box.top) / z
  }
  for (let i = 1; i < lineBottoms.length; i++) {
    const prev = lineBottoms[i - 1] ?? 0
    if ((lineBottoms[i] ?? 0) < prev + 0.5) lineBottoms[i] = prev + 0.5
  }
  let rowBottoms: Float32Array | undefined
  let cantSplitRows: Uint8Array | undefined
  let rowCuts: Float32Array | undefined
  let rowCutSplit: Uint8Array | undefined
  let rowMinPx: Float32Array | undefined
  const table = el.querySelector(':scope > table')
  if (table instanceof HTMLTableElement && table.rows.length > 0) {
    const measured = readTable(table, box.top, z)
    rowBottoms = measured.rowBottoms
    cantSplitRows = measured.cantSplitRows
    rowCuts = measured.rowCuts
    rowCutSplit = measured.rowCutSplit
    rowMinPx = measured.rowMinPx
  }
  const lastLine = lineBottoms.length > 0 ? (lineBottoms[lineBottoms.length - 1] ?? 0) : 0
  const lastRow = rowBottoms && rowBottoms.length > 0 ? (rowBottoms[rowBottoms.length - 1] ?? 0) : 0
  const headerRows = Number(el.dataset.headerRows ?? '0')
  const inlineBreaks = pageBreakBottoms(el, box.top, z)
  return {
    widthPx: el.offsetWidth || box.width / z,
    heightPx: Math.max(box.height / z, lastLine, lastRow),
    lineBottoms,
    ...(rowBottoms ? { rowBottoms, cantSplitRows } : {}),
    ...(rowCuts ? { rowCuts } : {}),
    ...(rowCutSplit ? { rowCutSplit } : {}),
    ...(rowMinPx ? { rowMinPx } : {}),
    headerRows: Number.isFinite(headerRows) ? headerRows : 0,
    spaceBeforePx,
    spaceAfterPx,
    keepNext: el.dataset.keepNext === '1',
    keepLines: el.dataset.keepLines === '1',
    pageBreakBefore: el.dataset.pageBreakBefore === '1',
    widowControl: el.dataset.widow === '1',
    ...(inlineBreaks ? { inlineBreaks } : {}),
  }
}

function pageBreakBottoms(el: HTMLElement, boxTop: number, zoom: number): Float32Array | undefined {
  if (!el.classList.contains('p')) return undefined
  const atoms = el.querySelectorAll(':scope > .atom[data-atom="pageBreak"]')
  if (atoms.length === 0) return undefined
  const breaks = new Float32Array(atoms.length)
  for (let i = 0; i < atoms.length; i++) {
    const atom = atoms[i]
    if (!(atom instanceof HTMLElement)) continue
    breaks[i] = (atom.getBoundingClientRect().bottom - boxTop) / zoom
  }
  return breaks
}

export interface Slice {
  readonly block: BlockId
  readonly y0: number
  readonly y1: number
  /** Header rows to repeat at the top of a table continuation. */
  readonly repeatHeaderRows: number
}

export interface Page {
  readonly section: number
  readonly slices: readonly Slice[]
  readonly header: StoryKey | null
  readonly footer: StoryKey | null
  /** Body top in px from the page top: the top margin, or lower when the header is taller than the space it has. */
  readonly bodyTopPx: number
  readonly bodyHeightPx: number
  readonly contentWidthPx: number
}

export interface Layout {
  readonly pages: readonly Page[]
  /** False while blocks after `pages` are still unmeasured (the first chunks of a large open). */
  readonly complete: boolean
  /**
   * Page size, margins, and break type of each section, in section order.
   * titlePage is not part of the key: it chooses a header and does not change a block's measure.
   */
  readonly sectionKeys: readonly string[]
}

type MeasureFn = (block: Block, section: Section) => Measure | undefined

interface Frame {
  paperWidth: number
  paperHeight: number
  contentWidth: number
  contentLeft: number
  bodyTop: number
  bodyHeight: number
  header: StoryKey | null
  footer: StoryKey | null
  headerTop: number
  footerBottom: number
}

function headerKey(section: Section, pageNumber: number): StoryKey | null {
  if (pageNumber === 1 && section.titlePage && section.headers.first) return section.headers.first
  if (pageNumber % 2 === 0 && section.headers.even) return section.headers.even
  return section.headers.default ?? null
}

function footerKey(section: Section, pageNumber: number): StoryKey | null {
  if (pageNumber === 1 && section.titlePage && section.footers.first) return section.footers.first
  if (pageNumber % 2 === 0 && section.footers.even) return section.footers.even
  return section.footers.default ?? null
}

function storyHeight(doc: Doc, key: StoryKey | null, section: Section, measure: MeasureFn): number {
  if (!key) return 0
  const blocks = doc.stories.get(key)
  if (!blocks) return 0
  let height = 0
  for (const block of blocks) {
    const measured = measure(block, section)
    if (!measured) return 0
    height += measured.heightPx
  }
  return height
}

function pageFrame(doc: Doc, section: Section, pageNumber: number, measure: MeasureFn): Frame {
  const page = section.page
  const margin = page.margin
  const paperWidth = page.widthPt * PT_TO_PX
  const paperHeight = page.heightPt * PT_TO_PX
  const contentLeft = (margin.left + margin.gutter) * PT_TO_PX
  const contentWidth = Math.max(0, paperWidth - contentLeft - margin.right * PT_TO_PX)
  const marginTop = margin.top * PT_TO_PX
  const marginBottom = margin.bottom * PT_TO_PX
  const headerTop = margin.header * PT_TO_PX
  const footerBottom = margin.footer * PT_TO_PX
  const header = headerKey(section, pageNumber)
  const footer = footerKey(section, pageNumber)
  const headerH = storyHeight(doc, header, section, measure)
  const footerH = storyHeight(doc, footer, section, measure)
  const bodyTop = Math.max(marginTop, headerTop + headerH)
  const bodyBottom = paperHeight - Math.max(marginBottom, footerBottom + footerH)
  return {
    paperWidth,
    paperHeight,
    contentWidth,
    contentLeft,
    bodyTop,
    bodyHeight: Math.max(1, bodyBottom - bodyTop),
    header,
    footer,
    headerTop,
    footerBottom,
  }
}

function contentWidthOf(section: Section): number {
  const margin = section.page.margin
  return Math.max(0, (section.page.widthPt - margin.left - margin.gutter - margin.right) * PT_TO_PX)
}

function sectionIndexes(body: readonly Block[]): number[] {
  const out = new Array<number>(body.length)
  let section = 0
  for (let i = 0; i < body.length; i++) {
    out[i] = section
    const block = body[i]
    if (block?.kind === 'p' && block.sectionEnd !== null) section = block.sectionEnd + 1
  }
  return out
}

function samePaper(a: Section, b: Section): boolean {
  return (
    a.page.widthPt === b.page.widthPt &&
    a.page.heightPt === b.page.heightPt &&
    a.page.margin.top === b.page.margin.top &&
    a.page.margin.bottom === b.page.margin.bottom &&
    a.page.margin.left === b.page.margin.left &&
    a.page.margin.right === b.page.margin.right
  )
}

interface Placed {
  slice: Slice
  consumed: number
  continueY: number | null
}

function fit(
  block: BlockId,
  y0: number,
  y1: number,
  repeatHeaderRows: number,
  consumed: number,
  continueY: number | null,
): Placed {
  const end = Math.max(y1, y0)
  return {
    slice: { block, y0, y1: end, repeatHeaderRows },
    consumed: Math.max(0, consumed),
    continueY,
  }
}

/** How much of the measured space before this slice skips. Page top drops all of it; elsewhere only the overlap with the previous block's space after. */
function spaceSkip(atTop: boolean, m: Measure, previousAfterPx: number): number {
  if (!(m.spaceBeforePx > 0)) return 0
  const wanted = atTop ? m.spaceBeforePx : Math.min(m.spaceBeforePx, Math.max(0, previousAfterPx))
  if (!(wanted > 0)) return 0
  return Math.min(wanted, Math.max(0, m.heightPx - 0.01))
}

function spaceAfterBefore(body: readonly Block[], sectionOf: readonly number[], index: number, measure: MeasureFn, doc: Doc): number {
  if (index === 0) return 0
  const prev = body[index - 1]
  const section = doc.sections[sectionOf[index - 1] ?? 0]
  if (!prev || !section) return 0
  return measure(prev, section)?.spaceAfterPx ?? 0
}

function firstBreakAfter(m: Measure, y: number): number | null {
  const breaks = m.inlineBreaks
  if (!breaks || breaks.length === 0) return null
  let found: number | null = null
  for (let i = 0; i < breaks.length; i++) {
    const at = breaks[i] ?? 0
    if (at <= y + EPS) continue
    if (found === null || at < found) found = at
  }
  return found
}

/** The slice finished the block and a page break sits at that end, so the next block starts a new page. */
function endsWithInlineBreak(m: Measure, placed: Placed): boolean {
  if (placed.continueY !== null) return false
  const at = firstBreakAfter(m, placed.slice.y0)
  if (at === null) return false
  return at >= m.heightPx - EPS && at <= placed.slice.y1 + EPS
}

function placeLines(
  id: BlockId,
  m: Measure,
  y0: number,
  available: number,
  atTop: boolean,
  skip: number,
  pageBody: number,
): Placed | null {
  const bottoms = m.lineBottoms
  const n = bottoms.length
  let start = y0
  if (skip > 0 && start <= EPS) start = skip
  // A page break ends the page even when later lines would fit. keepLines stops at the break.
  const breakAt = firstBreakAfter(m, start)
  const splitAt = breakAt !== null && breakAt < m.heightPx - EPS ? breakAt : null
  if (splitAt !== null && splitAt - start <= available + EPS) {
    return fit(id, start, splitAt, 0, splitAt - start, splitAt)
  }
  const limit = splitAt ?? m.heightPx
  if (m.keepLines && !atTop && start <= EPS && limit - start <= pageBody + EPS && limit - start > available + EPS) {
    return null
  }
  if (splitAt === null && m.heightPx - start <= available + EPS) return fit(id, start, m.heightPx, 0, m.heightPx - start, null)

  let first = 0
  while (first < n && (bottoms[first] ?? 0) <= start + EPS) first += 1
  if (first >= n) {
    if (splitAt !== null && splitAt - start <= available + EPS) return fit(id, start, splitAt, 0, splitAt - start, splitAt)
    if (atTop || m.heightPx - start <= available + EPS) return fit(id, start, m.heightPx, 0, Math.max(0, m.heightPx - start), null)
    return null
  }
  let last = first - 1
  for (let i = first; i < n; i++) {
    const bottom = bottoms[i] ?? 0
    if (splitAt !== null && bottom > splitAt + EPS) break
    if (bottom - start <= available + EPS) last = i
    else break
  }
  if (last < first) {
    if (!atTop) return null
    last = first
  }
  const atStart = start <= m.spaceBeforePx + EPS
  let count = last - first + 1
  if (m.widowControl && atStart && count < 2 && n - first >= 2 && !atTop) return null
  const after = n - (last + 1)
  if (m.widowControl && after === 1 && count >= 2) {
    last -= 1
    count -= 1
    if (count < 1) {
      if (!atTop) return null
      last += 1
    }
  }
  const end = bottoms[last] ?? start
  if (splitAt !== null && splitAt <= end + EPS) return fit(id, start, splitAt, 0, splitAt - start, splitAt)
  if (last >= n - 1) {
    if (splitAt !== null && splitAt > end + EPS) return fit(id, start, end, 0, Math.max(0, end - start), end)
    return fit(id, start, end, 0, Math.max(0, end - start), null)
  }
  if (end <= start + 0.25) return atTop ? fit(id, start, m.heightPx, 0, Math.max(0, m.heightPx - start), null) : null
  return fit(id, start, end, 0, end - start, end)
}

function placeTable(
  id: BlockId,
  m: Measure,
  y0: number,
  available: number,
  atTop: boolean,
  skip: number,
  pageBody: number,
): Placed | null {
  const rows = m.rowBottoms
  if (!rows || rows.length === 0) return null
  const n = rows.length
  const headerRows = Math.min(m.headerRows, n)
  const headerBottom = headerRows > 0 ? (rows[headerRows - 1] ?? 0) : 0
  let start = y0
  if (skip > 0 && start <= EPS) start = skip
  const continuation = headerRows > 0 && start >= headerBottom - EPS && start > EPS
  const repeat = continuation ? headerRows : 0
  const repeatH = continuation ? headerBottom : 0
  const whole = repeatH + (m.heightPx - start)
  if (whole <= available + EPS) return fit(id, start, m.heightPx, repeat, whole, null)

  let first = 0
  while (first < n && (rows[first] ?? 0) <= start + EPS) first += 1
  if (first >= n) {
    if (atTop || whole <= available + EPS) return fit(id, start, m.heightPx, repeat, Math.max(0, m.heightPx - start), null)
    return null
  }
  let last = first - 1
  for (let i = first; i < n; i++) {
    const visual = repeatH + ((rows[i] ?? 0) - start)
    if (visual <= available + EPS) last = i
    else break
  }
  const next = last + 1
  if (next < n) {
    const cut = cutInRow(m, next, start, available, repeatH, atTop, pageBody, headerRows)
    if (cut !== null) return finishTableCut(id, m, start, cut, repeat, repeatH)
  }
  if (last < first) {
    if (!atTop) return null
    last = first
  }
  // A header-only page moves the table, or takes the next row when that row itself fits on a page.
  if (headerRows > 0 && last === headerRows - 1 && n > headerRows && first < headerRows) {
    if (!atTop) return null
    const body = headerRows
    if (body !== next) {
      const cut = cutInRow(m, body, start, available, repeatH, atTop, pageBody, headerRows)
      if (cut !== null) return finishTableCut(id, m, start, cut, repeat, repeatH)
    }
    const bodyTop = rows[body - 1] ?? 0
    const bodyBottom = rows[body] ?? bodyTop
    if (bodyBottom - bodyTop <= pageBody + EPS) last = Math.min(n - 1, headerRows)
  }
  const end = rows[last] ?? start
  const consumed = repeatH + (end - start)
  if (last >= n - 1) {
    if (whole <= available + EPS) return fit(id, start, m.heightPx, repeat, whole, null)
    return fit(id, start, end, repeat, consumed, null)
  }
  if (end <= start + 0.25) return atTop ? fit(id, start, end + 1, repeat, consumed + 1, end + 1) : null
  return fit(id, start, end, repeat, consumed, end)
}

function finishTableCut(
  id: BlockId,
  m: Measure,
  start: number,
  y: number,
  repeat: number,
  repeatH: number,
): Placed {
  let end = y
  if (end < start + 0.5) end = Math.min(m.heightPx, start + 1)
  if (end > m.heightPx) end = m.heightPx
  if (end >= m.heightPx - EPS) return fit(id, start, m.heightPx, repeat, repeatH + Math.max(0, m.heightPx - start), null)
  return fit(id, start, end, repeat, repeatH + (end - start), end)
}

function cutInRow(
  m: Measure,
  row: number,
  start: number,
  available: number,
  repeatH: number,
  atTop: boolean,
  pageBody: number,
  headerRows: number,
): number | null {
  const rows = m.rowBottoms
  if (!rows || row < 0 || row >= rows.length || row < headerRows) return null
  const top = row <= 0 ? 0 : (rows[row - 1] ?? 0)
  const bottom = rows[row] ?? top
  // A row taller than the page is cut even when the kit would paint it whole. Anything else follows the kit.
  const taller = bottom - top > pageBody + EPS
  if (!taller && (m.cantSplitRows?.[row] ?? 0) !== 0) return null
  const cuts = m.rowCuts
  const flags = m.rowCutSplit
  let best = start
  if (cuts) {
    for (let i = 0; i < cuts.length; i++) {
      const y = cuts[i] ?? 0
      if (y <= start + EPS || y <= top + EPS || y >= bottom - EPS) continue
      if (repeatH + (y - start) > available + EPS) continue
      const legal = taller || !flags || i >= flags.length || (flags[i] ?? 1) !== 0
      if (!legal) continue
      if (y > best) best = y
    }
  }
  if (best > start + 0.25) {
    const fragment = best - Math.max(start, top)
    const firstPiece = start <= top + EPS
    const contentBefore = !atTop || repeatH > EPS
    const minH = !taller && firstPiece && contentBefore ? (m.rowMinPx?.[row] ?? 0) : 0
    if (fragment + EPS >= minH) return best
    return null
  }
  if (!taller) return null
  let end = start + Math.max(0, available - repeatH)
  if (end <= top + 0.25) return null
  if (end > bottom) end = bottom
  if (end <= start + 0.25) {
    if (!atTop) return null
    end = Math.min(bottom, start + 1)
  }
  if (end <= start + 0.25 || end <= top + 0.25) return null
  return end
}

function placeAtomic(id: BlockId, m: Measure, y0: number, available: number, atTop: boolean, skip: number): Placed | null {
  let start = y0
  if (skip > 0 && start <= EPS) start = skip
  const consumed = m.heightPx - start
  if (consumed <= available + EPS || atTop) return fit(id, start, m.heightPx, 0, Math.max(0, consumed), null)
  return null
}

function placeBlock(
  id: BlockId,
  m: Measure,
  y0: number,
  available: number,
  atTop: boolean,
  skip: number,
  pageBody: number,
): Placed | null {
  if (m.rowBottoms && m.rowBottoms.length > 0) return placeTable(id, m, y0, available, atTop, skip, pageBody)
  if (m.lineBottoms.length > 0 || (m.inlineBreaks !== undefined && m.inlineBreaks.length > 0)) {
    return placeLines(id, m, y0, available, atTop, skip, pageBody)
  }
  return placeAtomic(id, m, y0, available, atTop, skip)
}

function chainFit(
  body: readonly Block[],
  sectionOf: readonly number[],
  index: number,
  measure: MeasureFn,
  doc: Doc,
  pageBody: number,
  anchorAfter: number,
): { here: number; fresh: number } | null {
  let here = 0
  let fresh = 0
  let prevAfter = anchorAfter
  for (let i = index, guard = 0; i < body.length && guard < 40; i++, guard++) {
    if (i > index && (sectionOf[i] ?? 0) !== (sectionOf[index] ?? 0)) break
    const block = body[i]
    const section = doc.sections[sectionOf[i] ?? 0] ?? doc.sections[0]
    if (!block || !section) return null
    const measured = measure(block, section)
    if (!measured) return null
    if (i > index && measured.pageBreakBefore) break
    const before = Math.max(0, measured.spaceBeforePx)
    const overlap = Math.min(before, Math.max(0, prevAfter))
    const breakAt = firstBreakAfter(measured, 0)
    const used = breakAt === null ? measured.heightPx : Math.min(measured.heightPx, breakAt)
    here += Math.max(0, used - overlap)
    fresh += Math.max(0, used - (i === index ? before : overlap))
    prevAfter = measured.spaceAfterPx
    if (breakAt !== null || !measured.keepNext || fresh > pageBody) break
  }
  return { here, fresh }
}

class PageBuilder {
  used = 0
  readonly slices: Slice[] = []
  readonly sectionIndex: number
  readonly bodyHeight: number
  readonly bodyTop: number
  readonly header: StoryKey | null
  readonly footer: StoryKey | null
  readonly softTop: boolean
  readonly contentWidth: number

  constructor(
    sectionIndex: number,
    bodyHeight: number,
    bodyTop: number,
    header: StoryKey | null,
    footer: StoryKey | null,
    softTop: boolean,
    contentWidth: number,
  ) {
    this.sectionIndex = sectionIndex
    this.bodyHeight = bodyHeight
    this.bodyTop = bodyTop
    this.header = header
    this.footer = footer
    this.softTop = softTop
    this.contentWidth = contentWidth
  }

  remaining(): number {
    return Math.max(0, this.bodyHeight - this.used)
  }

  add(placed: Placed): void {
    this.slices.push(placed.slice)
    this.used += placed.consumed
  }

  toPage(): Page {
    return {
      section: this.sectionIndex,
      slices: this.slices,
      header: this.header,
      footer: this.footer,
      bodyTopPx: this.bodyTop,
      bodyHeightPx: this.bodyHeight,
      contentWidthPx: this.contentWidth,
    }
  }
}

function openBuilder(doc: Doc, section: Section, sectionIndex: number, pageNumber: number, measure: MeasureFn, softTop: boolean): PageBuilder {
  const frame = pageFrame(doc, section, pageNumber, measure)
  return new PageBuilder(
    sectionIndex,
    frame.bodyHeight,
    frame.bodyTop,
    frame.header,
    frame.footer,
    softTop,
    frame.contentWidth,
  )
}

function blankPage(doc: Doc, section: Section, sectionIndex: number, pageNumber: number, measure: MeasureFn): Page {
  const frame = pageFrame(doc, section, pageNumber, measure)
  return {
    section: sectionIndex,
    slices: [],
    header: frame.header,
    footer: frame.footer,
    bodyTopPx: frame.bodyTop,
    bodyHeightPx: frame.bodyHeight,
    contentWidthPx: frame.contentWidth,
  }
}

interface Cursor {
  index: number
  y0: number
}

/**
 * 0-based page of a block's first slice.
 * Null when the layout has not placed that id: a block past the measured part of an
 * incomplete layout, a nested paragraph, or an id the document does not have.
 * `PageView.pageOf` answers 0 when it cannot find a paragraph, so callers that need
 * "not placed" use this instead.
 */
export function pageOfBlock(layout: Layout, id: string): number | null {
  for (let i = 0; i < layout.pages.length; i++) {
    const page = layout.pages[i]
    if (!page) continue
    for (const slice of page.slices) {
      if (slice.block === id) return i
    }
  }
  return null
}

export interface SliceSpan {
  readonly page: number
  readonly y0: number
  readonly y1: number
}

/**
 * The page whose slice contains `y`. A y on the cut stays with the slice that ends there:
 * the last line of a cell ends on that y, and the next slice starts after it.
 */
export function pageAtSliceY(spans: readonly SliceSpan[], y: number): number | null {
  const hit = spans.find((span) => y >= span.y0 - 0.5 && y <= span.y1)
  return hit ? hit.page : null
}

/**
 * Y used to choose a slice. The caret after the last character of a paragraph is still on
 * that character's line; a range at the end can report a y on the following line, past the cut.
 */
export function caretSliceY(offset: number, length: number, yAt: (offset: number) => number | null): number | null {
  const y = yAt(offset)
  if (y === null || offset <= 0 || offset < length) return y
  const prev = yAt(offset - 1)
  if (prev === null || prev >= y) return y
  return prev
}

export interface VerticalLine {
  readonly globalMid: number
  readonly left: number
  readonly right: number
  readonly story: 'body' | 'header' | 'footer'
}

/** The next visual line in the same story. Header and footer lines are not steps of the body. */
export function nextVerticalLine<T extends VerticalLine>(
  lines: readonly T[],
  current: T,
  direction: 'up' | 'down',
  goal: number,
): T | null {
  const forward = direction === 'down'
  const candidates = lines.filter((line) => {
    if (line === current || line.story !== current.story) return false
    return forward ? line.globalMid > current.globalMid + 4 : line.globalMid < current.globalMid - 4
  })
  if (candidates.length === 0) return null
  const containing = candidates.filter((line) => goal >= line.left - 16 && goal <= line.right + 16)
  const pool = containing.length > 0 ? containing : candidates
  pool.sort((a, b) => Math.abs(a.globalMid - current.globalMid) - Math.abs(b.globalMid - current.globalMid))
  return pool[0] ?? null
}

function pageContaining(layout: Layout, id: BlockId): number {
  const page = pageOfBlock(layout, id)
  return page === null ? -1 : page
}

function changedAhead(body: readonly Block[], index: number, changed: ReadonlySet<BlockId>): boolean {
  if (changed.size === 0) return false
  for (let i = index; i < body.length; i++) {
    const block = body[i]
    if (block && changed.has(block.id)) return true
  }
  return false
}

function cursorFromPage(previous: Layout, body: readonly Block[], pageIndex: number): { pageIndex: number; cursor: Cursor; soft: boolean } {
  const first = previous.pages[pageIndex]?.slices[0]
  if (first) {
    const idx = body.findIndex((block) => block.id === first.block)
    if (idx >= 0) return { pageIndex, cursor: { index: idx, y0: first.y0 }, soft: false }
  }
  let index = 0
  const prev = pageIndex > 0 ? previous.pages[pageIndex - 1] : undefined
  const last = prev?.slices[prev.slices.length - 1]
  if (last) {
    const idx = body.findIndex((block) => block.id === last.block)
    if (idx >= 0) index = idx + 1
  }
  return { pageIndex, cursor: { index: Math.min(index, body.length), y0: 0 }, soft: true }
}

function pageOfEarlier(previous: Layout, body: readonly Block[], before: number): number {
  for (let i = before - 1; i >= 0; i--) {
    const id = body[i]?.id
    if (!id) continue
    const found = pageContaining(previous, id)
    if (found >= 0) return found
  }
  return -1
}

function earliestRemovedPage(previous: Layout, body: readonly Block[]): number {
  const live = new Set<BlockId>()
  for (const block of body) live.add(block.id)
  for (let i = 0; i < previous.pages.length; i++) {
    const page = previous.pages[i]
    if (!page) continue
    for (const slice of page.slices) {
      if (!live.has(slice.block)) return i
    }
  }
  return -1
}

function restart(previous: Layout, body: readonly Block[], changed: ReadonlySet<BlockId>): { pageIndex: number; cursor: Cursor; soft: boolean } {
  if (changed.size === 0) return cursorFromPage(previous, body, Math.max(0, previous.pages.length - 1))
  let earliest = -1
  for (let i = 0; i < body.length; i++) {
    const block = body[i]
    if (block && changed.has(block.id)) {
      earliest = i
      break
    }
  }
  let pageIndex = -1
  if (earliest >= 0) {
    const id = body[earliest]?.id ?? ('' as BlockId)
    pageIndex = pageContaining(previous, id)
    // A new block is on no previous page. Anchor on the nearest earlier block the
    // previous layout did place. If nothing before it was placed, the old first slice
    // now sits after this block, so the cursor starts at the beginning of the body.
    if (pageIndex < 0) pageIndex = pageOfEarlier(previous, body, earliest)
    if (pageIndex < 0) return { pageIndex: 0, cursor: { index: 0, y0: 0 }, soft: false }
  }
  // `changed` names blocks that are still in the body. A page that still holds a
  // removed block is an earlier anchor than the next surviving block, which may sit pages later.
  const removed = earliestRemovedPage(previous, body)
  if (removed >= 0 && (pageIndex < 0 || removed < pageIndex)) pageIndex = removed
  if (pageIndex < 0) pageIndex = Math.max(0, previous.pages.length - 1)
  else if (pageIndex > 0) pageIndex -= 1
  return cursorFromPage(previous, body, pageIndex)
}

function tryReuse(
  previous: Layout | null,
  pageIndex: number,
  doc: Doc,
  body: readonly Block[],
  cursor: Cursor,
  changed: ReadonlySet<BlockId>,
  sectionOf: readonly number[],
  measure: MeasureFn,
  keys: readonly string[],
): readonly Page[] | null {
  if (!previous) return null
  if (!previous.complete && pageIndex >= previous.pages.length - 1) return null
  if (pageIndex >= previous.pages.length) return null
  if (changedAhead(body, cursor.index, changed)) return null
  const prev = previous.pages[pageIndex]
  const first = prev?.slices[0]
  const block = body[cursor.index]
  if (!prev || !first || !block) return null
  if (first.block !== block.id || Math.abs(first.y0 - cursor.y0) > 0.25) return null
  const sectionIndex = sectionOf[cursor.index] ?? 0
  if (prev.section !== sectionIndex) return null
  if ((previous.sectionKeys[prev.section] ?? '') !== (keys[prev.section] ?? '')) return null
  const section = doc.sections[prev.section]
  if (!section) return null
  const frame = pageFrame(doc, section, pageIndex + 1, measure)
  if (frame.header !== prev.header || frame.footer !== prev.footer) return null
  if (Math.abs(frame.bodyTop - prev.bodyTopPx) > 0.25) return null
  if (Math.abs(frame.bodyHeight - prev.bodyHeightPx) > 0.25) return null
  if (Math.abs(frame.contentWidth - prev.contentWidthPx) > 0.25) return null
  return previous.pages.slice(pageIndex)
}

// titlePage picks which header prints. Leaving it out of the key keeps a title-page switch from measuring body blocks again.
function sectionFrameKey(section: Section): string {
  const page = section.page
  const margin = page.margin
  return [
    page.widthPt,
    page.heightPt,
    margin.top,
    margin.right,
    margin.bottom,
    margin.left,
    margin.header,
    margin.footer,
    margin.gutter,
    section.breakType,
  ].join(':')
}

function sectionKeysOf(doc: Doc): string[] {
  return doc.sections.map(sectionFrameKey)
}

function sameSectionKeys(a: readonly string[], b: readonly string[]): boolean {
  if (a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false
  return true
}

function finishLayout(doc: Doc, pages: readonly Page[], complete: boolean): Layout {
  return { pages, complete, sectionKeys: sectionKeysOf(doc) }
}

function firstDivergentPage(previous: Layout, keys: readonly string[]): number {
  for (let i = 0; i < previous.pages.length; i++) {
    const page = previous.pages[i]
    if (!page) return i
    if (previous.sectionKeys[page.section] !== keys[page.section]) return i
  }
  return previous.pages.length
}

function decorationKeysDiffer(doc: Doc, layout: Layout): boolean {
  for (let i = 0; i < layout.pages.length; i++) {
    const page = layout.pages[i]
    if (!page) return true
    const section = doc.sections[page.section]
    if (!section) return true
    if (headerKey(section, i + 1) !== page.header) return true
    if (footerKey(section, i + 1) !== page.footer) return true
  }
  return false
}

function patchDecorations(doc: Doc, layout: Layout, measure: MeasureFn): Page[] | null {
  const pages: Page[] = []
  for (let i = 0; i < layout.pages.length; i++) {
    const page = layout.pages[i]
    if (!page) return null
    const section = doc.sections[page.section]
    if (!section) return null
    const frame = pageFrame(doc, section, i + 1, measure)
    if (Math.abs(frame.bodyTop - page.bodyTopPx) > 0.25) return null
    if (Math.abs(frame.bodyHeight - page.bodyHeightPx) > 0.25) return null
    if (Math.abs(frame.contentWidth - page.contentWidthPx) > 0.25) return null
    if (frame.header === page.header && frame.footer === page.footer) {
      pages.push(page)
      continue
    }
    pages.push({ ...page, header: frame.header, footer: frame.footer })
  }
  return pages
}

function wantsBlank(breakType: 'evenPage' | 'oddPage', pageNumber: number): boolean {
  const even = pageNumber % 2 === 0
  return breakType === 'evenPage' ? !even : even
}

/**
 * Word's page rules over cached measures: pageBreakBefore, an inline page break, keepNext chains (capped at one page),
 * keepLines, widow and orphan control, section breaks, table rows (whole when they fit, cut inside a row when
 * they do not), header rows, and paragraph spacing. The top of a page drops space before. Between two blocks the gap is the larger of the first's
 * space after and the second's space before. The measured height stays the block's own; the skip is applied
 * when the slice is placed. A break at the end of a block starts the next block on a new page, and that one
 * transition absorbs the next block's pageBreakBefore.
 *
 * Incremental: with `previous` and `changed`, starts at the page holding the first changed block and stops at the first
 * page whose start (block, y0) equals the previous layout's page start at the same index, reusing the rest.
 * A changed section key (page size, margins, or break type) starts at that section's first page instead.
 */
export function paginate(doc: Doc, measure: MeasureFn, previous: Layout | null, changed: readonly BlockId[]): Layout {
  const body = doc.body
  const sectionOf = sectionIndexes(body)
  const changedSet = new Set(changed)
  const keys = sectionKeysOf(doc)
  if (body.length === 0) {
    const section = doc.sections[0]
    if (!section) return { pages: [], complete: true, sectionKeys: keys }
    return finishLayout(doc, [blankPage(doc, section, 0, 1, measure)], true)
  }
  const geometrySame = !previous || sameSectionKeys(previous.sectionKeys, keys)
  let forceStart = previous !== null && !geometrySame
  if (geometrySame && previous && changed.length === 0 && previous.complete) {
    if (!decorationKeysDiffer(doc, previous)) return previous
    const patched = patchDecorations(doc, previous, measure)
    if (patched) return finishLayout(doc, patched, true)
    forceStart = true
  }

  const pages: Page[] = []
  let cursor: Cursor = { index: 0, y0: 0 }
  let softNext = false
  if (previous && previous.pages.length > 0 && !forceStart) {
    const start = restart(previous, body, changedSet)
    for (let i = 0; i < start.pageIndex; i++) {
      const page = previous.pages[i]
      if (page) pages.push(page)
    }
    cursor = start.cursor
    softNext = start.soft
  } else if (previous && previous.pages.length > 0) {
    let from = geometrySame ? 0 : firstDivergentPage(previous, keys)
    let located = cursorFromPage(previous, body, from)
    if (changed.length > 0) {
      const start = restart(previous, body, changedSet)
      // cursorFromPage anchors on the old page's first slice. A block inserted before that
      // slice is on no previous page, so restart's cursor is the one that still places it.
      if (start.pageIndex < from || (start.pageIndex === from && start.cursor.index < located.cursor.index)) {
        from = start.pageIndex
        located = start
      }
    }
    for (let i = 0; i < from; i++) {
      const page = previous.pages[i]
      if (page) pages.push(page)
    }
    cursor = located.cursor
    softNext = located.soft
  }

  let pending: PageBuilder | null = null
  let inlineOpensPage = false
  let guard = 0
  const maxGuard = body.length * 6 + 8
  while (cursor.index < body.length && guard < maxGuard) {
    guard += 1
    if (!pending) {
      const reused = tryReuse(previous, pages.length, doc, body, cursor, changedSet, sectionOf, measure, keys)
      if (reused) {
        for (const page of reused) pages.push(page)
        return finishLayout(doc, pages, previous?.complete ?? true)
      }
    }

    const index = cursor.index
    const block = body[index]
    const sectionIndex = sectionOf[index] ?? 0
    const section = doc.sections[sectionIndex] ?? doc.sections[0]
    if (!block || !section) break

    if (cursor.y0 === 0 && index > 0 && sectionIndex !== (sectionOf[index - 1] ?? 0)) {
      const prevSection = doc.sections[sectionOf[index - 1] ?? 0]
      const continuous = section.breakType === 'continuous' && prevSection !== undefined && samePaper(prevSection, section)
      if (!(continuous && pending && pending.slices.length > 0)) {
        if (pending && pending.slices.length > 0) pages.push(pending.toPage())
        pending = null
        softNext = false
        if (section.breakType === 'evenPage' || section.breakType === 'oddPage') {
          const number = pages.length + 1
          if (wantsBlank(section.breakType, number)) pages.push(blankPage(doc, section, sectionIndex, number, measure))
        }
      }
    }

    const measured = measure(block, section)
    if (!measured) {
      if (pending && pending.slices.length > 0) pages.push(pending.toPage())
      return finishLayout(doc, pages, false)
    }

    if (!pending) {
      pending = openBuilder(doc, section, sectionIndex, pages.length + 1, measure, softNext)
      softNext = false
    }

    if (cursor.y0 === 0 && pending.slices.length > 0 && (measured.pageBreakBefore || inlineOpensPage)) {
      pages.push(pending.toPage())
      pending = openBuilder(doc, section, sectionIndex, pages.length + 1, measure, false)
    }
    if (cursor.y0 === 0) inlineOpensPage = false

    const atTop = pending.slices.length === 0
    const earlierAfter = atTop ? 0 : spaceAfterBefore(body, sectionOf, index, measure, doc)
    if (!atTop && cursor.y0 === 0 && measured.keepNext) {
      const chain = chainFit(body, sectionOf, index, measure, doc, pending.bodyHeight, earlierAfter)
      if (chain && chain.fresh <= pending.bodyHeight + EPS && chain.here > pending.remaining() + EPS) {
        pages.push(pending.toPage())
        pending = null
        softNext = true
        continue
      }
    }

    const skip = cursor.y0 === 0 ? spaceSkip(atTop, measured, earlierAfter) : 0
    let placed = placeBlock(block.id, measured, cursor.y0, pending.remaining(), atTop, skip, pending.bodyHeight)
    if (!placed) {
      if (!atTop) {
        pages.push(pending.toPage())
        pending = null
        softNext = true
        continue
      }
      placed = fit(block.id, cursor.y0, measured.heightPx, 0, measured.heightPx, null)
    }
    pending.add(placed)
    if (endsWithInlineBreak(measured, placed)) inlineOpensPage = true
    if (placed.continueY === null || placed.continueY <= cursor.y0 + 0.25) {
      cursor = { index: index + 1, y0: 0 }
    } else {
      cursor = { index, y0: placed.continueY }
      pages.push(pending.toPage())
      pending = null
      softNext = true
    }
  }

  if (pending && pending.slices.length > 0) pages.push(pending.toPage())
  if (pages.length === 0) {
    const section = doc.sections[0]
    if (section) pages.push(blankPage(doc, section, 0, 1, measure))
  }
  return finishLayout(doc, pages, cursor.index >= body.length && guard < maxGuard)
}

export interface CaretRect {
  readonly page: number
  readonly x: number
  readonly top: number
  readonly height: number
}

interface Place {
  para: Paragraph
  owner: Block
  row: number
  order: number
}

interface LineBox {
  page: number
  left: number
  right: number
  top: number
  bottom: number
  globalMid: number
  paraId: string
  story: 'body' | 'header' | 'footer'
}

function containsId(block: Block, id: string): boolean {
  if (block.id === id) return true
  if (block.kind === 'p') return false
  if (block.kind === 'opaque') return block.preview.some((para) => para.id === id)
  for (const row of block.rows) {
    for (const cell of row.cells) {
      for (const child of cell.blocks) {
        if (containsId(child, id)) return true
      }
    }
  }
  return false
}

function assignStyle(el: HTMLElement, prop: 'width' | 'height' | 'top' | 'left', value: string): void {
  if (el.style[prop] !== value) el.style[prop] = value
}

export interface ImageHitCandidate {
  readonly p: ParaId
  readonly base: number
  readonly left: number
  readonly right: number
  readonly top: number
  readonly bottom: number
  readonly cell: { readonly left: number; readonly right: number; readonly top: number; readonly bottom: number } | null
  readonly inkBefore: boolean
  readonly inkAfter: boolean
  readonly contentAbove: boolean
  readonly contentBelow: boolean
}

/**
 * Resolves a click/drag point against inline images in the nearest slice. Keeps hits inside the image's own table
 * cell, and sends clicks below the last image line (where `caretPositionFromPoint` sees no text node) to `base + 1`.
 */
export function resolveImageHit(
  candidates: readonly ImageHitCandidate[],
  x: number,
  y: number,
  rawX: number,
  rawY: number,
): Pos | null {
  for (const item of candidates) {
    if (item.right <= item.left || item.bottom <= item.top) continue
    if (item.cell) {
      if (x < item.cell.left || x > item.cell.right || y < item.cell.top || y > item.cell.bottom) continue
    }
    if (rawY >= item.top && rawY <= item.bottom) {
      if (rawX >= item.left && rawX <= item.right) {
        return { p: item.p, o: rawX > (item.left + item.right) / 2 ? item.base + 1 : item.base }
      }
      const after = rawX > item.right
      if (after ? !item.inkAfter : !item.inkBefore) {
        return { p: item.p, o: after ? item.base + 1 : item.base }
      }
      continue
    }
    if (rawY > item.bottom) {
      if (!item.contentBelow && !item.inkAfter) {
        return { p: item.p, o: item.base + 1 }
      }
      continue
    }
    if (!item.contentAbove && !item.inkBefore) {
      return { p: item.p, o: rawX > item.right && !item.inkAfter ? item.base + 1 : item.base }
    }
  }
  return null
}

/** Caret box on the left or right edge of an inline atom. Images sit `vertical-align: text-bottom`, so the caret sits at their bottom edge. */
export function atomCaretBox(
  rect: { readonly left: number; readonly right: number; readonly top: number; readonly bottom: number },
  after: boolean,
  isImage: boolean,
  lineHeightPx: number,
): { readonly left: number; readonly top: number; readonly height: number } {
  const left = after ? rect.right : rect.left
  const boxH = rect.bottom - rect.top
  if (isImage) {
    const height = Math.max(lineHeightPx, 1)
    const top = boxH >= height ? rect.bottom - height : rect.top
    return { left, top, height }
  }
  return { left, top: rect.top, height: Math.max(boxH, 1) }
}

function defaultCaretHeightPx(paraEl: HTMLElement): number {
  for (const child of paraEl.children) {
    if (!(child instanceof HTMLElement) || child.classList.contains('marker') || child.classList.contains('atom')) continue
    const size = parseFloat(child.style.fontSize)
    if (Number.isFinite(size) && size > 0) return Math.max(size * 1.15, 16)
  }
  return 16
}

function inlineAtomCaretViewport(
  point: { readonly node: Node; readonly offset: number },
  o: number,
  sliceRect: DOMRect | null,
  lineHeightPx: number,
): { readonly left: number; readonly top: number; readonly height: number } | null {
  if (point.node instanceof Text) return null
  const parent = point.node
  const prev = point.offset > 0 ? parent.childNodes[point.offset - 1] : null
  const next = parent.childNodes[point.offset] ?? null
  let atom: HTMLElement | null = null
  let after = false
  if (
    prev instanceof HTMLElement
    && prev.classList.contains('atom')
    && prev.dataset.atom !== 'pageBreak'
    && Number(prev.dataset.o) + 1 === o
  ) {
    atom = prev
    after = true
  } else if (
    next instanceof HTMLElement
    && next.classList.contains('atom')
    && next.dataset.atom !== 'pageBreak'
    && Number(next.dataset.o) === o
  ) {
    atom = next
    after = false
  }
  if (!atom) return null
  const rects = [...atom.getClientRects()].filter((item) => item.height > 0 || item.width > 0)
  const rect =
    rects.find((item) => !sliceRect || (item.bottom > sliceRect.top + 0.5 && item.top < sliceRect.bottom - 0.5)) ??
    rects[after ? rects.length - 1 : 0]
  if (!rect) return null
  return atomCaretBox(rect, after, atom.dataset.atom === 'image', lineHeightPx)
}

/**
 * The side of a picture a point on or around the picture's line names. acdert's caretPositionFromPoint never answers
 * with an image: on a line a picture fills, or in empty space below a picture, it misses the image altogether.
 */
function imageSideAt(slice: HTMLElement, x: number, y: number, rawX: number, rawY: number): Pos | null {
  const clip = slice.querySelector(':scope > .clip')
  const clipRect = clip instanceof HTMLElement ? clip.getBoundingClientRect() : null
  const candidates: ImageHitCandidate[] = []
  for (const img of slice.querySelectorAll<HTMLElement>('img.atom[data-atom="image"]')) {
    if (img.closest('.repeat-head')) continue
    const raw = img.getBoundingClientRect()
    if (raw.width <= 0 || raw.height <= 0) continue
    const top = clipRect ? Math.max(raw.top, clipRect.top) : raw.top
    const bottom = clipRect ? Math.min(raw.bottom, clipRect.bottom) : raw.bottom
    if (bottom <= top + 0.5) continue
    const block = img.closest('.blk.p')
    const id = block instanceof HTMLElement ? block.dataset.id : undefined
    const base = Number(img.dataset.o)
    if (!(block instanceof HTMLElement) || !id || !Number.isInteger(base)) continue
    const td = img.closest('td')
    const cellEl = td instanceof HTMLElement && slice.contains(td) ? td : null
    const scope = cellEl ?? slice
    const box = new DOMRect(raw.left, top, raw.width, bottom - top)
    candidates.push({
      p: id as ParaId,
      base,
      left: raw.left,
      right: raw.right,
      top,
      bottom,
      cell: cellEl ? cellEl.getBoundingClientRect() : null,
      inkBefore: lineHasInkBeside(img, box, false),
      inkAfter: lineHasInkBeside(img, box, true),
      contentAbove: scopeHasContentVertical(scope, block, img, box, clipRect, false),
      contentBelow: scopeHasContentVertical(scope, block, img, box, clipRect, true),
    })
  }
  return resolveImageHit(candidates, x, y, rawX, rawY)
}

/** True when the picture's line has text or another atom on the asked side of the picture. */
function lineHasInkBeside(img: HTMLElement, box: DOMRect, after: boolean): boolean {
  const parent = img.parentElement
  if (!parent) return false
  for (const child of parent.children) {
    if (child === img || child.classList.contains('marker')) continue
    for (const rect of child.getClientRects()) {
      if (rect.width <= 0 || rect.bottom <= box.top + 0.5 || rect.top >= box.bottom - 0.5) continue
      if (after ? rect.left >= box.right - 0.5 : rect.right <= box.left + 0.5) return true
    }
  }
  return false
}

function scopeHasContentVertical(
  scope: HTMLElement,
  block: HTMLElement,
  img: HTMLElement,
  box: DOMRect,
  clipRect: DOMRect | null,
  below: boolean,
): boolean {
  for (const child of block.children) {
    if (child === img || child.classList.contains('marker')) continue
    for (const rect of child.getClientRects()) {
      const top = clipRect ? Math.max(rect.top, clipRect.top) : rect.top
      const bottom = clipRect ? Math.min(rect.bottom, clipRect.bottom) : rect.bottom
      if ((rect.width <= 0 && bottom - top <= 0) || bottom <= top + 0.5) continue
      if (below ? bottom > box.bottom + 1 : top < box.top - 1) return true
    }
  }
  if (scope !== block) {
    for (const other of scope.children) {
      if (!(other instanceof HTMLElement) || other === block || !other.classList.contains('blk')) continue
      const rect = other.getBoundingClientRect()
      const top = clipRect ? Math.max(rect.top, clipRect.top) : rect.top
      const bottom = clipRect ? Math.min(rect.bottom, clipRect.bottom) : rect.bottom
      if (bottom <= top + 0.5) continue
      if (below ? top >= box.bottom - 1 : bottom <= box.top + 1) return true
    }
  }
  return false
}

function caretFromPoint(x: number, y: number): { node: Node; offset: number } | null {
  const doc = document as Document & {
    caretPositionFromPoint?: (px: number, py: number) => { offsetNode: Node; offset: number } | null
    caretRangeFromPoint?: (px: number, py: number) => Range | null
  }
  const pos = doc.caretPositionFromPoint?.(x, y)
  if (pos) return { node: pos.offsetNode, offset: pos.offset }
  const range = doc.caretRangeFromPoint?.(x, y)
  if (!range) return null
  return { node: range.startContainer, offset: range.startOffset }
}

export interface LayoutPlan {
  readonly remeasure: readonly BlockId[]
  readonly paint: readonly BlockId[]
  /** Ids `paginate` uses to restart. A removal names the surviving neighbor, which keeps its measure. */
  readonly changed: readonly BlockId[]
}

/**
 * What an edit invalidates. A block is measured again when its object is new, when its section's page
 * object changed, or when it now sits in a section whose content width differs. A paragraph whose marker
 * text changed is painted again and not measured: the marker sits outside the line. A marker that appears
 * or disappears on a hanging indent is measured again, because that is when the line-box padding changes.
 */
export function layoutPlan(
  previous: Doc,
  next: Doc,
  previousMarkers: ReadonlyMap<ParaId, string>,
  nextMarkers: ReadonlyMap<ParaId, string>,
): LayoutPlan {
  const remeasure: BlockId[] = []
  const seen = new Set<string>()
  const add = (id: BlockId): void => {
    if (seen.has(id)) return
    seen.add(id)
    remeasure.push(id)
  }
  const delta = bodyDelta(previous.body, next.body)
  if (delta === 'mixed') addMovedBlocks(previous, next, add)
  else if (delta !== 'same') add(delta.id)
  addSectionPageChanges(previous, next, add)
  addMeasureContext(previous, next, add)
  addMarkerBoxes(next, previousMarkers, nextMarkers, add)

  const structural = delta === 'mixed'
    ? structureIds(previous.body, next.body)
    : delta === 'same'
      ? []
      : [delta.id]
  const changed = mergeChanged(structural, remeasure)
  const paint = remeasure.slice()
  const painted = new Set<string>(paint)
  for (const id of markerPaintIds(next.body, previousMarkers, nextMarkers)) {
    if (painted.has(id)) continue
    painted.add(id)
    paint.push(id)
  }
  return { remeasure, paint, changed }
}

function sectionMark(block: Block | undefined): number | null {
  return block?.kind === 'p' ? block.sectionEnd : null
}

/** `'same'` when every block object matches. A `Block` when one slot was replaced and section marks match. */
function bodyDelta(prev: readonly Block[], next: readonly Block[]): 'same' | 'mixed' | Block {
  if (prev.length !== next.length) return 'mixed'
  let changed: Block | null = null
  for (let i = 0; i < next.length; i++) {
    const before = prev[i]
    const after = next[i]
    if (before === after) continue
    // A new section mark moves every later block. One pass cannot see that.
    if (changed || sectionMark(before) !== sectionMark(after) || !after) return 'mixed'
    changed = after
  }
  return changed ?? 'same'
}

function mergeChanged(structural: BlockId[], remeasure: readonly BlockId[]): BlockId[] {
  const seen = new Set<string>(structural)
  for (const id of remeasure) {
    if (seen.has(id)) continue
    seen.add(id)
    structural.push(id)
  }
  return structural
}

// measureBatch sizes the block from these. Page height and the vertical margins do not change the measure.
function sameContentWidth(before: Section | undefined, after: Section | undefined): boolean {
  if (before === after) return true
  const a = before?.page
  const b = after?.page
  if (!a || !b) return false
  if (a === b) return true
  return a.widthPt === b.widthPt
    && a.margin.left === b.margin.left
    && a.margin.right === b.margin.right
    && a.margin.gutter === b.margin.gutter
}

function bodyHasSectionEnd(body: readonly Block[]): boolean {
  for (const block of body) {
    if (block.kind === 'p' && block.sectionEnd !== null) return true
  }
  return false
}

function addMovedBlocks(previous: Doc, next: Doc, add: (id: BlockId) => void): void {
  // Without a section mark, later blocks stay on the same page. Only a new object can need a new measure.
  if (
    !bodyHasSectionEnd(previous.body)
    && !bodyHasSectionEnd(next.body)
    && sameContentWidth(previous.sections[0], next.sections[0])
  ) {
    const prevBlocks = new Set<Block>(previous.body)
    for (const block of next.body) {
      if (!prevBlocks.has(block)) add(block.id)
    }
    return
  }
  const prevSection = sectionIndexes(previous.body)
  const nextSection = sectionIndexes(next.body)
  const prevAt = new Map<Block, number>()
  for (let i = 0; i < previous.body.length; i++) {
    const block = previous.body[i]
    if (block) prevAt.set(block, prevSection[i] ?? 0)
  }
  for (let i = 0; i < next.body.length; i++) {
    const block = next.body[i]
    if (!block) continue
    const beforeAt = prevAt.get(block)
    if (beforeAt === undefined) {
      add(block.id)
      continue
    }
    if (!sameContentWidth(previous.sections[beforeAt], next.sections[nextSection[i] ?? 0])) add(block.id)
  }
}

function addSectionPageChanges(previous: Doc, next: Doc, add: (id: BlockId) => void): void {
  const limit = Math.max(previous.sections.length, next.sections.length)
  for (let i = 0; i < limit; i++) {
    const before = previous.sections[i]
    const after = next.sections[i]
    // A section inserted or dropped is not a page change of the blocks now at that index.
    if (!before || !after || before.page === after.page) continue
    for (const id of sectionMeasureIds(next, i)) add(id)
  }
}

function addMeasureContext(previous: Doc, next: Doc, add: (id: BlockId) => void): void {
  if (previous.sections !== next.sections && documentGridPitch(previous) !== documentGridPitch(next)) {
    for (const block of next.body) add(block.id)
    for (const blocks of next.stories.values()) {
      for (const block of blocks) add(block.id)
    }
    return
  }
  if (previous.lists === next.lists && previous.styles === next.styles) return
  const previousKey = new Map<Block, string>()
  const remember = (doc: Doc, block: Block): void => {
    previousKey.set(block, blockMeasureKey(doc, block))
  }
  for (const block of previous.body) remember(previous, block)
  for (const blocks of previous.stories.values()) {
    for (const block of blocks) remember(previous, block)
  }
  const consider = (doc: Doc, block: Block): void => {
    const before = previousKey.get(block)
    if (before === undefined || before === blockMeasureKey(doc, block)) return
    add(block.id)
  }
  for (const block of next.body) consider(next, block)
  for (const blocks of next.stories.values()) {
    for (const block of blocks) consider(next, block)
  }
}

function blockMeasureKey(doc: Doc, block: Block): string {
  if (block.kind === 'p') return resolvedMeasureKey(doc, block)
  if (block.kind === 'opaque') {
    let key = ''
    for (const para of block.preview) key += `${resolvedMeasureKey(doc, para)}\n`
    return key
  }
  let key = ''
  for (const row of block.rows) {
    for (const cell of row.cells) {
      for (const child of cell.blocks) key += `${blockMeasureKey(doc, child)}\n`
    }
  }
  return key
}

// A marker is absolutely positioned, so a new string does not change the line box. Its presence does:
// with a positive hanging indent the padding moves from left + hanging to left.
function addMarkerBoxes(
  doc: Doc,
  previousMarkers: ReadonlyMap<ParaId, string>,
  nextMarkers: ReadonlyMap<ParaId, string>,
  add: (id: BlockId) => void,
): void {
  const seen = new Set<string>()
  const consider = (id: ParaId): void => {
    if (seen.has(id)) return
    seen.add(id)
    const had = Boolean(previousMarkers.get(id))
    const has = Boolean(nextMarkers.get(id))
    if (had === has) return
    const found = ownedParagraph(doc, id)
    if (!found) return
    add(found.owner)
  }
  for (const id of previousMarkers.keys()) consider(id)
  for (const id of nextMarkers.keys()) consider(id)
}

function ownedParagraph(doc: Doc, id: string): { owner: BlockId, para: Paragraph } | null {
  const inBlocks = (blocks: readonly Block[]): { owner: BlockId, para: Paragraph } | null => {
    for (const block of blocks) {
      const para = paragraphInside(block, id)
      if (para) return { owner: block.id, para }
    }
    return null
  }
  const body = inBlocks(doc.body)
  if (body) return body
  for (const blocks of doc.stories.values()) {
    const found = inBlocks(blocks)
    if (found) return found
  }
  return null
}

function paragraphInside(block: Block, id: string): Paragraph | null {
  if (block.kind === 'p') return block.id === id ? block : null
  if (block.kind === 'opaque') {
    for (const para of block.preview) if (para.id === id) return para
    return null
  }
  for (const row of block.rows) {
    for (const cell of row.cells) {
      for (const child of cell.blocks) {
        const para = paragraphInside(child, id)
        if (para) return para
      }
    }
  }
  return null
}

function resolvedHanging(doc: Doc, para: Paragraph): number {
  const style = doc.styles.paragraph.get(para.props.style ?? doc.styles.defaultParagraph)?.look
  const props = para.props
  const numbering = props.numbering === undefined ? (style?.numbering ?? null) : props.numbering
  const level = numbering ? doc.lists.byNum.get(numbering.num)?.levels[numbering.level] : undefined
  return props.indent?.hanging ?? level?.hanging ?? style?.indent.hanging ?? 0
}

function lineSpacingKey(line: LineSpacing | undefined): string {
  if (!line) return ''
  if (line.rule === 'auto') return `auto:${line.multiple}`
  return `${line.rule}:${line.pt}`
}

// List level indent and the paragraph style are read at measure time. The paragraph object does not carry them.
function resolvedMeasureKey(doc: Doc, para: Paragraph): string {
  const style = doc.styles.paragraph.get(para.props.style ?? doc.styles.defaultParagraph)?.look
  const props = para.props
  const numbering = props.numbering === undefined ? (style?.numbering ?? null) : props.numbering
  const level = numbering ? doc.lists.byNum.get(numbering.num)?.levels[numbering.level] : undefined
  const left = props.indent?.left ?? level?.indentLeft ?? style?.indent.left
  const hanging = resolvedHanging(doc, para)
  const right = props.indent?.right ?? style?.indent.right
  const first = props.indent?.firstLine ?? style?.indent.firstLine
  const chars = props.indent?.firstLineChars ?? style?.indent.firstLineChars
  const before = props.spacing?.before ?? style?.spacing.before
  const after = props.spacing?.after ?? style?.spacing.after
  const line = lineSpacingKey(props.spacing?.line ?? style?.spacing.line)
  const keepNext = props.keepNext ?? style?.keepNext
  const keepLines = props.keepLines ?? style?.keepLines
  const pageBreakBefore = props.pageBreakBefore ?? style?.pageBreakBefore
  const widow = props.widowControl ?? style?.widowControl
  return `${left},${right},${first},${hanging},${chars}|${before},${after},${line}|${keepNext}${keepLines}${pageBreakBefore}${widow}`
}

function structureIds(prev: readonly Block[], next: readonly Block[]): BlockId[] {
  let prefix = 0
  const shared = Math.min(prev.length, next.length)
  while (prefix < shared && prev[prefix] === next[prefix]) prefix += 1
  let prevEnd = prev.length - 1
  let nextEnd = next.length - 1
  while (prevEnd >= prefix && nextEnd >= prefix && prev[prevEnd] === next[nextEnd]) {
    prevEnd -= 1
    nextEnd -= 1
  }
  const ids: BlockId[] = []
  const seen = new Set<string>()
  const push = (block: Block | undefined): void => {
    if (!block || seen.has(block.id)) return
    seen.add(block.id)
    ids.push(block.id)
  }
  if (nextEnd >= prefix) {
    for (let i = prefix; i <= nextEnd; i++) push(next[i])
    return ids
  }
  if (prefix === prev.length && prefix === next.length) return ids
  // A pure removal has no new block. Name the survivor at the hole, or the new last block when
  // the tail was removed, so pagination does not reuse the page that still shows the removed block.
  if (prefix < next.length) push(next[prefix])
  else if (prefix > 0) push(next[prefix - 1])
  return ids
}

function sectionMeasureIds(doc: Doc, sectionIndex: number): BlockId[] {
  const indexes = sectionIndexes(doc.body)
  const ids: BlockId[] = []
  for (let i = 0; i < doc.body.length; i++) {
    if ((indexes[i] ?? 0) !== sectionIndex) continue
    const block = doc.body[i]
    if (block && !ids.includes(block.id)) ids.push(block.id)
  }
  const section = doc.sections[sectionIndex]
  if (!section) return ids
  const keys = [
    section.headers.default,
    section.headers.first,
    section.headers.even,
    section.footers.default,
    section.footers.first,
    section.footers.even,
  ]
  for (const key of keys) {
    if (!key) continue
    const blocks = doc.stories.get(key)
    if (!blocks) continue
    for (const block of blocks) {
      if (!ids.includes(block.id)) ids.push(block.id)
    }
  }
  return ids
}

function markerPaintIds(
  body: readonly Block[],
  previousMarkers: ReadonlyMap<ParaId, string>,
  nextMarkers: ReadonlyMap<ParaId, string>,
): BlockId[] {
  const ids: BlockId[] = []
  const seen = new Set<string>()
  const consider = (id: ParaId): void => {
    if (previousMarkers.get(id) === nextMarkers.get(id)) return
    const owner = topBlockId(body, id)
    if (!owner || seen.has(owner)) return
    seen.add(owner)
    ids.push(owner)
  }
  for (const id of previousMarkers.keys()) consider(id)
  for (const id of nextMarkers.keys()) consider(id)
  return ids
}

function topBlockId(body: readonly Block[], id: string): BlockId | null {
  for (const block of body) {
    if (containsId(block, id)) return block.id
  }
  return null
}

function extractHeader(blockEl: HTMLElement, rows: number): HTMLTableElement | null {
  const source = blockEl.querySelector(':scope > table')
  if (!(source instanceof HTMLTableElement)) return null
  const copy = document.createElement('table')
  copy.className = 'repeat-head'
  copy.style.width = source.style.width
  const cols = source.querySelector('colgroup')
  if (cols) copy.append(cols.cloneNode(true))
  const body = document.createElement('tbody')
  for (let i = 0; i < rows && i < source.rows.length; i++) {
    const row = source.rows[i]
    if (row) body.append(row.cloneNode(true))
  }
  copy.append(body)
  return copy
}

function caretClientTop(el: HTMLElement, offset: number): number {
  const point = domPoint(el, offset)
  const range = scratchRange()
  range.setStart(point.node, point.offset)
  range.collapse(true)
  const rect = range.getClientRects()[0]
  return rect ? rect.top : el.getBoundingClientRect().bottom
}

function startsFromElement(el: HTMLElement, bottoms: Float32Array, length: number): Uint32Array {
  const starts = new Uint32Array(bottoms.length)
  const origin = el.getBoundingClientRect().top
  let prev = 0
  for (let line = 1; line < bottoms.length; line++) {
    const target = origin + (bottoms[line - 1] ?? 0)
    let lo = prev
    let hi = length
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (caretClientTop(el, mid) < target - 0.75) lo = mid + 1
      else hi = mid
    }
    starts[line] = lo
    prev = lo
  }
  return starts
}

/** Top of one block, in CSS pixels from the top of the scroller. */
interface ViewportHold {
  readonly id: string
  readonly offset: number
}

interface PreeditMetric {
  readonly deltaTop: number
  readonly height: number
}

function sampleCharAt(text: string, offset: number): string {
  if (text.length === 0) return '\u200b'
  if (offset > 0) {
    const chars = Array.from(text.slice(0, offset))
    return chars[chars.length - 1] ?? '\u200b'
  }
  return Array.from(text)[0] ?? '\u200b'
}

/** The runtime's inline composition, appended to `document.body` while an IME session is open. */
function hostPreeditElement(): HTMLElement | null {
  const children = document.body?.children
  if (!children) return null
  for (const node of children) {
    if (!(node instanceof HTMLElement) || node.tagName !== 'SPAN') continue
    const style = node.getAttribute('style') ?? ''
    if (node.dataset.hostPreedit !== '1' && !isHostPreeditStyle(style)) continue
    if ((node.textContent ?? '').length === 0 || node.style.display === 'none' || node.style.visibility === 'hidden') continue
    node.dataset.hostPreedit = '1'
    return node
  }
  return null
}

/**
 * Owns the pages container. Mounts the visible pages plus one on each side (IntersectionObserver); every other page
 * is an empty box of its exact size, so scroll geometry is exact without contents. All coordinates it returns are
 * page-container px at 100 percent zoom; zoom is a transform on the container, and this class divides it out.
 */
function sameHits(prev: readonly TextMatch[], prevCurrent: number, next: readonly TextMatch[], nextCurrent: number): boolean {
  if (prevCurrent !== nextCurrent || prev.length !== next.length) return false
  for (let i = 0; i < prev.length; i++) {
    const a = prev[i]
    const b = next[i]
    if (!a || !b || a.para !== b.para || a.start !== b.start || a.end !== b.end) return false
  }
  return true
}

export class PageView {
  private doc: Doc | null = null
  private ctx: () => RenderContext
  private host: HTMLElement
  private sizer: HTMLElement | null = null
  private stack: HTMLElement | null = null
  private measureHost: HTMLElement | null = null
  private observer: IntersectionObserver | null = null
  private pageEls: HTMLElement[] = []
  private measures = new WeakMap<Block, Measure>()
  private lineStartCache = new WeakMap<Block, Uint32Array>()
  private _layout: Layout = { pages: [], complete: false, sectionKeys: [] }
  private zoom = 1
  private cursor = 0
  private chunk = 0
  private dead = false
  private intersecting = new Set<number>()
  private mountedSig = new Map<number, string>()
  private syncQueued = false
  private watch = 0
  private seenScroll = -1
  private seenHeight = -1
  private indexMap: Map<string, Place> | null = null
  private indexDoc: Doc | null = null
  private changedIds: Set<BlockId> | null = null
  private deco: Decoration | null = null
  private decoMeasure: Measure | null = null
  private decoOwner: Block | null = null
  private composing = false
  private preeditMetrics = new Map<string, PreeditMetric>()
  private sink: HTMLTextAreaElement | null = null
  private lastPaint: { sel: Selection; sink: HTMLTextAreaElement; stored: Marks | null } | null = null
  private imagePreview: ImageGesture | null = null
  private findHits: readonly TextMatch[] = []
  private findCurrent = -1
  /** Set while a metrics pass remeasures in place. Pages above the hold are still estimates until it finishes. */
  private rewriting = false
  private hold: ViewportHold | null = null
  /** scrollTop captured with the hold, so a short intermediate layout cannot clamp it away. */
  private holdScroll = 0
  private readonly onLayout: () => void

  /** `onLayout` runs after each measured chunk, so a page count shown outside the view follows the layout. */
  constructor(host: HTMLElement, ctx: () => RenderContext, onLayout: () => void = () => {}) {
    this.host = host
    this.ctx = ctx
    this.onLayout = onLayout
  }

  get layout(): Layout {
    if (this.rewriting && this._layout.complete) {
      return { pages: this._layout.pages, complete: false, sectionKeys: this._layout.sectionKeys }
    }
    return this._layout
  }

  /** Full rebuild: measure every block in chunks of about 200 per task, mount pages as they fill. */
  show(doc: Doc): void {
    this.cancelChunk()
    this.finishRewrite()
    this.doc = doc
    this.measures = new WeakMap()
    this.lineStartCache = new WeakMap()
    this.preeditMetrics.clear()
    this.indexDoc = null
    this.indexMap = null
    this.deco = null
    this.decoMeasure = null
    this.decoOwner = null
    this.composing = false
    this._layout = { pages: [], complete: false, sectionKeys: [] }
    this.cursor = 0
    this.dead = false
    this.ensureChrome()
    this.resetPages()
    this.measureStories()
    this.pump()
  }

  /**
   * Measures again after font metrics arrive for an edit. The block that contains the selection head stays at the
   * same offset from the top of the scroller when that head was visible; otherwise the block at the top of the
   * viewport does. Each chunk puts that block back, because the pages above it are still the previous estimate.
   */
  relayout(doc: Doc): void {
    if (!this.doc || !this.stack?.isConnected || this.pageEls.length === 0) {
      this.show(doc)
      return
    }
    this.cancelChunk()
    this.hold = this.captureHold()
    this.holdScroll = this.host.scrollTop
    this.doc = doc
    this.lineStartCache = new WeakMap()
    this.preeditMetrics.clear()
    this.indexDoc = null
    this.indexMap = null
    this.deco = null
    this.decoMeasure = null
    this.decoOwner = null
    this.cursor = 0
    this.dead = false
    this.rewriting = true
    this.host.style.overflowAnchor = 'none'
    this.ensureChrome()
    this.measureStories()
    this.pump()
  }

  /**
   * After a transaction. Remeasures a block whose object is new, a section whose page object changed, and a block
   * that moved into a section with a different content width. A removal restarts pagination at the surviving
   * neighbor and keeps that neighbor's measure. A paragraph whose list marker changed is painted again without
   * a new measure. Unmounted pages only get new slice lists.
   */
  update(
    previous: Doc,
    next: Doc,
    previousMarkers: ReadonlyMap<ParaId, string>,
    nextMarkers: ReadonlyMap<ParaId, string>,
  ): void {
    this.cancelChunk()
    this.finishRewrite()
    const plan = layoutPlan(previous, next, previousMarkers, nextMarkers)
    this.doc = next
    this.indexDoc = null
    this.indexMap = null
    this.deco = null
    this.decoMeasure = null
    this.decoOwner = null
    for (const id of plan.remeasure) {
      const located = this.locateOwner(id)
      if (!located) continue
      this.measures.delete(located.block)
      this.lineStartCache.delete(located.block)
      this.measureOne(located.block, located.section)
    }
    this._layout = paginate(next, (block) => this.measures.get(block), this._layout, plan.changed)
    this.changedIds = new Set(plan.paint)
    this.syncPages()
    this.changedIds = null
    if (this.lastPaint) this.paintNow(this.lastPaint.sel, this.lastPaint.sink)
  }

  /** IME preedit. Re-renders and re-measures one block without touching the measure cache. */
  decorate(deco: Decoration | null): void {
    this.composing = deco !== null
    if (!this.doc) return
    const previousId = this.deco?.pos.p
    this.deco = deco
    const id = deco?.pos.p ?? previousId
    if (!id) return
    const located = this.locateOwner(id)
    if (!located) return
    if (deco) {
      const el = renderBlock(located.block, this.ctxFor(1), deco)
      el.style.width = `${contentWidthOf(located.section)}px`
      this.measureHost?.replaceChildren(el)
      this.decoMeasure = measureElement(el, 1)
      this.decoOwner = located.block
      this.measureHost?.replaceChildren()
    } else {
      this.decoMeasure = null
      this.decoOwner = null
    }
    this._layout = paginate(this.doc, (block) => this.measureFor(block), this._layout, [located.block.id])
    this.changedIds = new Set([located.block.id])
    this.syncPages()
    this.changedIds = null
    if (this.lastPaint) this.paintNow(this.lastPaint.sel, this.lastPaint.sink)
  }

  caretRect(pos: Pos): CaretRect {
    const pageCount = this._layout.pages.length
    const pageIndex = pageCount === 0 ? 0 : Math.min(this.pageOf(pos), pageCount - 1)
    this.mountPage(pageIndex)
    const pageEl = this.pageEls[pageIndex]
    if (!pageEl) return { page: pageIndex, x: 0, top: 0, height: 16 }
    const paraEl = this.paraElement(pageEl, pos)
    const origin = pageEl.getBoundingClientRect()
    const z = this.zoom
    if (!paraEl) return { page: pageIndex, x: 0, top: 0, height: 16 }
    const preedit = this.preeditEnd(paraEl, pos)
    if (preedit) {
      return {
        page: pageIndex,
        x: (preedit.right - origin.left) / z,
        top: (preedit.top - origin.top) / z,
        height: Math.max(preedit.height / z, 1),
      }
    }
    const slice = paraEl.closest('.slice')
    const sliceRect = slice instanceof HTMLElement ? slice.getBoundingClientRect() : null
    const point = domPoint(paraEl, pos.o)
    const atomCaret = inlineAtomCaretViewport(point, pos.o, sliceRect, defaultCaretHeightPx(paraEl) * z)
    if (atomCaret) {
      const rawX = (atomCaret.left - origin.left) / z
      const maxX = pageEl.offsetWidth > 1 ? pageEl.offsetWidth - 1 : rawX
      return {
        page: pageIndex,
        x: Math.min(rawX, maxX),
        top: (atomCaret.top - origin.top) / z,
        height: Math.max(atomCaret.height / z, 1),
      }
    }
    const range = scratchRange()
    range.setStart(point.node, point.offset)
    range.setEnd(point.node, point.offset)
    const rects = [...range.getClientRects()].filter((rect) => rect.height > 0)
    const rect =
      rects.find((item) => !sliceRect || (item.bottom > sliceRect.top + 0.5 && item.top < sliceRect.bottom - 0.5)) ??
      rects[0]
    if (!rect) {
      const box = paraEl.getBoundingClientRect()
      const padLeft = parseFloat(paraEl.style.paddingLeft) || 0
      const padTop = parseFloat(paraEl.style.paddingTop) || 0
      const indent = parseFloat(paraEl.style.textIndent) || 0
      return {
        page: pageIndex,
        x: (box.left - origin.left) / z + Math.max(0, padLeft + indent),
        top: (box.top - origin.top) / z + padTop,
        height: defaultCaretHeightPx(paraEl),
      }
    }
    return {
      page: pageIndex,
      x: (rect.left - origin.left) / z,
      top: (rect.top - origin.top) / z,
      height: Math.max(rect.height / z, 1),
    }
  }

  /** The model caret stays where the composition started. The IME window anchors at the end of the preedit. */
  private preeditEnd(paraEl: HTMLElement, pos: Pos): DOMRect | null {
    const deco = this.deco
    if (!deco || deco.text.length === 0 || deco.pos.p !== pos.p || deco.pos.o !== pos.o) return null
    const span = paraEl.querySelector('.preedit')
    if (!(span instanceof HTMLElement)) return null
    const rects = [...span.getClientRects()].filter((rect) => rect.width > 0 || rect.height > 0)
    return rects[rects.length - 1] ?? null
  }

  /**
   * Click and drag. Clamps (x, y) into the nearest slice's visible box on that page, then calls
   * document.caretPositionFromPoint and maps the DOM point with render.posFromDom.
   */
  posAtPoint(clientX: number, clientY: number): Pos | null {
    return this.withHits(() => {
      const pageIndex = this.pageIndexAt(clientX, clientY)
      if (pageIndex < 0) return null
      this.mountPage(pageIndex)
      const pageEl = this.pageEls[pageIndex]
      if (!pageEl) return null
      const slices = [...pageEl.querySelectorAll<HTMLElement>('.slice')]
      if (slices.length === 0) return null
      let best: { slice: HTMLElement; rect: DOMRect; dist: number } | null = null
      for (const slice of slices) {
        const rect = slice.getBoundingClientRect()
        const dx = clientX < rect.left ? rect.left - clientX : clientX > rect.right ? clientX - rect.right : 0
        const dy = clientY < rect.top ? rect.top - clientY : clientY > rect.bottom ? clientY - rect.bottom : 0
        const dist = dx * dx + dy * dy
        if (!best || dist < best.dist) best = { slice, rect, dist }
      }
      if (!best) return null
      let x = clientX
      let y = clientY
      if (best.dist > 0) {
        x = Math.min(Math.max(clientX, best.rect.left + 1), Math.max(best.rect.left + 1, best.rect.right - 1))
        y = Math.min(Math.max(clientY, best.rect.top + 1), Math.max(best.rect.top + 1, best.rect.bottom - 1))
      }
      const beside = imageSideAt(best.slice, x, y, clientX, clientY)
      if (beside) return beside
      const hidden: HTMLElement[] = []
      try {
        // acdert hit testing ignores the clip, so another page's shifted copy can take the hit.
        for (let i = 0; i < this.pageEls.length; i++) {
          if (i === pageIndex) continue
          const other = this.pageEls[i]
          if (!other || other.dataset.mounted !== '1') continue
          const body = other.querySelector('.page-body')
          if (body instanceof HTMLElement && body.style.visibility !== 'hidden') {
            body.style.visibility = 'hidden'
            hidden.push(body)
          }
        }
        const probe = (px: number, py: number): { node: Node; offset: number } | null => {
          const next = caretFromPoint(px, py)
          if (!next) return null
          const el = next.node instanceof Element ? next.node : next.node.parentElement
          // A repeated header sits in the slice, above the clip. It is not the shifted copy.
          if (el?.closest('.repeat-head')) return next
          const slice = el?.closest('.slice')
          if (!(slice instanceof HTMLElement) || !pageEl.contains(slice)) return next
          const clip = slice.querySelector(':scope > .clip')
          if (!(clip instanceof HTMLElement)) return next
          const box = clip.getBoundingClientRect()
          return py < box.top || py > box.bottom ? null : next
        }
        const hit = probe(x, y) ?? probe(x, y + 1) ?? probe(x + 0.5, y)
        if (!hit) return this.sliceParaFallback(best.slice, x, y)
        const pos = posFromDom(hit.node, hit.offset)
        if (!pos) return null
        return this.adjustAtom(hit.node, pos, x)
      } finally {
        for (const el of hidden) el.style.visibility = ''
      }
    })
  }

  private sliceParaFallback(slice: HTMLElement, x: number, y: number): Pos | null {
    const clip = slice.querySelector(':scope > .clip')
    const clipRect = clip instanceof HTMLElement ? clip.getBoundingClientRect() : null
    if (clipRect && (y < clipRect.top - 1 || y > clipRect.bottom + 1)) return null
    let best: { el: HTMLElement; rect: DOMRect; dist: number } | null = null
    for (const paraEl of slice.querySelectorAll<HTMLElement>('.blk.p')) {
      if (paraEl.closest('.repeat-head')) continue
      const td = paraEl.closest('td')
      if (td instanceof HTMLElement && slice.contains(td)) {
        const cellRect = td.getBoundingClientRect()
        if (x < cellRect.left || x > cellRect.right || y < cellRect.top || y > cellRect.bottom) continue
      }
      const raw = paraEl.getBoundingClientRect()
      const top = clipRect ? Math.max(raw.top, clipRect.top) : raw.top
      const bottom = clipRect ? Math.min(raw.bottom, clipRect.bottom) : raw.bottom
      if (bottom <= top + 0.5) continue
      const dy = y < top ? top - y : y > bottom ? y - bottom : 0
      if (!best || dy < best.dist) best = { el: paraEl, rect: raw, dist: dy }
    }
    if (!best) return null
    const id = best.el.dataset.id
    if (!id) return null
    const place = this.index().get(id)
    if (!place) return { p: id as ParaId, o: 0 }
    const length = flatText(place.para).length
    if (length === 0) return { p: place.para.id, o: 0 }
    const padLeft = (parseFloat(best.el.style.paddingLeft) || 0) * this.zoom
    return { p: place.para.id, o: x <= best.rect.left + padLeft + 2 ? 0 : length }
  }

  /**
   * Up and Down. Finds the next line box below or above the caret's (crossing slices and pages), then hit-tests at
   * goalX on that line's middle. Returns the goal to keep for the next vertical move.
   */
  lineMove(pos: Pos, direction: 'up' | 'down', goalX: number | null): { pos: Pos; goalX: number } {
    const caret = this.caretRect(pos)
    const goal = goalX ?? caret.x
    const neighbor = direction === 'down' ? caret.page + 1 : caret.page - 1
    this.mountPage(caret.page)
    if (neighbor >= 0 && neighbor < this._layout.pages.length) this.mountPage(neighbor)
    const lines = this.collectLines(neighbor >= 0 && neighbor < this._layout.pages.length ? [caret.page, neighbor] : [caret.page])
    const current = this.pickCurrent(lines, caret.page, caret.x, caret.top + caret.height / 2)
    if (!current) return { pos, goalX: goal }
    const next = nextVerticalLine(lines, current, direction, goal)
    if (!next) return { pos, goalX: goal }
    return { pos: this.posOnLine(next, goal) ?? pos, goalX: goal }
  }

  /** PageUp and PageDown: scrolls one viewport and hit-tests at the same viewport offset and goalX. */
  pageMove(pos: Pos, direction: 'up' | 'down', goalX: number | null): { pos: Pos; goalX: number } {
    const caret = this.caretRect(pos)
    const goal = goalX ?? caret.x
    const pageEl = this.pageEls[caret.page]
    if (!pageEl) return { pos, goalX: goal }
    const host = this.host.getBoundingClientRect()
    const origin = pageEl.getBoundingClientRect()
    const offset = origin.top + (caret.top + caret.height / 2) * this.zoom - host.top
    const clientX = origin.left + goal * this.zoom
    this.host.scrollTop += this.host.clientHeight * (direction === 'down' ? 1 : -1)
    const y = this.host.getBoundingClientRect().top + offset
    const pageIndex = this.pageIndexAt(clientX, y)
    if (pageIndex >= 0) this.mountPage(pageIndex)
    return { pos: this.posAtPoint(clientX, y) ?? pos, goalX: goal }
  }

  /** Home and End: the first or last offset on the caret's visual line. */
  lineEdge(pos: Pos, edge: 'start' | 'end'): Pos {
    const caret = this.caretRect(pos)
    const line = this.pickCurrent(this.collectLines([caret.page]), caret.page, caret.x, caret.top + caret.height / 2)
    if (!line) return pos
    const place = this.index().get(line.paraId)
    if (!place) return pos
    const length = flatText(place.para).length
    const mid = (line.top + line.bottom) / 2
    const probe = this.posAtLocal(line.page, edge === 'start' ? line.left + 1 : Math.max(line.left + 1, line.right - 1), mid)
    if (!probe || probe.p !== place.para.id) return pos
    const on = (offset: number): boolean => {
      const next = this.caretRect({ p: place.para.id, o: offset })
      const y = next.top + next.height / 2
      return next.page === line.page && y >= line.top - 3 && y <= line.bottom + 3
    }
    let o = probe.o
    let steps = 0
    if (edge === 'start') {
      while (o > 0 && steps < 200 && on(o - 1)) {
        o -= 1
        steps += 1
      }
    } else {
      while (o < length && steps < 200 && on(o + 1)) {
        o += 1
        steps += 1
      }
    }
    return { p: place.para.id, o }
  }

  /** Draws the caret div and selection rects in per-page overlays, clipped to slice boxes. Moves the input sink. */
  paint(sel: Selection, sink: HTMLTextAreaElement, stored: Marks | null = null): void {
    this.sink = sink
    this.lastPaint = { sel, sink, stored }
    this.paintNow(sel, sink)
  }

  /** Tints on pages that are already mounted. Passing the same list does not paint again. */
  setFindHits(hits: readonly TextMatch[], current: number): void {
    if (sameHits(this.findHits, this.findCurrent, hits, current)) return
    this.findHits = hits
    this.findCurrent = current
    if (this.lastPaint) this.paintNow(this.lastPaint.sel, this.lastPaint.sink)
  }

  scrollIntoView(pos: Pos): void {
    const caret = this.caretRect(pos)
    const pageEl = this.pageEls[caret.page]
    if (!pageEl) return
    const origin = pageEl.getBoundingClientRect()
    const host = this.host.getBoundingClientRect()
    const top = origin.top + caret.top * this.zoom
    const bottom = top + caret.height * this.zoom
    if (top < host.top) this.host.scrollTop -= host.top - top + 8
    else if (bottom > host.bottom) this.host.scrollTop += bottom - host.bottom + 8
  }

  /** Changes the container transform only. Measures and pagination stay valid. The held block keeps its offset. */
  setZoom(zoom: number): void {
    if (!(zoom > 0)) return
    const hold = zoom === this.zoom ? null : this.captureHold()
    const suppressAnchor = hold !== null && this.host.style.overflowAnchor !== 'none'
    if (suppressAnchor) this.host.style.overflowAnchor = 'none'
    this.zoom = zoom
    this.updateSizer()
    this.restoreHold(hold)
    if (suppressAnchor) this.clearOverflowAnchor()
    this.syncFromGeometry()
    if (this.lastPaint) this.paintNow(this.lastPaint.sel, this.lastPaint.sink)
  }

  currentZoom(): number {
    return this.zoom
  }

  /** The picture's border box in page pixels, or null when that page is not mounted. */
  imageBox(pos: Pos): PageBox | null {
    return this.imageFrame(pos)?.box ?? null
  }

  setImagePreview(preview: ImageGesture | null): void {
    this.imagePreview = preview
    if (this.lastPaint) this.paintNow(this.lastPaint.sel, this.lastPaint.sink)
  }

  /**
   * The selection head's block when that head is inside the scroller; otherwise the block at the top of the viewport.
   * Zoom and a metrics relayout both read this, then put the same top back after the heights change.
   */
  private captureHold(): ViewportHold | null {
    const host = this.host.getBoundingClientRect()
    if (host.height < 1) return null
    const sel = this.lastPaint?.sel
    if (sel?.kind === 'text') {
      const held = this.holdCaret(sel.head, host)
      if (held) return held
    }
    return this.blockAtViewportTop(host)
  }

  private holdCaret(pos: Pos, host: DOMRect): ViewportHold | null {
    if (!this.index().has(pos.p) || this._layout.pages.length === 0) return null
    const caret = this.caretRect(pos)
    const pageEl = this.pageEls[caret.page]
    if (!pageEl) return null
    const origin = pageEl.getBoundingClientRect()
    const top = origin.top + caret.top * this.zoom
    const bottom = top + Math.max(caret.height, 1) * this.zoom
    if (bottom <= host.top + 0.5 || top >= host.bottom - 0.5) return null
    const paraEl = this.paraElement(pageEl, pos)
    if (!paraEl) return null
    return { id: pos.p, offset: paraEl.getBoundingClientRect().top - host.top }
  }

  private blockAtViewportTop(host: DOMRect): ViewportHold | null {
    const edge = host.top + 1
    let hit: { id: string; top: number; span: number } | null = null
    let nearest: { id: string; top: number; distance: number } | null = null
    for (const pageEl of this.pageEls) {
      if (pageEl.dataset.mounted !== '1') continue
      const pageBox = pageEl.getBoundingClientRect()
      if (pageBox.bottom <= host.top + 0.5 || pageBox.top >= host.bottom - 0.5) continue
      for (const el of pageEl.querySelectorAll<HTMLElement>('.blk')) {
        if (el.closest('.repeat-head') || el.getAttribute('aria-hidden') === 'true') continue
        const id = el.dataset.id
        if (!id) continue
        const clip = el.closest('.clip')
        const visible = clip instanceof HTMLElement ? clip.getBoundingClientRect() : el.getBoundingClientRect()
        if (visible.bottom <= host.top + 0.5 || visible.top >= host.bottom - 0.5) continue
        const top = el.getBoundingClientRect().top
        if (visible.top <= edge && visible.bottom > edge) {
          const span = visible.height
          if (!hit || span < hit.span) hit = { id, top, span }
          continue
        }
        if (visible.top > edge) {
          const distance = visible.top - host.top
          if (!nearest || distance < nearest.distance) nearest = { id, top, distance }
        }
      }
    }
    const found = hit ?? nearest
    if (!found) return null
    return { id: found.id, offset: found.top - host.top }
  }

  /** Moves scrollTop so the held block's top matches the captured offset. Returns false when nothing moved. */
  private restoreHold(hold: ViewportHold | null): boolean {
    if (!hold || !this.doc || this.dead) return false
    const top = this.blockClientTop(hold.id, hold.offset)
    if (top === null) return false
    const delta = top - this.host.getBoundingClientRect().top - hold.offset
    if (Math.abs(delta) < 0.5) return false
    this.host.scrollTop += delta
    this.seenScroll = this.host.scrollTop
    return true
  }

  private blockClientTop(id: string, offset: number): number | null {
    const owner = this.locateOwner(id)
    if (owner) {
      for (const spot of this.spots(owner.block.id)) this.mountPage(spot.page)
    }
    const desired = this.host.getBoundingClientRect().top + offset
    let best: number | null = null
    let bestDist = Infinity
    for (const pageEl of this.pageEls) {
      if (pageEl.dataset.mounted !== '1') continue
      for (const el of blockElements(pageEl, id)) {
        if (el.closest('.repeat-head')) continue
        const rect = el.getBoundingClientRect()
        if (rect.width < 0.5 && rect.height < 0.5) continue
        const dist = Math.abs(rect.top - desired)
        if (dist < bestDist) {
          bestDist = dist
          best = rect.top
        }
      }
    }
    return best
  }

  private finishRewrite(): void {
    this.rewriting = false
    this.hold = null
    this.holdScroll = 0
    this.clearOverflowAnchor()
  }

  private clearOverflowAnchor(): void {
    if (this.host.style.overflowAnchor === 'none') this.host.style.overflowAnchor = ''
  }

  /**
   * Height in points of each body block, from the measure cache the paginator reads.
   * Does not measure. A block the paginator has not reached is null. Blocks already
   * measured into the hidden host are present even when their page is not mounted.
   */
  blockHeights(): readonly { index: number; kind: 'p' | 'table' | 'opaque'; heightPt: number | null }[] {
    const doc = this.doc
    if (!doc) return []
    return doc.body.map((block, index) => {
      const heightPx = this.measures.get(block)?.heightPx
      return {
        index,
        kind: block.kind,
        heightPt: heightPx !== undefined && Number.isFinite(heightPx) ? heightPx / PT_TO_PX : null,
      }
    })
  }

  /**
   * Page index (0-based) holding a position. A paragraph uses its measured lines.
   * A cell paragraph is not in that cache, so a row that continues onto another page
   * reads one mounted copy and picks the slice that contains the caret.
   */
  pageOf(pos: Pos): number {
    const place = this.index().get(pos.p)
    if (!place) return 0
    const spots = this.spots(place.owner.id)
    const first = spots[0]
    if (!first) return 0
    if (spots.length === 1) return first.page
    const y = this.sliceY(place, pos.o)
    const page = pageAtSliceY(
      spots.map((spot) => ({ page: spot.page, y0: spot.slice.y0, y1: spot.slice.y1 })),
      y,
    )
    return page ?? spots[spots.length - 1]?.page ?? 0
  }

  destroy(): void {
    this.dead = true
    this.cancelChunk()
    this.finishRewrite()
    this.observer?.disconnect()
    this.observer = null
    if (this.watch !== 0) cancelAnimationFrame(this.watch)
    this.watch = 0
    this.pageEls = []
    this.mountedSig.clear()
    this.intersecting.clear()
    this.host.replaceChildren()
    this.host.classList.remove('page-view')
    this.doc = null
  }

  private measureFor(block: Block): Measure | undefined {
    if (this.decoOwner === block && this.decoMeasure) return this.decoMeasure
    return this.measures.get(block)
  }

  private pump(): void {
    if (this.dead || !this.doc) return
    const body = this.doc.body
    const end = Math.min(body.length, this.cursor + 200)
    const sections = sectionIndexes(body)
    const batch: { block: Block; section: Section }[] = []
    for (let i = this.cursor; i < end; i++) {
      const block = body[i]
      const section = this.doc.sections[sections[i] ?? 0] ?? this.doc.sections[0]
      if (block && section) batch.push({ block, section })
    }
    this.measureBatch(batch)
    const changed = batch.map((item) => item.block.id)
    const previous = this._layout.pages.length > 0 ? this._layout : null
    this._layout = paginate(this.doc, (block) => this.measures.get(block), previous, previous ? changed : [])
    this.cursor = end
    if (this.rewriting) this.changedIds = new Set(changed)
    this.syncPages()
    this.changedIds = null
    if (this.restoreHold(this.hold)) {
      this.syncFromGeometry()
      this.holdScroll = this.host.scrollTop
    }
    if (this.rewriting && this.lastPaint) this.paintNow(this.lastPaint.sel, this.lastPaint.sink)
    if (this.cursor < body.length) {
      this.chunk = requestAnimationFrame(() => this.pump())
      this.onLayout()
      return
    }
    if (!this.rewriting) {
      this.onLayout()
      return
    }
    const hold = this.hold
    this.rewriting = false
    this.hold = null
    this.holdScroll = 0
    this.syncPages()
    if (this.restoreHold(hold)) this.syncFromGeometry()
    if (hold && this.lastPaint) this.paintNow(this.lastPaint.sel, this.lastPaint.sink)
    this.clearOverflowAnchor()
    this.onLayout()
  }

  private measureStories(): void {
    if (!this.doc) return
    const section = this.doc.sections[0]
    if (!section) return
    const batch: { block: Block; section: Section }[] = []
    for (const blocks of this.doc.stories.values()) {
      for (const block of blocks) batch.push({ block, section })
    }
    this.measureBatch(batch)
  }

  private measureBatch(items: readonly { block: Block; section: Section }[]): void {
    const host = this.measureHost
    if (!host) return
    const pending: { block: Block; el: HTMLElement }[] = []
    const frag = document.createDocumentFragment()
    for (const item of items) {
      if (this.measures.has(item.block) && !this.rewriting) continue
      const el = renderBlock(item.block, this.ctxFor(1))
      el.style.width = `${contentWidthOf(item.section)}px`
      frag.append(el)
      pending.push({ block: item.block, el })
    }
    if (pending.length === 0) return
    host.append(frag)
    for (const item of pending) {
      this.lineStartCache.delete(item.block)
      this.measures.set(item.block, measureElement(item.el, 1))
    }
    host.replaceChildren()
  }

  private measureOne(block: Block, section: Section): void {
    this.measureBatch([{ block, section }])
  }

  private ensureChrome(): void {
    if (this.stack?.isConnected) return
    this.host.classList.add('page-view')
    const sizer = document.createElement('div')
    sizer.className = 'page-sizer'
    const stack = document.createElement('div')
    stack.className = 'page-stack'
    const measureHost = document.createElement('div')
    measureHost.className = 'measure-host'
    sizer.append(stack)
    this.host.replaceChildren(sizer, measureHost)
    this.sizer = sizer
    this.stack = stack
    this.measureHost = measureHost
    this.observer = new IntersectionObserver((entries) => this.onIntersect(entries), { root: this.host })
    this.startWatch()
  }

  private resetPages(): void {
    this.observer?.disconnect()
    this.pageEls = []
    this.mountedSig.clear()
    this.intersecting.clear()
    this.stack?.replaceChildren()
    this.observer = new IntersectionObserver((entries) => this.onIntersect(entries), { root: this.host })
  }

  private cancelChunk(): void {
    if (this.chunk !== 0) cancelAnimationFrame(this.chunk)
    this.chunk = 0
  }

  private startWatch(): void {
    if (this.watch !== 0) return
    const tick = (): void => {
      if (this.dead) return
      const top = this.host.scrollTop
      const height = this.host.clientHeight
      if (top !== this.seenScroll || height !== this.seenHeight) {
        this.seenScroll = top
        this.seenHeight = height
        this.syncFromGeometry()
      }
      this.watch = requestAnimationFrame(tick)
    }
    this.watch = requestAnimationFrame(tick)
  }

  private onIntersect(_entries: IntersectionObserverEntry[]): void {
    // acdert reports the first observation, then nothing when the scroll root moves. startWatch covers that.
    if (this.syncQueued) return
    this.syncQueued = true
    requestAnimationFrame(() => {
      this.syncQueued = false
      if (!this.dead) this.syncFromGeometry()
    })
  }

  private syncFromGeometry(): void {
    const host = this.host.getBoundingClientRect()
    const n = this.pageEls.length
    const next = new Set<number>()
    let lo = 0
    let hi = n - 1
    let first = n
    while (lo <= hi) {
      const mid = (lo + hi) >> 1
      const el = this.pageEls[mid]
      if (!el) break
      if (el.getBoundingClientRect().bottom <= host.top + 0.5) lo = mid + 1
      else {
        first = mid
        hi = mid - 1
      }
    }
    for (let i = first; i < n; i++) {
      const el = this.pageEls[i]
      if (!el) continue
      const rect = el.getBoundingClientRect()
      if (rect.top >= host.bottom - 0.5) break
      if (rect.bottom > host.top && rect.right > host.left && rect.left < host.right) next.add(i)
    }
    this.intersecting = next
    this.applyWindow()
  }

  private syncPages(): void {
    this.ensureShells()
    this.updateSizer()
    this.syncFromGeometry()
  }

  private ensureShells(): void {
    const stack = this.stack
    const pages = this._layout.pages
    if (!stack || !this.doc) return
    while (this.pageEls.length > pages.length) {
      const el = this.pageEls.pop()
      if (!el) continue
      this.observer?.unobserve(el)
      this.intersecting.delete(Number(el.dataset.page))
      this.mountedSig.delete(Number(el.dataset.page))
      el.remove()
    }
    for (let i = 0; i < pages.length; i++) {
      const page = pages[i]
      if (!page) continue
      let el = this.pageEls[i]
      if (!el) {
        el = document.createElement('div')
        el.className = 'page'
        el.dataset.page = String(i)
        el.dataset.mounted = '0'
        const header = document.createElement('div')
        header.className = 'page-header'
        const body = document.createElement('div')
        body.className = 'page-body'
        const footer = document.createElement('div')
        footer.className = 'page-footer'
        const overlay = document.createElement('div')
        overlay.className = 'overlay'
        const num = document.createElement('div')
        num.className = 'page-num'
        el.append(header, body, footer, overlay, num)
        stack.append(el)
        this.pageEls.push(el)
        this.observer?.observe(el)
      }
      const section = this.doc.sections[page.section] ?? this.doc.sections[0]
      if (!section) continue
      const box = pageBoxStyle(section)
      assignStyle(el, 'width', box.page.width ?? '')
      assignStyle(el, 'height', box.page.height ?? '')
      const label = el.querySelector('.page-num')
      const number = String(i + 1)
      if (label && label.textContent !== number) label.textContent = number
      const frame = pageFrame(this.doc, section, i + 1, (block) => this.measures.get(block))
      const body = el.querySelector('.page-body')
      if (body instanceof HTMLElement) {
        assignStyle(body, 'top', `${page.bodyTopPx}px`)
        assignStyle(body, 'left', `${frame.contentLeft}px`)
        assignStyle(body, 'width', `${frame.contentWidth}px`)
        assignStyle(body, 'height', `${frame.bodyHeight}px`)
      }
    }
  }

  private applyWindow(): void {
    const n = this._layout.pages.length
    if (n === 0) return
    let from = 0
    let to = Math.min(1, n - 1)
    if (this.intersecting.size === 0 && this.mountedSig.size > 0) return
    if (this.intersecting.size > 0) {
      let min = n
      let max = -1
      for (const index of this.intersecting) {
        if (index < 0 || index >= n) continue
        if (index < min) min = index
        if (index > max) max = index
      }
      if (max >= 0) {
        from = Math.max(0, min - 1)
        to = Math.min(n - 1, max + 1)
      }
    }
    for (let i = 0; i < this.pageEls.length; i++) {
      if (i >= from && i <= to) {
        if (this.mountedSig.has(i)) this.refreshPage(i)
        else this.fillPage(i)
      } else if (this.mountedSig.has(i)) this.clearPage(i)
    }
  }

  private updateSizer(): void {
    const sizer = this.sizer
    const stack = this.stack
    if (!sizer || !stack) return
    const width = Math.max(1, this.maxPaper())
    let height = this.stackHeight()
    if (this.rewriting && this.zoom > 0) {
      const minHeight = (this.holdScroll + this.host.clientHeight) / this.zoom
      if (height < minHeight) height = minHeight
    }
    assignStyle(sizer, 'width', `${width * this.zoom}px`)
    assignStyle(sizer, 'height', `${height * this.zoom}px`)
    assignStyle(stack, 'width', `${width}px`)
    const transform = this.zoom === 1 ? '' : `scale(${this.zoom})`
    if (stack.style.transform !== transform) stack.style.transform = transform
  }

  private stackHeight(): number {
    let height = STACK_PAD
    for (let i = 0; i < this._layout.pages.length; i++) height += this.paperPx(i) + PAGE_GAP
    return height
  }

  private paperPx(i: number): number {
    const page = this._layout.pages[i]
    const section = page && this.doc ? (this.doc.sections[page.section] ?? this.doc.sections[0]) : undefined
    return section ? section.page.heightPt * PT_TO_PX : 0
  }

  private maxPaper(): number {
    let width = 0
    if (!this.doc) return 0
    for (const page of this._layout.pages) {
      const section = this.doc.sections[page.section] ?? this.doc.sections[0]
      if (section) width = Math.max(width, section.page.widthPt * PT_TO_PX)
    }
    return width
  }

  private pageOffsets(): number[] {
    const offsets: number[] = []
    let y = STACK_PAD
    for (let i = 0; i < this._layout.pages.length; i++) {
      offsets.push(y)
      y += this.paperPx(i) + PAGE_GAP
    }
    return offsets
  }

  private signature(page: Page): string {
    let text = `${page.section}:${page.bodyTopPx}:${page.header ?? ''}:${page.footer ?? ''}:`
    for (const slice of page.slices) text += `${slice.block}@${slice.y0},${slice.y1},${slice.repeatHeaderRows};`
    return text
  }

  private mountPage(i: number): void {
    if (i < 0 || i >= this._layout.pages.length) return
    this.ensureShells()
    if (this.mountedSig.has(i)) this.refreshPage(i)
    else this.fillPage(i)
  }

  private fillPage(i: number): void {
    const page = this._layout.pages[i]
    const el = this.pageEls[i]
    const doc = this.doc
    if (!page || !el || !doc) return
    const section = doc.sections[page.section] ?? doc.sections[0]
    if (!section) return
    const frame = pageFrame(doc, section, i + 1, (block) => this.measures.get(block))
    this.fillStory(el.querySelector('.page-header'), frame.header, frame, 'header', i)
    this.fillStory(el.querySelector('.page-footer'), frame.footer, frame, 'footer', i)
    const body = el.querySelector('.page-body')
    if (!(body instanceof HTMLElement)) return
    body.replaceChildren()
    const ctx = this.ctxFor(i + 1)
    for (const slice of page.slices) {
      const located = this.locateOwner(slice.block)
      if (!located) continue
      const measured = this.measures.get(located.block)
      const blockEl = renderBlock(located.block, ctx, this.decoFor(located.block))
      const headerH =
        slice.repeatHeaderRows > 0 && measured?.rowBottoms ? (measured.rowBottoms[slice.repeatHeaderRows - 1] ?? 0) : 0
      const wrap = document.createElement('div')
      wrap.className = 'slice'
      wrap.dataset.block = slice.block
      wrap.dataset.y0 = String(slice.y0)
      wrap.dataset.y1 = String(slice.y1)
      wrap.style.height = `${headerH + (slice.y1 - slice.y0)}px`
      if (slice.repeatHeaderRows > 0 && located.block.kind === 'table') {
        const head = extractHeader(blockEl, slice.repeatHeaderRows)
        if (head) wrap.append(head)
      }
      const clip = document.createElement('div')
      clip.className = 'clip'
      clip.style.height = `${slice.y1 - slice.y0}px`
      const shift = document.createElement('div')
      shift.className = 'shift'
      shift.style.transform = `translateY(${-slice.y0}px)`
      shift.append(blockEl)
      clip.append(shift)
      wrap.append(clip)
      body.append(wrap)
    }
    el.dataset.mounted = '1'
    this.mountedSig.set(i, this.signature(page))
    this.hideExtraCopies()
  }

  private fillStory(
    host: Element | null,
    key: StoryKey | null,
    frame: Frame,
    edge: 'header' | 'footer',
    pageIndex: number,
  ): void {
    if (!(host instanceof HTMLElement) || !this.doc) return
    host.replaceChildren()
    if (!key) return
    const blocks = this.doc.stories.get(key)
    if (!blocks) return
    host.style.left = `${frame.contentLeft}px`
    host.style.width = `${frame.contentWidth}px`
    if (edge === 'header') host.style.top = `${frame.headerTop}px`
    else host.style.bottom = `${frame.footerBottom}px`
    const ctx = this.ctxFor(pageIndex + 1)
    for (const block of blocks) host.append(renderBlock(block, ctx))
  }

  private refreshPage(i: number): void {
    const page = this._layout.pages[i]
    const el = this.pageEls[i]
    if (!page || !el) return
    const next = this.signature(page)
    if (this.mountedSig.get(i) !== next) {
      this.fillPage(i)
      return
    }
    if (!this.changedIds) return
    const ctx = this.ctxFor(i + 1)
    const painted = new Set<string>()
    for (const slice of page.slices) {
      const id = slice.block
      if (!this.changedIds.has(id) || painted.has(id)) continue
      painted.add(id)
      const located = this.locateOwner(id)
      if (!located) continue
      for (const node of blockElements(el, id)) {
        if (node.closest('.repeat-head')) continue
        node.replaceWith(renderBlock(located.block, ctx, this.decoFor(located.block)))
      }
    }
    if (painted.size > 0) this.hideExtraCopies()
  }

  private clearPage(i: number): void {
    const el = this.pageEls[i]
    if (!el) return
    el.querySelector('.page-body')?.replaceChildren()
    el.querySelector('.page-header')?.replaceChildren()
    el.querySelector('.page-footer')?.replaceChildren()
    el.querySelector('.overlay')?.replaceChildren()
    delete el.dataset.mounted
    el.dataset.mounted = '0'
    this.mountedSig.delete(i)
    this.hideExtraCopies()
  }

  // A straddling block is mounted twice. The later copy is the clipped remainder, so hiding it keeps the paragraph from being read twice.
  private hideExtraCopies(): void {
    const seen = new Set<string>()
    for (const pageEl of this.pageEls) {
      if (pageEl.dataset.mounted !== '1') continue
      for (const blockEl of pageEl.querySelectorAll<HTMLElement>('.page-body .shift > .blk')) {
        const id = blockEl.dataset.id
        if (!id) continue
        if (seen.has(id)) blockEl.setAttribute('aria-hidden', 'true')
        else {
          blockEl.removeAttribute('aria-hidden')
          seen.add(id)
        }
      }
    }
  }

  private ctxFor(pageNumber: number): RenderContext {
    const base = this.ctx()
    return { ...base, pageField: { page: pageNumber, pages: Math.max(this._layout.pages.length, pageNumber) } }
  }

  private decoFor(block: Block): Decoration | undefined {
    if (!this.deco) return undefined
    return containsId(block, this.deco.pos.p) ? this.deco : undefined
  }

  private index(): Map<string, Place> {
    if (this.indexDoc === this.doc && this.indexMap) return this.indexMap
    const map = new Map<string, Place>()
    let order = 0
    const walk = (block: Block, owner: Block, row: number): void => {
      if (block.kind === 'p') {
        map.set(block.id, { para: block, owner, row, order })
        order += 1
        return
      }
      if (block.kind === 'table') {
        block.rows.forEach((tableRow, rowIndex) => {
          for (const cell of tableRow.cells) {
            for (const child of cell.blocks) walk(child, block, rowIndex)
          }
        })
        return
      }
      for (const para of block.preview) walk(para, block, 0)
    }
    if (this.doc) {
      for (const block of this.doc.body) walk(block, block, 0)
      for (const blocks of this.doc.stories.values()) {
        for (const block of blocks) walk(block, block, 0)
      }
    }
    this.indexMap = map
    this.indexDoc = this.doc
    return map
  }

  private locateOwner(id: string): { block: Block; section: Section } | null {
    if (!this.doc) return null
    const sections = sectionIndexes(this.doc.body)
    for (let i = 0; i < this.doc.body.length; i++) {
      const block = this.doc.body[i]
      if (!block || !containsId(block, id)) continue
      const section = this.doc.sections[sections[i] ?? 0] ?? this.doc.sections[0]
      if (!section) return null
      return { block, section }
    }
    const section = this.doc.sections[0]
    if (!section) return null
    for (const blocks of this.doc.stories.values()) {
      for (const block of blocks) {
        if (containsId(block, id)) return { block, section }
      }
    }
    return null
  }

  private spots(id: BlockId): { page: number; slice: Slice }[] {
    const out: { page: number; slice: Slice }[] = []
    this._layout.pages.forEach((page, pageIndex) => {
      for (const slice of page.slices) {
        if (slice.block === id) out.push({ page: pageIndex, slice })
      }
    })
    return out
  }

  private sliceY(place: Place, offset: number): number {
    if (place.owner.kind === 'p') return this.yForOffset(place.owner, offset)
    if (place.owner.kind !== 'table') return 0
    const table = place.owner
    if (this.rowContinues(table, place.row)) {
      const length = flatText(place.para).length
      const y = caretSliceY(offset, length, (at) => this.yInRow(table, place.para.id, at))
      if (y !== null) return y
    }
    return this.rowMiddle(place.owner, place.row)
  }

  private rowContinues(table: Table, row: number): boolean {
    const bottoms = this.measures.get(table)?.rowBottoms
    if (!bottoms || row < 0 || row >= bottoms.length) return false
    const top = row <= 0 ? 0 : (bottoms[row - 1] ?? 0)
    const bottom = bottoms[row] ?? top
    for (const spot of this.spots(table.id)) {
      if (spot.slice.y1 <= top + EPS || spot.slice.y0 >= bottom - EPS) continue
      if (spot.slice.y0 > top + EPS || spot.slice.y1 < bottom - EPS) return true
    }
    return false
  }

  private yInRow(table: Table, paraId: string, offset: number): number | null {
    const spots = this.spots(table.id)
    let pageIndex = -1
    for (const spot of spots) {
      const pageEl = this.pageEls[spot.page]
      if (!pageEl || pageEl.dataset.mounted !== '1') continue
      if (this.paraCopy(pageEl, paraId)) {
        pageIndex = spot.page
        break
      }
    }
    if (pageIndex < 0) {
      const first = spots[0]
      if (!first) return null
      this.mountPage(first.page)
      pageIndex = first.page
    }
    const pageEl = this.pageEls[pageIndex]
    if (!pageEl) return null
    const paraEl = this.paraCopy(pageEl, paraId)
    if (!paraEl) return null
    const shift = paraEl.closest('.shift')
    const blockEl = shift instanceof HTMLElement ? shift.firstElementChild : null
    if (!(blockEl instanceof HTMLElement)) return null
    const range = rangeIn(paraEl, offset, offset)
    const rects = [...range.getClientRects()].filter((rect) => rect.height > 0)
    const rect = rects[0] ?? paraEl.getBoundingClientRect()
    const blockRect = blockEl.getBoundingClientRect()
    const z = this.zoom > 0 ? this.zoom : 1
    const y = (rect.top + Math.min(rect.height / 2, 1) - blockRect.top) / z
    return Number.isFinite(y) ? y : null
  }

  private paraCopy(pageEl: HTMLElement, paraId: string): HTMLElement | null {
    const found = blockElements(pageEl, paraId).find((el) => el.classList.contains('p') && !el.closest('.repeat-head'))
    return found ?? null
  }

  private mountContinuedRows(start: Place, end: Place): void {
    const seen = new Set<string>()
    for (const place of this.index().values()) {
      if (place.order < start.order || place.order > end.order) continue
      if (place.owner.kind !== 'table') continue
      const key = `${place.owner.id}:${place.row}`
      if (seen.has(key)) continue
      seen.add(key)
      if (!this.rowContinues(place.owner, place.row)) continue
      const bottoms = this.measures.get(place.owner)?.rowBottoms
      const top = place.row <= 0 ? 0 : (bottoms?.[place.row - 1] ?? 0)
      const bottom = bottoms?.[place.row] ?? top
      for (const spot of this.spots(place.owner.id)) {
        if (spot.slice.y1 <= top + EPS || spot.slice.y0 >= bottom - EPS) continue
        this.mountPage(spot.page)
      }
    }
  }

  private rowMiddle(table: Table, row: number): number {
    const bottoms = this.measures.get(table)?.rowBottoms
    if (!bottoms || bottoms.length === 0) return 0
    const top = row <= 0 ? 0 : (bottoms[row - 1] ?? 0)
    const bottom = bottoms[row] ?? top
    return (top + bottom) / 2
  }

  private yForOffset(para: Paragraph, offset: number): number {
    const measured = this.measures.get(para)
    if (!measured || measured.lineBottoms.length === 0) return 0
    const starts = this.lineStarts(para)
    let line = 0
    if (starts) {
      for (let i = 0; i < starts.length; i++) {
        if ((starts[i] ?? 0) <= offset) line = i
      }
    }
    const bottom = measured.lineBottoms[line] ?? measured.heightPx
    const prev = line === 0 ? 0 : (measured.lineBottoms[line - 1] ?? 0)
    return (prev + bottom) / 2
  }

  private lineStarts(para: Paragraph): Uint32Array | null {
    const cached = this.lineStartCache.get(para)
    if (cached) return cached
    const measured = this.measures.get(para)
    const located = this.locateOwner(para.id)
    const host = this.measureHost
    if (!measured || measured.lineBottoms.length === 0 || !located || !host) return null
    const el = renderBlock(para, this.ctxFor(1))
    el.style.width = `${contentWidthOf(located.section)}px`
    host.replaceChildren(el)
    const starts = startsFromElement(el, measured.lineBottoms, flatText(para).length)
    this.lineStartCache.set(para, starts)
    host.replaceChildren()
    return starts
  }

  private paraElement(pageEl: HTMLElement, pos: Pos): HTMLElement | null {
    const all = blockElements(pageEl, pos.p).filter((el) => el.classList.contains('p') && !el.closest('.repeat-head'))
    const only = all[0]
    if (!only) return null
    if (all.length === 1) return only
    const place = this.index().get(pos.p)
    const y = place ? this.sliceY(place, pos.o) : 0
    for (const el of all) {
      const slice = el.closest('.slice')
      if (!(slice instanceof HTMLElement)) continue
      const y0 = Number(slice.dataset.y0)
      const y1 = Number(slice.dataset.y1)
      if (y >= y0 - 0.5 && y < y1 + 0.5) return el
    }
    return only
  }

  private pageIndexAt(clientX: number, clientY: number): number {
    let best = -1
    let bestDist = Infinity
    for (let i = 0; i < this.pageEls.length; i++) {
      const el = this.pageEls[i]
      if (!el) continue
      const rect = el.getBoundingClientRect()
      const dx = clientX < rect.left ? rect.left - clientX : clientX > rect.right ? clientX - rect.right : 0
      const dy = clientY < rect.top ? rect.top - clientY : clientY > rect.bottom ? clientY - rect.bottom : 0
      const dist = dx * dx + dy * dy
      if (dist < bestDist) {
        bestDist = dist
        best = i
      }
    }
    return best
  }

  private adjustAtom(node: Node, pos: Pos, clientX: number): Pos {
    const el = node instanceof Element ? node : node.parentElement
    const atom = el?.closest('.atom')
    if (!(atom instanceof HTMLElement)) return pos
    const base = Number(atom.dataset.o)
    if (!Number.isFinite(base)) return pos
    const rect = atom.getBoundingClientRect()
    return { p: pos.p, o: clientX > (rect.left + rect.right) / 2 ? base + 1 : base }
  }

  private withHits<T>(fn: () => T): T {
    const hidden: HTMLElement[] = []
    const hide = (el: Element | null | undefined): void => {
      if (el instanceof HTMLElement && el.style.visibility !== 'hidden') {
        el.style.visibility = 'hidden'
        hidden.push(el)
      }
    }
    if (this.sink || this.lastPaint) {
      hide(this.sink)
      for (const page of this.pageEls) hide(page.querySelector('.overlay'))
    }
    try {
      return fn()
    } finally {
      for (const el of hidden) el.style.visibility = ''
    }
  }

  private posAtLocal(pageIndex: number, x: number, y: number): Pos | null {
    const pageEl = this.pageEls[pageIndex]
    if (!pageEl) return null
    const origin = pageEl.getBoundingClientRect()
    return this.posAtPoint(origin.left + x * this.zoom, origin.top + y * this.zoom)
  }

  private collectLines(indexes: readonly number[]): LineBox[] {
    const offsets = this.pageOffsets()
    const lines: LineBox[] = []
    const z = this.zoom
    for (const pageIndex of indexes) {
      const pageEl = this.pageEls[pageIndex]
      if (!pageEl || pageEl.dataset.mounted !== '1') continue
      const origin = pageEl.getBoundingClientRect()
      const base = offsets[pageIndex] ?? 0
      for (const para of pageEl.querySelectorAll<HTMLElement>('.blk.p')) {
        if (para.closest('.repeat-head')) continue
        const id = para.dataset.id
        if (!id) continue
        const clip = para.closest('.clip') ?? para.closest('.slice')
        const clipRect = clip instanceof HTMLElement ? clip.getBoundingClientRect() : null
        const rects: DOMRect[] = []
        for (const child of para.children) {
          if (!(child instanceof HTMLElement) || child.classList.contains('marker')) continue
          rects.push(...childContentRects(child))
        }
        const story = para.closest('.page-header') ? 'header' : para.closest('.page-footer') ? 'footer' : 'body'
        for (const line of groupLineRects(rects)) {
          const visTop = clipRect ? Math.max(line.top, clipRect.top) : line.top
          const visBottom = clipRect ? Math.min(line.bottom, clipRect.bottom) : line.bottom
          if (visBottom - visTop < 8) continue
          const top = (visTop - origin.top) / z
          const bottom = (visBottom - origin.top) / z
          lines.push({
            page: pageIndex,
            left: (line.left - origin.left) / z,
            right: (line.right - origin.left) / z,
            top,
            bottom,
            globalMid: base + (top + bottom) / 2,
            paraId: id,
            story,
          })
        }
      }
    }
    return lines
  }

  private pickCurrent(lines: readonly LineBox[], page: number, x: number, y: number): LineBox | null {
    let best: LineBox | null = null
    let bestDist = Infinity
    for (const line of lines) {
      if (line.page !== page) continue
      const inside = y >= line.top - 2 && y <= line.bottom + 2 && x >= line.left - 8 && x <= line.right + 8
      const dy = y < line.top ? line.top - y : y > line.bottom ? y - line.bottom : 0
      const dx = x < line.left ? line.left - x : x > line.right ? x - line.right : 0
      const dist = inside ? 0 : dx * dx + dy * dy
      if (dist < bestDist) {
        bestDist = dist
        best = line
      }
    }
    return best
  }

  private posOnLine(line: LineBox, goalX: number): Pos | null {
    const place = this.index().get(line.paraId)
    if (!place) return null
    const id = place.para.id
    const length = flatText(place.para).length
    const caretAt = (offset: number): CaretRect => this.caretRect({ p: id, o: offset })
    const onLine = (offset: number): boolean => {
      if (offset < 0 || offset > length) return false
      const caret = caretAt(offset)
      const y = caret.top + caret.height / 2
      return caret.page === line.page && y >= line.top - 3 && y <= line.bottom + 3
    }
    const beforeLine = (offset: number): boolean => {
      const caret = caretAt(offset)
      const y = caret.top + caret.height / 2
      return caret.page < line.page || (caret.page === line.page && y < line.top - 3)
    }
    let lo = 0
    let hi = length
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (beforeLine(mid)) lo = mid + 1
      else hi = mid
    }
    let seed = onLine(lo) ? lo : lo > 0 && onLine(lo - 1) ? lo - 1 : -1
    if (seed < 0) {
      const midY = (line.top + line.bottom) / 2
      const x = Math.min(Math.max(goalX, line.left + 0.5), Math.max(line.left + 0.5, line.right - 0.5))
      const direct = this.posAtLocal(line.page, x, midY)
      if (!direct || direct.p !== id) return null
      seed = direct.o
    }
    let left = seed
    let right = seed
    while (left > 0 && onLine(left - 1)) left -= 1
    while (right < length && onLine(right + 1)) right += 1
    let best = seed
    let bestDist = Math.abs(caretAt(seed).x - goalX)
    while (left <= right) {
      const mid = (left + right) >> 1
      const x = caretAt(mid).x
      const dist = Math.abs(x - goalX)
      if (dist < bestDist) {
        bestDist = dist
        best = mid
      }
      if (x < goalX - 0.25) left = mid + 1
      else right = mid - 1
    }
    return { p: id, o: best }
  }

  private paintNow(sel: Selection, sink: HTMLTextAreaElement): void {
    // Clearing an empty overlay still costs a mutation per page in acdert, on every keystroke.
    for (const el of this.pageEls) {
      const overlay = el.querySelector('.overlay')
      if (overlay && overlay.firstChild) overlay.replaceChildren()
    }
    this.drawFindHits(false)
    if (sel.kind === 'cells') {
      this.paintCells(sel, sink)
      this.drawFindHits(true)
      return
    }
    const head = this.caretRect(sel.head)
    // A selected picture shows its frame and handles. A caret beside it and a tint over it would only hide it.
    const picture = this.doc !== null && selectedImage(this.doc, sel) !== null
    // The current match is the selection. Its own tint shows it, and the selection tint over it would turn it muddy.
    const hit = this.findHits[this.findCurrent]
    const onHit = hit !== undefined && sel.anchor.p === hit.para && sel.head.p === hit.para
      && Math.min(sel.anchor.o, sel.head.o) === hit.start && Math.max(sel.anchor.o, sel.head.o) === hit.end
    if (!picture && !onHit) this.drawCaret(head)
    this.placeSink(head, sink)
    if (!picture && !onHit && (sel.anchor.p !== sel.head.p || sel.anchor.o !== sel.head.o)) this.drawSelection(sel.anchor, sel.head)
    this.paintImageChrome(sel)
    if (!picture && this.composing) this.syncHostPreeditCaret()
    this.drawFindHits(true)
  }

  private drawFindHits(currentOnly: boolean): void {
    if (this.findHits.length === 0) return
    const byPara = new Map<string, TextMatch[]>()
    for (let i = 0; i < this.findHits.length; i++) {
      const hit = this.findHits[i]
      if (!hit) continue
      if ((i === this.findCurrent) !== currentOnly) continue
      const list = byPara.get(hit.para)
      if (list) list.push(hit)
      else byPara.set(hit.para, [hit])
    }
    if (byPara.size === 0) return
    const z = this.zoom
    for (const pageEl of this.pageEls) {
      if (pageEl.dataset.mounted !== '1') continue
      const overlay = pageEl.querySelector('.overlay')
      if (!(overlay instanceof HTMLElement)) continue
      const origin = pageEl.getBoundingClientRect()
      for (const node of pageEl.querySelectorAll<HTMLElement>('.blk.p')) {
        if (node.closest('.repeat-head')) continue
        const id = node.dataset.id
        if (!id) continue
        const hits = byPara.get(id)
        if (!hits) continue
        const slice = node.closest('.slice')
        const sliceRect = slice instanceof HTMLElement ? slice.getBoundingClientRect() : null
        for (const hit of hits) {
          const range = rangeIn(node, hit.start, hit.end)
          const boxes: { left: number, top: number, right: number, bottom: number }[] = []
          for (const rect of range.getClientRects()) {
            boxes.push({ left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom })
          }
          for (const rect of boxes) {
            if (rect.right - rect.left < 0.5 || rect.bottom - rect.top < 0.5) continue
            let left = rect.left
            let top = rect.top
            let right = rect.right
            let bottom = rect.bottom
            if (sliceRect) {
              left = Math.max(left, sliceRect.left)
              top = Math.max(top, sliceRect.top)
              right = Math.min(right, sliceRect.right)
              bottom = Math.min(bottom, sliceRect.bottom)
            }
            if (right - left < 0.5 || bottom - top < 0.5) continue
            const mark = document.createElement('div')
            mark.className = currentOnly ? 'find-hit current' : 'find-hit'
            mark.setAttribute('aria-hidden', 'true')
            mark.style.left = `${(left - origin.left) / z}px`
            mark.style.top = `${(top - origin.top) / z}px`
            mark.style.width = `${(right - left) / z}px`
            mark.style.height = `${(bottom - top) / z}px`
            overlay.append(mark)
          }
        }
      }
    }
  }

  /**
   * The runtime's composition span sits on `document.body` and is not part of the paragraph. Align its font and
   * baseline with the paragraph run at the caret, then place the visible caret at its right edge.
   */
  syncHostPreeditCaret(): void {
    const span = hostPreeditElement()
    if (!span) return
    this.composing = true
    const mark = this.host.querySelector('.caret')
    if (!(mark instanceof HTMLElement)) return
    const overlay = mark.parentElement
    if (!(overlay instanceof HTMLElement)) return
    const overlayBox = overlay.getBoundingClientRect()
    const scale = overlay.offsetWidth > 0 ? overlayBox.width / overlay.offsetWidth : this.zoom
    const z = scale > 0 ? scale : 1
    this.alignHostPreedit(span, z)
    const spanBox = span.getBoundingClientRect()
    const placed = caretAtPreeditEnd(spanBox, {
      left: overlayBox.left,
      top: overlayBox.top,
      width: overlayBox.width,
      offsetWidth: overlay.offsetWidth,
    })
    mark.style.left = `${placed.left}px`
    mark.style.top = `${placed.top}px`
    mark.style.height = `${placed.height}px`
  }

  private alignHostPreedit(span: HTMLElement, z: number): void {
    const last = this.lastPaint
    if (!last || last.sel.kind !== 'text' || !this.doc) return
    const pos = last.sel.head
    const place = this.index().get(pos.p)
    if (!place) return
    const pageCount = this._layout.pages.length
    const pageIndex = pageCount === 0 ? 0 : Math.min(this.pageOf(pos), pageCount - 1)
    this.mountPage(pageIndex)
    const pageEl = this.pageEls[pageIndex]
    if (!pageEl) return
    const paraEl = this.paraElement(pageEl, pos)
    if (!paraEl) return
    const marks = last.stored ?? marksAt(place.para, pos.o).marks
    const run = preeditRunStyle(this.ctxFor(1), place.para, marks, span.textContent ?? '')
    const point = domPoint(paraEl, pos.o)
    const slice = paraEl.closest('.slice')
    const sliceRect = slice instanceof HTMLElement ? slice.getBoundingClientRect() : null
    const inlineH = this.preeditInlineHeight(run)
    const atomCaret = inlineAtomCaretViewport(point, pos.o, sliceRect, inlineH * z)
    let placed: { readonly left: number; readonly top: number; readonly height: number }
    if (atomCaret) {
      placed = atomCaret
    } else {
      const range = scratchRange()
      range.setStart(point.node, point.offset)
      range.setEnd(point.node, point.offset)
      const rects = [...range.getClientRects()].filter((item) => item.height > 0)
      const rect =
        rects.find((item) => !sliceRect || (item.bottom > sliceRect.top + 0.5 && item.top < sliceRect.bottom - 0.5)) ??
        rects[0]
      if (!rect) {
        const box = paraEl.getBoundingClientRect()
        const padLeft = parseFloat(paraEl.style.paddingLeft) || 0
        const padTop = parseFloat(paraEl.style.paddingTop) || 0
        const indent = parseFloat(paraEl.style.textIndent) || 0
        placed = {
          left: box.left + Math.max(0, padLeft + indent) * z,
          top: box.top + padTop * z,
          height: Math.max(inlineH * z, 1),
        }
      } else if (point.node instanceof Text && point.node.parentElement instanceof HTMLElement) {
        const sample = sampleCharAt(point.node.data, point.offset)
        const metric = this.preeditMetric(point.node.parentElement, sample, point.offset > 0, run)
        placed = hostPreeditBox({ left: rect.left, top: rect.top }, metric, z)
      } else {
        const height = Math.max(inlineH * z, 1)
        placed = { left: rect.left, top: rect.bottom - height, height }
      }
    }
    span.style.fontFamily = run.fontFamily
    span.style.fontSize = `${run.fontSizePx * z}px`
    span.style.fontWeight = run.fontWeight
    span.style.fontStyle = run.fontStyle
    span.style.letterSpacing = run.letterSpacingPx > 0 ? `${run.letterSpacingPx * z}px` : 'normal'
    span.style.color = run.color
    span.style.background = '#fff'
    span.style.left = `${placed.left}px`
    span.style.top = `${placed.top}px`
    span.style.height = `${placed.height}px`
    span.style.lineHeight = `${placed.height}px`
  }

  private preeditMetric(anchorSpan: HTMLElement, sample: string, atEnd: boolean, run: PreeditRunStyle): PreeditMetric {
    const preSample = run.slot === 'eastAsia' ? '中' : 'a'
    const key = [
      anchorSpan.style.fontFamily,
      anchorSpan.style.fontSize,
      anchorSpan.style.lineHeight,
      anchorSpan.style.fontWeight,
      anchorSpan.style.fontStyle,
      anchorSpan.style.verticalAlign,
      sample,
      atEnd ? '1' : '0',
      run.fontFamily,
      run.fontSizePx,
      run.fontWeight,
      run.fontStyle,
      run.vertAlign ?? '',
      preSample,
    ].join('\0')
    const cached = this.preeditMetrics.get(key)
    if (cached) return cached
    const host = this.measureHost
    if (!host) return { deltaTop: 0, height: run.fontSizePx }
    const probe = document.createElement('div')
    probe.style.cssText = 'position:absolute;left:0;top:0;white-space:pre'
    const a = document.createElement('span')
    a.style.fontFamily = anchorSpan.style.fontFamily
    a.style.fontSize = anchorSpan.style.fontSize
    a.style.lineHeight = anchorSpan.style.lineHeight
    a.style.fontWeight = anchorSpan.style.fontWeight
    a.style.fontStyle = anchorSpan.style.fontStyle
    if (anchorSpan.style.verticalAlign) a.style.verticalAlign = anchorSpan.style.verticalAlign
    a.textContent = sample
    const b = document.createElement('span')
    b.style.fontFamily = run.fontFamily
    b.style.fontSize = `${run.fontSizePx}px`
    b.style.lineHeight = 'normal'
    b.style.fontWeight = run.fontWeight
    b.style.fontStyle = run.fontStyle
    if (run.vertAlign) b.style.verticalAlign = run.vertAlign
    b.textContent = preSample
    probe.append(a, b)
    host.replaceChildren(probe)
    const textNode = a.firstChild
    let aTop: number
    if (textNode instanceof Text) {
      const range = scratchRange()
      const off = atEnd ? textNode.data.length : 0
      range.setStart(textNode, off)
      range.setEnd(textNode, off)
      const r = [...range.getClientRects()].find((item) => item.height > 0)
      aTop = r ? r.top : a.getBoundingClientRect().top
    } else {
      aTop = a.getBoundingClientRect().top
    }
    const bRect = b.getBoundingClientRect()
    host.replaceChildren()
    const metric: PreeditMetric = {
      deltaTop: bRect.top - aTop,
      height: bRect.height > 0 ? bRect.height : run.fontSizePx,
    }
    if (this.preeditMetrics.size >= 128) this.preeditMetrics.clear()
    this.preeditMetrics.set(key, metric)
    return metric
  }

  private preeditInlineHeight(run: PreeditRunStyle): number {
    const preSample = run.slot === 'eastAsia' ? '中' : 'a'
    const key = `h\0${run.fontFamily}\0${run.fontSizePx}\0${run.fontWeight}\0${run.fontStyle}\0${preSample}`
    const cached = this.preeditMetrics.get(key)
    if (cached) return cached.height
    const host = this.measureHost
    if (!host) return run.fontSizePx
    const probe = document.createElement('div')
    probe.style.cssText = 'position:absolute;left:0;top:0;white-space:pre'
    const b = document.createElement('span')
    b.style.fontFamily = run.fontFamily
    b.style.fontSize = `${run.fontSizePx}px`
    b.style.lineHeight = 'normal'
    b.style.fontWeight = run.fontWeight
    b.style.fontStyle = run.fontStyle
    b.textContent = preSample
    probe.append(b)
    host.replaceChildren(probe)
    const h = b.getBoundingClientRect().height
    host.replaceChildren()
    const height = h > 0 ? h : run.fontSizePx
    if (this.preeditMetrics.size >= 128) this.preeditMetrics.clear()
    this.preeditMetrics.set(key, { deltaTop: 0, height })
    return height
  }

  private drawCaret(caret: CaretRect): void {
    const overlay = this.pageEls[caret.page]?.querySelector('.overlay')
    if (!(overlay instanceof HTMLElement)) return
    const mark = document.createElement('div')
    mark.className = 'caret'
    mark.setAttribute('aria-hidden', 'true')
    mark.style.left = `${caret.x}px`
    mark.style.top = `${caret.top}px`
    mark.style.height = `${caret.height}px`
    overlay.append(mark)
  }

  private placeSink(caret: CaretRect, sink: HTMLTextAreaElement): void {
    const page = this.pageEls[caret.page]
    if (!page) return
    if (sink.parentElement !== page) page.append(sink)
    sink.classList.add('ime-sink')
    sink.style.left = `${caret.x}px`
    sink.style.top = `${caret.top}px`
    const linePx = `${caret.height}px`
    sink.style.height = `${Math.max(caret.height, 16)}px`
    sink.style.fontSize = linePx
    sink.style.lineHeight = linePx
    sink.style.width = '1px'
  }

  private drawSelection(anchor: Pos, head: Pos): void {
    const index = this.index()
    const a = index.get(anchor.p)
    const b = index.get(head.p)
    if (!a || !b) return
    const reverse = a.order > b.order || (a.order === b.order && anchor.o > head.o)
    const start = reverse ? head : anchor
    const end = reverse ? anchor : head
    const startPlace = index.get(start.p)
    const endPlace = index.get(end.p)
    if (!startPlace || !endPlace) return
    this.mountPage(this.pageOf(start))
    this.mountPage(this.pageOf(end))
    this.mountContinuedRows(startPlace, endPlace)
    for (const place of index.values()) {
      if (place.order < startPlace.order || place.order > endPlace.order) continue
      const from = place.para.id === start.p ? start.o : 0
      const to = place.para.id === end.p ? end.o : flatText(place.para).length
      if (from === to) continue
      for (let i = 0; i < this.pageEls.length; i++) {
        const pageEl = this.pageEls[i]
        if (!pageEl || pageEl.dataset.mounted !== '1') continue
        const paraEl = [...blockElements(pageEl, place.para.id)].find((el) => el.classList.contains('p') && !el.closest('.repeat-head'))
        if (!(paraEl instanceof HTMLElement)) continue
        const range = rangeIn(paraEl, from, to)
        const origin = pageEl.getBoundingClientRect()
        const slice = paraEl.closest('.slice')
        const sliceRect = slice instanceof HTMLElement ? slice.getBoundingClientRect() : null
        const overlay = pageEl.querySelector('.overlay')
        if (!(overlay instanceof HTMLElement)) continue
        for (const rect of range.getClientRects()) {
          if (rect.width < 0.5 || rect.height < 0.5) continue
          let left = rect.left
          let top = rect.top
          let right = rect.right
          let bottom = rect.bottom
          if (sliceRect) {
            left = Math.max(left, sliceRect.left)
            top = Math.max(top, sliceRect.top)
            right = Math.min(right, sliceRect.right)
            bottom = Math.min(bottom, sliceRect.bottom)
          }
          if (right - left < 0.5 || bottom - top < 0.5) continue
          const mark = document.createElement('div')
          mark.className = 'sel'
          mark.setAttribute('aria-hidden', 'true')
          const z = this.zoom
          mark.style.left = `${(left - origin.left) / z}px`
          mark.style.top = `${(top - origin.top) / z}px`
          mark.style.width = `${(right - left) / z}px`
          mark.style.height = `${(bottom - top) / z}px`
          overlay.append(mark)
        }
      }
    }
  }

  private paintCells(sel: Extract<Selection, { kind: 'cells' }>, sink: HTMLTextAreaElement): void {
    const doc = this.doc
    if (!doc) return
    const table = doc.body.find((block) => block.kind === 'table' && block.id === sel.table)
    if (!table || table.kind !== 'table') return
    const r0 = Math.min(sel.anchor.row, sel.head.row)
    const r1 = Math.max(sel.anchor.row, sel.head.row)
    const c0 = Math.min(sel.anchor.cell, sel.head.cell)
    const c1 = Math.max(sel.anchor.cell, sel.head.cell)
    const row = table.rows[r0]
    const cell = row?.cells[c0]
    const para = cell?.blocks.find((block) => block.kind === 'p')
    if (para && para.kind === 'p') {
      const caret = this.caretRect({ p: para.id, o: 0 })
      this.placeSink(caret, sink)
    }
    for (const pageEl of this.pageEls) {
      if (pageEl.dataset.mounted !== '1') continue
      const tableEl = pageEl.querySelector(`.blk.t[data-id="${CSS.escape(sel.table)}"] table`)
      if (!(tableEl instanceof HTMLTableElement)) continue
      const origin = pageEl.getBoundingClientRect()
      const overlay = pageEl.querySelector('.overlay')
      if (!(overlay instanceof HTMLElement)) continue
      for (let r = r0; r <= r1; r++) {
        const tr = tableEl.rows[r]
        if (!tr) continue
        for (let c = c0; c <= c1 && c < tr.cells.length; c++) {
          const td = tr.cells[c]
          if (!td) continue
          const rect = td.getBoundingClientRect()
          const mark = document.createElement('div')
          mark.className = 'cell-sel'
          mark.setAttribute('aria-hidden', 'true')
          const z = this.zoom
          mark.style.left = `${(rect.left - origin.left) / z}px`
          mark.style.top = `${(rect.top - origin.top) / z}px`
          mark.style.width = `${rect.width / z}px`
          mark.style.height = `${rect.height / z}px`
          overlay.append(mark)
        }
      }
    }
  }

  private paintImageChrome(sel: Selection): void {
    for (const img of this.host.querySelectorAll('img.atom.dim')) img.classList.remove('dim')
    const doc = this.doc
    if (!doc || sel.kind !== 'text') return
    const hit = selectedImage(doc, sel)
    if (!hit) return
    const frame = this.imageFrame({ p: hit.para.id, o: hit.offset })
    if (!frame) return
    const overlay = frame.pageEl.querySelector('.overlay')
    if (!(overlay instanceof HTMLElement)) return
    this.drawImageFrame(overlay, frame.box)
    if (imageChrome(doc, sel) === 'handles') this.drawImageHandles(overlay, frame.box)
    const preview = this.imagePreview
    if (preview?.kind === 'resize') this.drawImageGhost(overlay, preview.box, preview.label)
    if (preview?.kind === 'move') {
      frame.img.classList.add('dim')
      if (preview.target) this.drawCaret(this.caretRect(preview.target))
    }
  }

  private imageFrame(pos: Pos): { pageEl: HTMLElement, img: HTMLElement, box: PageBox } | null {
    if (!this.doc) return null
    const pageIndex = this.pageOf(pos)
    this.mountPage(pageIndex)
    const pageEl = this.pageEls[pageIndex]
    if (!pageEl) return null
    const paraEl = this.paraElement(pageEl, pos)
    if (!paraEl) return null
    const img = paraEl.querySelector(`img.atom[data-atom="image"][data-o="${CSS.escape(String(pos.o))}"]`)
    if (!(img instanceof HTMLElement)) return null
    const origin = pageEl.getBoundingClientRect()
    const rect = img.getBoundingClientRect()
    const z = this.zoom > 0 ? this.zoom : 1
    return {
      pageEl,
      img,
      box: {
        x: (rect.left - origin.left) / z,
        y: (rect.top - origin.top) / z,
        width: rect.width / z,
        height: rect.height / z,
      },
    }
  }

  private drawImageFrame(overlay: HTMLElement, box: PageBox): void {
    const frame = document.createElement('div')
    frame.className = 'image-frame'
    frame.setAttribute('aria-hidden', 'true')
    const border = 1 / (this.zoom > 0 ? this.zoom : 1)
    frame.style.left = `${box.x}px`
    frame.style.top = `${box.y}px`
    frame.style.width = `${box.width}px`
    frame.style.height = `${box.height}px`
    frame.style.borderWidth = `${border}px`
    overlay.append(frame)
  }

  private drawImageHandles(overlay: HTMLElement, box: PageBox): void {
    const z = this.zoom > 0 ? this.zoom : 1
    for (const handle of RESIZE_HANDLES) {
      const origin = handleOrigin(box, handle, HANDLE_HIT_SCREEN_PX, z)
      const button = document.createElement('button')
      button.type = 'button'
      button.className = 'image-handle'
      button.dataset.handle = handle
      button.tabIndex = -1
      button.setAttribute('aria-label', t(`image.handle.${handle}`))
      button.style.left = `${origin.x}px`
      button.style.top = `${origin.y}px`
      button.style.width = `${HANDLE_HIT_SCREEN_PX / z}px`
      button.style.height = `${HANDLE_HIT_SCREEN_PX / z}px`
      const mark = document.createElement('span')
      mark.setAttribute('aria-hidden', 'true')
      const size = HANDLE_SCREEN_PX / z
      const border = 1 / z
      mark.style.width = `${size}px`
      mark.style.height = `${size}px`
      mark.style.borderWidth = `${border}px`
      mark.style.left = `${(HANDLE_HIT_SCREEN_PX - HANDLE_SCREEN_PX) / 2 / z}px`
      mark.style.top = `${(HANDLE_HIT_SCREEN_PX - HANDLE_SCREEN_PX) / 2 / z}px`
      button.append(mark)
      overlay.append(button)
    }
  }

  private drawImageGhost(overlay: HTMLElement, box: PageBox, label: string): void {
    const z = this.zoom > 0 ? this.zoom : 1
    const ghost = document.createElement('div')
    ghost.className = 'image-ghost'
    ghost.setAttribute('aria-hidden', 'true')
    ghost.style.left = `${box.x}px`
    ghost.style.top = `${box.y}px`
    ghost.style.width = `${box.width}px`
    ghost.style.height = `${box.height}px`
    ghost.style.borderWidth = `${1 / z}px`
    overlay.append(ghost)
    const tag = document.createElement('div')
    tag.className = 'image-size'
    tag.setAttribute('aria-hidden', 'true')
    tag.textContent = label
    tag.style.left = `${box.x}px`
    tag.style.top = `${box.y + box.height + 4 / z}px`
    tag.style.fontSize = `${11 / z}px`
    tag.style.lineHeight = `${16 / z}px`
    tag.style.padding = `0 ${5 / z}px`
    tag.style.borderWidth = `${1 / z}px`
    tag.style.borderRadius = `${3 / z}px`
    overlay.append(tag)
  }
}
