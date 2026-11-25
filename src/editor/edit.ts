// How the model changes. Pure functions only: no DOM, no rdocx, no clock reads except the timestamp apply() is handed.
import { findMatches } from './find.ts'
import type { FindOptions, TextMatch } from './find.ts'
import { MIN_IMAGE_PT } from './image-geom.ts'
import {
  comparePos,
  flatText,
  indexOf,
  marksAt,
  newParaId,
  normalizeInlines,
} from '../model.ts'
import type {
  Atom,
  Block,
  BorderLine,
  Cell,
  CellMargins,
  Doc,
  DocIndex,
  Hex,
  HighlightName,
  Inline,
  LineSpacing,
  ListDef,
  ListLevel,
  Marks,
  MediaKey,
  NumId,
  PageSetup,
  ParaId,
  ParaLoc,
  ParaLook,
  ParaProps,
  Paragraph,
  Pos,
  Pt,
  Row,
  RunLook,
  RunSrc,
  Selection,
  StyleId,
  StyleRole,
  Table,
  TableBorders,
  TableId,
  TextRun,
} from '../model.ts'

export interface EditorState {
  readonly doc: Doc
  readonly sel: Selection
  /** Marks toggled at a collapsed caret, applied to the next typed text. Cleared by any selection change. */
  readonly stored: Marks | null
  readonly history: History
}

export type Intent = 'typing' | 'deleting' | 'discrete' | 'none'

export interface Transaction {
  readonly doc: Doc
  readonly sel: Selection
  readonly stored: Marks | null
  readonly intent: Intent
}

export type Command = (state: EditorState) => Transaction | null

export interface Snapshot {
  readonly doc: Doc
  readonly sel: Selection
}

export interface HistoryEntry {
  readonly before: Snapshot
  readonly after: Snapshot
  readonly intent: Intent
  readonly para: ParaId | null
  readonly at: number
}

export interface History {
  readonly past: readonly HistoryEntry[]
  readonly future: readonly HistoryEntry[]
}

export const HISTORY_LIMIT = 500
export const COALESCE_MS = 1000

const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' })
const words = new Intl.Segmenter('zh', { granularity: 'word' })
const ROLES: readonly StyleRole[] = ['normal', 'title', 'heading1', 'heading2', 'heading3']
let tableSerial = 1

export function initialState(doc: Doc): EditorState {
  const first = movableParas(doc)[0] ?? bodyLocations(doc)[0]
  if (!first) throw new Error('document has no paragraph')
  const sel = caret({ p: first.para.id, o: 0 })
  return { doc, sel, stored: null, history: { past: [], future: [] } }
}

export function apply(state: EditorState, tr: Transaction, now: number): EditorState {
  if (tr.intent === 'none') {
    return { doc: tr.doc, sel: tr.sel, stored: tr.stored, history: state.history }
  }
  const para = singleTouched(state.doc, tr.doc)
  const past = state.history.past
  const top = past[past.length - 1]
  const coalesce = state.history.future.length === 0
    && top !== undefined
    && (tr.intent === 'typing' || tr.intent === 'deleting')
    && top.intent === tr.intent
    && top.para !== null
    && para !== null
    && top.para === para
    && now - top.at <= COALESCE_MS
  const entry: HistoryEntry = coalesce
    ? { before: top.before, after: { doc: tr.doc, sel: tr.sel }, intent: tr.intent, para, at: now }
    : { before: { doc: state.doc, sel: state.sel }, after: { doc: tr.doc, sel: tr.sel }, intent: tr.intent, para, at: now }
  const next = coalesce ? past.slice(0, -1).concat(entry) : past.concat(entry)
  const trimmed = next.length > HISTORY_LIMIT ? next.slice(next.length - HISTORY_LIMIT) : next
  return { doc: tr.doc, sel: tr.sel, stored: tr.stored, history: { past: trimmed, future: [] } }
}

export function undo(state: EditorState): EditorState | null {
  const past = state.history.past
  const entry = past[past.length - 1]
  if (!entry) return null
  return {
    doc: entry.before.doc,
    sel: entry.before.sel,
    stored: null,
    history: { past: past.slice(0, -1), future: state.history.future.concat(entry) },
  }
}

export function redo(state: EditorState): EditorState | null {
  const future = state.history.future
  const entry = future[future.length - 1]
  if (!entry) return null
  return {
    doc: entry.after.doc,
    sel: entry.after.sel,
    stored: null,
    history: { past: state.history.past.concat(entry), future: future.slice(0, -1) },
  }
}

export function setSelection(sel: Selection): Command {
  return (state) => {
    if (!selectionFits(state.doc, sel)) return null
    if (state.stored === null && sameSelection(state.sel, sel)) return null
    return { doc: state.doc, sel, stored: null, intent: 'none' }
  }
}

export type LogicalUnit = 'char' | 'word' | 'paragraph' | 'document'

export function move(direction: 'backward' | 'forward', unit: LogicalUnit, extend: boolean): Command {
  return (state) => {
    if (state.sel.kind === 'cells') {
      if (extend) return null
      const edge = direction === 'backward' ? cellEdge(state.doc, state.sel, 'start') : cellEdge(state.doc, state.sel, 'end')
      if (!edge) return null
      return select(state, caret(edge))
    }
    if (state.sel.kind !== 'text') return null
    const index = indexOf(state.doc)
    const ordered = orderRange(index, state.sel.anchor, state.sel.head)
    if (!extend && comparePos(index, ordered.from, ordered.to) !== 0) {
      return select(state, caret(direction === 'backward' ? ordered.from : ordered.to))
    }
    const origin = state.sel.head
    const next = stepPos(state.doc, origin, direction, unit)
    if (!next || (next.p === origin.p && next.o === origin.o)) return null
    const sel: Selection = extend
      ? { kind: 'text', anchor: state.sel.anchor, head: next }
      : caret(next)
    return select(state, sel)
  }
}

export const selectAll: Command = (state) => {
  const paras = movableParas(state.doc)
  const first = paras[0]
  const last = paras[paras.length - 1]
  if (!first || !last) return null
  const sel: Selection = {
    kind: 'text',
    anchor: { p: first.para.id, o: 0 },
    head: { p: last.para.id, o: flatText(last.para).length },
  }
  return select(state, sel)
}

export function selectAround(pos: Pos, unit: 'word' | 'paragraph'): Command {
  return (state) => {
    const index = indexOf(state.doc)
    let loc: ParaLoc
    try {
      loc = index.get(pos.p)
    } catch {
      return null
    }
    if (loc.story !== 'body' || isPreview(state.doc, loc)) return null
    const text = flatText(loc.para)
    const o = clamp(pos.o, text.length)
    if (unit === 'paragraph') return select(state, rangeSel(loc.para.id, 0, text.length))
    const span = wordSpan(text, o)
    return select(state, rangeSel(loc.para.id, span.start, span.end))
  }
}

export function insertText(text: string): Command {
  return (state) => {
    const cleared = replaceSelection(state)
    if (!cleared) return null
    if (text.length === 0) {
      if (cleared.deleted) return done(cleared.doc, caret(cleared.pos), null, 'deleting')
      return null
    }
    const loc = indexOf(cleared.doc).get(cleared.pos.p)
    if (loc.para.lock || loc.story !== 'body' || isPreview(cleared.doc, loc)) return null
    const at = marksAt(loc.para, cleared.pos.o)
    const marks = cleared.stored ?? at.marks
    const inserted = insertPlain(cleared.doc, cleared.pos, text, marks, at.src)
    if (!inserted) return null
    const intent = text.includes('\n') ? 'discrete' : 'typing'
    return done(inserted.doc, caret(inserted.pos), null, intent)
  }
}

export interface ReplaceCounts {
  readonly replaced: number
  readonly skipped: number
}

/**
 * How many of `matches` a replacement would change. Same text is neither replaced nor skipped.
 * A locked paragraph, a preview, a hyperlink, or a span that is not plain text is skipped.
 */
export function replaceCounts(doc: Doc, matches: readonly TextMatch[], replacement: string): ReplaceCounts {
  let replaced = 0
  let skipped = 0
  for (const match of matches) {
    const action = spanAction(doc, match, replacement)
    if (action === 'replace') replaced += 1
    else if (action === 'skip') skipped += 1
  }
  return { replaced, skipped }
}

/** Replace one match and select `next` when it is given, or a caret at the end of the replacement. One undo entry. */
export function replaceMatch(match: TextMatch, replacement: string, next: TextMatch | null): Command {
  return (state) => applyReplacement(state, [match], replacement, next, false)
}

/** Replace every match of `needle` in one transaction. A replacement that contains `needle` is not searched again. */
export function replaceAllText(needle: string, replacement: string, options: FindOptions): Command {
  return (state) => applyReplacement(state, findMatches(state.doc, needle, options), replacement, null, true)
}

function applyReplacement(
  state: EditorState,
  matches: readonly TextMatch[],
  replacement: string,
  next: TextMatch | null,
  all: boolean,
): Transaction | null {
  const planned = matches.flatMap((match) => {
    const action = spanAction(state.doc, match, replacement)
    return action === 'replace' ? [match] : []
  })
  if (planned.length === 0) return null
  const doc = applySpans(state.doc, planned, replacement)
  if (doc === state.doc) return null
  const first = planned[0]
  if (!first) return null
  const sel = !all && next
    ? rangeSel(next.para, next.start, next.end)
    : caret({ p: first.para, o: first.start + replacement.length })
  return done(doc, sel, null, 'discrete')
}

function applySpans(doc: Doc, matches: readonly TextMatch[], replacement: string): Doc {
  const byPara = new Map<ParaId, TextMatch[]>()
  for (const match of matches) {
    const list = byPara.get(match.para)
    if (list) list.push(match)
    else byPara.set(match.para, [match])
  }
  let next = doc
  for (const [id, list] of byPara) {
    const loc = safeLoc(indexOf(next), id)
    if (!loc) continue
    let para = loc.para
    const ordered = list.slice().sort((a, b) => b.start - a.start)
    for (const match of ordered) {
      const rewritten = rewriteSpan(para, match.start, match.end, replacement)
      if (rewritten) para = rewritten
    }
    next = replaceParagraph(next, para)
  }
  return next
}

type SpanAction = 'replace' | 'skip' | 'same'

function spanAction(doc: Doc, match: TextMatch, replacement: string): SpanAction {
  const loc = safeLoc(indexOf(doc), match.para)
  if (!loc || loc.story !== 'body' || isPreview(doc, loc) || loc.para.lock) return 'skip'
  const text = flatText(loc.para)
  if (!plainSpan(text, match.start, match.end)) return 'skip'
  if (touchesLink(loc.para, match.start, match.end)) return 'skip'
  if (text.slice(match.start, match.end) === replacement) return 'same'
  return 'replace'
}

function plainSpan(text: string, start: number, end: number): boolean {
  if (start < 0 || end > text.length || start >= end) return false
  if (splitsPair(text, start) || splitsPair(text, end)) return false
  return !text.slice(start, end).includes('\uFFFC')
}

function splitsPair(text: string, offset: number): boolean {
  if (offset <= 0 || offset >= text.length) return false
  const before = text.charCodeAt(offset - 1)
  const after = text.charCodeAt(offset)
  return before >= 0xD800 && before <= 0xDBFF && after >= 0xDC00 && after <= 0xDFFF
}

function touchesLink(para: Paragraph, start: number, end: number): boolean {
  let cursor = 0
  for (const inline of para.inlines) {
    const len = inline.kind === 'text' ? inline.text.length : 1
    const from = cursor
    const to = cursor + len
    cursor = to
    if (to <= start || from >= end) continue
    if (inline.kind === 'text' && inline.marks.link !== undefined) return true
  }
  return false
}

function runCovering(para: Paragraph, offset: number): TextRun | null {
  let cursor = 0
  for (const inline of para.inlines) {
    const len = inline.kind === 'text' ? inline.text.length : 1
    if (offset < cursor + len) return inline.kind === 'text' ? inline : null
    cursor += len
  }
  return null
}

function rewriteSpan(para: Paragraph, start: number, end: number, replacement: string): Paragraph | null {
  const run = runCovering(para, start)
  const marks = run ? run.marks : {}
  const src = run ? run.src : null
  let inlines = deleteInlines(para.inlines, start, end)
  if (replacement.length > 0) {
    const inserted: TextRun = { kind: 'text', text: replacement, marks, src }
    inlines = insertInlines(inlines, start, [inserted])
  }
  const normalized = normalizeInlines(inlines)
  if (sameInlines(para.inlines, normalized)) return null
  return { ...para, inlines: normalized }
}

export const insertLineBreak: Command = (state) => {
  const cleared = replaceSelection(state)
  if (!cleared) return null
  const loc = indexOf(cleared.doc).get(cleared.pos.p)
  if (loc.para.lock || isPreview(cleared.doc, loc)) return null
  const at = marksAt(loc.para, cleared.pos.o)
  const marks = cleared.stored ?? at.marks
  const inserted = insertString(cleared.doc, cleared.pos, '\n', marks, at.src)
  if (!inserted) return null
  return done(inserted.doc, caret(inserted.pos), null, 'typing')
}

export const insertPageBreak: Command = (state) => {
  const cleared = replaceSelection(state)
  if (!cleared) return null
  const loc = indexOf(cleared.doc).get(cleared.pos.p)
  if (loc.para.lock || isPreview(cleared.doc, loc)) return null
  const atom: Atom = { kind: 'atom', src: null, atom: { type: 'pageBreak' } }
  const inserted = insertAtom(cleared.doc, cleared.pos, atom)
  if (!inserted) return null
  return done(inserted.doc, caret(inserted.pos), null, 'discrete')
}

export const splitParagraph: Command = (state) => {
  const cleared = replaceSelection(state)
  if (!cleared) return null
  const split = splitAt(cleared.doc, cleared.pos, true)
  if (!split) return null
  return done(split.doc, caret(split.pos), null, 'discrete')
}

export function mergeSide(first: Paragraph, second: Paragraph, deletionStart: number): 'first' | 'second' {
  // Word keeps the following paragraph when the deletion starts at the first paragraph, or the first is empty.
  void second
  if (deletionStart === 0 || flatText(first).length === 0) return 'second'
  return 'first'
}

export function deleteBackward(unit: 'char' | 'word'): Command {
  return (state) => {
    if (state.sel.kind === 'cells') return deleteSelection(state)
    if (state.sel.kind !== 'text') return null
    const index = indexOf(state.doc)
    if (comparePos(index, state.sel.anchor, state.sel.head) !== 0) return deleteSelection(state)
    const pos = state.sel.head
    const loc = safeLoc(index, pos.p)
    if (!loc || loc.story !== 'body' || isPreview(state.doc, loc) || loc.para.lock) return null
    if (pos.o === 0) {
      if (listInfo(state.doc, loc.para)) {
        const para = withNumbering(loc.para, null)
        return done(replaceParagraph(state.doc, para), caret(pos), null, 'discrete')
      }
      return mergeSibling(state, loc, 'backward')
    }
    const text = flatText(loc.para)
    const start = unit === 'char' ? prevBoundary(text, pos.o) : prevWordStop(text, pos.o)
    const end = Math.max(pos.o, snapForward(text, pos.o))
    return deleteInside(state, loc.para, start, end, start)
  }
}

export function deleteForward(unit: 'char' | 'word'): Command {
  return (state) => {
    if (state.sel.kind === 'cells') return deleteSelection(state)
    if (state.sel.kind !== 'text') return null
    const index = indexOf(state.doc)
    if (comparePos(index, state.sel.anchor, state.sel.head) !== 0) return deleteSelection(state)
    const pos = state.sel.head
    const loc = safeLoc(index, pos.p)
    if (!loc || loc.story !== 'body' || isPreview(state.doc, loc) || loc.para.lock) return null
    const text = flatText(loc.para)
    if (pos.o >= text.length) return mergeSibling(state, loc, 'forward')
    const end = unit === 'char' ? nextBoundary(text, pos.o) : nextWordStop(text, pos.o)
    const start = snapBack(text, pos.o)
    return deleteInside(state, loc.para, start, end, start)
  }
}

export const deleteSelection: Command = (state) => {
  if (state.sel.kind === 'cells') return clearSelectedCells(state)
  if (state.sel.kind !== 'text') return null
  const index = indexOf(state.doc)
  if (comparePos(index, state.sel.anchor, state.sel.head) === 0) return null
  const removed = deleteRange(state.doc, state.sel.anchor, state.sel.head)
  if (!removed) return null
  return done(removed.doc, caret(removed.pos), null, 'deleting')
}

/** Drops a non-empty selection at the start of an IME composition, as a typing step so the commit coalesces with it. */
export function compositionDelete(): Command {
  return (state) => {
    const tr = deleteSelection(state)
    if (!tr) return null
    return { ...tr, intent: 'typing' }
  }
}

export type ToggleMark = 'bold' | 'italic' | 'underline' | 'strike'

export function toggleMark(mark: ToggleMark): Command {
  return (state) => {
    if (state.sel.kind !== 'text') return null
    const index = indexOf(state.doc)
    const ordered = orderRange(index, state.sel.anchor, state.sel.head)
    if (comparePos(index, ordered.from, ordered.to) === 0) {
      const loc = ordered.fromLoc
      if (loc.story !== 'body') return null
      const base = state.stored ?? marksAt(loc.para, ordered.from.o).marks
      const stored = base[mark] === true ? { ...base, [mark]: false } : { ...base, [mark]: true }
      if (state.stored && marksEqual(state.stored, stored)) return null
      return { doc: state.doc, sel: state.sel, stored, intent: 'none' }
    }
    const on = coveredText(state.doc, ordered).some(sample => sample.marks[mark] !== true)
    const doc = mapCovered(state.doc, ordered, (marks) => ({ ...marks, [mark]: on }))
    if (!doc) return null
    return done(doc, state.sel, null, 'discrete')
  }
}

export type MarkValue =
  | { readonly mark: 'font'; readonly value: string | undefined }
  | { readonly mark: 'fontEastAsia'; readonly value: string | undefined }
  | { readonly mark: 'sizePt'; readonly value: Pt | undefined }
  | { readonly mark: 'color'; readonly value: Hex | undefined }
  | { readonly mark: 'highlight'; readonly value: HighlightName | undefined }
  | { readonly mark: 'vertAlign'; readonly value: 'super' | 'sub' | undefined }

export function setMark(change: MarkValue): Command {
  return (state) => {
    if (state.sel.kind !== 'text') return null
    const index = indexOf(state.doc)
    const ordered = orderRange(index, state.sel.anchor, state.sel.head)
    if (comparePos(index, ordered.from, ordered.to) === 0) {
      const loc = ordered.fromLoc
      if (loc.story !== 'body') return null
      const base = state.stored ?? marksAt(loc.para, ordered.from.o).marks
      const stored = applyMarkValue(base, change)
      return { doc: state.doc, sel: state.sel, stored, intent: 'none' }
    }
    const doc = mapCovered(state.doc, ordered, (marks) => applyMarkValue(marks, change))
    if (!doc) return null
    return done(doc, state.sel, null, 'discrete')
  }
}

export function setParaStyle(style: StyleId): Command {
  return (state) => {
    const doc0 = ensureStyle(state.doc, style)
    const doc = mapTouched(doc0, state.sel, (para) => (
      para.props.style === style ? para : { ...para, props: { ...para.props, style } }
    ))
    if (!doc) return null
    return done(doc, state.sel, null, 'discrete')
  }
}

export function setAlign(align: NonNullable<ParaProps['align']>): Command {
  return (state) => {
    const doc = mapTouched(state.doc, state.sel, (para) => (
      para.props.align === align ? para : { ...para, props: { ...para.props, align } }
    ))
    if (!doc) return null
    return done(doc, state.sel, null, 'discrete')
  }
}

export function setLineSpacing(line: LineSpacing): Command {
  return (state) => {
    const doc = mapTouched(state.doc, state.sel, (para) => {
      const spacing = para.props.spacing
      if (spacing?.line && sameLine(spacing.line, line)) return para
      return { ...para, props: { ...para.props, spacing: { ...spacing, line } } }
    })
    if (!doc) return null
    return done(doc, state.sel, null, 'discrete')
  }
}

const DIRECT_MARKS = [
  'bold', 'italic', 'underline', 'strike', 'font', 'fontEastAsia', 'sizePt', 'color', 'highlight', 'vertAlign', 'charStyle',
] as const

/**
 * Removes direct character formatting from the selection and leaves the hyperlink.
 * Paragraph properties stay as they are: unlike Word's button, this does not reset the paragraph style.
 * A collapsed caret stores empty marks, so the next typed text does not inherit the caret's formatting.
 */
export const clearMarks: Command = (state) => {
  if (state.sel.kind !== 'text') return null
  const index = indexOf(state.doc)
  const ordered = orderRange(index, state.sel.anchor, state.sel.head)
  if (comparePos(index, ordered.from, ordered.to) === 0) {
    if (ordered.fromLoc.story !== 'body') return null
    const stored: Marks = {}
    if (state.stored && marksEqual(state.stored, stored)) return null
    return { doc: state.doc, sel: state.sel, stored, intent: 'none' }
  }
  const doc = mapCovered(state.doc, ordered, clearedMarks)
  if (!doc) return null
  return done(doc, state.sel, null, 'discrete')
}

/** Superscript and subscript share one mark, so turning one on replaces the other. Choosing the active one clears it. */
export function toggleVertAlign(which: 'super' | 'sub'): Command {
  return (state) => {
    if (state.sel.kind !== 'text') return null
    const index = indexOf(state.doc)
    const ordered = orderRange(index, state.sel.anchor, state.sel.head)
    if (comparePos(index, ordered.from, ordered.to) === 0) {
      if (ordered.fromLoc.story !== 'body') return null
      const base = state.stored ?? marksAt(ordered.fromLoc.para, ordered.from.o).marks
      const stored = applyMarkValue(base, { mark: 'vertAlign', value: base.vertAlign === which ? undefined : which })
      if (state.stored && marksEqual(state.stored, stored)) return null
      return { doc: state.doc, sel: state.sel, stored, intent: 'none' }
    }
    const on = coveredText(state.doc, ordered).some(sample => sample.marks.vertAlign !== which)
    const doc = mapCovered(state.doc, ordered, (marks) => applyMarkValue(marks, {
      mark: 'vertAlign',
      value: on ? which : undefined,
    }))
    if (!doc) return null
    return done(doc, state.sel, null, 'discrete')
  }
}

/** null removes that direct value. An omitted field is left alone. */
export interface IndentEdit {
  readonly left?: Pt
  readonly right?: Pt
  readonly firstLine?: Pt | null
  readonly hanging?: Pt | null
  readonly firstLineChars?: number | null
}

export interface SpacingEdit {
  readonly before?: Pt
  readonly after?: Pt
  readonly line?: LineSpacing
}

/** The paragraph dialog applies one of these as a single undo entry. */
export interface ParaFormat {
  readonly align?: NonNullable<ParaProps['align']>
  readonly indent?: IndentEdit
  readonly spacing?: SpacingEdit
}

/** Sets spacing before and after. An omitted edge is left alone, including the line rule. */
export function setSpacing(spacing: { readonly before?: Pt, readonly after?: Pt }): Command {
  return (state) => {
    if (spacing.before === undefined && spacing.after === undefined) return null
    return applyParaFormat(state, { spacing })
  }
}

/** Alignment, indent and spacing in one history entry. An empty patch changes nothing. */
export function setParaFormat(format: ParaFormat): Command {
  return (state) => applyParaFormat(state, format)
}

export function setIndent(indent: NonNullable<ParaProps['indent']>): Command {
  return (state) => {
    const doc = mapTouched(state.doc, state.sel, (para) => {
      const prev = para.props.indent
      const next = { ...prev, ...indent }
      if (indent.firstLineChars === undefined && prev?.firstLineChars !== undefined) {
        next.firstLineChars = prev.firstLineChars
      }
      return { ...para, props: { ...para.props, indent: next } }
    })
    if (!doc) return null
    return done(doc, state.sel, null, 'discrete')
  }
}

export function changeIndentLevel(delta: 1 | -1): Command {
  return (state) => {
    const doc = mapTouched(state.doc, state.sel, (para) => changeLevel(state.doc, para, delta))
    if (!doc) return null
    return done(doc, state.sel, null, 'discrete')
  }
}

/**
 * Tab moves to the next table cell, or raises a list level, or inserts a tab character.
 * A list inside a cell follows the cell, so the table is decided first.
 * Shift+Tab outside a table or a list changes nothing; the chord stays bound so focus stays in the editor.
 */
export function tabKey(direction: 1 | -1): Command {
  return (state) => {
    const hit = tableHit(state)
    if (hit) return moveTabCell(state, hit, direction)
    if (touchedParagraphs(state.doc, state.sel).some((para) => listInfo(state.doc, para) !== null)) {
      return changeIndentLevel(direction)(state)
    }
    if (direction < 0) return null
    return insertText('\t')(state)
  }
}

function moveTabCell(
  state: EditorState,
  hit: { table: Table, rows: { start: number, end: number }, cell: number },
  direction: 1 | -1,
): Transaction | null {
  const at = tabCell(state, hit)
  const next = neighborCell(hit.table, at.row, at.cell, direction)
  if (!next) {
    if (direction < 0) return null
    return insertRow('below')(state)
  }
  const blocks = hit.table.rows[next.row]?.cells[next.cell]?.blocks
  if (!blocks) return null
  const pos = firstPosIn(blocks)
  if (!pos) return null
  return setSelection(caret(pos))(state)
}

function tabCell(
  state: EditorState,
  hit: { rows: { start: number, end: number }, cell: number },
): { row: number, cell: number } {
  if (state.sel.kind === 'cells') {
    const anchor = state.sel.anchor
    const head = state.sel.head
    const headLater = head.row > anchor.row || (head.row === anchor.row && head.cell >= anchor.cell)
    const later = headLater ? head : anchor
    return { row: later.row, cell: later.cell }
  }
  return { row: hit.rows.end, cell: hit.cell }
}

function neighborCell(
  table: Table,
  row: number,
  cell: number,
  direction: 1 | -1,
): { row: number, cell: number } | null {
  const rows = table.rows
  if (direction > 0) {
    const current = rows[row]
    if (current && cell + 1 < current.cells.length) return { row, cell: cell + 1 }
    if (row + 1 < rows.length && (rows[row + 1]?.cells.length ?? 0) > 0) return { row: row + 1, cell: 0 }
    return null
  }
  if (cell > 0) return { row, cell: cell - 1 }
  if (row <= 0) return null
  const prev = rows[row - 1]
  if (!prev || prev.cells.length === 0) return null
  return { row: row - 1, cell: prev.cells.length - 1 }
}

export function toggleList(kind: 'bullet' | 'decimal'): Command {
  return (state) => {
    const touched = touchedParagraphs(state.doc, state.sel).filter(para => !para.lock)
    if (touched.length === 0) return null
    const allOn = touched.every(para => listKindOf(state.doc, para) === kind)
    if (allOn) {
      const doc = mapTouched(state.doc, state.sel, (para) => withNumbering(para, null))
      if (!doc) return null
      return done(doc, state.sel, null, 'discrete')
    }
    const neighbor = neighborList(state.doc, touched, kind)
    let doc = state.doc
    let num: NumId
    let level: number
    if (neighbor) {
      num = neighbor.num
      level = neighbor.level
    } else {
      num = freshNumId(doc)
      level = 0
      const byNum = new Map(doc.lists.byNum)
      byNum.set(num, freshList(kind))
      doc = { ...doc, lists: { byNum } }
    }
    const next = mapTouched(doc, state.sel, (para) => withNumbering(para, { num, level }))
    if (!next) return null
    return done(next, state.sel, null, 'discrete')
  }
}

export function insertTable(rows: number, cols: number): Command {
  return (state) => {
    if (rows < 1 || cols < 1) return null
    const cleared = replaceSelection(state)
    if (!cleared) return null
    const loc = safeLoc(indexOf(cleared.doc), cleared.pos.p)
    if (!loc || loc.para.lock || isPreview(cleared.doc, loc)) return null
    const table = makeTable(cleared.doc, loc, rows, cols)
    const placed = placeBlock(cleared.doc, loc, cleared.pos.o, table)
    if (!placed) return null
    const pos = firstPosIn(table.rows[0]?.cells[0]?.blocks ?? [])
    if (!pos) return null
    return done(placed, caret(pos), null, 'discrete')
  }
}

export function insertRow(where: 'above' | 'below'): Command {
  return (state) => editRows(state, (table, rows) => {
    const at = where === 'above' ? rows.start : rows.end + 1
    const template = table.rows[where === 'above' ? rows.start : rows.end] ?? table.rows[0]
    if (!template) return null
    const row = blankRow(template, table.grid.length)
    const nextRows = table.rows.slice()
    nextRows.splice(at, 0, row)
    const para = row.cells[0]?.blocks[0]
    if (!para || para.kind !== 'p') return null
    return { table: { ...table, rows: nextRows }, pos: { p: para.id, o: 0 } }
  })
}

export function deleteRows(): Command {
  return (state) => editRows(state, (table, rows) => {
    const kept = table.rows.filter((row, index) => index < rows.start || index > rows.end || rowHasProtected(row))
    if (kept.length === table.rows.length) return null
    if (kept.length === 0) return { table: null, pos: null }
    return { table: { ...table, rows: kept }, pos: firstPosIn(kept[0]?.cells[0]?.blocks ?? []) }
  })
}

export function insertColumn(where: 'left' | 'right'): Command {
  return (state) => editColumns(state, (table, cols) => {
    const at = where === 'left' ? cols.start : cols.end + 1
    const next = insertGridColumn(table, at)
    const pos = firstPosIn(next.rows[0]?.cells[0]?.blocks ?? [])
    return pos ? { table: next, pos } : null
  })
}

export function deleteColumns(): Command {
  return (state) => editColumns(state, (table, cols) => {
    let next: Table | null = table
    for (let col = cols.end; col >= cols.start; col--) {
      if (!next) break
      next = deleteGridColumn(next, col)
    }
    if (next === table) return null
    if (!next) return { table: null, pos: null }
    return { table: next, pos: firstPosIn(next.rows[0]?.cells[0]?.blocks ?? []) }
  })
}

export function deleteTable(table: TableId): Command {
  return (state) => {
    const edited = removeTable(state.doc, table)
    if (!edited) return null
    return done(edited.doc, caret(edited.pos), null, 'discrete')
  }
}

export function deleteEnclosingTable(): Command {
  return (state) => {
    const hit = tableHit(state)
    if (!hit) return null
    return deleteTable(hit.table.id)(state)
  }
}

export interface MediaRef {
  readonly key: MediaKey
  readonly contentType: string
  readonly widthPx: number
  readonly heightPx: number
}

export function insertImage(media: MediaRef): Command {
  return (state) => {
    const cleared = replaceSelection(state)
    if (!cleared) return null
    const loc = safeLoc(indexOf(cleared.doc), cleared.pos.p)
    if (!loc || loc.para.lock || isPreview(cleared.doc, loc)) return null
    const scaled = scaleImage(cleared.doc, loc, media)
    const atom: Atom = {
      kind: 'atom',
      src: null,
      atom: { type: 'image', media: media.key, widthPt: asPt(scaled.width), heightPt: asPt(scaled.height) },
    }
    let doc = cleared.doc
    if (!doc.media.byKey.has(media.key)) {
      const byKey = new Map(doc.media.byKey)
      byKey.set(media.key, { contentType: media.contentType, origin: 'added' })
      doc = { ...doc, media: { byKey } }
    }
    const inserted = insertAtom(doc, cleared.pos, atom)
    if (!inserted) return null
    return done(inserted.doc, caret(inserted.pos), null, 'discrete')
  }
}

export type ImageMoveIssue = 'missing' | 'file' | 'locked' | 'preview' | 'story' | 'same'

export function findImage(doc: Doc, pos: Pos): Atom | null {
  const loc = safeLoc(indexOf(doc), pos.p)
  if (!loc) return null
  return imageAtomAt(loc.para, pos.o)
}

export function selectedImage(doc: Doc, sel: Selection): { para: Paragraph, offset: number, atom: Atom } | null {
  if (sel.kind !== 'text') return null
  let ordered: Ordered
  try {
    ordered = orderRange(indexOf(doc), sel.anchor, sel.head)
  } catch {
    return null
  }
  if (ordered.from.p !== ordered.to.p || ordered.to.o !== ordered.from.o + 1) return null
  const atom = imageAtomAt(ordered.fromLoc.para, ordered.from.o)
  if (!atom) return null
  return { para: ordered.fromLoc.para, offset: ordered.from.o, atom }
}

/** Outline while one picture is selected. Handles only when that picture can be edited. */
export function imageChrome(doc: Doc, sel: Selection): 'none' | 'outline' | 'handles' {
  const hit = selectedImage(doc, sel)
  if (!hit) return 'none'
  const loc = safeLoc(indexOf(doc), hit.para.id)
  if (!loc || loc.story !== 'body' || loc.para.lock || isPreview(doc, loc)) return 'outline'
  return 'handles'
}

export function imageMaxWidth(doc: Doc, id: ParaId): number {
  const loc = safeLoc(indexOf(doc), id)
  if (!loc) return contentWidth(doc, 0)
  return maxImageWidth(doc, loc)
}

export function imageResizeIssue(doc: Doc, pos: Pos, widthPt: number, heightPt: number): 'missing' | 'range' | null {
  const image = findImage(doc, pos)
  if (!image || image.atom.type !== 'image') return 'missing'
  if (!imageSizeAllowed(image.atom.widthPt, image.atom.heightPt, widthPt, heightPt, imageMaxWidth(doc, pos.p))) {
    return 'range'
  }
  return null
}

export function imageMoveIssue(doc: Doc, from: Pos, to: Pos): ImageMoveIssue | null {
  const image = findImage(doc, from)
  if (!image) return 'missing'
  if (image.src !== null) return 'file'
  const index = indexOf(doc)
  const fromLoc = safeLoc(index, from.p)
  const toLoc = safeLoc(index, to.p)
  if (!fromLoc || !toLoc) return 'missing'
  if (fromLoc.story !== 'body' || toLoc.story !== 'body') return 'story'
  if (isPreview(doc, fromLoc) || isPreview(doc, toLoc)) return 'preview'
  if (fromLoc.para.lock || toLoc.para.lock) return 'locked'
  const dest = imageInsertAt(toLoc.para, to.o)
  if (from.p === to.p && (dest === from.o || dest === from.o + 1)) return 'same'
  return null
}

export function imageDropTarget(doc: Doc, from: Pos, to: Pos): boolean {
  return imageMoveIssue(doc, from, to) === null
}

export function resizeImage(pos: Pos, widthPt: number, heightPt: number): Command {
  return (state) => {
    const loc = safeLoc(indexOf(state.doc), pos.p)
    if (!loc || loc.story !== 'body' || loc.para.lock || isPreview(state.doc, loc)) return null
    const image = imageAtomAt(loc.para, pos.o)
    if (!image || image.atom.type !== 'image') return null
    const max = maxImageWidth(state.doc, loc)
    if (!imageSizeAllowed(image.atom.widthPt, image.atom.heightPt, widthPt, heightPt, max)) return null
    const range = imageRange(loc.para.id, pos.o)
    if (Math.abs(widthPt - image.atom.widthPt) < 1e-4 && Math.abs(heightPt - image.atom.heightPt) < 1e-4) {
      return done(state.doc, range, null, 'none')
    }
    const next: Atom = {
      ...image,
      atom: { type: 'image', media: image.atom.media, widthPt: asPt(widthPt), heightPt: asPt(heightPt) },
    }
    const inlines = replaceAtomAt(loc.para.inlines, pos.o, next)
    const doc = replaceParagraph(state.doc, { ...loc.para, inlines })
    return done(doc, range, null, 'discrete')
  }
}

export function moveAtom(from: Pos, to: Pos): Command {
  return (state) => {
    if (imageMoveIssue(state.doc, from, to)) return null
    const index = indexOf(state.doc)
    const fromLoc = safeLoc(index, from.p)
    const toLoc = safeLoc(index, to.p)
    const image = fromLoc ? imageAtomAt(fromLoc.para, from.o) : null
    if (!fromLoc || !toLoc || !image || image.atom.type !== 'image') return null
    const dest = imageInsertAt(toLoc.para, to.o)
    const fitted = fitMovedImage(state.doc, toLoc, image)
    if (from.p === to.p) {
      const removed = deleteInlines(fromLoc.para.inlines, from.o, from.o + 1)
      const insertAt = dest > from.o ? dest - 1 : dest
      const inlines = normalizeInlines(insertInlines(removed, insertAt, [fitted]))
      if (sameInlines(fromLoc.para.inlines, inlines)) return null
      const doc = replaceParagraph(state.doc, { ...fromLoc.para, inlines })
      return done(doc, imageRange(fromLoc.para.id, insertAt), null, 'discrete')
    }
    const stripped = normalizeInlines(deleteInlines(fromLoc.para.inlines, from.o, from.o + 1))
    let doc = replaceParagraph(state.doc, { ...fromLoc.para, inlines: stripped })
    const target = safeLoc(indexOf(doc), to.p)
    if (!target) return null
    const inlines = normalizeInlines(insertInlines(target.para.inlines, dest, [fitted]))
    doc = replaceParagraph(doc, { ...target.para, inlines })
    return done(doc, imageRange(target.para.id, dest), null, 'discrete')
  }
}

function imageAtomAt(para: Paragraph, offset: number): Atom | null {
  let cursor = 0
  for (const inline of para.inlines) {
    if (inline.kind === 'atom' && inline.atom.type === 'image' && cursor === offset) return inline
    cursor += inlineLen(inline)
  }
  return null
}

function imageRange(id: ParaId, offset: number): Selection {
  return { kind: 'text', anchor: { p: id, o: offset }, head: { p: id, o: offset + 1 } }
}

function imageInsertAt(para: Paragraph, offset: number): number {
  const length = inlineLength(para.inlines)
  return snapBack(flatText(para), clampOffset(offset, length))
}

function clampOffset(offset: number, length: number): number {
  if (offset < 0) return 0
  if (offset > length) return length
  return offset
}

function maxImageWidth(doc: Doc, loc: ParaLoc): number {
  const cell = cellWidth(doc, loc)
  if (cell !== null && cell > 0) return cell
  return contentWidth(doc, sectionIndex(doc, loc))
}

function imageSizeAllowed(
  currentWidth: number,
  currentHeight: number,
  width: number,
  height: number,
  maxWidth: number,
): boolean {
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return false
  if (width > maxWidth + 1e-6) return false
  const widthOk = width + 1e-6 >= MIN_IMAGE_PT || Math.abs(width - currentWidth) < 1e-4
  const heightOk = height + 1e-6 >= MIN_IMAGE_PT || Math.abs(height - currentHeight) < 1e-4
  return widthOk && heightOk
}

function replaceAtomAt(inlines: readonly Inline[], offset: number, atom: Atom): readonly Inline[] {
  const out: Inline[] = []
  let cursor = 0
  let replaced = false
  for (const inline of inlines) {
    if (!replaced && inline.kind === 'atom' && cursor === offset) {
      out.push(atom)
      replaced = true
    } else {
      out.push(inline)
    }
    cursor += inlineLen(inline)
  }
  return replaced ? out : inlines
}

function fitMovedImage(doc: Doc, loc: ParaLoc, atom: Atom): Atom {
  if (atom.atom.type !== 'image') return atom
  const max = maxImageWidth(doc, loc)
  const width = atom.atom.widthPt
  const height = atom.atom.heightPt
  if (!(width > max) || !(width > 0)) return atom
  const scale = max / width
  return {
    ...atom,
    atom: { type: 'image', media: atom.atom.media, widthPt: asPt(width * scale), heightPt: asPt(height * scale) },
  }
}

export function setPageSetup(section: number, page: PageSetup): Command {
  return (state) => {
    const current = state.doc.sections[section]
    if (!current) return null
    const sections = state.doc.sections.slice()
    sections[section] = { ...current, page }
    return done({ ...state.doc, sections }, state.sel, null, 'discrete')
  }
}

export type Tri = boolean | 'mixed'

export interface ToolbarState {
  readonly bold: Tri
  readonly italic: Tri
  readonly underline: Tri
  readonly strike: Tri
  readonly font: string | 'mixed'
  readonly sizePt: Pt | 'mixed'
  readonly color: Hex | 'auto' | 'mixed'
  readonly highlight: HighlightName | 'none' | 'mixed'
  readonly vertAlign: 'super' | 'sub' | 'none' | 'mixed'
  /** The shared auto multiple of every touched paragraph, or null when they do not share one. */
  readonly lineMultiple: number | null
  /** Direct spacing of the paragraph at the caret, in points. A missing value is 0. */
  readonly spaceBefore: number
  readonly spaceAfter: number
  readonly style: StyleId | 'mixed'
  readonly align: NonNullable<ParaProps['align']> | 'mixed'
  readonly list: 'bullet' | 'decimal' | 'none' | 'mixed'
  readonly inTable: boolean
  readonly rowTools: boolean
  readonly columnTools: boolean
  readonly canUndo: boolean
  readonly canRedo: boolean
}

// insertColumn and deleteColumns rewrite span when the edited column falls inside a cell that already
// covers more than one column, so one such cell turns column tools off for the whole table. insertRow
// copies span onto a new row and clears vMerge there; deleteRows drops rows. Neither writes span or
// vMerge on a cell that stays, so a vertical merge does not turn tools off.
export function tableEditTools(table: Table): { readonly rows: boolean, readonly columns: boolean } {
  let columns = true
  for (const row of table.rows) {
    for (const cell of row.cells) {
      if (cell.span > 1) columns = false
    }
  }
  return { rows: true, columns }
}

export function toolbarState(state: EditorState, look: (para: Paragraph, marks: Marks) => RunLook): ToolbarState {
  const paras = touchedParagraphs(state.doc, state.sel)
  const looks = runLooks(state, look)
  const bold = tri(looks.map(item => item.bold))
  const italic = tri(looks.map(item => item.italic))
  const underline = tri(looks.map(item => item.underline))
  const strike = tri(looks.map(item => item.strike))
  const fontKey = looks.map(item => `${item.fonts.ascii}\0${item.fonts.eastAsia}`)
  const font = uniform(fontKey) === 'mixed' ? 'mixed' : (looks[0]?.fonts.eastAsia ?? 'mixed')
  const sizePt = uniform(looks.map(item => item.sizePt))
  const color = uniform(looks.map(item => item.color))
  const highlight = uniform(looks.map(item => item.highlight ?? 'none'))
  const vertAlign = vertAlignOf(markSamples(state))
  const lineMultiple = sharedMultiple(paras)
  const caret = caretParagraph(state)
  const spaceBefore = Number(caret?.props.spacing?.before ?? 0)
  const spaceAfter = Number(caret?.props.spacing?.after ?? 0)
  const style = uniform(paras.map(para => para.props.style ?? state.doc.styles.defaultParagraph))
  const align = uniform(paras.map(para => effectiveAlign(state.doc, para)))
  const list = uniform(paras.map(para => listKindOf(state.doc, para)))
  const inTable = state.sel.kind === 'cells' || paras.some(para => paragraphInTable(state.doc, para.id))
  const hit = tableHit(state)
  const tools = hit ? tableEditTools(hit.table) : { rows: false, columns: false }
  return {
    bold, italic, underline, strike, font, sizePt, color, highlight, vertAlign, lineMultiple,
    spaceBefore, spaceAfter, style, align, list, inTable,
    rowTools: tools.rows,
    columnTools: tools.columns,
    canUndo: state.history.past.length > 0,
    canRedo: state.history.future.length > 0,
  }
}

export function selectedText(state: EditorState): string {
  if (state.sel.kind === 'cells') return cellsText(state.doc, state.sel)
  const index = indexOf(state.doc)
  const ordered = orderRange(index, state.sel.anchor, state.sel.head)
  if (comparePos(index, ordered.from, ordered.to) === 0) return ''
  const parts: { row: string, cell: string, text: string }[] = []
  for (const loc of bodyLocations(state.doc)) {
    if (!paragraphSelected(loc, ordered)) continue
    const covered = coveredOffsets(loc, ordered)
    const text = covered ? plainSlice(loc.para, covered[0], covered[1]) : ''
    const addr = tableAddress(state.doc, loc)
    if (!addr) parts.push({ row: `p:${loc.para.id}`, cell: '0', text })
    else parts.push({ row: `${addr.table}:${addr.row}`, cell: String(addr.cell), text })
  }
  let text = joinParts(parts)
  if (ordered.from.p !== ordered.to.p && ordered.to.o === 0) text += '\n'
  return text
}

const wordCounts = new WeakMap<Paragraph, number>()

export function wordCount(doc: Doc): number {
  let total = 0
  const walk = (blocks: readonly Block[]): void => {
    for (const block of blocks) {
      if (block.kind === 'p') total += countParagraph(block)
      else if (block.kind === 'table') {
        for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks)
      } else {
        for (const para of block.preview) total += countParagraph(para)
      }
    }
  }
  walk(doc.body)
  return total
}

function countParagraph(para: Paragraph): number {
  const cached = wordCounts.get(para)
  if (cached !== undefined) return cached
  const count = countWords(flatText(para))
  wordCounts.set(para, count)
  return count
}

function countWords(text: string): number {
  let count = 0
  let latin = false
  for (const ch of text) {
    if (ch === '\uFFFC' || /\s/u.test(ch)) {
      latin = false
      continue
    }
    const cp = ch.codePointAt(0) ?? 0
    if (isCjk(cp)) {
      latin = false
      count += 1
      continue
    }
    if (/^[\p{L}\p{N}]$/u.test(ch)) {
      if (!latin) count += 1
      latin = true
    }
  }
  return count
}

function isCjk(cp: number): boolean {
  return (cp >= 0x3400 && cp <= 0x9fff) || (cp >= 0xf900 && cp <= 0xfaff) || (cp >= 0x20000 && cp <= 0x3ffff)
}

interface Ordered {
  readonly from: Pos
  readonly to: Pos
  readonly fromLoc: ParaLoc
  readonly toLoc: ParaLoc
}

interface EditResult {
  readonly blocks: readonly Block[]
  readonly pos: Pos
}

function done(doc: Doc, sel: Selection, stored: Marks | null, intent: Intent): Transaction {
  return { doc, sel, stored, intent }
}

function select(state: EditorState, sel: Selection): Transaction | null {
  if (state.stored === null && sameSelection(state.sel, sel)) return null
  return { doc: state.doc, sel, stored: null, intent: 'none' }
}

function caret(pos: Pos): Selection {
  return { kind: 'text', anchor: pos, head: pos }
}

function rangeSel(id: ParaId, from: number, to: number): Selection {
  return { kind: 'text', anchor: { p: id, o: from }, head: { p: id, o: to } }
}

function sameSelection(a: Selection, b: Selection): boolean {
  if (a.kind !== b.kind) return false
  if (a.kind === 'text' && b.kind === 'text') {
    return a.anchor.p === b.anchor.p && a.anchor.o === b.anchor.o && a.head.p === b.head.p && a.head.o === b.head.o
  }
  if (a.kind === 'cells' && b.kind === 'cells') {
    return a.table === b.table
      && a.anchor.row === b.anchor.row && a.anchor.cell === b.anchor.cell
      && a.head.row === b.head.row && a.head.cell === b.head.cell
  }
  return false
}

function selectionFits(doc: Doc, sel: Selection): boolean {
  if (sel.kind === 'cells') {
    const found = findTable(doc.body, sel.table, [])
    if (!found) return false
    return cellFits(found.table, sel.anchor.row, sel.anchor.cell) && cellFits(found.table, sel.head.row, sel.head.cell)
  }
  const index = indexOf(doc)
  const anchor = safeLoc(index, sel.anchor.p)
  const head = safeLoc(index, sel.head.p)
  if (!anchor || !head || anchor.story !== 'body' || head.story !== 'body') return false
  return sel.anchor.o >= 0 && sel.anchor.o <= flatText(anchor.para).length
    && sel.head.o >= 0 && sel.head.o <= flatText(head.para).length
}

function cellFits(table: Table, row: number, cell: number): boolean {
  return row >= 0 && row < table.rows.length && cell >= 0 && cell < (table.rows[row]?.cells.length ?? 0)
}

function safeLoc(index: DocIndex, id: ParaId): ParaLoc | null {
  try {
    return index.get(id)
  } catch {
    return null
  }
}

function orderRange(index: DocIndex, a: Pos, b: Pos): Ordered {
  const aLoc = index.get(a.p)
  const bLoc = index.get(b.p)
  if (aLoc.order < bLoc.order || (aLoc.order === bLoc.order && a.o <= b.o)) {
    return { from: a, to: b, fromLoc: aLoc, toLoc: bLoc }
  }
  return { from: b, to: a, fromLoc: bLoc, toLoc: aLoc }
}

function bodyLocations(doc: Doc): ParaLoc[] {
  const index = indexOf(doc)
  const out: ParaLoc[] = []
  for (const id of index.ordered) {
    const loc = index.get(id)
    if (loc.story !== 'body') break
    out.push(loc)
  }
  return out
}

function movableParas(doc: Doc): ParaLoc[] {
  return bodyLocations(doc).filter(loc => !isPreview(doc, loc))
}

function isPreview(doc: Doc, loc: ParaLoc): boolean {
  let blocks: readonly Block[] = doc.body
  const path = loc.path
  let i = 0
  while (i < path.length) {
    const index = path[i]
    if (index === undefined) return false
    const block = blocks[index]
    if (!block) return false
    if (block.kind === 'opaque') return true
    if (block.kind === 'p') return false
    const row = path[i + 1]
    const cell = path[i + 2]
    if (row === undefined || cell === undefined) return false
    const inner = block.rows[row]?.cells[cell]?.blocks
    if (!inner) return false
    blocks = inner
    i += 3
  }
  return false
}

function paragraphInTable(doc: Doc, id: ParaId): boolean {
  const loc = safeLoc(indexOf(doc), id)
  return loc !== null && tableAddress(doc, loc) !== null
}

function tableAddress(doc: Doc, loc: ParaLoc): { table: TableId, row: number, cell: number } | null {
  let blocks: readonly Block[] = doc.body
  const path = loc.path
  let i = 0
  while (i + 3 < path.length) {
    const blockIndex = path[i]
    const row = path[i + 1]
    const cell = path[i + 2]
    if (blockIndex === undefined || row === undefined || cell === undefined) return null
    const block = blocks[blockIndex]
    if (!block || block.kind !== 'table') return null
    const inner = block.rows[row]?.cells[cell]?.blocks
    if (!inner) return null
    if (i + 4 === path.length) return { table: block.id, row, cell }
    blocks = inner
    i += 3
  }
  return null
}

function replaceSelection(state: EditorState): { doc: Doc, pos: Pos, stored: Marks | null, deleted: boolean } | null {
  if (state.sel.kind !== 'text') return null
  const index = indexOf(state.doc)
  const ordered = orderRange(index, state.sel.anchor, state.sel.head)
  if (ordered.fromLoc.story !== 'body' || ordered.toLoc.story !== 'body') return null
  if (comparePos(index, ordered.from, ordered.to) === 0) {
    return { doc: state.doc, pos: ordered.from, stored: state.stored, deleted: false }
  }
  const removed = deleteRange(state.doc, ordered.from, ordered.to)
  if (!removed) return null
  return { doc: removed.doc, pos: removed.pos, stored: null, deleted: true }
}

function deleteInside(state: EditorState, para: Paragraph, from: number, to: number, caretAt: number): Transaction | null {
  const next = editParagraphText(para, from, to)
  if (!next) return null
  return done(replaceParagraph(state.doc, next), caret({ p: next.id, o: caretAt }), null, 'deleting')
}

function editParagraphText(para: Paragraph, from: number, to: number): Paragraph | null {
  if (para.lock) return null
  const inlines = normalizeInlines(deleteInlines(para.inlines, from, to))
  if (sameInlines(para.inlines, inlines)) return null
  return { ...para, inlines }
}

function mergeSibling(state: EditorState, loc: ParaLoc, direction: 'backward' | 'forward'): Transaction | null {
  const paras = movableParas(state.doc)
  const at = paras.findIndex(item => item.para.id === loc.para.id)
  const other = paras[direction === 'backward' ? at - 1 : at + 1]
  if (!other || !siblingPair(direction === 'backward' ? other : loc, direction === 'backward' ? loc : other)) return null
  const first = direction === 'backward' ? other : loc
  const second = direction === 'backward' ? loc : other
  if (first.para.lock || second.para.lock || first.para.sectionEnd !== null) return null
  const removed = deleteRange(
    state.doc,
    { p: first.para.id, o: flatText(first.para).length },
    { p: second.para.id, o: 0 },
  )
  if (!removed) return null
  return done(removed.doc, caret(removed.pos), null, 'deleting')
}

function siblingPair(first: ParaLoc, second: ParaLoc): boolean {
  if (first.path.length !== second.path.length) return false
  for (let i = 0; i < first.path.length - 1; i++) if (first.path[i] !== second.path[i]) return false
  const a = first.path[first.path.length - 1]
  const b = second.path[second.path.length - 1]
  return a !== undefined && b !== undefined && b === a + 1
}

function deleteRange(doc: Doc, anchor: Pos, head: Pos): { doc: Doc, pos: Pos } | null {
  const index = indexOf(doc)
  const ordered = snapOrdered(orderRange(index, anchor, head))
  if (ordered.fromLoc.story !== 'body' || ordered.toLoc.story !== 'body') return null
  if (isPreview(doc, ordered.fromLoc) || isPreview(doc, ordered.toLoc)) return null
  if (ordered.from.p === ordered.to.p && ordered.from.o === ordered.to.o) return null
  const edited = dispatch(doc.body, ordered.fromLoc.path, ordered.toLoc.path, ordered.from, ordered.to, false)
  if (!edited || edited.blocks === doc.body) return null
  const placed = placeCaret(edited.blocks, edited.pos)
  return { doc: { ...doc, body: placed.blocks }, pos: placed.pos }
}

function snapOrdered(ordered: Ordered): Ordered {
  const fromText = flatText(ordered.fromLoc.para)
  const toText = flatText(ordered.toLoc.para)
  const fromO = snapBack(fromText, clamp(ordered.from.o, fromText.length))
  const toO = snapForward(toText, clamp(ordered.to.o, toText.length))
  return {
    from: { p: ordered.from.p, o: fromO },
    to: { p: ordered.to.p, o: toO },
    fromLoc: ordered.fromLoc,
    toLoc: ordered.toLoc,
  }
}

function dispatch(
  blocks: readonly Block[],
  fromPath: readonly number[],
  toPath: readonly number[],
  from: Pos,
  to: Pos,
  inCell: boolean,
): EditResult | null {
  const start = fromPath[0]
  const end = toPath[0]
  if (start === undefined || end === undefined) return null
  if (start === end) {
    const block = blocks[start]
    if (!block) return null
    if (block.kind === 'p' && fromPath.length === 1 && toPath.length === 1) {
      const para = editParagraphText(block, from.o, to.o)
      if (!para) return null
      return { blocks: replaceAt(blocks, start, para), pos: { p: para.id, o: from.o } }
    }
    if (block.kind !== 'table') return null
    const rowA = fromPath[1]
    const cellA = fromPath[2]
    const rowB = toPath[1]
    const cellB = toPath[2]
    if (rowA === undefined || cellA === undefined || rowB === undefined || cellB === undefined) return null
    if (rowA === rowB && cellA === cellB) {
      const row = block.rows[rowA]
      const cell = row?.cells[cellA]
      if (!row || !cell) return null
      const inner = dispatch(cell.blocks, fromPath.slice(3), toPath.slice(3), from, to, true)
      if (!inner) return null
      const newCell: Cell = { ...cell, blocks: inner.blocks }
      const newRow: Row = { ...row, cells: replaceAt(row.cells, cellA, newCell) }
      const newTable: Table = { ...block, rows: replaceAt(block.rows, rowA, newRow) }
      return { blocks: replaceAt(blocks, start, newTable), pos: inner.pos }
    }
    return clearDivergedCells(blocks, start, block, rowA, cellA, rowB, cellB)
  }
  const lo = Math.min(start, end)
  const hi = Math.max(start, end)
  const forward = start <= end
  return deleteSpan(
    blocks,
    lo,
    hi,
    forward ? fromPath : toPath,
    forward ? from : to,
    forward ? toPath : fromPath,
    forward ? to : from,
    inCell,
  )
}

function clearDivergedCells(
  blocks: readonly Block[],
  tableIndex: number,
  table: Table,
  rowA: number,
  cellA: number,
  rowB: number,
  cellB: number,
): EditResult | null {
  const start = rowA < rowB || (rowA === rowB && cellA <= cellB) ? { row: rowA, cell: cellA } : { row: rowB, cell: cellB }
  const end = start.row === rowA && start.cell === cellA ? { row: rowB, cell: cellB } : { row: rowA, cell: cellA }
  let pos: Pos | null = null
  let changed = false
  const rows = table.rows.map((row, r) => {
    let rowChanged = false
    const cells = row.cells.map((cell, c) => {
      const afterStart = r > start.row || (r === start.row && c >= start.cell)
      const beforeEnd = r < end.row || (r === end.row && c <= end.cell)
      if (!afterStart || !beforeEnd || containsProtected(cell.blocks)) return cell
      const cleared = clearCell(cell)
      changed = true
      rowChanged = true
      if (!pos) pos = { p: cleared.para.id, o: 0 }
      return cleared.cell
    })
    return rowChanged ? { ...row, cells } : row
  })
  if (!changed || !pos) return null
  return { blocks: replaceAt(blocks, tableIndex, { ...table, rows }), pos }
}

function deleteSpan(
  blocks: readonly Block[],
  start: number,
  end: number,
  fromPath: readonly number[],
  from: Pos,
  toPath: readonly number[],
  to: Pos,
  inCell: boolean,
): EditResult | null {
  const left = blocks[start]
  const right = blocks[end]
  if (!left || !right) return null
  const middle = blocks.slice(start + 1, end)
  if (
    fromPath.length === 1 && toPath.length === 1
    && left.kind === 'p' && right.kind === 'p'
    && !left.lock && !right.lock
    && left.sectionEnd === null
    && middle.every(canRemoveBlock)
  ) {
    return mergeSpan(blocks, start, end - start + 1, left, right, from.o, to.o, inCell)
  }
  const pieces: Block[] = []
  let changed = false
  let pos: Pos | null = null
  const leftEdit = editEndpoint(left, fromPath.slice(1), from, 'to-end')
  if (leftEdit.block !== left) changed = true
  if (leftEdit.block) pieces.push(leftEdit.block)
  if (leftEdit.pos) pos = leftEdit.pos
  for (const block of middle) {
    const kept = shrinkBlock(block)
    if (kept !== block) changed = true
    if (kept) pieces.push(kept)
  }
  const rightEdit = editEndpoint(right, toPath.slice(1), to, 'from-start')
  if (rightEdit.block !== right) changed = true
  if (rightEdit.block) pieces.push(rightEdit.block)
  if (!pos) pos = rightEdit.pos
  if (!changed) return null
  const finished = finishBlocks(blocks.slice(0, start).concat(pieces, blocks.slice(end + 1)), inCell)
  const placed = placeCaret(finished, pos)
  return { blocks: placed.blocks, pos: placed.pos }
}

function mergeSpan(
  blocks: readonly Block[],
  start: number,
  count: number,
  first: Paragraph,
  second: Paragraph,
  fromO: number,
  toO: number,
  inCell: boolean,
): EditResult {
  const head = sliceInlines(first.inlines, 0, fromO)
  const tail = sliceInlines(second.inlines, toO, inlineLength(second.inlines))
  const kept = fixedBetween(first, fromO, inlineLength(first.inlines)).concat(fixedBetween(second, 0, toO))
  const side = mergeSide(first, second, fromO)
  const props = side === 'first' ? first.props : second.props
  const preferred = side === 'first' ? first.src : second.src
  const fallback = side === 'first' ? second.src : first.src
  const merged: Paragraph = {
    kind: 'p',
    id: side === 'first' ? first.id : second.id,
    src: preferred ?? fallback,
    props,
    inlines: normalizeInlines(head.concat(kept, tail)),
    lock: null,
    sectionEnd: second.sectionEnd,
  }
  const next = finishBlocks(spliceBlocks(blocks, start, count, [merged]), inCell)
  return { blocks: next, pos: { p: merged.id, o: inlineLength(head) } }
}

interface Endpoint {
  readonly block: Block | null
  readonly pos: Pos | null
}

function editEndpoint(block: Block, rest: readonly number[], pos: Pos, mode: 'to-end' | 'from-start'): Endpoint {
  if (block.kind === 'p') {
    return editParagraphEndpoint(block, pos.o, mode)
  }
  if (block.kind === 'opaque') return { block, pos: null }
  const row = rest[0]
  const cell = rest[1]
  if (row === undefined || cell === undefined) return { block, pos: null }
  const edited = clearHalf(block, row, cell, mode)
  if (edited.table === block) return { block, pos: null }
  return { block: edited.table, pos: edited.pos }
}

function editParagraphEndpoint(para: Paragraph, offset: number, mode: 'to-end' | 'from-start'): Endpoint {
  const len = flatText(para).length
  const at = clamp(offset, len)
  if (para.lock) return { block: para, pos: { p: para.id, o: mode === 'to-end' ? at : 0 } }
  if (mode === 'to-end' && at >= len) return { block: para, pos: { p: para.id, o: len } }
  if (mode === 'from-start' && at <= 0) return { block: para, pos: { p: para.id, o: 0 } }
  const from = mode === 'to-end' ? at : 0
  const to = mode === 'to-end' ? len : at
  const inlines = normalizeInlines(deleteInlines(para.inlines, from, to))
  if (inlines.length === 0 && para.sectionEnd === null && !hasFixedAtom(para)) return { block: null, pos: null }
  const next = sameInlines(para.inlines, inlines) ? para : { ...para, inlines }
  const caretAt = mode === 'to-end' ? Math.min(at, flatText(next).length) : 0
  return { block: next, pos: { p: next.id, o: caretAt } }
}

function clearHalf(table: Table, row: number, cell: number, mode: 'to-end' | 'from-start'): { table: Table, pos: Pos | null } {
  let pos: Pos | null = null
  let changed = false
  const rows = table.rows.map((entry, r) => {
    let rowChanged = false
    const cells = entry.cells.map((item, c) => {
      const after = r > row || (r === row && c >= cell)
      const before = r < row || (r === row && c <= cell)
      const hit = mode === 'to-end' ? after : before
      if (!hit || containsProtected(item.blocks)) return item
      const cleared = clearCell(item)
      changed = true
      rowChanged = true
      if (!pos) pos = { p: cleared.para.id, o: 0 }
      return cleared.cell
    })
    return rowChanged ? { ...entry, cells } : entry
  })
  if (!changed) return { table, pos: null }
  return { table: { ...table, rows }, pos }
}

function shrinkBlock(block: Block): Block | null {
  if (canRemoveBlock(block)) return null
  if (block.kind === 'p') {
    if (block.lock) return block
    const inlines = normalizeInlines(block.inlines.filter(isFixedAtom))
    return sameInlines(block.inlines, inlines) ? block : { ...block, inlines }
  }
  if (block.kind === 'table') {
    const cleared = clearUnlocked(block)
    return cleared
  }
  return block
}

function clearUnlocked(table: Table): Table {
  let changed = false
  const rows = table.rows.map(row => {
    let rowChanged = false
    const cells = row.cells.map(cell => {
      if (containsProtected(cell.blocks)) return cell
      changed = true
      rowChanged = true
      return clearCell(cell).cell
    })
    return rowChanged ? { ...row, cells } : row
  })
  return changed ? { ...table, rows } : table
}

function canRemoveBlock(block: Block): boolean {
  if (block.kind === 'opaque') return block.removable
  if (block.kind === 'p') return block.lock === null && block.sectionEnd === null && !hasFixedAtom(block)
  return !containsProtected(block.rows.flatMap(row => row.cells.flatMap(cell => cell.blocks)))
}

function containsProtected(blocks: readonly Block[]): boolean {
  for (const block of blocks) {
    if (block.kind === 'p') {
      if (block.lock || block.sectionEnd !== null || hasFixedAtom(block)) return true
    } else if (block.kind === 'table') {
      for (const row of block.rows) for (const cell of row.cells) {
        if (containsProtected(cell.blocks)) return true
      }
    } else if (!block.removable) return true
  }
  return false
}

function hasFixedAtom(para: Paragraph): boolean {
  return para.inlines.some(isFixedAtom)
}

function isFixedAtom(inline: Inline): inline is Atom {
  return inline.kind === 'atom' && inline.atom.type === 'opaque' && !inline.atom.removable
}

function fixedBetween(para: Paragraph, from: number, to: number): Atom[] {
  const out: Atom[] = []
  let o = 0
  for (const inline of para.inlines) {
    const len = inlineLen(inline)
    if (isFixedAtom(inline) && o >= from && o < to) out.push(inline)
    o += len
  }
  return out
}

function clearCell(cell: Cell): { cell: Cell, para: Paragraph } {
  const para = emptyParagraph()
  return { cell: { ...cell, blocks: [para] }, para }
}

function clearSelectedCells(state: EditorState): Transaction | null {
  if (state.sel.kind !== 'cells') return null
  const found = findTable(state.doc.body, state.sel.table, [])
  if (!found) return null
  const r0 = Math.min(state.sel.anchor.row, state.sel.head.row)
  const r1 = Math.max(state.sel.anchor.row, state.sel.head.row)
  const c0 = Math.min(state.sel.anchor.cell, state.sel.head.cell)
  const c1 = Math.max(state.sel.anchor.cell, state.sel.head.cell)
  let pos: Pos | null = null
  let changed = false
  const rows = found.table.rows.map((row, r) => {
    if (r < r0 || r > r1) return row
    let rowChanged = false
    const cells = row.cells.map((cell, c) => {
      if (c < c0 || c > c1 || containsProtected(cell.blocks)) return cell
      const cleared = clearCell(cell)
      changed = true
      rowChanged = true
      if (!pos) pos = { p: cleared.para.id, o: 0 }
      return cleared.cell
    })
    return rowChanged ? { ...row, cells } : row
  })
  if (!changed || !pos) return null
  const table: Table = { ...found.table, rows }
  const doc = writeTable(state.doc, found.path, table)
  if (!doc) return null
  return done(doc, caret(pos), null, 'deleting')
}

function insertPlain(doc: Doc, pos: Pos, text: string, marks: Marks, src: RunSrc | null): { doc: Doc, pos: Pos } | null {
  const parts = text.split('\n')
  let current = doc
  let caretPos = pos
  for (let i = 0; i < parts.length; i++) {
    if (i > 0) {
      const split = splitAt(current, caretPos, false)
      if (!split) return null
      current = split.doc
      caretPos = split.pos
    }
    const part = parts[i] ?? ''
    if (part.length === 0) continue
    const inserted = insertString(current, caretPos, part, marks, src)
    if (!inserted) return null
    current = inserted.doc
    caretPos = inserted.pos
  }
  return { doc: current, pos: caretPos }
}

function insertString(doc: Doc, pos: Pos, text: string, marks: Marks, src: RunSrc | null): { doc: Doc, pos: Pos } | null {
  const loc = safeLoc(indexOf(doc), pos.p)
  if (!loc || loc.para.lock) return null
  const run: TextRun = { kind: 'text', text, marks, src }
  const inlines = normalizeInlines(insertInlines(loc.para.inlines, snapBack(flatText(loc.para), clamp(pos.o, flatText(loc.para).length)), [run]))
  const para: Paragraph = { ...loc.para, inlines }
  const at = snapBack(flatText(loc.para), clamp(pos.o, flatText(loc.para).length))
  return { doc: replaceParagraph(doc, para), pos: { p: para.id, o: at + text.length } }
}

function insertAtom(doc: Doc, pos: Pos, atom: Atom): { doc: Doc, pos: Pos } | null {
  const loc = safeLoc(indexOf(doc), pos.p)
  if (!loc || loc.para.lock) return null
  const text = flatText(loc.para)
  const at = snapBack(text, clamp(pos.o, text.length))
  const inlines = normalizeInlines(insertInlines(loc.para.inlines, at, [atom]))
  const para: Paragraph = { ...loc.para, inlines }
  return { doc: replaceParagraph(doc, para), pos: { p: para.id, o: at + 1 } }
}

function splitAt(doc: Doc, pos: Pos, useNext: boolean): { doc: Doc, pos: Pos } | null {
  const loc = safeLoc(indexOf(doc), pos.p)
  if (!loc || loc.story !== 'body' || loc.para.lock || isPreview(doc, loc)) return null
  const para = loc.para
  const text = flatText(para)
  const o = snapBack(text, clamp(pos.o, text.length))
  if (o === 0 && text.length > 0) {
    const fresh = paragraphLike(para, [], null)
    const body = spliceParagraph(doc, loc, [fresh, para])
    return { doc: body, pos: { p: fresh.id, o: 0 } }
  }
  const headInlines = sliceInlines(para.inlines, 0, o)
  const tailInlines = sliceInlines(para.inlines, o, text.length)
  const head: Paragraph = headInlines === para.inlines && para.sectionEnd === null
    ? para
    : { ...para, inlines: headInlines, sectionEnd: null }
  const tail = paragraphLike(para, tailInlines, para.sectionEnd)
  if (o === text.length && useNext) {
    const props = nextProps(doc, para)
    const styled: Paragraph = props === para.props ? tail : { ...tail, props }
    return { doc: spliceParagraph(doc, loc, [head, styled]), pos: { p: styled.id, o: 0 } }
  }
  return { doc: spliceParagraph(doc, loc, [head, tail]), pos: { p: tail.id, o: 0 } }
}

function paragraphLike(para: Paragraph, inlines: readonly Inline[], sectionEnd: number | null): Paragraph {
  return {
    kind: 'p',
    id: newParaId(),
    src: para.src,
    props: para.props,
    inlines: normalizeInlines(inlines),
    lock: null,
    sectionEnd,
  }
}

function nextProps(doc: Doc, para: Paragraph): ParaProps {
  const style = para.props.style
  if (!style) return para.props
  const next = doc.styles.paragraph.get(style)?.next
  if (!next) return para.props
  return { style: next }
}

function placeBlock(doc: Doc, loc: ParaLoc, offset: number, block: Block): Doc | null {
  const para = loc.para
  const text = flatText(para)
  const o = snapBack(text, clamp(offset, text.length))
  const inCell = loc.path.length > 1
  if (o <= 0) return spliceParagraph(doc, loc, finishBlocks([block, para], inCell))
  if (o >= text.length) return spliceParagraph(doc, loc, finishBlocks([para, block], inCell))
  const split = splitAt(doc, { p: para.id, o }, false)
  if (!split) return null
  const head = safeLoc(indexOf(split.doc), para.id)
  const tail = safeLoc(indexOf(split.doc), split.pos.p)
  if (!head || !tail) return null
  const container = head.path.slice(0, -1)
  const index = head.path[head.path.length - 1] ?? 0
  const body = rewrite(split.doc.body, container, (siblings) => (
    spliceBlocks(siblings, index, 2, [head.para, block, tail.para])
  ))
  const blocks = finishContainer(body, container, head.path.length > 1)
  return { ...split.doc, body: blocks }
}

function spliceParagraph(doc: Doc, loc: ParaLoc, blocks: readonly Block[]): Doc {
  const container = loc.path.slice(0, -1)
  const index = loc.path[loc.path.length - 1] ?? 0
  const body = rewrite(doc.body, container, (siblings) => spliceBlocks(siblings, index, 1, blocks))
  return { ...doc, body }
}

function replaceParagraph(doc: Doc, para: Paragraph): Doc {
  const loc = indexOf(doc).get(para.id)
  if (loc.para === para) return doc
  const body = writeParagraph(doc.body, loc.path, para)
  if (body === doc.body) return doc
  return { ...doc, body }
}

function writeParagraph(blocks: readonly Block[], path: readonly number[], para: Paragraph): readonly Block[] {
  const index = path[0]
  if (index === undefined) return blocks
  if (path.length === 1) return replaceAt(blocks, index, para)
  const block = blocks[index]
  if (!block || block.kind !== 'table') return blocks
  const rowIndex = path[1]
  const cellIndex = path[2]
  if (rowIndex === undefined || cellIndex === undefined) return blocks
  const row = block.rows[rowIndex]
  const cell = row?.cells[cellIndex]
  if (!row || !cell) return blocks
  const nested = writeParagraph(cell.blocks, path.slice(3), para)
  if (nested === cell.blocks) return blocks
  const newCell: Cell = { ...cell, blocks: nested }
  const newRow: Row = { ...row, cells: replaceAt(row.cells, cellIndex, newCell) }
  const newTable: Table = { ...block, rows: replaceAt(block.rows, rowIndex, newRow) }
  return replaceAt(blocks, index, newTable)
}

function rewrite(blocks: readonly Block[], path: readonly number[], fn: (blocks: readonly Block[]) => readonly Block[]): readonly Block[] {
  if (path.length === 0) return fn(blocks)
  const index = path[0]
  if (index === undefined) return blocks
  const block = blocks[index]
  if (!block || block.kind !== 'table') return blocks
  const rowIndex = path[1]
  const cellIndex = path[2]
  if (rowIndex === undefined || cellIndex === undefined) return blocks
  const row = block.rows[rowIndex]
  const cell = row?.cells[cellIndex]
  if (!row || !cell) return blocks
  const nested = rewrite(cell.blocks, path.slice(3), fn)
  if (nested === cell.blocks) return blocks
  const newCell: Cell = { ...cell, blocks: nested }
  const newRow: Row = { ...row, cells: replaceAt(row.cells, cellIndex, newCell) }
  const newTable: Table = { ...block, rows: replaceAt(block.rows, rowIndex, newRow) }
  return replaceAt(blocks, index, newTable)
}

function replaceAt<T>(items: readonly T[], index: number, item: T): readonly T[] {
  if (items[index] === item) return items
  const next = items.slice()
  next[index] = item
  return next
}

function spliceBlocks(blocks: readonly Block[], start: number, count: number, insert: readonly Block[]): readonly Block[] {
  const next = blocks.slice(0, start).concat(insert, blocks.slice(start + count))
  return next
}

function finishBlocks(blocks: readonly Block[], inCell: boolean): readonly Block[] {
  if (blocks.length === 0) return [emptyParagraph()]
  if (inCell && blocks[blocks.length - 1]?.kind !== 'p') return blocks.concat(emptyParagraph())
  return blocks
}

function emptyParagraph(props: ParaProps = {}): Paragraph {
  return { kind: 'p', id: newParaId(), src: null, props, inlines: [], lock: null, sectionEnd: null }
}

function emptyCell(): Cell {
  return { src: null, span: 1, vMerge: null, blocks: [emptyParagraph()] }
}

function inlineLen(inline: Inline): number {
  return inline.kind === 'text' ? inline.text.length : 1
}

function inlineLength(inlines: readonly Inline[]): number {
  let n = 0
  for (const inline of inlines) n += inlineLen(inline)
  return n
}

function sameInlines(a: readonly Inline[], b: readonly Inline[]): boolean {
  if (a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false
  return true
}

function sliceInlines(inlines: readonly Inline[], from: number, to: number): readonly Inline[] {
  const total = inlineLength(inlines)
  const lo = clamp(from, total)
  const hi = clamp(to, total)
  if (lo <= 0 && hi >= total) return inlines
  if (lo >= hi) return []
  const out: Inline[] = []
  let o = 0
  for (const inline of inlines) {
    const len = inlineLen(inline)
    const start = o
    const end = o + len
    const cutLo = Math.max(lo, start)
    const cutHi = Math.min(hi, end)
    if (cutLo < cutHi) {
      if (inline.kind === 'atom' || (cutLo === start && cutHi === end)) out.push(inline)
      else out.push({ ...inline, text: inline.text.slice(cutLo - start, cutHi - start) })
    }
    o = end
  }
  return out
}

function deleteInlines(inlines: readonly Inline[], from: number, to: number): readonly Inline[] {
  if (from >= to) return inlines
  const out: Inline[] = []
  let changed = false
  let o = 0
  for (const inline of inlines) {
    const len = inlineLen(inline)
    const start = o
    const end = o + len
    o = end
    if (end <= from || start >= to) {
      out.push(inline)
      continue
    }
    if (inline.kind === 'atom') {
      if (isFixedAtom(inline)) out.push(inline)
      else changed = true
      continue
    }
    changed = true
    if (start < from) {
      const text = inline.text.slice(0, from - start)
      if (text.length > 0) out.push({ ...inline, text })
    }
    if (end > to) {
      const text = inline.text.slice(to - start)
      if (text.length > 0) out.push({ ...inline, text })
    }
  }
  return changed ? out : inlines
}

function insertInlines(inlines: readonly Inline[], o: number, extra: readonly Inline[]): Inline[] {
  const total = inlineLength(inlines)
  const at = clamp(o, total)
  return sliceInlines(inlines, 0, at).concat(extra, sliceInlines(inlines, at, total))
}

function placeCaret(blocks: readonly Block[], pos: Pos | null): { blocks: readonly Block[], pos: Pos } {
  const found = pos ? findPara(blocks, pos.p) : null
  if (found && pos) return { blocks, pos: { p: found.id, o: clamp(pos.o, flatText(found).length) } }
  const first = firstPosIn(blocks)
  if (first) return { blocks, pos: first }
  const para = emptyParagraph()
  return { blocks: blocks.concat(para), pos: { p: para.id, o: 0 } }
}

function firstPosIn(blocks: readonly Block[]): Pos | null {
  for (const block of blocks) {
    if (block.kind === 'p') return { p: block.id, o: 0 }
    if (block.kind === 'table') {
      for (const row of block.rows) for (const cell of row.cells) {
        const pos = firstPosIn(cell.blocks)
        if (pos) return pos
      }
    }
  }
  return null
}

function findPara(blocks: readonly Block[], id: ParaId): Paragraph | null {
  for (const block of blocks) {
    if (block.kind === 'p' && block.id === id) return block
    if (block.kind === 'table') {
      for (const row of block.rows) for (const cell of row.cells) {
        const found = findPara(cell.blocks, id)
        if (found) return found
      }
    }
  }
  return null
}

function findTable(blocks: readonly Block[], id: TableId, prefix: number[]): { table: Table, path: number[] } | null {
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i]
    if (!block || block.kind !== 'table') continue
    if (block.id === id) return { table: block, path: prefix.concat(i) }
    for (let r = 0; r < block.rows.length; r++) {
      const row = block.rows[r]
      if (!row) continue
      for (let c = 0; c < row.cells.length; c++) {
        const cell = row.cells[c]
        if (!cell) continue
        const found = findTable(cell.blocks, id, prefix.concat(i, r, c))
        if (found) return found
      }
    }
  }
  return null
}

function writeTable(doc: Doc, path: readonly number[], table: Table): Doc | null {
  const container = path.slice(0, -1)
  const index = path[path.length - 1]
  if (index === undefined) return null
  const body = rewrite(doc.body, container, (blocks) => replaceAt(blocks, index, table))
  if (body === doc.body) return null
  return { ...doc, body }
}

function removeTable(doc: Doc, id: TableId): { doc: Doc, pos: Pos } | null {
  const found = findTable(doc.body, id, [])
  if (!found) return null
  const container = found.path.slice(0, -1)
  const index = found.path[found.path.length - 1]
  if (index === undefined) return null
  const inCell = container.length > 0
  let pos: Pos | null = null
  const body = rewrite(doc.body, container, (blocks) => {
    const next = finishBlocks(blocks.filter((_, i) => i !== index), inCell)
    pos = firstPosIn(next)
    return next
  })
  if (!pos) return null
  return { doc: { ...doc, body }, pos }
}

const insertedEdge: BorderLine = { style: 'single', widthPt: asPt(0.5), color: '000000' as Hex }

/** 3px 5px at 96dpi. An inserted table keeps the padding the view used before cell margins came from the file. */
export const INSERTED_CELL_MARGINS: CellMargins = {
  top: asPt(2.25),
  right: asPt(3.75),
  bottom: asPt(2.25),
  left: asPt(3.75),
}

/** The six edges write.rs gives a table the editor inserts: single, half a point, black. */
export const INSERTED_TABLE_BORDERS: TableBorders = {
  top: insertedEdge,
  left: insertedEdge,
  bottom: insertedEdge,
  right: insertedEdge,
  insideH: insertedEdge,
  insideV: insertedEdge,
}

function makeTable(doc: Doc, loc: ParaLoc, rows: number, cols: number): Table {
  const width = contentWidth(doc, sectionIndex(doc, loc))
  const colWidth = asPt(Math.max(36, width / cols))
  const grid = Array.from({ length: cols }, () => colWidth)
  const body: Row[] = []
  for (let r = 0; r < rows; r++) {
    const cells: Cell[] = []
    for (let c = 0; c < cols; c++) cells.push(emptyCell())
    body.push({ src: null, props: {}, cells })
  }
  return { kind: 'table', id: newTableId(), src: null, props: { borders: INSERTED_TABLE_BORDERS, margins: INSERTED_CELL_MARGINS }, grid, rows: body }
}

function newTableId(): TableId {
  tableSerial += 1
  return `tbl-${tableSerial}` as TableId
}

function blankRow(template: Row, columns: number): Row {
  const cells = template.cells.length > 0
    ? template.cells.map(cell => ({ src: null, span: cell.span, vMerge: null, widthPt: cell.widthPt, blocks: [emptyParagraph()] }))
    : Array.from({ length: columns }, () => emptyCell())
  return {
    src: null,
    props: {
      ...(template.props.heightPt !== undefined ? { heightPt: template.props.heightPt } : {}),
      ...(template.props.heightRule !== undefined ? { heightRule: template.props.heightRule } : {}),
    },
    cells,
  }
}

function rowHasProtected(row: Row): boolean {
  return row.cells.some(cell => containsProtected(cell.blocks))
}

interface RowSpan { readonly start: number, readonly end: number }
interface ColSpan { readonly start: number, readonly end: number }

function editRows(state: EditorState, fn: (table: Table, rows: RowSpan) => { table: Table | null, pos: Pos | null } | null): Transaction | null {
  const hit = tableHit(state)
  if (!hit) return null
  const edited = fn(hit.table, hit.rows)
  if (!edited) return null
  if (edited.table === null) {
    const removed = removeTable(state.doc, hit.table.id)
    if (!removed) return null
    return done(removed.doc, caret(removed.pos), null, 'discrete')
  }
  const doc = writeTable(state.doc, hit.path, edited.table)
  if (!doc || !edited.pos) return null
  return done(doc, caret(edited.pos), null, 'discrete')
}

function editColumns(state: EditorState, fn: (table: Table, cols: ColSpan) => { table: Table | null, pos: Pos | null } | null): Transaction | null {
  return editRows(state, (table, rows) => {
    const cols = columnSpan(table, rows, tableHitColumns(state, table))
    return fn(table, cols)
  })
}

function tableHitColumns(state: EditorState, table: Table): { start: number, end: number } {
  if (state.sel.kind === 'cells' && state.sel.table === table.id) {
    const a = columnOf(table.rows[state.sel.anchor.row], state.sel.anchor.cell)
    const b = columnOf(table.rows[state.sel.head.row], state.sel.head.cell)
    const aEnd = a + spanOf(table.rows[state.sel.anchor.row], state.sel.anchor.cell) - 1
    const bEnd = b + spanOf(table.rows[state.sel.head.row], state.sel.head.cell) - 1
    return { start: Math.min(a, b), end: Math.max(aEnd, bEnd) }
  }
  const hit = tableHit(state)
  const row = table.rows[hit?.rows.start ?? 0]
  const cell = hit?.cell ?? 0
  const start = columnOf(row, cell)
  return { start, end: start + spanOf(row, cell) - 1 }
}

function columnSpan(table: Table, _rows: RowSpan, cols: { start: number, end: number }): ColSpan {
  const max = Math.max(0, table.grid.length - 1)
  return { start: clamp(cols.start, max), end: clamp(cols.end, max) }
}

function columnOf(row: Row | undefined, cellIndex: number): number {
  if (!row) return 0
  let col = 0
  for (let i = 0; i < cellIndex; i++) col += row.cells[i]?.span ?? 1
  return col
}

function spanOf(row: Row | undefined, cellIndex: number): number {
  return row?.cells[cellIndex]?.span ?? 1
}

function tableHit(state: EditorState): { table: Table, path: number[], rows: RowSpan, cell: number } | null {
  if (state.sel.kind === 'cells') {
    const found = findTable(state.doc.body, state.sel.table, [])
    if (!found) return null
    const start = Math.min(state.sel.anchor.row, state.sel.head.row)
    const end = Math.max(state.sel.anchor.row, state.sel.head.row)
    return { table: found.table, path: found.path, rows: { start, end }, cell: state.sel.anchor.cell }
  }
  if (state.sel.kind !== 'text') return null
  const loc = safeLoc(indexOf(state.doc), state.sel.head.p)
  if (!loc) return null
  const addr = tableAddress(state.doc, loc)
  if (!addr) return null
  const found = findTable(state.doc.body, addr.table, [])
  if (!found) return null
  return { table: found.table, path: found.path, rows: { start: addr.row, end: addr.row }, cell: addr.cell }
}

function insertGridColumn(table: Table, col: number): Table {
  const grid = table.grid.slice()
  const width = grid[col] ?? grid[col - 1] ?? asPt(72)
  const at = clamp(col, grid.length)
  grid.splice(at, 0, width)
  const rows = table.rows.map(row => ({ ...row, cells: insertCellAt(row, at) }))
  return { ...table, grid, rows }
}

function insertCellAt(row: Row, col: number): Cell[] {
  const cells = row.cells.slice()
  let cursor = 0
  for (let i = 0; i <= cells.length; i++) {
    if (cursor === col || i === cells.length) {
      cells.splice(i, 0, emptyCell())
      return cells
    }
    const cell = cells[i]
    if (!cell) break
    if (cursor < col && cursor + cell.span > col) {
      cells[i] = { ...cell, span: cell.span + 1 }
      return cells
    }
    cursor += cell.span
  }
  cells.push(emptyCell())
  return cells
}

function deleteGridColumn(table: Table, col: number): Table | null {
  if (table.grid.length <= 1) return null
  const grid = table.grid.filter((_, index) => index !== col)
  const rows = table.rows.map(row => {
    const cells: Cell[] = []
    let cursor = 0
    for (const cell of row.cells) {
      const start = cursor
      cursor += cell.span
      if (cursor <= col || start > col) {
        cells.push(cell)
        continue
      }
      if (cell.span > 1) cells.push({ ...cell, span: cell.span - 1 })
    }
    if (cells.length === 0) cells.push(emptyCell())
    return cells === row.cells ? row : { ...row, cells }
  })
  return { ...table, grid, rows }
}

function scaleImage(doc: Doc, loc: ParaLoc, media: MediaRef): { width: number, height: number } {
  let width = media.widthPx * 0.75
  let height = media.heightPx * 0.75
  const cell = cellWidth(doc, loc)
  const max = cell !== null && cell > 0 ? cell : contentWidth(doc, sectionIndex(doc, loc))
  if (width > max && width > 0) {
    const scale = max / width
    width *= scale
    height *= scale
  }
  return { width, height }
}

function cellWidth(doc: Doc, loc: ParaLoc): number | null {
  let blocks: readonly Block[] = doc.body
  let width: number | null = null
  const path = loc.path
  let i = 0
  while (i + 3 < path.length) {
    const blockIndex = path[i]
    const rowIndex = path[i + 1]
    const cellIndex = path[i + 2]
    if (blockIndex === undefined || rowIndex === undefined || cellIndex === undefined) return width
    const block = blocks[blockIndex]
    if (!block || block.kind !== 'table') return width
    const row = block.rows[rowIndex]
    const cell = row?.cells[cellIndex]
    if (!row || !cell) return width
    const start = columnOf(row, cellIndex)
    let sum = 0
    for (let column = 0; column < cell.span; column += 1) sum += block.grid[start + column] ?? 0
    if (sum > 0) width = sum
    if (i + 4 === path.length) return width
    blocks = cell.blocks
    i += 3
  }
  return width
}

function sectionIndex(doc: Doc, loc: ParaLoc): number {
  let section = 0
  const top = loc.path[0] ?? 0
  for (let i = 0; i < doc.body.length; i++) {
    if (i === top) return section
    const block = doc.body[i]
    if (block?.kind === 'p' && block.sectionEnd !== null) section = block.sectionEnd + 1
  }
  return section
}

function contentWidth(doc: Doc, section: number): number {
  const page = doc.sections[section]?.page ?? doc.sections[0]?.page
  if (!page) return 468
  return Math.max(72, page.widthPt - page.margin.left - page.margin.gutter - page.margin.right)
}

function mapTouched(doc: Doc, sel: Selection, fn: (para: Paragraph) => Paragraph): Doc | null {
  const targets = touchedParagraphs(doc, sel)
  let next = doc
  let changed = false
  for (const para of targets) {
    if (para.lock) continue
    const loc = safeLoc(indexOf(next), para.id)
    if (!loc || isPreview(next, loc)) continue
    const updated = fn(loc.para)
    if (updated === loc.para) continue
    next = replaceParagraph(next, updated)
    changed = true
  }
  if (!changed && next === doc) return null
  return changed ? next : null
}

export function touchedParagraphs(doc: Doc, sel: Selection): Paragraph[] {
  if (sel.kind === 'cells') {
    const found = findTable(doc.body, sel.table, [])
    if (!found) return []
    const r0 = Math.min(sel.anchor.row, sel.head.row)
    const r1 = Math.max(sel.anchor.row, sel.head.row)
    const c0 = Math.min(sel.anchor.cell, sel.head.cell)
    const c1 = Math.max(sel.anchor.cell, sel.head.cell)
    const out: Paragraph[] = []
    for (let r = r0; r <= r1; r++) {
      const row = found.table.rows[r]
      if (!row) continue
      for (let c = c0; c <= c1; c++) {
        const cell = row.cells[c]
        if (cell) collectParas(cell.blocks, out)
      }
    }
    return out
  }
  const index = indexOf(doc)
  const ordered = orderRange(index, sel.anchor, sel.head)
  const out: Paragraph[] = []
  for (const loc of bodyLocations(doc)) {
    if (!paragraphSelected(loc, ordered)) continue
    if (isPreview(doc, loc)) continue
    out.push(loc.para)
  }
  return out
}

function paragraphSelected(loc: ParaLoc, ordered: Ordered): boolean {
  if (loc.order < ordered.fromLoc.order || loc.order > ordered.toLoc.order) return false
  if (loc.order === ordered.toLoc.order && loc.order !== ordered.fromLoc.order && ordered.to.o === 0) return false
  return true
}

function coveredOffsets(loc: ParaLoc, ordered: Ordered): [number, number] | null {
  const len = flatText(loc.para).length
  if (loc.order < ordered.fromLoc.order || loc.order > ordered.toLoc.order) return null
  let start = 0
  let end = len
  if (loc.order === ordered.fromLoc.order) start = ordered.from.o
  if (loc.order === ordered.toLoc.order) end = ordered.to.o
  if (start >= end) return null
  return [start, end]
}

function collectParas(blocks: readonly Block[], out: Paragraph[]): void {
  for (const block of blocks) {
    if (block.kind === 'p') out.push(block)
    else if (block.kind === 'table') {
      for (const row of block.rows) for (const cell of row.cells) collectParas(cell.blocks, out)
    }
  }
}

interface Covered {
  readonly para: Paragraph
  readonly marks: Marks
}

function coveredText(doc: Doc, ordered: Ordered): Covered[] {
  const out: Covered[] = []
  for (const loc of bodyLocations(doc)) {
    const range = coveredOffsets(loc, ordered)
    if (!range || loc.para.lock || isPreview(doc, loc)) continue
    let o = 0
    for (const inline of loc.para.inlines) {
      const len = inlineLen(inline)
      if (inline.kind === 'text' && o < range[1] && o + len > range[0]) out.push({ para: loc.para, marks: inline.marks })
      o += len
    }
  }
  return out
}

function mapCovered(doc: Doc, ordered: Ordered, fn: (marks: Marks) => Marks): Doc | null {
  let next = doc
  let changed = false
  for (const loc of bodyLocations(doc)) {
    const range = coveredOffsets(loc, ordered)
    if (!range || loc.para.lock || isPreview(doc, loc)) continue
    const current = safeLoc(indexOf(next), loc.para.id)
    if (!current) continue
    const text = flatText(current.para)
    const inlines = mapTextRange(current.para.inlines, snapBack(text, range[0]), snapForward(text, range[1]), fn)
    const normalized = normalizeInlines(inlines)
    if (sameInlines(current.para.inlines, normalized)) continue
    next = replaceParagraph(next, { ...current.para, inlines: normalized })
    changed = true
  }
  return changed ? next : null
}

function mapTextRange(inlines: readonly Inline[], from: number, to: number, fn: (marks: Marks) => Marks): readonly Inline[] {
  const out: Inline[] = []
  let changed = false
  let o = 0
  for (const inline of inlines) {
    if (inline.kind === 'atom') {
      out.push(inline)
      o += 1
      continue
    }
    const start = o
    const end = o + inline.text.length
    o = end
    if (end <= from || start >= to) {
      out.push(inline)
      continue
    }
    const lo = Math.max(from, start)
    const hi = Math.min(to, end)
    if (start < lo) out.push({ ...inline, text: inline.text.slice(0, lo - start) })
    const mapped = fn(inline.marks)
    const mid = inline.text.slice(lo - start, hi - start)
    if (marksEqual(mapped, inline.marks) && lo === start && hi === end) out.push(inline)
    else {
      changed = true
      out.push({ ...inline, text: mid, marks: mapped })
    }
    if (hi < end) out.push({ ...inline, text: inline.text.slice(hi - start) })
    if (start < lo || hi < end) changed = true
  }
  return changed ? out : inlines
}

function clearedMarks(marks: Marks): Marks {
  let direct = false
  for (const key of DIRECT_MARKS) {
    if (marks[key] !== undefined) direct = true
  }
  if (!direct) return marks
  return marks.link === undefined ? {} : { link: marks.link }
}

function applyParaFormat(state: EditorState, format: ParaFormat): Transaction | null {
  if (format.align === undefined && format.indent === undefined && format.spacing === undefined) return null
  const doc = mapTouched(state.doc, state.sel, (para) => formatParagraph(para, format))
  if (!doc) return null
  return done(doc, state.sel, null, 'discrete')
}

function formatParagraph(para: Paragraph, format: ParaFormat): Paragraph {
  let props = para.props
  if (format.align !== undefined && props.align !== format.align) props = { ...props, align: format.align }
  if (format.indent) {
    const indent = mergeIndent(props.indent, format.indent)
    if (indent !== props.indent) props = indent === undefined ? without(props, 'indent') : { ...props, indent }
  }
  if (format.spacing) {
    const spacing = mergeSpacing(props.spacing, format.spacing)
    if (spacing !== props.spacing) props = spacing === undefined ? without(props, 'spacing') : { ...props, spacing }
  }
  if (props === para.props) return para
  return { ...para, props }
}

function without(props: ParaProps, key: 'indent' | 'spacing'): ParaProps {
  const next: ParaProps = { ...props }
  delete next[key]
  return next
}

function mergeIndent(prev: ParaProps['indent'], edit: IndentEdit): ParaProps['indent'] | undefined {
  const next: {
    left?: Pt
    right?: Pt
    firstLine?: Pt
    hanging?: Pt
    firstLineChars?: number
  } = { ...prev }
  let touched = false
  if (edit.left !== undefined && next.left !== edit.left) {
    next.left = edit.left
    touched = true
  }
  if (edit.right !== undefined && next.right !== edit.right) {
    next.right = edit.right
    touched = true
  }
  if (edit.firstLine === null && next.firstLine !== undefined) {
    delete next.firstLine
    touched = true
  } else if (edit.firstLine !== undefined && edit.firstLine !== null && next.firstLine !== edit.firstLine) {
    next.firstLine = edit.firstLine
    touched = true
  }
  if (edit.hanging === null && next.hanging !== undefined) {
    delete next.hanging
    touched = true
  } else if (edit.hanging !== undefined && edit.hanging !== null && next.hanging !== edit.hanging) {
    next.hanging = edit.hanging
    touched = true
  }
  if (edit.firstLineChars === null && next.firstLineChars !== undefined) {
    delete next.firstLineChars
    touched = true
  } else if (
    edit.firstLineChars !== undefined && edit.firstLineChars !== null && next.firstLineChars !== edit.firstLineChars
  ) {
    next.firstLineChars = edit.firstLineChars
    touched = true
  }
  if (!touched) return prev
  if (
    next.left === undefined && next.right === undefined && next.firstLine === undefined
    && next.hanging === undefined && next.firstLineChars === undefined
  ) return undefined
  return next
}

function mergeSpacing(prev: ParaProps['spacing'], edit: SpacingEdit): ParaProps['spacing'] | undefined {
  const next: { before?: Pt, after?: Pt, line?: LineSpacing } = { ...prev }
  let touched = false
  if (edit.before !== undefined && next.before !== edit.before) {
    next.before = edit.before
    touched = true
  }
  if (edit.after !== undefined && next.after !== edit.after) {
    next.after = edit.after
    touched = true
  }
  if (edit.line !== undefined && !(next.line && sameLine(next.line, edit.line))) {
    next.line = edit.line
    touched = true
  }
  if (!touched) return prev
  if (next.before === undefined && next.after === undefined && next.line === undefined) return undefined
  return next
}

function applyMarkValue(marks: Marks, change: MarkValue): Marks {
  if (change.value === undefined) {
    if (marks[change.mark] === undefined) return marks
    const next: Marks = { ...marks }
    delete next[change.mark]
    return next
  }
  if (marks[change.mark] === change.value) return marks
  switch (change.mark) {
    case 'font': return { ...marks, font: change.value }
    case 'fontEastAsia': return { ...marks, fontEastAsia: change.value }
    case 'sizePt': return { ...marks, sizePt: change.value }
    case 'color': return { ...marks, color: change.value }
    case 'highlight': return { ...marks, highlight: change.value }
    case 'vertAlign': return { ...marks, vertAlign: change.value }
  }
}

function marksEqual(a: Marks, b: Marks): boolean {
  return a.bold === b.bold && a.italic === b.italic && a.underline === b.underline && a.strike === b.strike
    && a.font === b.font && a.fontEastAsia === b.fontEastAsia && a.sizePt === b.sizePt && a.color === b.color
    && a.highlight === b.highlight && a.vertAlign === b.vertAlign && a.charStyle === b.charStyle && a.link === b.link
}

function withoutListIndent(indent: ParaProps['indent']): ParaProps['indent'] {
  if (!indent || (indent.left === undefined && indent.hanging === undefined)) return indent
  const { left: _left, hanging: _hanging, ...rest } = indent
  return Object.keys(rest).length > 0 ? rest : undefined
}

function withNumbering(para: Paragraph, numbering: ParaProps['numbering']): Paragraph {
  const indent = withoutListIndent(para.props.indent)
  if (para.props.numbering === numbering && indent === para.props.indent) return para
  const { indent: _dropped, ...rest } = para.props
  const props: ParaProps = indent === undefined ? { ...rest, numbering } : { ...rest, indent, numbering }
  return { ...para, props }
}

function listInfo(doc: Doc, para: Paragraph): { num: NumId, level: number } | null {
  if (para.props.numbering === null) return null
  if (para.props.numbering) return para.props.numbering
  const style = para.props.style
  if (!style) return null
  return doc.styles.paragraph.get(style)?.look.numbering ?? null
}

function listKindOf(doc: Doc, para: Paragraph): 'bullet' | 'decimal' | 'none' {
  const info = listInfo(doc, para)
  if (!info) return 'none'
  const format = doc.lists.byNum.get(info.num)?.levels[info.level]?.format
  if (format === 'bullet' || format === 'decimal') return format
  return 'none'
}

function neighborList(doc: Doc, touched: readonly Paragraph[], kind: 'bullet' | 'decimal'): { num: NumId, level: number } | null {
  const index = indexOf(doc)
  const orders = touched.map(para => index.get(para.id).order)
  const first = Math.min(...orders)
  const last = Math.max(...orders)
  const paras = movableParas(doc)
  const before = paras.filter(loc => loc.order < first).reverse()[0]
  const after = paras.find(loc => loc.order > last)
  for (const loc of [before, after]) {
    if (!loc) continue
    if (listKindOf(doc, loc.para) !== kind) continue
    const info = listInfo(doc, loc.para)
    if (info) return info
  }
  return null
}

function freshNumId(doc: Doc): NumId {
  let min = 0
  for (const id of doc.lists.byNum.keys()) if (id < min) min = id
  return (min - 1) as NumId
}

function freshList(kind: 'bullet' | 'decimal'): ListDef {
  const levels: ListLevel[] = []
  for (let level = 0; level < 9; level++) {
    levels.push({
      format: kind,
      text: kind === 'bullet' ? '•' : `%${level + 1}.`,
      start: 1,
      // Word's default list: the marker sits on the margin (left - hanging = 0) and the text starts 21pt in;
      // each deeper level steps in by the same 21pt.
      indentLeft: asPt(21 * (level + 1)),
      hanging: asPt(21),
    })
  }
  return { fresh: true, levels }
}

function changeLevel(doc: Doc, para: Paragraph, delta: 1 | -1): Paragraph {
  const info = listInfo(doc, para)
  if (info) {
    const count = doc.lists.byNum.get(info.num)?.levels.length ?? 9
    const next = info.level + delta
    if (next < 0) return withNumbering(para, null)
    const level = Math.min(next, Math.max(0, count - 1))
    if (para.props.numbering && para.props.numbering.num === info.num && para.props.numbering.level === level) return para
    return withNumbering(para, { num: info.num, level })
  }
  const left = Math.max(0, (para.props.indent?.left ?? 0) + delta * 36)
  return { ...para, props: { ...para.props, indent: { ...para.props.indent, left: asPt(left) } } }
}

function ensureStyle(doc: Doc, style: StyleId): Doc {
  if (doc.styles.paragraph.has(style)) return doc
  const role = ROLES.find(item => item === style) ?? null
  if (!role) return doc
  const paragraph = new Map(doc.styles.paragraph)
  const next = role === 'normal' ? null : doc.styles.defaultParagraph
  paragraph.set(style, { name: role, role, next, look: plainLook() })
  const added = doc.styles.added.includes(role) ? doc.styles.added : doc.styles.added.concat(role)
  return { ...doc, styles: { ...doc.styles, paragraph, added } }
}

function plainLook(): ParaLook {
  const zero = asPt(0)
  return {
    align: 'left',
    indent: { left: zero, right: zero, firstLine: zero, hanging: zero, firstLineChars: 0 },
    spacing: { before: zero, after: zero, line: { rule: 'auto', multiple: 1 }, contextual: false },
    numbering: null,
    keepNext: false,
    keepLines: false,
    pageBreakBefore: false,
    widowControl: true,
  }
}

function effectiveAlign(doc: Doc, para: Paragraph): NonNullable<ParaProps['align']> {
  if (para.props.align) return para.props.align
  const style = para.props.style
  return (style ? doc.styles.paragraph.get(style)?.look.align : undefined) ?? 'left'
}

function sameLine(a: LineSpacing, b: LineSpacing): boolean {
  if (a.rule !== b.rule) return false
  if (a.rule === 'auto' && b.rule === 'auto') return a.multiple === b.multiple
  if (a.rule !== 'auto' && b.rule !== 'auto') return a.pt === b.pt
  return false
}

function stepPos(doc: Doc, pos: Pos, direction: 'backward' | 'forward', unit: LogicalUnit): Pos | null {
  const loc = safeLoc(indexOf(doc), pos.p)
  if (!loc || loc.story !== 'body') return null
  const paras = movableParas(doc)
  const text = flatText(loc.para)
  const o = clamp(pos.o, text.length)
  if (unit === 'document') {
    const target = direction === 'backward' ? paras[0] : paras[paras.length - 1]
    if (!target) return null
    return { p: target.para.id, o: direction === 'backward' ? 0 : flatText(target.para).length }
  }
  if (unit === 'paragraph') {
    if (direction === 'forward') {
      if (o < text.length) return { p: pos.p, o: text.length }
      const next = paras[paras.findIndex(item => item.para.id === pos.p) + 1]
      return next ? { p: next.para.id, o: flatText(next.para).length } : { p: pos.p, o }
    }
    if (o > 0) return { p: pos.p, o: 0 }
    const prev = paras[paras.findIndex(item => item.para.id === pos.p) - 1]
    return prev ? { p: prev.para.id, o: 0 } : { p: pos.p, o }
  }
  if (unit === 'word') {
    if (direction === 'forward') {
      const next = nextWordStop(text, o)
      if (next !== o) return { p: pos.p, o: next }
      const neighbor = paras[paras.findIndex(item => item.para.id === pos.p) + 1]
      return neighbor ? { p: neighbor.para.id, o: 0 } : { p: pos.p, o }
    }
    const prevStop = prevWordStop(text, o)
    if (prevStop !== o) return { p: pos.p, o: prevStop }
    const neighbor = paras[paras.findIndex(item => item.para.id === pos.p) - 1]
    return neighbor ? { p: neighbor.para.id, o: flatText(neighbor.para).length } : { p: pos.p, o }
  }
  if (direction === 'forward') {
    if (o < text.length) return { p: pos.p, o: nextBoundary(text, o) }
    const neighbor = paras[paras.findIndex(item => item.para.id === pos.p) + 1]
    return neighbor ? { p: neighbor.para.id, o: 0 } : { p: pos.p, o }
  }
  if (o > 0) return { p: pos.p, o: prevBoundary(text, o) }
  const neighbor = paras[paras.findIndex(item => item.para.id === pos.p) - 1]
  return neighbor ? { p: neighbor.para.id, o: flatText(neighbor.para).length } : { p: pos.p, o }
}

function cellEdge(doc: Doc, sel: Extract<Selection, { kind: 'cells' }>, edge: 'start' | 'end'): Pos | null {
  const found = findTable(doc.body, sel.table, [])
  if (!found) return null
  const row = edge === 'start' ? Math.min(sel.anchor.row, sel.head.row) : Math.max(sel.anchor.row, sel.head.row)
  const cell = edge === 'start' ? Math.min(sel.anchor.cell, sel.head.cell) : Math.max(sel.anchor.cell, sel.head.cell)
  const blocks = found.table.rows[row]?.cells[cell]?.blocks
  if (!blocks) return null
  if (edge === 'start') return firstPosIn(blocks)
  const paras: Paragraph[] = []
  collectParas(blocks, paras)
  const last = paras[paras.length - 1]
  return last ? { p: last.id, o: flatText(last).length } : null
}

function wordStops(text: string): number[] {
  const stops = [0]
  for (const part of words.segment(text)) {
    if (part.isWordLike === true && part.index !== 0) stops.push(part.index)
  }
  if (stops[stops.length - 1] !== text.length) stops.push(text.length)
  return stops
}

function nextWordStop(text: string, o: number): number {
  for (const stop of wordStops(text)) if (stop > o) return stop
  return text.length
}

function prevWordStop(text: string, o: number): number {
  let prev = 0
  for (const stop of wordStops(text)) {
    if (stop >= o) break
    prev = stop
  }
  return prev
}

/** The word a double click selects. An offset on whitespace uses the nearest word, preferring the one before. */
export function wordRange(text: string, offset: number): { start: number, end: number } {
  if (text.length === 0) return { start: 0, end: 0 }
  const o = offset < 0 ? 0 : offset > text.length ? text.length : offset
  return wordSpan(text, o)
}

function wordSpan(text: string, o: number): { start: number, end: number } {
  const parts = [...words.segment(text)]
  if (parts.length === 0 || text.length === 0) return { start: 0, end: 0 }
  let index = parts.findIndex(part => part.index <= o && o < part.index + part.segment.length)
  if (index < 0) index = parts.length - 1
  const pick = (at: number): { start: number, end: number } | null => {
    const part = parts[at]
    if (!part?.isWordLike) return null
    return { start: part.index, end: part.index + part.segment.length }
  }
  for (let i = index; i >= 0; i--) {
    const span = pick(i)
    if (span) return span
  }
  for (let i = index + 1; i < parts.length; i++) {
    const span = pick(i)
    if (span) return span
  }
  const part = parts[index]
  if (!part) return { start: 0, end: text.length }
  return { start: part.index, end: part.index + part.segment.length }
}

function prevBoundary(text: string, o: number): number {
  if (o <= 0) return 0
  let start = 0
  for (const part of graphemes.segment(text)) {
    const end = part.index + part.segment.length
    if (part.index >= o) break
    if (end >= o) return part.index
    start = part.index
  }
  return start
}

function nextBoundary(text: string, o: number): number {
  if (o >= text.length) return text.length
  for (const part of graphemes.segment(text)) {
    const end = part.index + part.segment.length
    if (end > o) return end
  }
  return text.length
}

function snapBack(text: string, o: number): number {
  if (o <= 0) return 0
  if (o >= text.length) return text.length
  for (const part of graphemes.segment(text)) {
    const end = part.index + part.segment.length
    if (o === part.index || o === end) return o
    if (part.index < o && o < end) return part.index
  }
  return o
}

function snapForward(text: string, o: number): number {
  if (o <= 0) return 0
  if (o >= text.length) return text.length
  for (const part of graphemes.segment(text)) {
    const end = part.index + part.segment.length
    if (o === part.index || o === end) return o
    if (part.index < o && o < end) return end
  }
  return o
}

function clamp(o: number, len: number): number {
  if (!Number.isFinite(o) || o < 0) return 0
  if (o > len) return len
  return o
}

function plainSlice(para: Paragraph, from: number, to: number): string {
  let out = ''
  let o = 0
  for (const inline of para.inlines) {
    if (inline.kind === 'text') {
      const start = Math.max(from, o)
      const end = Math.min(to, o + inline.text.length)
      if (start < end) out += inline.text.slice(start - o, end - o)
      o += inline.text.length
    } else {
      if (o >= from && o < to) out += atomPlain(inline)
      o += 1
    }
  }
  return out
}

function atomPlain(atom: Atom): string {
  if (atom.atom.type === 'field' || atom.atom.type === 'opaque') return atom.atom.display
  return ''
}

function joinParts(parts: readonly { row: string, cell: string, text: string }[]): string {
  const rows: string[] = []
  let rowKey = ''
  let cells: { key: string, texts: string[] }[] = []
  const flush = (): void => {
    if (cells.length === 0) return
    rows.push(cells.map(cell => cell.texts.join('\n')).join('\t'))
    cells = []
  }
  for (const part of parts) {
    if (part.row !== rowKey) {
      flush()
      rowKey = part.row
    }
    const last = cells[cells.length - 1]
    if (!last || last.key !== part.cell) cells.push({ key: part.cell, texts: [part.text] })
    else last.texts.push(part.text)
  }
  flush()
  return rows.join('\n')
}

function cellsText(doc: Doc, sel: Extract<Selection, { kind: 'cells' }>): string {
  const found = findTable(doc.body, sel.table, [])
  if (!found) return ''
  const r0 = Math.min(sel.anchor.row, sel.head.row)
  const r1 = Math.max(sel.anchor.row, sel.head.row)
  const c0 = Math.min(sel.anchor.cell, sel.head.cell)
  const c1 = Math.max(sel.anchor.cell, sel.head.cell)
  const lines: string[] = []
  for (let r = r0; r <= r1; r++) {
    const cols: string[] = []
    for (let c = c0; c <= c1; c++) {
      const cell = found.table.rows[r]?.cells[c]
      cols.push(cell ? blockPlain(cell.blocks) : '')
    }
    lines.push(cols.join('\t'))
  }
  return lines.join('\n')
}

function blockPlain(blocks: readonly Block[]): string {
  const parts: string[] = []
  const walk = (items: readonly Block[]): void => {
    for (const block of items) {
      if (block.kind === 'p') parts.push(plainSlice(block, 0, flatText(block).length))
      else if (block.kind === 'table') for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks)
    }
  }
  walk(blocks)
  return parts.join('\n')
}

function markSamples(state: EditorState): Marks[] {
  if (state.sel.kind === 'cells') return touchedParagraphs(state.doc, state.sel).map(() => ({}))
  const index = indexOf(state.doc)
  const ordered = orderRange(index, state.sel.anchor, state.sel.head)
  if (comparePos(index, ordered.from, ordered.to) === 0) {
    return [state.stored ?? marksAt(ordered.fromLoc.para, ordered.from.o).marks]
  }
  const covered = coveredText(state.doc, ordered)
  if (covered.length === 0) return [marksAt(ordered.fromLoc.para, ordered.from.o).marks]
  return covered.map(sample => sample.marks)
}

function vertAlignOf(marks: readonly Marks[]): ToolbarState['vertAlign'] {
  if (marks.length === 0) return 'none'
  const values = marks.map((item) => item.vertAlign === 'super' || item.vertAlign === 'sub' ? item.vertAlign : 'none')
  const first = values[0]
  if (first === undefined) return 'none'
  for (const value of values) if (value !== first) return 'mixed'
  return first
}

function caretParagraph(state: EditorState): Paragraph | null {
  if (state.sel.kind === 'text') return safeLoc(indexOf(state.doc), state.sel.head.p)?.para ?? null
  return touchedParagraphs(state.doc, state.sel)[0] ?? null
}

function sharedMultiple(paras: readonly Paragraph[]): number | null {
  if (paras.length === 0) return null
  let found: number | null = null
  for (const para of paras) {
    const line = para.props.spacing?.line
    if (!line || line.rule !== 'auto') return null
    if (found === null) found = line.multiple
    else if (found !== line.multiple) return null
  }
  return found
}

function runLooks(state: EditorState, look: (para: Paragraph, marks: Marks) => RunLook): RunLook[] {
  if (state.sel.kind === 'cells') {
    return touchedParagraphs(state.doc, state.sel).map(para => look(para, {}))
  }
  const index = indexOf(state.doc)
  const ordered = orderRange(index, state.sel.anchor, state.sel.head)
  if (comparePos(index, ordered.from, ordered.to) === 0) {
    const marks = state.stored ?? marksAt(ordered.fromLoc.para, ordered.from.o).marks
    return [look(ordered.fromLoc.para, marks)]
  }
  const covered = coveredText(state.doc, ordered)
  if (covered.length === 0) {
    const marks = marksAt(ordered.fromLoc.para, ordered.from.o).marks
    return [look(ordered.fromLoc.para, marks)]
  }
  return covered.map(sample => look(sample.para, sample.marks))
}

function tri(values: readonly boolean[]): Tri {
  if (values.length === 0) return false
  if (values.every(value => value)) return true
  if (values.every(value => !value)) return false
  return 'mixed'
}

function uniform<T>(values: readonly T[]): T | 'mixed' {
  const first = values[0]
  if (first === undefined) return 'mixed'
  for (const value of values) if (value !== first) return 'mixed'
  return first
}

function singleTouched(before: Doc, after: Doc): ParaId | null {
  if (before.body === after.body) return null
  const changed = diffBlocks(before.body, after.body)
  if (!changed || changed.length !== 1) return null
  return changed[0] ?? null
}

function diffBlocks(before: readonly Block[], after: readonly Block[]): ParaId[] | null {
  if (before.length !== after.length) return null
  const changed: ParaId[] = []
  for (let i = 0; i < before.length; i++) {
    const left = before[i]
    const right = after[i]
    if (left === right) continue
    if (!left || !right || left.kind !== right.kind) return null
    if (left.kind === 'p' && right.kind === 'p') {
      if (left.id !== right.id) return null
      changed.push(right.id)
      continue
    }
    if (left.kind === 'table' && right.kind === 'table') {
      if (left.id !== right.id || left.rows.length !== right.rows.length) return null
      for (let r = 0; r < left.rows.length; r++) {
        const lr = left.rows[r]
        const rr = right.rows[r]
        if (!lr || !rr) return null
        if (lr === rr) continue
        if (lr.cells.length !== rr.cells.length) return null
        for (let c = 0; c < lr.cells.length; c++) {
          const lc = lr.cells[c]
          const rc = rr.cells[c]
          if (!lc || !rc) return null
          if (lc === rc) continue
          if (lc.blocks.length !== rc.blocks.length) return null
          const inner = diffBlocks(lc.blocks, rc.blocks)
          if (!inner) return null
          changed.push(...inner)
        }
      }
      continue
    }
    return null
  }
  return changed
}

function asPt(value: number): Pt {
  return value as Pt
}

function finishContainer(blocks: readonly Block[], container: readonly number[], inCell: boolean): readonly Block[] {
  if (!inCell) {
    return firstPosIn(blocks) ? blocks : blocks.concat(emptyParagraph())
  }
  return rewrite(blocks, container, (siblings) => finishBlocks(siblings, true))
}
