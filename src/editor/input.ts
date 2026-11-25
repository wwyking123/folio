// The pure half of the input layer: the keymap, chord naming, and the input mode machine. No DOM, no editor state.
import type { Pos } from '../model.ts'
import {
  deleteBackward,
  deleteForward,
  insertLineBreak,
  insertPageBreak,
  insertText,
  move,
  selectAll,
  splitParagraph,
  clearMarks,
  tabKey,
  toggleMark,
  toggleVertAlign,
  wordRange,
} from './edit.ts'
import type { Command, LogicalUnit } from './edit.ts'

export type Platform = 'mac' | 'other'

/** Motion that needs line geometry, so edit.ts cannot express it as a Command. PageView resolves the target Pos. */
export type VisualMotion = 'lineUp' | 'lineDown' | 'pageUp' | 'pageDown' | 'lineStart' | 'lineEnd'

export type Action =
  | { readonly kind: 'command'; readonly command: Command }
  | { readonly kind: 'visual'; readonly motion: VisualMotion; readonly extend: boolean }
  | { readonly kind: 'history'; readonly step: 'undo' | 'redo' }
  | { readonly kind: 'clipboard'; readonly op: 'copy' | 'cut' | 'paste' }
  | { readonly kind: 'file'; readonly op: FileOp }
  | { readonly kind: 'find'; readonly op: FindOp }

/** Requests the editor passes up to the app: the file on this tab, and the tab strip around it. */
export type FileOp = 'save' | 'saveAs' | 'new' | 'open' | 'closeTab' | 'nextTab' | 'prevTab'

export type FindOp = 'open' | 'replace' | 'next' | 'previous'

export type Binding = readonly [chord: string, action: Action]

/**
 * One table, read by the keydown handler and by the menus, so a shortcut cannot exist in one place and be missing from
 * the other. Keys are chord names as chordOf spells them.
 */
export type Keymap = ReadonlyMap<string, Action>

export type DragUnit = 'char' | 'word' | 'paragraph'

export type InputMode =
  | { readonly kind: 'idle' }
  /** `at` is the caret where the preedit is drawn, taken after compositionstart deleted any selection. */
  | { readonly kind: 'composing'; readonly at: Pos }
  /** `unit` is 'word' or 'paragraph' after a double or triple click, and the drag then extends by that unit. */
  | { readonly kind: 'dragging'; readonly anchor: Pos; readonly unit: DragUnit }

export type ModeEvent =
  | { readonly kind: 'compositionStart'; readonly at: Pos }
  | { readonly kind: 'compositionEnd' }
  | { readonly kind: 'pointerDown'; readonly at: Pos; readonly unit: DragUnit }
  | { readonly kind: 'pointerUp' }
  /** Blur, a new document, or destroy: whatever was in flight is abandoned. */
  | { readonly kind: 'reset' }

export type KeyFields = Pick<KeyboardEvent, 'key' | 'ctrlKey' | 'metaKey' | 'altKey' | 'shiftKey'>

export const IDLE: InputMode = { kind: 'idle' }

/**
 * Every (mode, event) pair has an answer. An event that means nothing in the current mode returns `mode` itself, so a
 * caller can tell "ignored" from "moved" by identity.
 */
export function transition(mode: InputMode, event: ModeEvent): InputMode {
  switch (mode.kind) {
    case 'idle':
      switch (event.kind) {
        case 'compositionStart': return { kind: 'composing', at: event.at }
        case 'pointerDown': return { kind: 'dragging', anchor: event.at, unit: event.unit }
        case 'compositionEnd':
        case 'pointerUp':
        case 'reset': return mode
      }
    case 'composing':
      switch (event.kind) {
        case 'compositionEnd':
        case 'reset': return IDLE
        // A click cannot start a drag until the IME commits, or the committed text would land at the old caret.
        case 'pointerDown':
        case 'compositionStart':
        case 'pointerUp': return mode
      }
    case 'dragging':
      switch (event.kind) {
        // Typing through an IME while the button is down wins over the drag, so the committed text is not lost.
        case 'compositionStart': return { kind: 'composing', at: event.at }
        // A press while dragging means the earlier mouseup never arrived (released outside the window).
        case 'pointerDown': return { kind: 'dragging', anchor: event.at, unit: event.unit }
        case 'pointerUp':
        case 'reset': return IDLE
        case 'compositionEnd': return mode
      }
  }
}

type ModifierField = Exclude<keyof KeyFields, 'key'>

// "mod" is the platform's primary modifier, so Cmd on macOS and Ctrl elsewhere both read "mod". The other one of the
// pair keeps its own name. The order of each list is the canonical modifier order in a chord.
const MODIFIERS: Readonly<Record<Platform, readonly (readonly [name: string, field: ModifierField])[]>> = {
  mac: [['mod', 'metaKey'], ['ctrl', 'ctrlKey'], ['alt', 'altKey'], ['shift', 'shiftKey']],
  other: [['mod', 'ctrlKey'], ['meta', 'metaKey'], ['alt', 'altKey'], ['shift', 'shiftKey']],
}

// A space would split a chord's words and "+" would be read as the separator.
const KEY_NAMES: ReadonlyMap<string, string> = new Map([[' ', 'space'], ['+', 'plus']])

/** "mod+b", "mod+shift+z", "shift+arrowdown": modifiers in canonical order, then the lower-cased key, joined by "+". */
export function chordOf(event: KeyFields, platform: Platform): string {
  const held = MODIFIERS[platform].filter(([, field]) => event[field]).map(([name]) => name)
  return [...held, KEY_NAMES.get(event.key) ?? event.key.toLowerCase()].join('+')
}

/** Refuses a chord bound twice, which a Map would otherwise settle silently in favor of the later one. */
export function buildKeymap(bindings: Iterable<Binding>): Keymap {
  const map = new Map<string, Action>()
  for (const [chord, action] of bindings) {
    if (map.has(chord)) throw new RangeError(`chord "${chord}" is bound twice`)
    map.set(chord, action)
  }
  return map
}

function run(command: Command): Action {
  return { kind: 'command', command }
}

// One command instance per mark, shared with the keymap, so a toolbar title can find the chord by identity.
const markCommand = {
  bold: toggleMark('bold'),
  italic: toggleMark('italic'),
  underline: toggleMark('underline'),
} as const

// Shift and "=" arrive as key "+", named "plus" so a chord can be split on "+".
const formatCommand = {
  superscript: toggleVertAlign('super'),
  subscript: toggleVertAlign('sub'),
  clearFormatting: clearMarks,
} as const

export type ShortcutName =
  | 'bold'
  | 'italic'
  | 'underline'
  | 'undo'
  | 'redo'
  | 'save'
  | 'saveAs'
  | 'new'
  | 'open'
  | 'superscript'
  | 'subscript'
  | 'clearFormatting'
  | 'find'
  | 'replace'
  | 'findNext'
  | 'findPrevious'

/** The chord `defaultKeymap` binds for a toolbar command, or null when the keymap has none. Redo is the first binding. */
export function shortcutChord(platform: Platform, name: ShortcutName): string | null {
  for (const [chord, action] of defaultKeymap(platform)) {
    if (action.kind === 'history' && action.step === name) return chord
    if (action.kind === 'file' && action.op === name) return chord
    if (action.kind === 'find' && findName(action.op) === name) return chord
    if (
      action.kind === 'command'
      && (name === 'bold' || name === 'italic' || name === 'underline')
      && action.command === markCommand[name]
    ) return chord
    if (action.kind === 'command' && name === 'superscript' && action.command === formatCommand.superscript) return chord
    if (action.kind === 'command' && name === 'subscript' && action.command === formatCommand.subscript) return chord
    if (action.kind === 'command' && name === 'clearFormatting' && action.command === formatCommand.clearFormatting) {
      return chord
    }
  }
  return null
}

function findName(op: FindOp): ShortcutName | null {
  if (op === 'open') return 'find'
  if (op === 'replace') return 'replace'
  if (op === 'next') return 'findNext'
  return 'findPrevious'
}

function step(direction: 'backward' | 'forward', unit: LogicalUnit): (extend: boolean) => Action {
  return (extend) => run(move(direction, unit, extend))
}

function glide(motion: VisualMotion): (extend: boolean) => Action {
  return (extend) => ({ kind: 'visual', motion, extend })
}

/** The chord as given, and the chord with Shift added, which extends the selection. Shift sorts last before the key. */
function extendable(chord: string, make: (extend: boolean) => Action): Binding[] {
  const parts = chord.split('+')
  const key = parts.pop()
  return [[chord, make(false)], [[...parts, 'shift', key].join('+'), make(true)]]
}

/**
 * Word motion and word deletion sit on Alt on macOS and on Ctrl elsewhere. macOS then spends Cmd on line and document
 * edges and Alt+Up/Down on paragraphs. Elsewhere those are Home, End and Ctrl+Home/End.
 */
export function defaultKeymap(platform: Platform): Keymap {
  const word = platform === 'mac' ? 'alt' : 'mod'
  const common: Binding[] = [
    ...extendable('arrowleft', step('backward', 'char')),
    ...extendable('arrowright', step('forward', 'char')),
    ...extendable('arrowup', glide('lineUp')),
    ...extendable('arrowdown', glide('lineDown')),
    ...extendable('pageup', glide('pageUp')),
    ...extendable('pagedown', glide('pageDown')),
    ...extendable('home', glide('lineStart')),
    ...extendable('end', glide('lineEnd')),
    ...extendable('mod+home', step('backward', 'document')),
    ...extendable('mod+end', step('forward', 'document')),
    ...extendable(`${word}+arrowleft`, step('backward', 'word')),
    ...extendable(`${word}+arrowright`, step('forward', 'word')),
    ['enter', run(splitParagraph)],
    ['shift+enter', run(insertLineBreak)],
    ['mod+enter', run(insertPageBreak)],
    ['tab', run(tabKey(1))],
    ['shift+tab', run(tabKey(-1))],
    ['backspace', run(deleteBackward('char'))],
    ['delete', run(deleteForward('char'))],
    [`${word}+backspace`, run(deleteBackward('word'))],
    [`${word}+delete`, run(deleteForward('word'))],
    ['mod+a', run(selectAll)],
    ['mod+b', run(markCommand.bold)],
    ['mod+i', run(markCommand.italic)],
    ['mod+u', run(markCommand.underline)],
    ['mod+shift+plus', run(formatCommand.superscript)],
    ['mod+=', run(formatCommand.subscript)],
    ['mod+space', run(formatCommand.clearFormatting)],
    ['mod+z', { kind: 'history', step: 'undo' }],
    ['mod+y', { kind: 'history', step: 'redo' }],
    ['mod+shift+z', { kind: 'history', step: 'redo' }],
    ['mod+s', { kind: 'file', op: 'save' }],
    ['mod+shift+s', { kind: 'file', op: 'saveAs' }],
    ['mod+n', { kind: 'file', op: 'new' }],
    ['mod+o', { kind: 'file', op: 'open' }],
    ['mod+w', { kind: 'file', op: 'closeTab' }],
    ['mod+c', { kind: 'clipboard', op: 'copy' }],
    ['mod+x', { kind: 'clipboard', op: 'cut' }],
    ['mod+v', { kind: 'clipboard', op: 'paste' }],
    // ⌘H hides the app on macOS and never reaches the page, so replace is ctrl+h there.
    ['mod+f', { kind: 'find', op: 'open' }],
    [platform === 'mac' ? 'ctrl+h' : 'mod+h', { kind: 'find', op: 'replace' }],
    ['mod+g', { kind: 'find', op: 'next' }],
    ['mod+shift+g', { kind: 'find', op: 'previous' }],
    ['f3', { kind: 'find', op: 'next' }],
    ['shift+f3', { kind: 'find', op: 'previous' }],
  ]
  const mac: Binding[] = [
    ...extendable('mod+arrowleft', glide('lineStart')),
    ...extendable('mod+arrowright', glide('lineEnd')),
    ...extendable('mod+arrowup', step('backward', 'document')),
    ...extendable('mod+arrowdown', step('forward', 'document')),
    ...extendable('alt+arrowup', step('backward', 'paragraph')),
    ...extendable('alt+arrowdown', step('forward', 'paragraph')),
  ]
  const ctrl = platform === 'mac' ? 'ctrl' : 'mod'
  const tabs: Binding[] = [
    [`${ctrl}+tab`, { kind: 'file', op: 'nextTab' }],
    [`${ctrl}+shift+tab`, { kind: 'file', op: 'prevTab' }],
  ]
  return buildKeymap(platform === 'mac' ? [...common, ...tabs, ...mac] : [...common, ...tabs])
}

export { wordRange }

export interface TextRange {
  readonly start: number
  readonly end: number
}

/** The whole paragraph, which is what a triple click selects. */
export function paragraphRange(text: string): TextRange {
  return { start: 0, end: text.length }
}

/** Click count on a primary button: one click is a caret, two a word, three or more a paragraph. */
export function dragUnit(detail: number): DragUnit {
  if (detail >= 3) return 'paragraph'
  if (detail === 2) return 'word'
  return 'char'
}

function clampOffset(offset: number, length: number): number {
  if (offset < 0) return 0
  if (offset > length) return length
  return offset
}

function unitRange(text: string, offset: number, unit: DragUnit): TextRange {
  if (unit === 'paragraph') return paragraphRange(text)
  if (unit === 'word') return wordRange(text, offset)
  const o = clampOffset(offset, text.length)
  return { start: o, end: o }
}

/**
 * Offsets for a press or a drag inside one paragraph. A forward drag anchors at the start of the press's unit
 * and ends at the end of the unit under the pointer, so a double click selects the word and a drag grows by words.
 */
export function gestureOffsets(
  text: string,
  anchor: number,
  head: number,
  unit: DragUnit,
): { readonly anchor: number; readonly head: number } {
  const a = clampOffset(anchor, text.length)
  const h = clampOffset(head, text.length)
  if (unit === 'char') return { anchor: a, head: h }
  const from = unitRange(text, a, unit)
  const to = unitRange(text, h, unit)
  if (h < a) return { anchor: from.end, head: to.start }
  return { anchor: from.start, head: to.end }
}

/**
 * Endpoints when the drag crosses a paragraph boundary. `headBeforeAnchor` is document order, not string order.
 * The caller keeps the paragraph ids; only the offsets change.
 */
export function dragEndpoints(
  anchorText: string,
  anchorOffset: number,
  headText: string,
  headOffset: number,
  sameParagraph: boolean,
  headBeforeAnchor: boolean,
  unit: DragUnit,
): { readonly anchor: number; readonly head: number } {
  if (sameParagraph) return gestureOffsets(anchorText, anchorOffset, headOffset, unit)
  if (unit === 'char') {
    return {
      anchor: clampOffset(anchorOffset, anchorText.length),
      head: clampOffset(headOffset, headText.length),
    }
  }
  const from = unitRange(anchorText, anchorOffset, unit)
  const to = unitRange(headText, headOffset, unit)
  if (headBeforeAnchor) return { anchor: from.end, head: to.start }
  return { anchor: from.start, head: to.end }
}

const VERTICAL: ReadonlySet<VisualMotion> = new Set(['lineUp', 'lineDown', 'pageUp', 'pageDown'])

export function isVerticalMotion(motion: VisualMotion): boolean {
  return VERTICAL.has(motion)
}

/**
 * The goal column survives only consecutive vertical motions. `caretX` is the column measured when a run starts
 * (`kept` is null); later moves in the run ignore a new measurement.
 */
export function trackGoal(kept: number | null, action: Action, caretX: number | null): number | null {
  if (action.kind !== 'visual' || !isVerticalMotion(action.motion)) return null
  return kept ?? caretX
}

export type BeforeInputPlan =
  | { readonly kind: 'run'; readonly command: Command }
  | { readonly kind: 'compose' }
  | { readonly kind: 'ignore' }

const COMPOSITION_INPUT: ReadonlySet<string> = new Set([
  'insertCompositionText',
  'deleteCompositionText',
  'insertFromComposition',
  'deleteByComposition',
])

/**
 * The acde runtime paints the composition itself: a fixed span at the textarea caret, solid underline,
 * `background:Canvas`. The editor used to paint a second, dotted copy in the paragraph and then move the
 * textarea to the end of that copy, so the runtime drew the same pinyin again. The runtime's span is the one
 * that matches the system candidate window.
 */
export function isHostPreeditStyle(style: string): boolean {
  const compact = style.replace(/\s+/g, '')
  return compact.includes('background:Canvas')
    && compact.includes('text-decoration:underline')
    && compact.includes('white-space:pre')
}

/**
 * Places the runtime's fixed composition span on the paragraph's baseline. `anchor` is the collapsed caret rect in
 * viewport pixels; `metric` is the 1x inline offset (`preedit.top - anchor.top`) and content height measured side by
 * side on one baseline.
 */
export function hostPreeditBox(
  anchor: { readonly left: number; readonly top: number },
  metric: { readonly deltaTop: number; readonly height: number },
  zoom: number,
): { readonly left: number; readonly top: number; readonly height: number } {
  const z = zoom > 0 ? zoom : 1
  return {
    left: anchor.left,
    top: anchor.top + metric.deltaTop * z,
    height: Math.max(metric.height * z, 1),
  }
}

/** Caret at the right edge of the runtime preedit, in the overlay's own pixels. `width` is the laid-out box. */
export function caretAtPreeditEnd(
  span: { readonly right: number; readonly top: number; readonly height: number },
  overlay: { readonly left: number; readonly top: number; readonly width: number; readonly offsetWidth: number },
): { readonly left: number; readonly top: number; readonly height: number } {
  const scale = overlay.offsetWidth > 0 ? overlay.width / overlay.offsetWidth : 1
  const z = scale > 0 ? scale : 1
  return {
    left: (span.right - overlay.left) / z,
    top: (span.top - overlay.top) / z,
    height: Math.max(span.height / z, 1),
  }
}

/**
 * What a `beforeinput` should do. Composition is left to the composition events, including when a commit's
 * `insertText` arrives after `compositionend` has already been classified as a composition type.
 * While `composing` is true every type is left alone, so the textarea can hold the preedit.
 */
export function planBeforeInput(inputType: string, data: string | null, composing: boolean): BeforeInputPlan {
  if (composing || COMPOSITION_INPUT.has(inputType)) return { kind: 'compose' }
  switch (inputType) {
    case 'insertText':
      return data !== null && data.length > 0 ? { kind: 'run', command: insertText(data) } : { kind: 'ignore' }
    case 'insertLineBreak':
      return { kind: 'run', command: insertLineBreak }
    case 'insertParagraph':
      return { kind: 'run', command: splitParagraph }
    case 'deleteContentBackward':
      return { kind: 'run', command: deleteBackward('char') }
    case 'deleteContentForward':
      return { kind: 'run', command: deleteForward('char') }
    default:
      return { kind: 'ignore' }
  }
}

/** `MacIntel` and a Macintosh user agent are mac; Dopame OS reports a Linux platform and is `other`. */
export function detectPlatform(platform: string, userAgent = ''): Platform {
  return /mac/i.test(platform) || /mac/i.test(userAgent) ? 'mac' : 'other'
}

/** What the clipboard's plain text holds for one copied picture: the object replacement character. */
export const PICTURE_TEXT = '\uFFFC'

/** True for a file the engine can embed as a picture. The engine still checks the bytes. */
export function isPictureFile(file: { readonly type: string; readonly name?: string }): boolean {
  if (file.type === 'image/png' || file.type === 'image/jpeg') return true
  if (file.type !== '') return false
  return /\.(png|jpe?g)$/i.test(file.name ?? '')
}

/**
 * What a paste inserts. Picture files on the clipboard win. Text that is the mark this app wrote for a copied
 * picture brings that picture back, since the runtime's clipboard carries text only. Any other text is typed.
 */
export function pastePlan(
  text: string,
  pictureFiles: number,
  held: { readonly text: string } | null,
): 'files' | 'held' | 'text' | 'none' {
  if (pictureFiles > 0) return 'files'
  if (held && text === held.text) return 'held'
  return text.length > 0 ? 'text' : 'none'
}
