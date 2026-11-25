import { isPrintBoldKey } from '../fonts.ts'
import { PrintError } from './print-error.ts'

export { PrintError }
import init, { PrintLayout as WasmLayout, Session, type MediaFile } from './pkg/engine.js'
import type {
  Atom,
  Block,
  BorderEdge,
  Cell,
  CellBorders,
  CellMargins,
  Doc,
  Hex,
  HighlightName,
  Inline,
  LineSpacing,
  ListDef,
  ListLevel,
  ListTable,
  Lock,
  Marks,
  MediaKey,
  NumId,
  OpaqueBlock,
  OpaqueId,
  ParaId,
  ParaSrc,
  ParaLook,
  ParaProps,
  Paragraph,
  Pt,
  Row,
  RowSrc,
  RunLook,
  RunSrc,
  Section,
  SectionSrc,
  StoryKey,
  StyleId,
  StyleRole,
  Table,
  TableBorders,
  TableId,
  TableSrc,
  CellSrc,
} from '../model.ts'
import type * as wire from './wire.ts'

export class OpenError extends Error {
  readonly reason: 'notDocx' | 'encrypted' | 'corrupt' | 'tooLarge'

  constructor(reason: 'notDocx' | 'encrypted' | 'corrupt' | 'tooLarge', detail: string) {
    super(detail)
    this.name = 'OpenError'
    this.reason = reason
  }
}

/** The written document did not re-import to the model, or lost an unmodeled item. Nothing was written. */
export class SaveVerifyError extends Error {
  readonly firstDifference: string

  constructor(firstDifference: string) {
    super(firstDifference)
    this.name = 'SaveVerifyError'
    this.firstDifference = firstDifference
  }
}

/** `write` refused the patch. `detail` is the contract or kit message, shown as-is. */
export class SaveWriteError extends Error {
  readonly detail: string

  constructor(detail: string) {
    super(detail)
    this.name = 'SaveWriteError'
    this.detail = detail
  }
}

export type PrintFontInfo = {
  readonly family: string
  readonly embedded: boolean
}

/** One model body block. Pages are 0-based. Both are null when the block was not placed. */
export type BlockPageInfo = {
  readonly index: number
  readonly firstPage: number | null
  readonly lastPage: number | null
}

/** One placed fragment. `page` is 0-based. `y` and `height` are points. */
export type BlockFragmentInfo = {
  readonly page: number
  readonly y: number
  readonly height: number
}

/** Fragments of one body block, in page order. Empty when the block was not placed. */
export type BlockBoxInfo = {
  readonly index: number
  readonly fragments: readonly BlockFragmentInfo[]
}

/** One layout. Preview and PDF both read it. `free` drops the wasm object. */
export type PrintLayout = {
  pageCount(): number
  pageSize(index: number): readonly [number, number]
  pagePng(index: number, dpi: number): Uint8Array
  pdf(): Uint8Array
  fonts(): readonly PrintFontInfo[]
  blockPages(): readonly BlockPageInfo[]
  blockBoxes(): readonly BlockBoxInfo[]
  free(): void
}

/** Bytes of an image added this session. Every save starts from the opened document, so the bytes go out again. */
export type AddedMedia = {
  readonly key: MediaKey
  readonly contentType: string
  readonly widthPt: Pt
  readonly heightPt: Pt
}

export type Engine = {
  open(bytes: Uint8Array): Opened
  blank(): Opened
}

export type Opened = {
  readonly doc: Doc
  readonly source: Source
}

export type Source = {
  save(doc: Doc): Uint8Array
  /** Lay out `doc` once. The caller frees the layout. */
  printLayout(doc: Doc): PrintLayout
  fontsUsed(doc: Doc): readonly string[]
  /** Store `bytes` as `family` for the rest of the session. A second add replaces the first. */
  addFont(family: string, bytes: Uint8Array): void
  /** Map a requested family onto a family passed to `addFont`. */
  addFontAlias(requested: string, loaded: string): void
  addMedia(file: Blob): Promise<AddedMedia>
  imageUrl(key: MediaKey): string
  /** The bytes of one picture, from the file or added this session. Null when the key names none. */
  mediaBlob(key: MediaKey): Blob | null
  /** Resolved run look for the paragraph's styles, with these direct marks overlaid. */
  look(para: Paragraph, marks: Marks): RunLook
  /** Resolved paragraph look: the style, then the level's left and hanging indent from the given list table, then direct props. */
  paraLook(para: Paragraph, lists: ListTable): ParaLook
  free(): void
}

type LookTables = {
  run: ReadonlyMap<string, RunLook>
  para: ReadonlyMap<string, ParaLook>
}

type HeldMedia = { bytes: Uint8Array; contentType: string }

type SaveBaseline = {
  readonly blocks: ReadonlyMap<number, Block>
  readonly sections: readonly wire.Section[]
  readonly sectionList: readonly Section[]
}

const baselines = new WeakMap<Source, SaveBaseline>()
const tamperSaves = new WeakMap<Source, (doc: Doc) => Uint8Array>()
const sessionMedia = new WeakMap<Source, Map<string, HeldMedia>>()
const sessionFaceLogOf = new WeakMap<Source, { family: string; bytes: number }[]>()

/** Faces `addFont` has stored on this session, in the order they were first added. */
export function sessionFaces(source: Source): readonly { readonly family: string; readonly bytes: number }[] {
  return (sessionFaceLogOf.get(source) ?? []).map((face) => ({ family: face.family, bytes: face.bytes }))
}

let bootMs = 0
let wasmMemory: WebAssembly.Memory | undefined

/** How long `loadEngine` took to fetch and instantiate `engine.wasm`. */
export function engineBootMs(): number {
  return bootMs
}

/** The wasm memory size now. Grows as layouts run; it does not shrink. */
export function wasmMemoryBytes(): number {
  return wasmMemory?.buffer.byteLength ?? 0
}

const ENGINE_URL = '/engine.wasm'
const ENGINE_RETRY_MS = 100

type LoadAttempt =
  | { ok: true; response: Response }
  | { ok: false; status: string }

async function attemptLoad(load: (url: string) => Promise<Response>, url: string): Promise<LoadAttempt> {
  try {
    const response = await load(url)
    if (response.ok) return { ok: true, response }
    return { ok: false, status: String(response.status) }
  } catch (error) {
    const status = error instanceof Error && error.message !== '' ? error.message : String(error)
    return { ok: false, status: status === '' ? 'error' : status }
  }
}

/**
 * One more try after `wait` when `load` throws or the response is not ok.
 * The rejection names `url` and the status of the attempt that was kept.
 */
export async function readWithRetry(
  url: string,
  load: (url: string) => Promise<Response>,
  wait: () => Promise<void>,
): Promise<Response> {
  const first = await attemptLoad(load, url)
  if (first.ok) return first.response
  await wait()
  const second = await attemptLoad(load, url)
  if (second.ok) return second.response
  throw new Error(`${url}: ${second.status}`)
}

function wait(ms: number): Promise<void> {
  return new Promise((resolve) => { setTimeout(resolve, ms) })
}

export async function loadEngine(): Promise<Engine> {
  const started = performance.now()
  const response = await readWithRetry(ENGINE_URL, (url) => fetch(url), () => wait(ENGINE_RETRY_MS))
  const bytes = await response.arrayBuffer()
  const ready = await init({ module_or_path: bytes })
  bootMs = performance.now() - started
  wasmMemory = ready.memory
  return createEngine()
}

/** An engine bound to a wasm instance that the caller already initialized. */
export function createEngine(): Engine {
  return {
    open(bytes) {
      let session: Session
      try {
        session = Session.open(bytes)
      } catch (error) {
        throw toOpenError(error)
      }
      return openSession(session)
    },
    blank() {
      return openSession(Session.blank())
    },
  }
}

function openSession(session: Session): Opened {
  let parsed: unknown
  try {
    parsed = JSON.parse(session.import())
  } catch (error) {
    session.free()
    const detail = error instanceof Error ? error.message : String(error)
    throw new OpenError('corrupt', detail)
  }
  let imported: wire.Imported
  try {
    imported = parseImported(parsed)
  } catch (error) {
    session.free()
    const detail = error instanceof Error ? error.message : String(error)
    throw new OpenError('corrupt', detail)
  }
  const built = toDoc(imported)
  const blocks = new Map<number, Block>()
  for (const block of built.doc.body) {
    if (block.src !== null) blocks.set(block.src, block)
  }
  return {
    doc: built.doc,
    source: bindSource(session, built.doc, built.looks, imported.fonts, {
      blocks,
      sections: imported.sections,
      sectionList: built.doc.sections,
    }),
  }
}

function bindSource(
  session: Session,
  opened: Doc,
  looks: LookTables,
  importedFonts: readonly string[],
  baseline: SaveBaseline,
): Source {
  const urls = new Map<string, string>()
  const added = new Map<string, HeldMedia>()
  const extraFonts: string[] = []
  const sessionFaceLog: { family: string; bytes: number }[] = []
  let addedCount = 0
  const source: Source = {
    save(doc) {
      try {
        return copyBytes(session.save(JSON.stringify(toPatch(doc, baseline, added))))
      } catch (error) {
        throw toSaveError(error)
      }
    },
    printLayout(doc) {
      try {
        return wrapLayout(session.print_layout(JSON.stringify(toPatch(doc, baseline, added))))
      } catch (error) {
        throw asPrintError(error)
      }
    },
    fontsUsed(current) {
      return fontsUsedIn(current, looks.run, [...importedFonts, ...extraFonts])
    },
    addFont(family, bytes) {
      const trimmed = family.trim()
      if (trimmed === '') return
      session.add_font(trimmed, bytes)
      const recorded = sessionFaceLog.find((face) => face.family === trimmed)
      if (recorded === undefined) sessionFaceLog.push({ family: trimmed, bytes: bytes.byteLength })
      else recorded.bytes = bytes.byteLength
      // The bold companion is the same family. Listing `${family} Bold` would put it in the font menu.
      if (isPrintBoldKey(trimmed) || extraFonts.includes(trimmed)) return
      extraFonts.push(trimmed)
    },
    addFontAlias(requested, loaded) {
      const from = requested.trim()
      const to = loaded.trim()
      if (from === '' || to === '') return
      session.add_font_alias(from, to)
    },
    async addMedia(file) {
      const bytes = new Uint8Array(await file.arrayBuffer())
      const contentType = sniff(bytes)
      if (contentType !== 'image/png' && contentType !== 'image/jpeg') {
        throw new Error('not a png or jpeg')
      }
      const size = await pixelSize(bytes, contentType)
      const widthPt = pt(size.width * 0.75)
      const heightPt = pt(size.height * 0.75)
      const key = mediaKey(`m${addedCount}`)
      addedCount += 1
      added.set(key, { bytes, contentType })
      return { key, contentType, widthPt, heightPt }
    },
    imageUrl(key) {
      const existing = urls.get(key)
      if (existing !== undefined) return existing
      const held = added.get(key) ?? takeSessionMedia(session, key)
      if (held === undefined) throw new Error(`missing media ${key}`)
      const url = URL.createObjectURL(new Blob([new Uint8Array(held.bytes)], { type: held.contentType }))
      urls.set(key, url)
      return url
    },
    mediaBlob(key) {
      const held = added.get(key) ?? takeSessionMedia(session, key)
      if (held === undefined) return null
      return new Blob([new Uint8Array(held.bytes)], { type: held.contentType })
    },
    look(para, marks) {
      return overlayRun(runLookFor(looks, opened, para, marks.charStyle), marks)
    },
    paraLook(para, lists) {
      return overlayPara(paraLookFor(looks, opened, para), para.props, lists)
    },
    free() {
      for (const url of urls.values()) URL.revokeObjectURL(url)
      urls.clear()
      session.free()
    },
  }
  baselines.set(source, baseline)
  sessionMedia.set(source, added)
  sessionFaceLogOf.set(source, sessionFaceLog)
  tamperSaves.set(source, (doc) => {
    try {
      return copyBytes(session.save_tampered(JSON.stringify(toPatch(doc, baseline, added))))
    } catch (error) {
      throw toSaveError(error)
    }
  })
  return source
}

/** The patch `source.save` would send. Tests check keep, src, new lists and added styles. */
export function documentPatch(source: Source, doc: Doc): wire.Patch {
  const baseline = baselines.get(source)
  if (!baseline) throw new Error('save is not available')
  return toPatch(doc, baseline, sessionMedia.get(source) ?? new Map())
}

/**
 * `save`, then a bookmark is dropped before verification. The dev hook is the only caller.
 * A document with no bookmark still fails verification.
 */
export function saveTampered(source: Source, doc: Doc): Uint8Array {
  const run = tamperSaves.get(source)
  if (!run) throw new Error('save is not available')
  return run(doc)
}

function takeSessionMedia(session: Session, key: string): HeldMedia | undefined {
  const file: MediaFile | undefined = session.media(key)
  if (file === undefined) return undefined
  const held = { bytes: file.bytes.slice(), contentType: file.contentType }
  file.free()
  return held
}

function runLookFor(looks: LookTables, doc: Doc, para: Paragraph, charStyle: StyleId | null | undefined): RunLook {
  const paraStyle = para.props.style ?? doc.styles.defaultParagraph
  const character = charStyle ?? ''
  return looks.run.get(`${paraStyle}\0${character}`) ?? looks.run.get(`${paraStyle}\0`) ?? defaultRunLook()
}

function paraLookFor(looks: LookTables, doc: Doc, para: Paragraph): ParaLook {
  const paraStyle = para.props.style ?? doc.styles.defaultParagraph
  return looks.para.get(paraStyle) ?? defaultParaLook()
}

function overlayRun(base: RunLook, marks: Marks): RunLook {
  const fonts =
    marks.font !== undefined
      ? {
          ascii: marks.font,
          hAnsi: marks.font,
          eastAsia: marks.fontEastAsia ?? marks.font,
          cs: marks.font,
          ...(base.fonts.hint !== undefined ? { hint: base.fonts.hint } : {}),
        }
      : marks.fontEastAsia !== undefined
        ? { ...base.fonts, eastAsia: marks.fontEastAsia }
        : base.fonts
  return {
    fonts,
    sizePt: marks.sizePt ?? base.sizePt,
    bold: marks.bold ?? base.bold,
    italic: marks.italic ?? base.italic,
    underline: marks.underline ?? base.underline,
    strike: marks.strike ?? base.strike,
    color: marks.color ?? base.color,
    ...(marks.highlight !== undefined || base.highlight !== undefined
      ? { highlight: marks.highlight ?? base.highlight }
      : {}),
    caps: base.caps,
    charSpacingPt: base.charSpacingPt,
  }
}

function overlayPara(base: ParaLook, props: ParaProps, lists: ListTable): ParaLook {
  const indent = props.indent
  const spacing = props.spacing
  const numbering = props.numbering === undefined ? base.numbering : props.numbering
  // Direct fields are still optional here. A list level's left and hanging sit under them and over the style,
  // which is what keeps a bullet off the first letter when the paragraph itself stores no indent.
  const level = numbering === null ? undefined : lists.byNum.get(numbering.num)?.levels[numbering.level]
  return {
    align: props.align ?? base.align,
    indent: {
      left: indent?.left ?? level?.indentLeft ?? base.indent.left,
      right: indent?.right ?? base.indent.right,
      firstLine: indent?.firstLine ?? base.indent.firstLine,
      hanging: indent?.hanging ?? level?.hanging ?? base.indent.hanging,
      firstLineChars: indent?.firstLineChars ?? base.indent.firstLineChars,
    },
    spacing: {
      before: spacing?.before ?? base.spacing.before,
      after: spacing?.after ?? base.spacing.after,
      line: spacing?.line ?? base.spacing.line,
      contextual: base.spacing.contextual,
    },
    numbering,
    keepNext: props.keepNext ?? base.keepNext,
    keepLines: props.keepLines ?? base.keepLines,
    pageBreakBefore: props.pageBreakBefore ?? base.pageBreakBefore,
    widowControl: props.widowControl ?? base.widowControl,
  }
}

function defaultRunLook(): RunLook {
  return {
    fonts: { ascii: 'Calibri', hAnsi: 'Calibri', eastAsia: 'Calibri', cs: 'Calibri' },
    sizePt: pt(11),
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    color: 'auto',
    caps: false,
    charSpacingPt: pt(0),
  }
}

function defaultParaLook(): ParaLook {
  return {
    align: 'left',
    indent: { left: pt(0), right: pt(0), firstLine: pt(0), hanging: pt(0), firstLineChars: 0 },
    spacing: { before: pt(0), after: pt(0), line: { rule: 'auto', multiple: 1 }, contextual: false },
    numbering: null,
    keepNext: false,
    keepLines: false,
    pageBreakBefore: false,
    widowControl: true,
  }
}

/**
 * Families a print should load: the default face, the looks of text runs, and the
 * paragraph styles used in the body, cells, headers and footers, including each
 * `basedOn` parent. Styles that nothing in the document uses are left out.
 * `extra` is the import-time list plus families the font box added later.
 */
export function fontsUsedIn(doc: Doc, runLooks: ReadonlyMap<string, RunLook>, extra: readonly string[] = []): string[] {
  const names = new Set<string>()
  const add = (name: string | undefined) => {
    if (name !== undefined && name !== '') names.add(name)
  }
  const addLook = (look: RunLook | undefined) => {
    if (look === undefined) return
    add(look.fonts.ascii)
    add(look.fonts.hAnsi)
    add(look.fonts.eastAsia)
    add(look.fonts.cs)
  }
  addLook(defaultRunLook())
  const basedOn = new Map<string, string>()
  for (const [id, style] of doc.styles.paragraph) {
    if (style.basedOn) basedOn.set(id, style.basedOn)
  }
  const used = new Set<string>()
  const addStyle = (id: string | undefined) => {
    let current = id
    while (current !== undefined && current !== '' && !used.has(current)) {
      used.add(current)
      current = basedOn.get(current)
    }
  }
  addStyle(doc.styles.defaultParagraph)
  const visitPara = (para: Paragraph) => {
    addStyle(para.props.style ?? doc.styles.defaultParagraph)
    for (const inline of para.inlines) {
      if (inline.kind === 'text') {
        add(inline.marks.font)
        add(inline.marks.fontEastAsia)
        continue
      }
      if (inline.kind !== 'atom' || inline.atom.type !== 'opaque' || inline.atom.label !== 'symbol') continue
      const marker = inline.atom.display.lastIndexOf(' U+')
      if (marker > 0) add(inline.atom.display.slice(0, marker))
    }
  }
  const visit = (blocks: readonly Block[]) => {
    for (const block of blocks) {
      if (block.kind === 'p') visitPara(block)
      else if (block.kind === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) visit(cell.blocks)
        }
      } else {
        for (const para of block.preview) visitPara(para)
      }
    }
  }
  visit(doc.body)
  for (const blocks of doc.stories.values()) visit(blocks)
  for (const style of used) {
    const prefix = `${style}\0`
    for (const [key, look] of runLooks) {
      if (key.startsWith(prefix)) addLook(look)
    }
  }
  for (const name of extra) add(name)
  return [...names]
}

/** Whether any text run in the document contains a Han character. */
export function documentHasHan(doc: Doc): boolean {
  const han = /\p{Script=Han}/u
  const visit = (blocks: readonly Block[]): boolean => {
    for (const block of blocks) {
      if (block.kind === 'p') {
        for (const inline of block.inlines) {
          if (inline.kind === 'text' && han.test(inline.text)) return true
        }
      } else if (block.kind === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            if (visit(cell.blocks)) return true
          }
        }
      } else if (visit(block.preview)) {
        return true
      }
    }
    return false
  }
  if (visit(doc.body)) return true
  for (const blocks of doc.stories.values()) {
    if (visit(blocks)) return true
  }
  return false
}

/** Families and the Han flag `printFonts` needs, from one `session.import()` JSON string. */
export function printFontRequest(importedJson: string): { families: readonly string[]; han: boolean } {
  const imported = parseImported(JSON.parse(importedJson))
  const built = toDoc(imported)
  return {
    families: fontsUsedIn(built.doc, built.looks.run, imported.fonts),
    han: documentHasHan(built.doc),
  }
}

type Ids = { para: number; table: number; opaque: number }

export function toDoc(imported: wire.Imported): { doc: Doc; looks: LookTables } {
  const ids: Ids = { para: 0, table: 0, opaque: 0 }
  const paraLooks = new Map<string, ParaLook>()
  for (const entry of imported.paraLooks) paraLooks.set(entry.style, modelParaLook(entry.look))
  const runLooks = new Map<string, RunLook>()
  for (const entry of imported.runLooks) {
    runLooks.set(`${entry.para}\0${entry.character ?? ''}`, modelRunLook(entry.look))
  }
  const paragraph = new Map<StyleId, { name: string; role: StyleRole | null; next: StyleId | null; basedOn?: StyleId | null; look: ParaLook }>()
  for (const style of imported.styles.paragraph) {
    const id = styleId(style.id)
    const basedOn = style.basedOn
    paragraph.set(id, {
      name: style.name,
      role: styleRole(style.role),
      next: style.next === undefined || style.next === null ? null : styleId(style.next),
      ...(typeof basedOn === 'string' && basedOn !== '' ? { basedOn: styleId(basedOn) } : {}),
      look: paraLooks.get(style.id) ?? defaultParaLook(),
    })
  }
  const character = new Map<StyleId, { name: string }>()
  for (const [id, name] of imported.styles.character) character.set(styleId(id), { name })
  const byNum = new Map<NumId, ListDef>()
  for (const [id, def] of imported.lists) byNum.set(numId(id), modelList(def))
  const byKey = new Map<MediaKey, { contentType: string; origin: 'pristine' | 'added' }>()
  for (const [key, contentType] of imported.media) byKey.set(mediaKey(key), { contentType, origin: 'pristine' })
  const stories = new Map<StoryKey, readonly Block[]>()
  for (const [key, blocks] of imported.stories) stories.set(storyKey(key), blocks.map((block) => modelBlock(block, ids)))
  const doc: Doc = {
    body: imported.body.map((block) => modelBlock(block, ids)),
    sections: imported.sections.map(modelSection),
    stories,
    styles: {
      paragraph,
      character,
      defaultParagraph: styleId(imported.styles.defaultParagraph),
      added: [],
    },
    lists: { byNum },
    media: { byKey },
  }
  return { doc, looks: { run: runLooks, para: paraLooks } }
}

/**
 * One model block as the import wire shape.
 * `toDoc` then this is the identity on every block of a fixture import.
 */
export function toWireBlock(block: Block): wire.Block {
  if (block.kind === 'p') return toWirePara(block)
  if (block.kind === 'table') return toWireTable(block)
  return toWireOpaque(block)
}

function toPatch(doc: Doc, baseline: SaveBaseline, added: ReadonlyMap<string, HeldMedia>): wire.Patch {
  const newLists: [number, wire.ListDef][] = []
  for (const [id, def] of doc.lists.byNum) {
    if (id < 0) newLists.push([id, toWireList(def)])
  }
  newLists.sort((left, right) => left[0] - right[0])
  return {
    body: doc.body.map((block) => toPatchBlock(block, baseline.blocks)),
    sections: doc.sections === baseline.sectionList ? [...baseline.sections] : doc.sections.map(toWireSection),
    newLists,
    addedStyles: [...doc.styles.added],
    media: freshMedia(doc, added),
  }
}

function freshMedia(doc: Doc, added: ReadonlyMap<string, HeldMedia>): wire.MediaPart[] {
  const keys: string[] = []
  const seen = new Set<string>()
  const visit = (blocks: readonly Block[]) => {
    for (const block of blocks) {
      if (block.kind === 'p') {
        for (const inline of block.inlines) {
          if (inline.kind !== 'atom' || inline.src !== null || inline.atom.type !== 'image') continue
          const key = inline.atom.media
          if (seen.has(key)) continue
          seen.add(key)
          keys.push(key)
        }
      } else if (block.kind === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) visit(cell.blocks)
        }
      }
    }
  }
  visit(doc.body)
  return keys.map((key) => {
    const held = added.get(key)
    if (!held) throw new Error(`missing media ${key}`)
    return { key, fileName: mediaFileName(key, held.contentType), base64: encodeBase64(held.bytes) }
  })
}

function mediaFileName(key: string, contentType: string): string {
  const stem = key.replace(/[^A-Za-z0-9_-]/g, '')
  const name = stem.length > 0 ? stem : 'image'
  return `${name}.${contentType === 'image/jpeg' ? 'jpg' : 'png'}`
}

function encodeBase64(bytes: Uint8Array): string {
  let binary = ''
  const chunk = 0x8000
  for (let i = 0; i < bytes.length; i += chunk) {
    const piece = bytes.subarray(i, i + chunk)
    binary += String.fromCharCode(...piece)
  }
  return btoa(binary)
}

function toPatchBlock(block: Block, kept: ReadonlyMap<number, Block>): wire.PatchBlock {
  if (block.src !== null && kept.get(block.src) === block) return { keep: block.src }
  return toWireBlock(block)
}

function toWirePara(para: Paragraph): wire.Para {
  const out: wire.Para = { kind: 'p', props: toWireProps(para.props), inlines: para.inlines.map(toWireInline) }
  if (para.src !== null) out.src = para.src
  if (para.lock) out.lock = para.lock.reason
  if (para.sectionEnd !== null) out.sectionEnd = para.sectionEnd
  return out
}

function toWireInline(inline: Inline): wire.Inline {
  if (inline.kind === 'text') {
    const text: wire.Inline = { kind: 'text', text: inline.text, marks: toWireMarks(inline.marks) }
    if (inline.src !== null) text.src = inline.src
    return text
  }
  const atom: wire.Inline = { kind: 'atom', atom: toWireAtom(inline.atom) }
  if (inline.src !== null) atom.src = inline.src
  return atom
}

function toWireAtom(atom: Atom['atom']): wire.Atom {
  switch (atom.type) {
    case 'image':
      return { type: 'image', media: atom.media, widthPt: atom.widthPt, heightPt: atom.heightPt }
    case 'field':
      return { type: 'field', field: atom.field, display: atom.display }
    case 'pageBreak':
      return { type: 'pageBreak' }
    case 'columnBreak':
      return { type: 'columnBreak' }
    case 'opaque':
      return { type: 'opaque', label: atom.label, display: atom.display, removable: atom.removable }
  }
}

function toWireMarks(marks: Marks): wire.Marks {
  const out: wire.Marks = {}
  if (marks.bold !== undefined) out.bold = marks.bold
  if (marks.italic !== undefined) out.italic = marks.italic
  if (marks.underline !== undefined) out.underline = marks.underline
  if (marks.strike !== undefined) out.strike = marks.strike
  if (marks.font !== undefined) out.font = marks.font
  if (marks.fontEastAsia !== undefined) out.fontEastAsia = marks.fontEastAsia
  if (marks.sizePt !== undefined) out.sizePt = marks.sizePt
  if (marks.color !== undefined) out.color = marks.color
  if (marks.highlight !== undefined) out.highlight = marks.highlight
  if (marks.vertAlign !== undefined) out.vertAlign = marks.vertAlign
  if (marks.charStyle !== undefined) out.charStyle = marks.charStyle
  if (marks.link !== undefined) out.link = marks.link
  return out
}

function toWireProps(props: ParaProps): wire.ParaProps {
  const out: wire.ParaProps = {}
  if (props.style !== undefined) out.style = props.style
  if (props.align !== undefined) out.align = props.align
  const indent = props.indent
  if (indent) {
    if (indent.left !== undefined) out.indentLeft = indent.left
    if (indent.right !== undefined) out.indentRight = indent.right
    if (indent.firstLine !== undefined) out.firstLine = indent.firstLine
    if (indent.hanging !== undefined) out.hanging = indent.hanging
    if (indent.firstLineChars !== undefined) out.firstLineChars = indent.firstLineChars
  }
  const spacing = props.spacing
  if (spacing) {
    if (spacing.before !== undefined) out.spaceBefore = spacing.before
    if (spacing.after !== undefined) out.spaceAfter = spacing.after
    if (spacing.line !== undefined) out.line = toWireLine(spacing.line)
  }
  if (props.numbering === null) out.numbering = null
  else if (props.numbering) out.numbering = [props.numbering.num, props.numbering.level]
  if (props.keepNext !== undefined) out.keepNext = props.keepNext
  if (props.keepLines !== undefined) out.keepLines = props.keepLines
  if (props.pageBreakBefore !== undefined) out.pageBreakBefore = props.pageBreakBefore
  if (props.widowControl !== undefined) out.widowControl = props.widowControl
  return out
}

function toWireLine(line: LineSpacing): wire.LineSpacing {
  if (line.rule === 'auto') return { rule: 'auto', multiple: line.multiple }
  if (line.rule === 'exact') return { rule: 'exact', pt: line.pt }
  return { rule: 'atLeast', pt: line.pt }
}

function toWireTable(table: Table): wire.Table {
  const out: wire.Table = { kind: 'table', gridPt: [...table.grid], rows: table.rows.map(toWireRow) }
  if (table.src !== null) out.src = table.src
  if (table.props.borders) out.borders = toWireTableBorders(table.props.borders)
  if (table.props.margins) out.margins = toWireMargins(table.props.margins)
  return out
}

function toWireTableBorders(borders: TableBorders): wire.TableBorders {
  return {
    top: toWireEdge(borders.top),
    left: toWireEdge(borders.left),
    bottom: toWireEdge(borders.bottom),
    right: toWireEdge(borders.right),
    insideH: toWireEdge(borders.insideH),
    insideV: toWireEdge(borders.insideV),
  }
}

function toWireEdge(edge: BorderEdge): wire.BorderEdge {
  if (edge === null) return null
  return { style: edge.style, widthPt: edge.widthPt, color: edge.color }
}

function toWireRow(row: Row): wire.Row {
  const out: wire.Row = {
    header: row.props.header === true,
    cantSplit: row.props.cantSplit === true,
    cells: row.cells.map(toWireCell),
  }
  if (row.src !== null) out.src = row.src
  if (row.props.heightPt !== undefined) out.heightPt = row.props.heightPt
  if (row.props.heightRule !== undefined) out.heightRule = row.props.heightRule
  return out
}

function toWireCell(cell: Cell): wire.Cell {
  const out: wire.Cell = { span: cell.span, blocks: cell.blocks.map(toWireBlock) }
  if (cell.src !== null) out.src = cell.src
  if (cell.vMerge) out.vMerge = cell.vMerge
  if (cell.borders) out.borders = toWireCellBorders(cell.borders)
  if (cell.margins) out.margins = toWireMargins(cell.margins)
  return out
}

function toWireMargins(margins: CellMargins): wire.CellMargins {
  const out: wire.CellMargins = {}
  if (margins.top !== undefined) out.top = margins.top
  if (margins.right !== undefined) out.right = margins.right
  if (margins.bottom !== undefined) out.bottom = margins.bottom
  if (margins.left !== undefined) out.left = margins.left
  return out
}

function toWireCellBorders(borders: CellBorders): wire.CellBorders {
  const out: wire.CellBorders = {}
  if (borders.top !== undefined) out.top = toWireEdge(borders.top)
  if (borders.left !== undefined) out.left = toWireEdge(borders.left)
  if (borders.bottom !== undefined) out.bottom = toWireEdge(borders.bottom)
  if (borders.right !== undefined) out.right = toWireEdge(borders.right)
  return out
}

function toWireOpaque(block: OpaqueBlock): wire.Opaque {
  return {
    kind: 'opaque',
    src: block.src,
    label: block.label,
    preview: block.preview.map(toWirePara),
    removable: block.removable,
  }
}

function toWireSection(section: Section): wire.Section {
  const margin = section.page.margin
  const out: wire.Section = {
    widthPt: section.page.widthPt,
    heightPt: section.page.heightPt,
    marginsPt: [margin.top, margin.right, margin.bottom, margin.left, margin.header, margin.footer, margin.gutter],
    breakType: section.breakType,
    titlePage: section.titlePage,
    headers: storyPairs(section.headers),
    footers: storyPairs(section.footers),
  }
  if (section.src !== null) out.src = section.src
  if (section.grid) out.grid = [section.grid.type, section.grid.linePitch, section.grid.charSpace ?? null]
  return out
}

function storyPairs(refs: Section['headers']): [string, string][] {
  const out: [string, string][] = []
  if (refs.default) out.push(['default', refs.default])
  if (refs.first) out.push(['first', refs.first])
  if (refs.even) out.push(['even', refs.even])
  return out
}

function toWireList(def: ListDef): wire.ListDef {
  return {
    fresh: def.fresh,
    levels: def.levels.map((level) => ({
      format: level.format,
      text: level.text,
      start: level.start,
      indentLeftPt: level.indentLeft,
      hangingPt: level.hanging,
    })),
  }
}

function modelBlock(block: wire.Block, ids: Ids): Block {
  if (block.kind === 'p') return modelPara(block, ids)
  if (block.kind === 'table') return modelTable(block, ids)
  return modelOpaque(block, ids)
}

function modelPara(para: wire.Para, ids: Ids): Paragraph {
  ids.para += 1
  return {
    kind: 'p',
    id: paraId(`p${ids.para}`),
    src: typeof para.src === 'number' ? (para.src as ParaSrc) : null,
    props: modelProps(para.props),
    inlines: para.inlines.map(modelInline),
    lock: para.lock === undefined || para.lock === null ? null : modelLock(para.lock),
    sectionEnd: para.sectionEnd === undefined || para.sectionEnd === null ? null : para.sectionEnd,
  }
}

function modelInline(inline: wire.Inline): Inline {
  const src = inline.src === undefined || inline.src === null ? null : runSrc(inline.src)
  if (inline.kind === 'text') {
    return { kind: 'text', text: inline.text, marks: modelMarks(inline.marks), src }
  }
  return { kind: 'atom', atom: modelAtom(inline.atom), src }
}

function modelAtom(atom: wire.Atom): Atom['atom'] {
  if (atom.type === 'image') {
    return { type: 'image', media: mediaKey(atom.media), widthPt: pt(atom.widthPt), heightPt: pt(atom.heightPt) }
  }
  if (atom.type === 'field') return { type: 'field', field: fieldKind(atom.field), display: atom.display }
  if (atom.type === 'pageBreak') return { type: 'pageBreak' }
  if (atom.type === 'columnBreak') return { type: 'columnBreak' }
  return { type: 'opaque', label: atom.label, display: atom.display, removable: atom.removable }
}

function modelMarks(marks: wire.Marks): Marks {
  const out: {
    bold?: boolean
    italic?: boolean
    underline?: boolean
    strike?: boolean
    font?: string
    fontEastAsia?: string
    sizePt?: Pt
    color?: Hex
    highlight?: HighlightName
    vertAlign?: 'super' | 'sub'
    charStyle?: StyleId
    link?: Marks['link']
  } = {}
  if (typeof marks.bold === 'boolean') out.bold = marks.bold
  if (typeof marks.italic === 'boolean') out.italic = marks.italic
  if (typeof marks.underline === 'boolean') out.underline = marks.underline
  if (typeof marks.strike === 'boolean') out.strike = marks.strike
  if (typeof marks.font === 'string') out.font = marks.font
  if (typeof marks.fontEastAsia === 'string') out.fontEastAsia = marks.fontEastAsia
  if (typeof marks.sizePt === 'number') out.sizePt = pt(marks.sizePt)
  if (typeof marks.color === 'string') out.color = hex(marks.color)
  if (typeof marks.highlight === 'string') out.highlight = highlightName(marks.highlight)
  if (marks.vertAlign === 'super' || marks.vertAlign === 'sub') out.vertAlign = marks.vertAlign
  if (typeof marks.charStyle === 'string') out.charStyle = styleId(marks.charStyle)
  if (typeof marks.link === 'string') out.link = marks.link as Marks['link']
  return out
}

function modelProps(props: wire.ParaProps): ParaProps {
  const indent = modelIndent(props)
  const spacing = modelSpacing(props)
  const out: {
    style?: StyleId
    align?: ParaProps['align']
    indent?: ParaProps['indent']
    spacing?: ParaProps['spacing']
    numbering?: ParaProps['numbering']
    keepNext?: boolean
    keepLines?: boolean
    pageBreakBefore?: boolean
    widowControl?: boolean
  } = {}
  if (typeof props.style === 'string') out.style = styleId(props.style)
  if (typeof props.align === 'string') out.align = alignName(props.align)
  if (indent !== undefined) out.indent = indent
  if (spacing !== undefined) out.spacing = spacing
  if (props.numbering === null) out.numbering = null
  else if (Array.isArray(props.numbering)) {
    const num = props.numbering[0]
    const level = props.numbering[1]
    if (num === undefined || level === undefined) fail('props.numbering', 'expected a pair')
    out.numbering = { num: numId(num), level }
  }
  if (typeof props.keepNext === 'boolean') out.keepNext = props.keepNext
  if (typeof props.keepLines === 'boolean') out.keepLines = props.keepLines
  if (typeof props.pageBreakBefore === 'boolean') out.pageBreakBefore = props.pageBreakBefore
  if (typeof props.widowControl === 'boolean') out.widowControl = props.widowControl
  return out
}

function modelIndent(props: wire.ParaProps): ParaProps['indent'] | undefined {
  const left = optPt(props.indentLeft)
  const right = optPt(props.indentRight)
  const firstLine = optPt(props.firstLine)
  const hanging = optPt(props.hanging)
  const firstLineChars = typeof props.firstLineChars === 'number' ? props.firstLineChars : undefined
  if (
    left === undefined &&
    right === undefined &&
    firstLine === undefined &&
    hanging === undefined &&
    firstLineChars === undefined
  ) {
    return undefined
  }
  return { left, right, firstLine, hanging, firstLineChars }
}

function modelSpacing(props: wire.ParaProps): ParaProps['spacing'] | undefined {
  const before = optPt(props.spaceBefore)
  const after = optPt(props.spaceAfter)
  const line = props.line === undefined || props.line === null ? undefined : modelLine(props.line)
  if (before === undefined && after === undefined && line === undefined) return undefined
  return { before, after, line }
}

function modelTable(table: wire.Table, ids: Ids): Table {
  ids.table += 1
  const borders = table.borders ? modelTableBorders(table.borders) : undefined
  const margins = table.margins ? modelMargins(table.margins) : undefined
  return {
    kind: 'table',
    id: tableId(`t${ids.table}`),
    src: typeof table.src === 'number' ? (table.src as TableSrc) : null,
    props: {
      ...(borders === undefined ? {} : { borders }),
      ...(margins === undefined ? {} : { margins }),
    },
    grid: table.gridPt.map(pt),
    rows: table.rows.map((row) => modelRow(row, ids)),
  }
}

function modelTableBorders(borders: wire.TableBorders): TableBorders {
  return {
    top: modelEdge(borders.top),
    left: modelEdge(borders.left),
    bottom: modelEdge(borders.bottom),
    right: modelEdge(borders.right),
    insideH: modelEdge(borders.insideH),
    insideV: modelEdge(borders.insideV),
  }
}

function modelEdge(edge: wire.BorderEdge): BorderEdge {
  if (edge === null) return null
  return { style: edge.style, widthPt: pt(edge.widthPt), color: hex(edge.color) }
}

function modelRow(row: wire.Row, ids: Ids): Row {
  const heightPt = typeof row.heightPt === 'number' ? pt(row.heightPt) : undefined
  const heightRule = row.heightRule === 'exact' || row.heightRule === 'atLeast' ? row.heightRule : undefined
  return {
    src: typeof row.src === 'number' ? (row.src as RowSrc) : null,
    props: {
      header: row.header,
      cantSplit: row.cantSplit,
      ...(heightPt === undefined ? {} : { heightPt }),
      ...(heightRule === undefined ? {} : { heightRule }),
    },
    cells: row.cells.map((cell) => modelCell(cell, ids)),
  }
}

function modelCell(cell: wire.Cell, ids: Ids): Cell {
  const borders = cell.borders ? modelCellBorders(cell.borders) : undefined
  const margins = cell.margins ? modelMargins(cell.margins) : undefined
  return {
    src: typeof cell.src === 'number' ? (cell.src as CellSrc) : null,
    span: cell.span,
    vMerge: cell.vMerge === 'restart' || cell.vMerge === 'continue' ? cell.vMerge : null,
    ...(borders === undefined ? {} : { borders }),
    ...(margins === undefined ? {} : { margins }),
    blocks: cell.blocks.map((block) => modelBlock(block, ids)),
  }
}

function modelMargins(margins: wire.CellMargins): CellMargins | undefined {
  const out: { top?: Pt; right?: Pt; bottom?: Pt; left?: Pt } = {}
  if (typeof margins.top === 'number') out.top = pt(margins.top)
  if (typeof margins.right === 'number') out.right = pt(margins.right)
  if (typeof margins.bottom === 'number') out.bottom = pt(margins.bottom)
  if (typeof margins.left === 'number') out.left = pt(margins.left)
  if (out.top === undefined && out.right === undefined && out.bottom === undefined && out.left === undefined) return undefined
  return out
}

function modelCellBorders(borders: wire.CellBorders): CellBorders | undefined {
  const out: { top?: BorderEdge; left?: BorderEdge; bottom?: BorderEdge; right?: BorderEdge } = {}
  if (borders.top !== undefined) out.top = modelEdge(borders.top)
  if (borders.left !== undefined) out.left = modelEdge(borders.left)
  if (borders.bottom !== undefined) out.bottom = modelEdge(borders.bottom)
  if (borders.right !== undefined) out.right = modelEdge(borders.right)
  if (out.top === undefined && out.left === undefined && out.bottom === undefined && out.right === undefined) {
    return undefined
  }
  return out
}

function modelOpaque(block: wire.Opaque, ids: Ids): OpaqueBlock {
  ids.opaque += 1
  return {
    kind: 'opaque',
    id: opaqueId(`o${ids.opaque}`),
    src: block.src as OpaqueBlock['src'],
    label: block.label,
    preview: block.preview.map((para) => modelPara(para, ids)),
    removable: block.removable,
  }
}

function modelSection(section: wire.Section): Section {
  const margin = section.marginsPt
  const top = margin[0]
  const right = margin[1]
  const bottom = margin[2]
  const left = margin[3]
  const header = margin[4]
  const footer = margin[5]
  const gutter = margin[6]
  if (
    top === undefined ||
    right === undefined ||
    bottom === undefined ||
    left === undefined ||
    header === undefined ||
    footer === undefined ||
    gutter === undefined
  ) {
    fail('section.marginsPt', 'expected 7 margins')
  }
  return {
    src: section.src === undefined || section.src === null ? null : (section.src as SectionSrc),
    page: {
      widthPt: pt(section.widthPt),
      heightPt: pt(section.heightPt),
      margin: {
        top: pt(top),
        right: pt(right),
        bottom: pt(bottom),
        left: pt(left),
        header: pt(header),
        footer: pt(footer),
        gutter: pt(gutter),
      },
    },
    breakType: breakType(section.breakType),
    titlePage: section.titlePage,
    headers: storyRefs(section.headers),
    footers: storyRefs(section.footers),
    grid: modelGrid(section.grid),
  }
}

function storyRefs(pairs: [string, string][]): Section['headers'] {
  const out: { default?: StoryKey; first?: StoryKey; even?: StoryKey } = {}
  for (const [kind, key] of pairs) {
    if (kind === 'default') out.default = storyKey(key)
    else if (kind === 'first') out.first = storyKey(key)
    else if (kind === 'even') out.even = storyKey(key)
    else fail('section.story', `unknown story kind ${kind}`)
  }
  return out
}

function modelGrid(grid: wire.Section['grid']): Section['grid'] {
  if (grid === undefined || grid === null) return null
  const type = grid[0]
  const linePitch = grid[1]
  const charSpace = grid[2]
  if (type === undefined || linePitch === undefined) fail('section.grid', 'expected a grid triple')
  if (type !== 'lines' && type !== 'linesAndChars' && type !== 'snapToChars') {
    fail('section.grid', `unknown grid ${type}`)
  }
  return {
    type,
    linePitch: pt(linePitch),
    ...(typeof charSpace === 'number' ? { charSpace } : {}),
  }
}

function modelRunLook(look: wire.RunLook): RunLook {
  const hint = look.fonts.hint === 'eastAsia' || look.fonts.hint === 'cs' ? look.fonts.hint : undefined
  return {
    fonts: {
      ascii: look.fonts.ascii,
      hAnsi: look.fonts.hAnsi,
      eastAsia: look.fonts.eastAsia,
      cs: look.fonts.cs,
      ...(hint !== undefined ? { hint } : {}),
    },
    sizePt: pt(look.sizePt),
    bold: look.bold,
    italic: look.italic,
    underline: look.underline,
    strike: look.strike,
    color: look.color.toLowerCase() === 'auto' ? 'auto' : hex(look.color),
    ...(typeof look.highlight === 'string' ? { highlight: highlightName(look.highlight) } : {}),
    caps: look.caps,
    charSpacingPt: pt(look.charSpacingPt),
  }
}

function modelParaLook(look: wire.ParaLook): ParaLook {
  const numbering = look.numbering
  return {
    align: alignName(look.align),
    indent: {
      left: pt(look.indent.left),
      right: pt(look.indent.right),
      firstLine: pt(look.indent.firstLine),
      hanging: pt(look.indent.hanging),
      firstLineChars: typeof look.indent.firstLineChars === 'number' ? look.indent.firstLineChars : 0,
    },
    spacing: {
      before: pt(look.spacing.before),
      after: pt(look.spacing.after),
      line: modelLine(look.spacing.line),
      contextual: look.spacing.contextual,
    },
    numbering:
      numbering === null
        ? null
        : { num: numId(numbering[0] ?? fail('paraLook.numbering', 'expected a pair')), level: numbering[1] ?? 0 },
    keepNext: look.keepNext,
    keepLines: look.keepLines,
    pageBreakBefore: look.pageBreakBefore,
    widowControl: look.widowControl,
  }
}

function modelLine(line: wire.LineSpacing): LineSpacing {
  if (line.rule === 'auto') return { rule: 'auto', multiple: line.multiple }
  if (line.rule === 'exact') return { rule: 'exact', pt: pt(line.pt) }
  return { rule: 'atLeast', pt: pt(line.pt) }
}

function modelList(def: wire.ListDef): ListDef {
  return { fresh: def.fresh, levels: def.levels.map(modelLevel) }
}

function modelLevel(level: wire.ListLevel): ListLevel {
  return {
    format: listFormat(level.format),
    text: level.text,
    start: level.start,
    indentLeft: pt(level.indentLeftPt),
    hanging: pt(level.hangingPt),
  }
}

function modelLock(reason: string): Lock {
  if (reason === 'revisions' || reason === 'unbalancedField' || reason === 'preview') return { reason }
  fail('lock', `unknown lock ${reason}`)
}

export function parseImported(value: unknown): wire.Imported {
  const record = rec(value, 'imported')
  return {
    body: arr(record.body, 'body').map((item, index) => parseBlock(item, `body[${index}]`)),
    sections: arr(record.sections, 'sections').map((item, index) => parseSection(item, `sections[${index}]`)),
    stories: arr(record.stories, 'stories').map((item, index) => parseStory(item, `stories[${index}]`)),
    styles: parseStyles(record.styles),
    lists: arr(record.lists, 'lists').map((item, index) => parseList(item, `lists[${index}]`)),
    media: arr(record.media, 'media').map((item, index) => parsePair(item, `media[${index}]`, text, text)),
    fonts: arr(record.fonts, 'fonts').map((item, index) => text(item, `fonts[${index}]`)),
    runLooks: arr(record.runLooks, 'runLooks').map((item, index) => parseRunLook(item, `runLooks[${index}]`)),
    paraLooks: arr(record.paraLooks, 'paraLooks').map((item, index) => parseParaLook(item, `paraLooks[${index}]`)),
  }
}

function parseBlock(value: unknown, path: string): wire.Block {
  const record = rec(value, path)
  if (record.kind === 'p') return parsePara(record, path)
  if (record.kind === 'table') return parseTable(record, path)
  if (record.kind === 'opaque') return parseOpaque(record, path)
  fail(path, 'unknown block kind')
}

function parsePara(record: Record<string, unknown>, path: string): wire.Para {
  return {
    kind: 'p',
    src: optNum(record.src, `${path}.src`),
    props: parseProps(record.props, `${path}.props`),
    inlines: arr(record.inlines, `${path}.inlines`).map((item, index) => parseInline(item, `${path}.inlines[${index}]`)),
    lock: optText(record.lock, `${path}.lock`),
    sectionEnd: optNum(record.sectionEnd, `${path}.sectionEnd`),
  }
}

function parseInline(value: unknown, path: string): wire.Inline {
  const record = rec(value, path)
  const src = optNum(record.src, `${path}.src`)
  if (record.kind === 'text') {
    return { kind: 'text', text: text(record.text, `${path}.text`), marks: parseMarks(record.marks, `${path}.marks`), src }
  }
  if (record.kind === 'atom') return { kind: 'atom', atom: parseAtom(record.atom, `${path}.atom`), src }
  fail(path, 'unknown inline kind')
}

function parseAtom(value: unknown, path: string): wire.Atom {
  const record = rec(value, path)
  if (record.type === 'image') {
    return {
      type: 'image',
      media: text(record.media, `${path}.media`),
      widthPt: finite(record.widthPt, `${path}.widthPt`),
      heightPt: finite(record.heightPt, `${path}.heightPt`),
    }
  }
  if (record.type === 'field') {
    return { type: 'field', field: text(record.field, `${path}.field`), display: text(record.display, `${path}.display`) }
  }
  if (record.type === 'pageBreak') return { type: 'pageBreak' }
  if (record.type === 'columnBreak') return { type: 'columnBreak' }
  if (record.type === 'opaque') {
    return {
      type: 'opaque',
      label: text(record.label, `${path}.label`),
      display: text(record.display, `${path}.display`),
      removable: bool(record.removable, `${path}.removable`),
    }
  }
  fail(path, 'unknown atom type')
}

function parseMarks(value: unknown, path: string): wire.Marks {
  if (value === undefined || value === null) return {}
  const record = rec(value, path)
  return {
    bold: optBool(record.bold, `${path}.bold`),
    italic: optBool(record.italic, `${path}.italic`),
    underline: optBool(record.underline, `${path}.underline`),
    strike: optBool(record.strike, `${path}.strike`),
    font: optText(record.font, `${path}.font`),
    fontEastAsia: optText(record.fontEastAsia, `${path}.fontEastAsia`),
    sizePt: optNum(record.sizePt, `${path}.sizePt`),
    color: optText(record.color, `${path}.color`),
    highlight: optText(record.highlight, `${path}.highlight`),
    vertAlign: optText(record.vertAlign, `${path}.vertAlign`),
    charStyle: optText(record.charStyle, `${path}.charStyle`),
    link: optText(record.link, `${path}.link`),
  }
}

function parseProps(value: unknown, path: string): wire.ParaProps {
  if (value === undefined || value === null) return {}
  const record = rec(value, path)
  return {
    style: optText(record.style, `${path}.style`),
    align: optText(record.align, `${path}.align`),
    indentLeft: optNum(record.indentLeft, `${path}.indentLeft`),
    indentRight: optNum(record.indentRight, `${path}.indentRight`),
    firstLine: optNum(record.firstLine, `${path}.firstLine`),
    hanging: optNum(record.hanging, `${path}.hanging`),
    firstLineChars: optNum(record.firstLineChars, `${path}.firstLineChars`),
    spaceBefore: optNum(record.spaceBefore, `${path}.spaceBefore`),
    spaceAfter: optNum(record.spaceAfter, `${path}.spaceAfter`),
    line: record.line === undefined || record.line === null ? undefined : parseLine(record.line, `${path}.line`),
    numbering: parseNumbering(record.numbering, `${path}.numbering`),
    keepNext: optBool(record.keepNext, `${path}.keepNext`),
    keepLines: optBool(record.keepLines, `${path}.keepLines`),
    pageBreakBefore: optBool(record.pageBreakBefore, `${path}.pageBreakBefore`),
    widowControl: optBool(record.widowControl, `${path}.widowControl`),
  }
}

function parseNumbering(value: unknown, path: string): [number, number] | null | undefined {
  if (value === undefined) return undefined
  if (value === null) return null
  if (!Array.isArray(value) || value.length !== 2) fail(path, 'expected a pair or null')
  return [finite(value[0], `${path}[0]`), finite(value[1], `${path}[1]`)]
}

function parseLine(value: unknown, path: string): wire.LineSpacing {
  const record = rec(value, path)
  if (record.rule === 'auto') return { rule: 'auto', multiple: finite(record.multiple, `${path}.multiple`) }
  if (record.rule === 'exact') return { rule: 'exact', pt: finite(record.pt, `${path}.pt`) }
  if (record.rule === 'atLeast') return { rule: 'atLeast', pt: finite(record.pt, `${path}.pt`) }
  fail(path, 'unknown line rule')
}

function parseTable(record: Record<string, unknown>, path: string): wire.Table {
  const out: wire.Table = {
    kind: 'table',
    src: optNum(record.src, `${path}.src`),
    gridPt: arr(record.gridPt, `${path}.gridPt`).map((item, index) => finite(item, `${path}.gridPt[${index}]`)),
    rows: arr(record.rows, `${path}.rows`).map((item, index) => parseRow(item, `${path}.rows[${index}]`)),
  }
  if (record.borders !== undefined && record.borders !== null) {
    out.borders = parseTableBorders(record.borders, `${path}.borders`)
  }
  const margins = parseMargins(record.margins, `${path}.margins`)
  if (margins !== undefined) out.margins = margins
  return out
}

function parseTableBorders(value: unknown, path: string): wire.TableBorders {
  const record = rec(value, path)
  return {
    top: parseBorderEdge(record.top, `${path}.top`),
    left: parseBorderEdge(record.left, `${path}.left`),
    bottom: parseBorderEdge(record.bottom, `${path}.bottom`),
    right: parseBorderEdge(record.right, `${path}.right`),
    insideH: parseBorderEdge(record.insideH, `${path}.insideH`),
    insideV: parseBorderEdge(record.insideV, `${path}.insideV`),
  }
}

function parseBorderEdge(value: unknown, path: string): wire.BorderEdge {
  if (value === undefined || value === null) return null
  const record = rec(value, path)
  const color = text(record.color, `${path}.color`)
  if (!/^[0-9A-Fa-f]{6}$/.test(color)) fail(path, 'expected six hex digits')
  return {
    style: parseBorderStyle(record.style, `${path}.style`),
    widthPt: finite(record.widthPt, `${path}.widthPt`),
    color: color.toUpperCase(),
  }
}

function parseBorderStyle(value: unknown, path: string): wire.BorderStyleName {
  const style = text(value, path)
  if (style === 'single' || style === 'double' || style === 'dashed' || style === 'dotted') return style
  fail(path, 'unknown border style')
}

function parseRow(value: unknown, path: string): wire.Row {
  const record = rec(value, path)
  const out: wire.Row = {
    src: optNum(record.src, `${path}.src`),
    header: bool(record.header, `${path}.header`),
    cantSplit: bool(record.cantSplit, `${path}.cantSplit`),
    cells: arr(record.cells, `${path}.cells`).map((item, index) => parseCell(item, `${path}.cells[${index}]`)),
  }
  if (record.heightPt !== undefined && record.heightPt !== null) out.heightPt = finite(record.heightPt, `${path}.heightPt`)
  const rule = parseHeightRule(record.heightRule, `${path}.heightRule`)
  if (rule !== undefined) out.heightRule = rule
  return out
}

function parseHeightRule(value: unknown, path: string): 'exact' | 'atLeast' | undefined {
  if (value === undefined || value === null) return undefined
  const rule = text(value, path)
  if (rule === 'exact' || rule === 'atLeast') return rule
  fail(path, 'unknown row height rule')
}

function parseMargins(value: unknown, path: string): wire.CellMargins | undefined {
  if (value === undefined || value === null) return undefined
  const record = rec(value, path)
  const out: wire.CellMargins = {}
  const top = optNum(record.top, `${path}.top`)
  const right = optNum(record.right, `${path}.right`)
  const bottom = optNum(record.bottom, `${path}.bottom`)
  const left = optNum(record.left, `${path}.left`)
  if (typeof top === 'number') out.top = top
  if (typeof right === 'number') out.right = right
  if (typeof bottom === 'number') out.bottom = bottom
  if (typeof left === 'number') out.left = left
  if (out.top === undefined && out.right === undefined && out.bottom === undefined && out.left === undefined) return undefined
  return out
}

function parseCell(value: unknown, path: string): wire.Cell {
  const record = rec(value, path)
  const out: wire.Cell = {
    src: optNum(record.src, `${path}.src`),
    span: finite(record.span, `${path}.span`),
    vMerge: optText(record.vMerge, `${path}.vMerge`),
    blocks: arr(record.blocks, `${path}.blocks`).map((item, index) => parseBlock(item, `${path}.blocks[${index}]`)),
  }
  if (record.borders !== undefined && record.borders !== null) {
    out.borders = parseCellBorders(record.borders, `${path}.borders`)
  }
  const margins = parseMargins(record.margins, `${path}.margins`)
  if (margins !== undefined) out.margins = margins
  return out
}

function parseCellBorders(value: unknown, path: string): wire.CellBorders {
  const record = rec(value, path)
  const out: wire.CellBorders = {}
  if ('top' in record) out.top = parseBorderEdge(record.top, `${path}.top`)
  if ('left' in record) out.left = parseBorderEdge(record.left, `${path}.left`)
  if ('bottom' in record) out.bottom = parseBorderEdge(record.bottom, `${path}.bottom`)
  if ('right' in record) out.right = parseBorderEdge(record.right, `${path}.right`)
  return out
}

function parseOpaque(record: Record<string, unknown>, path: string): wire.Opaque {
  return {
    kind: 'opaque',
    src: finite(record.src, `${path}.src`),
    label: text(record.label, `${path}.label`),
    preview: arr(record.preview, `${path}.preview`).map((item, index) => parsePara(rec(item, `${path}.preview[${index}]`), `${path}.preview[${index}]`)),
    removable: bool(record.removable, `${path}.removable`),
  }
}

function parseSection(value: unknown, path: string): wire.Section {
  const record = rec(value, path)
  const grid = record.grid
  return {
    src: optNum(record.src, `${path}.src`),
    widthPt: finite(record.widthPt, `${path}.widthPt`),
    heightPt: finite(record.heightPt, `${path}.heightPt`),
    marginsPt: arr(record.marginsPt, `${path}.marginsPt`).map((item, index) => finite(item, `${path}.marginsPt[${index}]`)),
    breakType: text(record.breakType, `${path}.breakType`),
    titlePage: bool(record.titlePage, `${path}.titlePage`),
    headers: arr(record.headers, `${path}.headers`).map((item, index) => parsePair(item, `${path}.headers[${index}]`, text, text)),
    footers: arr(record.footers, `${path}.footers`).map((item, index) => parsePair(item, `${path}.footers[${index}]`, text, text)),
    grid: grid === undefined || grid === null ? grid : parseGrid(grid, `${path}.grid`),
  }
}

function parseGrid(value: unknown, path: string): [string, number, number | null] {
  if (!Array.isArray(value) || value.length !== 3) fail(path, 'expected a triple')
  const space = value[2]
  return [
    text(value[0], `${path}[0]`),
    finite(value[1], `${path}[1]`),
    space === null ? null : finite(space, `${path}[2]`),
  ]
}

function parseStory(value: unknown, path: string): [string, wire.Block[]] {
  if (!Array.isArray(value) || value.length !== 2) fail(path, 'expected a pair')
  const blocks = value[1]
  return [text(value[0], `${path}[0]`), arr(blocks, `${path}[1]`).map((item, index) => parseBlock(item, `${path}[1][${index}]`))]
}

function parseStyles(value: unknown): wire.StyleSheet {
  const record = rec(value, 'styles')
  return {
    paragraph: arr(record.paragraph, 'styles.paragraph').map((item, index) => parseParaStyle(item, `styles.paragraph[${index}]`)),
    character: arr(record.character, 'styles.character').map((item, index) =>
      parsePair(item, `styles.character[${index}]`, text, text),
    ),
    defaultParagraph: text(record.defaultParagraph, 'styles.defaultParagraph'),
  }
}

function parseParaStyle(value: unknown, path: string): wire.ParaStyle {
  const record = rec(value, path)
  const basedOn = optText(record.basedOn, `${path}.basedOn`)
  return {
    id: text(record.id, `${path}.id`),
    name: text(record.name, `${path}.name`),
    role: optText(record.role, `${path}.role`),
    next: optText(record.next, `${path}.next`),
    ...(typeof basedOn === 'string' ? { basedOn } : {}),
  }
}

function parseList(value: unknown, path: string): [number, wire.ListDef] {
  if (!Array.isArray(value) || value.length !== 2) fail(path, 'expected a pair')
  const def = rec(value[1], `${path}[1]`)
  return [
    finite(value[0], `${path}[0]`),
    {
      fresh: bool(def.fresh, `${path}.fresh`),
      levels: arr(def.levels, `${path}.levels`).map((item, index) => parseLevel(item, `${path}.levels[${index}]`)),
    },
  ]
}

function parseLevel(value: unknown, path: string): wire.ListLevel {
  const record = rec(value, path)
  return {
    format: text(record.format, `${path}.format`),
    text: text(record.text, `${path}.text`),
    start: finite(record.start, `${path}.start`),
    indentLeftPt: finite(record.indentLeftPt, `${path}.indentLeftPt`),
    hangingPt: finite(record.hangingPt, `${path}.hangingPt`),
  }
}

function parseRunLook(value: unknown, path: string): wire.RunLookEntry {
  const record = rec(value, path)
  const look = rec(record.look, `${path}.look`)
  const fonts = rec(look.fonts, `${path}.look.fonts`)
  return {
    para: text(record.para, `${path}.para`),
    character: optText(record.character, `${path}.character`),
    look: {
      fonts: {
        ascii: text(fonts.ascii, `${path}.ascii`),
        hAnsi: text(fonts.hAnsi, `${path}.hAnsi`),
        eastAsia: text(fonts.eastAsia, `${path}.eastAsia`),
        cs: text(fonts.cs, `${path}.cs`),
        hint: optText(fonts.hint, `${path}.hint`),
      },
      sizePt: finite(look.sizePt, `${path}.sizePt`),
      bold: bool(look.bold, `${path}.bold`),
      italic: bool(look.italic, `${path}.italic`),
      underline: bool(look.underline, `${path}.underline`),
      strike: bool(look.strike, `${path}.strike`),
      color: text(look.color, `${path}.color`),
      highlight: optText(look.highlight, `${path}.highlight`),
      caps: bool(look.caps, `${path}.caps`),
      charSpacingPt: finite(look.charSpacingPt, `${path}.charSpacingPt`),
    },
  }
}

function parseParaLook(value: unknown, path: string): wire.ParaLookEntry {
  const record = rec(value, path)
  const look = rec(record.look, `${path}.look`)
  const indent = rec(look.indent, `${path}.indent`)
  const spacing = rec(look.spacing, `${path}.spacing`)
  const numbering = look.numbering
  return {
    style: text(record.style, `${path}.style`),
    look: {
      align: text(look.align, `${path}.align`),
      indent: {
        left: finite(indent.left, `${path}.left`),
        right: finite(indent.right, `${path}.right`),
        firstLine: finite(indent.firstLine, `${path}.firstLine`),
        hanging: finite(indent.hanging, `${path}.hanging`),
        firstLineChars: optNum(indent.firstLineChars, `${path}.firstLineChars`),
      },
      spacing: {
        before: finite(spacing.before, `${path}.before`),
        after: finite(spacing.after, `${path}.after`),
        line: parseLine(spacing.line, `${path}.line`),
        contextual: bool(spacing.contextual, `${path}.contextual`),
      },
      numbering:
        numbering === null
          ? null
          : parseNumbering(numbering, `${path}.numbering`) ?? fail(path, 'numbering is required on a look'),
      keepNext: bool(look.keepNext, `${path}.keepNext`),
      keepLines: bool(look.keepLines, `${path}.keepLines`),
      pageBreakBefore: bool(look.pageBreakBefore, `${path}.pageBreakBefore`),
      widowControl: bool(look.widowControl, `${path}.widowControl`),
    },
  }
}

function parsePair<A, B>(
  value: unknown,
  path: string,
  left: (item: unknown, itemPath: string) => A,
  right: (item: unknown, itemPath: string) => B,
): [A, B] {
  if (!Array.isArray(value) || value.length !== 2) fail(path, 'expected a pair')
  return [left(value[0], `${path}[0]`), right(value[1], `${path}[1]`)]
}

function rec(value: unknown, path: string): Record<string, unknown> {
  if (typeof value === 'object' && value !== null && !Array.isArray(value)) return value as Record<string, unknown>
  fail(path, 'expected an object')
}

function arr(value: unknown, path: string): unknown[] {
  if (Array.isArray(value)) return value
  fail(path, 'expected an array')
}

function text(value: unknown, path: string): string {
  if (typeof value === 'string') return value
  fail(path, 'expected a string')
}

function finite(value: unknown, path: string): number {
  if (typeof value === 'number' && Number.isFinite(value)) return value
  fail(path, 'expected a number')
}

function bool(value: unknown, path: string): boolean {
  if (typeof value === 'boolean') return value
  fail(path, 'expected a boolean')
}

function optText(value: unknown, path: string): string | null | undefined {
  if (value === undefined || value === null) return value
  return text(value, path)
}

function optNum(value: unknown, path: string): number | null | undefined {
  if (value === undefined || value === null) return value
  return finite(value, path)
}

function optBool(value: unknown, path: string): boolean | null | undefined {
  if (value === undefined || value === null) return value
  return bool(value, path)
}

function optPt(value: number | null | undefined): Pt | undefined {
  return typeof value === 'number' ? pt(value) : undefined
}

function alignName(value: string): NonNullable<ParaProps['align']> {
  if (value === 'left' || value === 'center' || value === 'right' || value === 'justify' || value === 'distribute') {
    return value
  }
  fail('align', `unknown alignment ${value}`)
}

function breakType(value: string): Section['breakType'] {
  if (value === 'nextPage' || value === 'continuous' || value === 'evenPage' || value === 'oddPage') return value
  fail('breakType', `unknown break ${value}`)
}

function fieldKind(value: string): 'page' | 'numPages' | 'other' {
  if (value === 'page' || value === 'numPages' || value === 'other') return value
  fail('field', `unknown field ${value}`)
}

function highlightName(value: string): HighlightName {
  if (
    value === 'yellow' ||
    value === 'green' ||
    value === 'cyan' ||
    value === 'magenta' ||
    value === 'blue' ||
    value === 'red' ||
    value === 'darkBlue' ||
    value === 'darkCyan' ||
    value === 'darkGreen' ||
    value === 'darkMagenta' ||
    value === 'darkRed' ||
    value === 'darkYellow' ||
    value === 'darkGray' ||
    value === 'lightGray' ||
    value === 'black'
  ) {
    return value
  }
  fail('highlight', `unknown highlight ${value}`)
}

function listFormat(value: string): ListLevel['format'] {
  if (
    value === 'bullet' ||
    value === 'decimal' ||
    value === 'chineseCounting' ||
    value === 'ideographTraditional' ||
    value === 'decimalEnclosedCircle' ||
    value === 'lowerLetter' ||
    value === 'upperLetter' ||
    value === 'lowerRoman' ||
    value === 'upperRoman' ||
    value === 'none'
  ) {
    return value
  }
  return 'none'
}

function styleRole(value: string | null | undefined): StyleRole | null {
  if (value === undefined || value === null) return null
  if (value === 'normal' || value === 'title' || value === 'heading1' || value === 'heading2' || value === 'heading3') {
    return value
  }
  fail('role', `unknown role ${value}`)
}

function fail(path: string, message: string): never {
  throw new Error(`${path}: ${message}`)
}

function wrapLayout(layout: WasmLayout): PrintLayout {
  return {
    pageCount() {
      return layout.page_count()
    },
    pageSize(index) {
      try {
        return parsePageSize(layout.page_size(index))
      } catch (error) {
        throw asPrintError(error)
      }
    },
    pagePng(index, dpi) {
      try {
        return copyBytes(layout.page_png(index, dpi))
      } catch (error) {
        throw asPrintError(error)
      }
    },
    pdf() {
      try {
        return copyBytes(layout.pdf())
      } catch (error) {
        throw asPrintError(error)
      }
    },
    fonts() {
      try {
        return parsePrintFonts(layout.fonts())
      } catch (error) {
        throw asPrintError(error)
      }
    },
    blockPages() {
      try {
        return parseBlockPages(layout.block_pages())
      } catch (error) {
        throw asPrintError(error)
      }
    },
    blockBoxes() {
      try {
        return parseBlockBoxes(layout.block_boxes())
      } catch (error) {
        throw asPrintError(error)
      }
    },
    free() {
      layout.free()
    },
  }
}

function parsePageSize(json: string): readonly [number, number] {
  const parsed = parseJson(json)
  if (!Array.isArray(parsed) || parsed.length !== 2) throw renderError('page size')
  const width = parsed[0]
  const height = parsed[1]
  if (typeof width !== 'number' || typeof height !== 'number') throw renderError('page size')
  return [width, height]
}

function parseBlockPages(json: string): BlockPageInfo[] {
  const parsed = parseJson(json)
  if (!Array.isArray(parsed)) throw renderError('block pages')
  return parsed.map((item, index) => {
    if (!isRecord(item) || item.index !== index) throw renderError('block pages')
    return {
      index,
      firstPage: pageOrNull(item.firstPage),
      lastPage: pageOrNull(item.lastPage),
    }
  })
}

function parseBlockBoxes(json: string): BlockBoxInfo[] {
  const parsed = parseJson(json)
  if (!Array.isArray(parsed)) throw renderError('block boxes')
  return parsed.map((item, index) => {
    if (!isRecord(item) || item.index !== index || !Array.isArray(item.fragments)) throw renderError('block boxes')
    let previous = -1
    const fragments = item.fragments.map((fragment) => {
      if (!isRecord(fragment)) throw renderError('block boxes')
      const page = pageIndex(fragment.page)
      if (page < previous) throw renderError('block boxes')
      previous = page
      return { page, y: finiteNumber(fragment.y), height: finiteNumber(fragment.height) }
    })
    return { index, fragments }
  })
}

function pageIndex(value: unknown): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 0) throw renderError('block boxes')
  return value
}

function finiteNumber(value: unknown): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) throw renderError('block boxes')
  return value
}

function pageOrNull(value: unknown): number | null {
  if (value === null) return null
  if (typeof value === 'number' && Number.isInteger(value) && value >= 0) return value
  throw renderError('block pages')
}

function parsePrintFonts(json: string): PrintFontInfo[] {
  const parsed = parseJson(json)
  if (!Array.isArray(parsed)) throw renderError('fonts')
  return parsed.map((item) => {
    if (!isRecord(item) || typeof item.family !== 'string' || typeof item.embedded !== 'boolean') {
      throw renderError('fonts')
    }
    return { family: item.family, embedded: item.embedded }
  })
}

function parseJson(json: string): unknown {
  try {
    return JSON.parse(json)
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error)
    throw renderError(detail)
  }
}

function renderError(detail: string): Error {
  return new Error(JSON.stringify({ kind: 'render', detail }))
}

function asPrintError(error: unknown): never {
  if (error instanceof PrintError || error instanceof SaveWriteError) throw error
  throw toPrintError(error)
}

function toPrintError(error: unknown): PrintError | SaveWriteError {
  const detail = errorText(error)
  const parsed = parseErrorJson(detail)
  if (isRecord(parsed)) {
    if (parsed.kind === 'font' && typeof parsed.family === 'string') return new PrintError('font', parsed.family)
    if ((parsed.kind === 'layout' || parsed.kind === 'render') && typeof parsed.detail === 'string') {
      return new PrintError(parsed.kind, parsed.detail)
    }
    if (parsed.kind === 'write' && typeof parsed.detail === 'string') return new SaveWriteError(parsed.detail)
  }
  return new SaveWriteError(detail)
}

function errorText(error: unknown): string {
  if (typeof error === 'string') return error
  if (error instanceof Error) return error.message
  return String(error)
}

function parseErrorJson(detail: string): unknown {
  try {
    return JSON.parse(detail)
  } catch {
    const start = detail.indexOf('{')
    if (start < 0) return undefined
    try {
      return JSON.parse(detail.slice(start))
    } catch {
      return undefined
    }
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function toSaveError(error: unknown): SaveVerifyError | SaveWriteError {
  const detail = error instanceof Error ? error.message : String(error)
  try {
    const parsed: unknown = JSON.parse(detail)
    if (typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed)) {
      const record = parsed as Record<string, unknown>
      const difference = record.firstDifference ?? record.first_difference
      if (record.kind === 'verify' && typeof difference === 'string') {
        return new SaveVerifyError(difference)
      }
      if (record.kind === 'write' && typeof record.detail === 'string') {
        return new SaveWriteError(record.detail)
      }
    }
  } catch {
    // The wasm layer sometimes throws plain text.
  }
  return new SaveWriteError(detail)
}

function copyBytes(bytes: Uint8Array): Uint8Array {
  return new Uint8Array(bytes)
}

function toOpenError(error: unknown): OpenError {
  const detail = error instanceof Error ? error.message : String(error)
  try {
    const parsed: unknown = JSON.parse(detail)
    if (typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed)) {
      const record = parsed as Record<string, unknown>
      const reason = record.reason
      const message = record.detail
      if (
        record.kind === 'open' &&
        typeof message === 'string' &&
        (reason === 'notDocx' || reason === 'encrypted' || reason === 'corrupt' || reason === 'tooLarge')
      ) {
        return new OpenError(reason, message)
      }
    }
  } catch {
    // The wasm layer sometimes throws plain text.
  }
  return new OpenError('corrupt', detail)
}

async function pixelSize(bytes: Uint8Array, contentType: string): Promise<{ width: number; height: number }> {
  // acdert may expose one decoder and not the other. A throw from createImageBitmap is final only
  // when Image is absent; otherwise Image is what decides the file is not a picture.
  if (typeof createImageBitmap === 'function') {
    try {
      return await bitmapSize(bytes, contentType)
    } catch (error) {
      if (typeof Image !== 'function') throw error
    }
  }
  if (typeof Image === 'function') return elementSize(bytes, contentType)
  // node --test has neither decoder. A PNG or JPEG header carries the pixel size the decoder would.
  const header = imageSize(bytes)
  if (!header || header.width <= 0 || header.height <= 0) throw new Error('undecodable image')
  return header
}

async function bitmapSize(bytes: Uint8Array, contentType: string): Promise<{ width: number; height: number }> {
  const bitmap = await createImageBitmap(new Blob([bytes.slice()], { type: contentType }))
  try {
    if (bitmap.width <= 0 || bitmap.height <= 0) throw new Error('undecodable image')
    return { width: bitmap.width, height: bitmap.height }
  } finally {
    bitmap.close()
  }
}

function elementSize(bytes: Uint8Array, contentType: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(new Blob([bytes.slice()], { type: contentType }))
    const image = new Image()
    // acdert does not fetch or decode a detached <img>; mounting it off-screen drives onload and onerror.
    const body = typeof document !== 'undefined' ? document.body : null
    if (body) {
      image.setAttribute('aria-hidden', 'true')
      image.style.cssText = 'position:fixed;left:0;top:0;width:1px;height:1px;opacity:0;pointer-events:none'
      body.append(image)
    }
    const cleanup = (): void => {
      if (image.isConnected) image.remove()
      URL.revokeObjectURL(url)
    }
    image.onload = () => {
      const width = image.naturalWidth
      const height = image.naturalHeight
      cleanup()
      if (width <= 0 || height <= 0) {
        reject(new Error('undecodable image'))
        return
      }
      resolve({ width, height })
    }
    image.onerror = () => {
      cleanup()
      reject(new Error('undecodable image'))
    }
    image.src = url
  })
}

function sniff(bytes: Uint8Array): string {
  if (bytes.length >= 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return 'image/png'
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'image/jpeg'
  if (bytes.length >= 4 && bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) return 'image/gif'
  return 'application/octet-stream'
}

function imageSize(bytes: Uint8Array): { width: number; height: number } | undefined {
  if (bytes.length >= 24 && bytes[0] === 0x89 && bytes[1] === 0x50) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    return { width: view.getUint32(16), height: view.getUint32(20) }
  }
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return undefined
  let offset = 2
  while (offset + 9 < bytes.length) {
    if (bytes[offset] !== 0xff) return undefined
    const marker = bytes[offset + 1] ?? 0
    if (marker === 0xc0 || marker === 0xc1 || marker === 0xc2) {
      const height = ((bytes[offset + 5] ?? 0) << 8) | (bytes[offset + 6] ?? 0)
      const width = ((bytes[offset + 7] ?? 0) << 8) | (bytes[offset + 8] ?? 0)
      return { width, height }
    }
    const length = ((bytes[offset + 2] ?? 0) << 8) | (bytes[offset + 3] ?? 0)
    if (length < 2) return undefined
    offset += 2 + length
  }
  return undefined
}

function pt(value: number): Pt {
  return value as Pt
}

function hex(value: string): Hex {
  return value as Hex
}

function paraId(value: string): ParaId {
  return value as ParaId
}

function tableId(value: string): TableId {
  return value as TableId
}

function opaqueId(value: string): OpaqueId {
  return value as OpaqueId
}

function runSrc(value: number): RunSrc {
  return value as RunSrc
}

function styleId(value: string): StyleId {
  return value as StyleId
}

function numId(value: number): NumId {
  return value as NumId
}

function mediaKey(value: string): MediaKey {
  return value as MediaKey
}

function storyKey(value: string): StoryKey {
  return value as StoryKey
}
