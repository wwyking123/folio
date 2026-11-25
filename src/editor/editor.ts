// The only editor module the app imports. Holds one EditorState, runs commands, and drives PageView.
import type { Opened, Source } from '../engine/engine.ts'
import { cssStack, fetchMetrics, rememberedFamilies } from '../fonts.ts'
import type { FontMetrics } from '../fonts.ts'
import { indexOf } from '../model.ts'
import type { Block, Doc, Paragraph, ParaId, ParaLook, Pos, Table } from '../model.ts'
import { apply, initialState, insertImage, redo as redoState, selectedImage, setSelection, toolbarState, undo as undoState, wordCount } from './edit.ts'
import type { Command, EditorState, MediaRef, ToolbarState } from './edit.ts'
import { attachInput } from './input-dom.ts'
import type { EditingSurface, InputBinding, PictureInput } from './input-dom.ts'
import { detectPlatform } from './input.ts'
import type { FileOp, FindOp, Platform } from './input.ts'
import type { TextMatch } from './find.ts'
import {
  cssGeneric,
  documentGridPitch,
  lineHeightPx,
  listMarkers,
  withCaretParagraph,
} from './derive.ts'
import type { ImageGesture, PageBox } from './image-geom.ts'
import { pageOfBlock, PageView } from './pages.ts'
import type { Decoration, RenderContext } from './render.ts'

export interface EditorOptions {
  readonly locale: 'zh' | 'en'
}

export interface Status {
  /** 1-based page holding the caret. */
  readonly page: number
  readonly pages: number
  /**
   * False while a large document is still being measured, and while this generation is still resolving font metrics.
   * A non-empty metrics result lays the document out again, so the first pass is not final.
   */
  readonly pagesFinal: boolean
  readonly words: number
  /** Section index holding the caret, for page setup. */
  readonly section: number
  /** A transient status-line message, or null. */
  readonly notice: string | null
}

const EMPTY_IMAGE =
  'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'

export class Editor implements EditingSurface {
  readonly platform: Platform
  private readonly host: HTMLElement
  private readonly view: PageView
  private readonly sink: HTMLTextAreaElement
  private input: InputBinding | null = null
  private readonly listeners = new Set<() => void>()
  private readonly fileListeners = new Set<(op: FileOp) => void>()
  private readonly findListeners = new Set<(op: FindOp) => void>()
  private readonly pictureErrorListeners = new Set<() => void>()
  private session: Source
  private current: EditorState
  private markers: ReadonlyMap<ParaId, string>
  private metrics: ReadonlyMap<string, FontMetrics> = new Map()
  private metricsSettled = false
  /** Families already resolved, including ones the index does not have. A catalog that could not be read does not add to this set. */
  private knownFonts = new Set<string>()
  /** Families with a metrics request in flight. */
  private pendingFonts = new Set<string>()
  private generation = 0
  /** The generation of the last mount or load. An edit does not advance it. */
  private loadGeneration = 0
  /** An edit since that load changed the document, so the metrics pass keeps the viewport. */
  private editedSinceLoad = false
  private changeFrame = 0
  private destroyed = false
  private notice: string | null = null

  /** Creates the page view inside `host` and shows the opened document. Input listeners belong to the input layer. */
  static mount(host: HTMLElement, opened: Opened, options: EditorOptions): Editor {
    return new Editor(host, opened, options)
  }

  private constructor(host: HTMLElement, opened: Opened, options: EditorOptions) {
    const doc = withCaretParagraph(opened.doc)
    const state = initialState(doc)
    this.host = host
    this.platform = detectPlatform(navigator.platform, navigator.userAgent)
    this.session = opened.source
    this.current = state
    this.markers = listMarkers(doc)
    const sink = document.createElement('textarea')
    sink.className = 'ime-sink'
    sink.lang = options.locale
    sink.autocomplete = 'off'
    sink.spellcheck = false
    this.sink = sink
    this.view = new PageView(host, () => this.renderContext(), () => this.scheduleChange())
    this.view.show(doc)
    this.paintSelection(true)
    this.input = attachInput(this, this.sink, this.host)
    this.scheduleChange()
    void this.applyMetrics(this.generation)
  }

  /** Replaces the document and the rdocx session; frees the previous session; resets history and selection. */
  load(opened: Opened): void {
    if (this.destroyed) {
      opened.source.free()
      return
    }
    this.input?.reset()
    this.generation += 1
    const generation = this.generation
    this.loadGeneration = generation
    this.editedSinceLoad = false
    const doc = withCaretParagraph(opened.doc)
    const state = initialState(doc)
    const previous = this.session
    this.session = opened.source
    this.current = state
    this.markers = listMarkers(doc)
    this.metrics = new Map()
    this.metricsSettled = false
    this.knownFonts.clear()
    this.pendingFonts.clear()
    this.notice = null
    this.view.show(doc)
    this.view.setImagePreview(null)
    this.paintSelection(true)
    if (previous !== opened.source) previous.free()
    this.scheduleChange()
    void this.applyMetrics(generation)
  }

  get state(): EditorState {
    return this.current
  }

  /** The rdocx session the current document came from: save, export, fonts, media. */
  get source(): Source {
    return this.session
  }

  /**
   * Runs a command. On a transaction: apply() with history, PageView.update when the document changed, paint the
   * caret, scroll it into view, schedule 'change' for the next animation frame. Returns false when the command did
   * not apply.
   */
  run(command: Command): boolean {
    if (this.destroyed) return false
    const tr = command(this.current)
    if (!tr) return false
    this.commit(apply(this.current, tr, Date.now()), true)
    return true
  }

  undo(): boolean {
    if (this.destroyed) return false
    const next = undoState(this.current)
    if (!next) return false
    this.commit(next, true)
    return true
  }

  redo(): boolean {
    if (this.destroyed) return false
    const next = redoState(this.current)
    if (!next) return false
    this.commit(next, true)
    return true
  }

  /** IME preedit from input.ts. View-only; never enters state or history. */
  setPreedit(deco: Decoration | null): void {
    if (this.destroyed) return
    this.view.decorate(deco)
    this.paintSelection(true)
    this.scheduleChange()
  }

  /** The runtime's composition span is painted after the composition event. Slide the caret to its end. */
  syncCompositionCaret(): void {
    if (this.destroyed) return
    this.view.syncHostPreeditCaret()
  }

  /** Called once per animation frame after changes, never per keystroke synchronously. */
  on(event: 'change', listener: () => void): () => void {
    void event
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  /** File and tab chords from the keymap. The app owns files and tabs; this only delivers the request. */
  onFile(listener: (op: FileOp) => void): () => void {
    this.fileListeners.add(listener)
    return () => {
      this.fileListeners.delete(listener)
    }
  }

  fileCommand(op: FileOp): void {
    if (this.destroyed) return
    for (const listener of this.fileListeners) listener(op)
  }

  /** Find, replace, and match navigation from the keymap. The find bar owns the panel. */
  onFind(listener: (op: FindOp) => void): () => void {
    this.findListeners.add(listener)
    return () => {
      this.findListeners.delete(listener)
    }
  }

  findCommand(op: FindOp): void {
    if (this.destroyed) return
    for (const listener of this.findListeners) listener(op)
  }

  /** Text drag or picture drag. Escape cancels the picture drag before it closes the find bar. */
  dragging(): boolean {
    return this.input?.dragging ?? false
  }

  /** Match tints on mounted pages. An empty list clears them. Does not mount pages. */
  setFindHits(hits: readonly TextMatch[], current: number): void {
    if (this.destroyed) return
    this.view.setFindHits(hits, current)
  }

  /** A pasted or dropped file that is not a picture the engine can embed. The app shows the message. */
  onPictureError(listener: () => void): () => void {
    this.pictureErrorListeners.add(listener)
    return () => {
      this.pictureErrorListeners.delete(listener)
    }
  }

  pictureError(): void {
    if (this.destroyed) return
    for (const listener of this.pictureErrorListeners) listener()
  }

  /** The one selected picture with its bytes, for the clipboard. Null when the selection is anything else. */
  selectedPicture(): PictureInput | null {
    if (this.destroyed) return null
    const hit = selectedImage(this.current.doc, this.current.sel)
    if (!hit || hit.atom.atom.type !== 'image') return null
    const blob = this.session.mediaBlob(hit.atom.atom.media)
    if (!blob) return null
    return { blob, widthPt: hit.atom.atom.widthPt, heightPt: hit.atom.atom.heightPt }
  }

  /**
   * Embeds the pictures and inserts them in order, at `at` or at the selection. Each is one undo entry. A blob the
   * engine refuses is skipped and reported once. A document replaced while the bytes decode gets nothing.
   */
  insertPictures(pictures: readonly PictureInput[], at: Pos | null): void {
    if (this.destroyed || pictures.length === 0) return
    void this.embedPictures(pictures, at)
  }

  private async embedPictures(pictures: readonly PictureInput[], at: Pos | null): Promise<void> {
    const session = this.session
    const refs: MediaRef[] = []
    let refused = false
    for (const picture of pictures) {
      try {
        const media = await session.addMedia(picture.blob)
        refs.push({
          key: media.key,
          contentType: media.contentType,
          widthPx: (picture.widthPt ?? media.widthPt) / 0.75,
          heightPx: (picture.heightPt ?? media.heightPt) / 0.75,
        })
      } catch {
        refused = true
      }
    }
    if (this.destroyed || this.session !== session) return
    if (refs.length > 0) {
      if (at) {
        try {
          this.setTextSelection(at, at)
        } catch {
          // The paragraph under the pointer went away while the bytes decoded. The selection is where they go.
        }
      }
      for (const ref of refs) this.run(insertImage(ref))
    }
    if (refused) this.pictureError()
  }

  toolbar(): ToolbarState {
    return toolbarState(this.current, (para, marks) => this.session.look(para, marks))
  }

  /** The look the page paints: direct properties over the style chain, including a list level's indent. */
  paraLook(para: Paragraph): ParaLook {
    return this.session.paraLook(para, this.current.doc.lists)
  }

  status(): Status {
    const pages = this.view.layout.pages.length
    const pos = caretPos(this.current)
    const page = pos !== null && pages > 0 ? this.view.pageOf(pos) + 1 : 1
    return {
      page,
      pages,
      pagesFinal: this.metricsSettled && this.view.layout.complete,
      words: wordCount(this.current.doc),
      section: pos !== null ? sectionIndex(this.current.doc, pos) : 0,
      notice: this.notice,
    }
  }

  /**
   * 0-based page of a top-level body block, read from the layout the view already holds.
   * Null when that block has no slice, which includes every block past the measured part
   * while the layout is incomplete.
   */
  pageOf(id: string): number | null {
    if (this.destroyed) return null
    return pageOfBlock(this.view.layout, id)
  }

  /**
   * Cached block heights in points. Reads the paginator's measure cache and does not
   * measure, so a block past the measured chunk is null.
   */
  blockHeights(): readonly { index: number; kind: 'p' | 'table' | 'opaque'; heightPt: number | null }[] {
    if (this.destroyed) return []
    return this.view.blockHeights()
  }

  setZoom(zoom: number): void {
    if (this.destroyed) return
    this.view.setZoom(zoom)
    this.scheduleChange()
  }

  zoom(): number {
    return this.view.currentZoom()
  }

  imageBox(pos: Pos): PageBox | null {
    if (this.destroyed) return null
    return this.view.imageBox(pos)
  }

  previewImage(preview: ImageGesture | null): void {
    if (this.destroyed) return
    this.view.setImagePreview(preview)
  }

  setNotice(text: string | null): void {
    if (this.destroyed || this.notice === text) return
    this.notice = text
    this.scheduleChange()
  }

  focus(): void {
    if (this.destroyed) return
    if (!this.sink.isConnected) this.host.append(this.sink)
    this.sink.focus()
  }

  /** True while the input layer has an IME composition open. */
  get composing(): boolean {
    return this.input?.composing ?? false
  }

  /** Where a pointer click lands, or null when it misses the pages. */
  positionAtPoint(clientX: number, clientY: number): Pos | null {
    if (this.destroyed) return null
    return this.view.posAtPoint(clientX, clientY)
  }

  /** The caret's x in page coordinates, kept as the goal column across vertical moves. */
  caretColumn(): number | null {
    if (this.destroyed) return null
    const sel = this.current.sel
    if (sel.kind !== 'text') return null
    return this.view.caretRect(sel.head).x
  }

  /**
   * Moves the caret one visual line. `goalX` is the column to hit; null measures the caret.
   * Returns the column to keep for the next vertical move.
   */
  moveLine(direction: 'up' | 'down', extend: boolean, goalX: number | null): number | null {
    if (this.destroyed) return goalX
    const sel = this.current.sel
    if (sel.kind !== 'text') return goalX
    const image = !extend ? selectedImage(this.current.doc, sel) : null
    const origin = image
      ? { p: image.para.id, o: direction === 'down' ? image.offset + 1 : image.offset }
      : sel.head
    const moved = this.view.lineMove(origin, direction, goalX)
    const anchor = extend ? sel.anchor : moved.pos
    this.run(setSelection({ kind: 'text', anchor, head: moved.pos }))
    return moved.goalX
  }

  /** Page Up and Page Down. Returns the goal column, like moveLine. */
  movePage(direction: 'up' | 'down', extend: boolean, goalX: number | null): number | null {
    if (this.destroyed) return goalX
    const sel = this.current.sel
    if (sel.kind !== 'text') return goalX
    const moved = this.view.pageMove(sel.head, direction, goalX)
    const anchor = extend ? sel.anchor : moved.pos
    this.run(setSelection({ kind: 'text', anchor, head: moved.pos }))
    return moved.goalX
  }

  /** Home and End: the first or last offset on the caret's visual line. */
  moveToLineEdge(edge: 'start' | 'end', extend: boolean): void {
    if (this.destroyed) return
    const sel = this.current.sel
    if (sel.kind !== 'text') return
    const next = this.view.lineEdge(sel.head, edge)
    const anchor = extend ? sel.anchor : next
    this.run(setSelection({ kind: 'text', anchor, head: next }))
  }

  /** Places the caret, or a selection when anchor and head differ. */
  setTextSelection(anchor: Pos, head: Pos): void {
    if (this.destroyed) return
    this.run(setSelection({ kind: 'text', anchor, head }))
  }

  destroy(): void {
    if (this.destroyed) return
    this.destroyed = true
    this.input?.detach()
    this.input = null
    this.generation += 1
    if (this.changeFrame !== 0) cancelAnimationFrame(this.changeFrame)
    this.changeFrame = 0
    this.listeners.clear()
    this.findListeners.clear()
    this.view.destroy()
    this.sink.remove()
    this.session.free()
  }

  private commit(next: EditorState, scroll: boolean): void {
    const previous = this.current.doc
    this.current = next
    if (next.doc !== previous) {
      this.editedSinceLoad = true
      const previousMarkers = this.markers
      const markers = listMarkers(next.doc)
      this.markers = markers
      // update() cancels the in-progress measure. A selection-only transaction keeps the pump running.
      this.view.update(previous, next.doc, previousMarkers, markers)
      void this.applyMetrics(this.generation)
    }
    this.paintSelection(scroll)
    this.scheduleChange()
  }

  private paintSelection(scroll: boolean): void {
    const sel = this.current.sel
    this.view.paint(sel, this.sink, this.current.stored)
    if (scroll && sel.kind === 'text') this.view.scrollIntoView(sel.head)
  }

  private scheduleChange(): void {
    if (this.changeFrame !== 0 || this.destroyed) return
    this.changeFrame = requestAnimationFrame(() => {
      this.changeFrame = 0
      if (this.destroyed) return
      for (const listener of [...this.listeners]) listener()
    })
  }

  private renderContext(): RenderContext {
    return {
      look: (para, marks) => this.session.look(para, marks),
      paraLook: (para) => this.session.paraLook(para, this.current.doc.lists),
      fontStack: (family) => cssStack(family, cssGeneric(family)),
      lineHeightPx: (family, sizePt, rule) => lineHeightPx(this.metrics.get(family), sizePt, rule),
      cjkGlyphs: (family) => this.metrics.get(family)?.cjk,
      marker: (id) => this.markers.get(id) ?? null,
      imageUrl: (key) => {
        try {
          return this.session.imageUrl(key)
        } catch {
          return EMPTY_IMAGE
        }
      },
      gridPitchPx: documentGridPitch(this.current.doc),
    }
  }

  private async applyMetrics(generation: number): Promise<void> {
    if (this.destroyed || generation !== this.generation) return
    const missing: string[] = []
    for (const family of this.session.fontsUsed(this.current.doc)) {
      if (this.knownFonts.has(family) || this.pendingFonts.has(family)) continue
      this.pendingFonts.add(family)
      missing.push(family)
    }
    if (missing.length === 0) {
      if (!this.metricsSettled && this.pendingFonts.size === 0) {
        this.metricsSettled = true
        this.scheduleChange()
      }
      return
    }
    this.metricsSettled = false
    this.scheduleChange()
    let fetched: ReadonlyMap<string, FontMetrics> = new Map()
    let available = false
    try {
      const read = await fetchMetrics(missing)
      available = read.available
      fetched = read.faces
    } catch {
      available = false
    }
    // A load() that started while this was in flight cleared the sets and owns them now.
    if (this.destroyed || generation !== this.generation) return
    for (const family of missing) this.pendingFonts.delete(family)
    for (const family of rememberedFamilies(missing, available, fetched.keys())) this.knownFonts.add(family)
    if (available && fetched.size > 0) {
      const merged = new Map(this.metrics)
      for (const [family, value] of fetched) merged.set(family, value)
      this.metrics = merged
      // The paint before this fetch used the fallback line box. A resolved face changes the line heights.
      // The load's own fetch starts at the top. A fetch for a family an edit introduced keeps the viewport.
      const keepPlace = this.editedSinceLoad && generation === this.loadGeneration
      if (keepPlace) this.view.relayout(this.current.doc)
      else this.view.show(this.current.doc)
      this.paintSelection(false)
    }
    this.metricsSettled = true
    this.scheduleChange()
  }
}

export type { Doc, Pos }

function caretPos(state: EditorState): Pos | null {
  const sel = state.sel
  if (sel.kind === 'text') return sel.head
  const table = findTable(state.doc.body, sel.table)
  const cell = table?.rows[sel.anchor.row]?.cells[sel.anchor.cell]
  const para = cell ? firstParagraph(cell.blocks) : null
  if (!para) return null
  return { p: para.id, o: 0 }
}

function findTable(blocks: readonly Block[], id: Table['id']): Table | null {
  for (const block of blocks) {
    if (block.kind !== 'table') continue
    if (block.id === id) return block
    for (const row of block.rows) {
      for (const cell of row.cells) {
        const found = findTable(cell.blocks, id)
        if (found) return found
      }
    }
  }
  return null
}

function firstParagraph(blocks: readonly Block[]): Paragraph | null {
  for (const block of blocks) {
    if (block.kind === 'p') return block
    if (block.kind === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          const found = firstParagraph(cell.blocks)
          if (found) return found
        }
      }
    } else if (block.preview[0]) return block.preview[0]
  }
  return null
}

function sectionIndex(doc: Doc, pos: Pos): number {
  let index = 0
  try {
    const loc = indexOf(doc).get(pos.p)
    if (loc.story !== 'body') return 0
    index = loc.path[0] ?? 0
  } catch {
    return 0
  }
  let section = 0
  for (let i = 0; i < index; i++) {
    const block = doc.body[i]
    if (block?.kind === 'p' && block.sectionEnd !== null) section = block.sectionEnd + 1
  }
  const count = doc.sections.length
  if (count === 0) return 0
  return section < count ? section : count - 1
}
