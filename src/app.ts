// Owns the open documents, one tab each, and the window chrome. A tab is dirty when its editor.state.doc !== file.saved.
import { window as appWindow } from 'acde:window'
import { createInterface } from './automation/api.ts'
import type { Word } from './automation/api.ts'
import type { Host } from './automation/host.ts'
import { Editor } from './editor/editor.ts'
import { changeIndentLevel, clearMarks, deleteColumns, deleteEnclosingTable, deleteRows, insertColumn, insertImage, insertRow, insertTable, setAlign, setLineSpacing, setMark, setPageSetup, setParaFormat, setParaStyle, setSpacing, toggleList, toggleMark, toggleVertAlign, touchedParagraphs } from './editor/edit.ts'
import { listMarkers } from './editor/derive.ts'
import type { ToggleMark } from './editor/edit.ts'
import { detectPlatform } from './editor/input.ts'
import { drift, heightKind } from './editor/drift.ts'
import type { DriftReport, HeightKind } from './editor/drift.ts'
import { OpenError, SaveVerifyError, SaveWriteError, documentHasHan, engineBootMs, loadEngine, saveTampered, sessionFaces, wasmMemoryBytes } from './engine/engine.ts'
import type { Opened, PrintFontInfo, PrintLayout } from './engine/engine.ts'
import { ensurePrintFonts } from './fonts.ts'
import { locale, t } from './i18n.ts'
import type { Doc, Hex, HighlightName, ParaId, Pt, StyleId } from './model.ts'
import { createStatus } from './ui/status.ts'
import { createTabs } from './ui/tabs.ts'
import { createToolbar } from './ui/toolbar.ts'
import { afterClose, canSave, cycleTab, fileHandleFromRequest, formatTitle, isAbort, isDirty, nextUntitled } from './ui/text.ts'
import { isDialogOpen } from './ui/dialog.ts'
import { openPageSetup } from './ui/page-setup.ts'
import { openParagraph } from './ui/paragraph.ts'
import { openInsertTable } from './ui/table-tools.ts'
import { createUnsavedDialog } from './ui/unsaved.ts'
import type { UnsavedChoice } from './ui/unsaved.ts'
import { createFindBar } from './ui/find-bar.ts'
import type { FindBar } from './ui/find-bar.ts'
import type { FileOp } from './editor/input.ts'
import { describePrintFailure, mountPreview, pdfFileName, previewDpi, PreviewMachine } from './ui/preview.ts'

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
const PDF_MIME = 'application/pdf'

export interface OpenFile {
  readonly handle: FileSystemFileHandle | null
  readonly name: string
  readonly saved: Doc
}

export interface WordApp {
  /** The active tab's editor. It changes when another tab is chosen. */
  readonly editor: Editor
  /** `window.word`. JSON in, JSON out, in every build. */
  readonly word: Word
  openBytes(bytes: Uint8Array, name?: string): Promise<void>
  /** Bytes from `source.save`, timed around that call only. Does not write a file or clear the dirty bit. */
  save(): { base64: string; ms: number }
  /**
   * Dev hook: save, then fail verification. Shows the error band, writes nothing and stays dirty.
   * Returns the detail. There is no separate production build, so the hook lives on `window.__word`.
   */
  saveTampered(): string
  /** The `beforeclose` handler. False means the window stays and the dialog is up. */
  requestClose(): boolean
  /** How long the last page-setup confirm blocked, including the relayout. Zero before the first one. */
  pageSetupMs(): number
  /** The wasm memory size now. The print preview e2e reads it after opening and closing. */
  memoryBytes(): number
  /**
   * Dev hook: fetch the CJK faces, lay out the open document, and time the blocking calls.
   * Does not write a file. The PDF is base64. The preview does not call this.
   */
  printLayout(): Promise<PrintMeasure>
  /**
   * Dev hook: one print layout of the open document, compared with the editing
   * view's page per body block. Nothing calls it until the hook does.
   */
  drift(): Promise<DriftReport>
  /**
   * Dev hook: cached body-block heights in points, and the kind. Does not measure.
   * A block the paginator has not reached is a null height.
   */
  blockHeights(): readonly BlockHeight[]
  /**
   * Dev hook: one print layout, and the summed fragment height of each body block.
   * Null when that block was not placed. Nothing calls it until the hook does.
   */
  printHeights(): Promise<readonly PrintBlockHeight[]>
}

export type BlockHeight = {
  readonly index: number
  readonly kind: HeightKind
  readonly heightPt: number | null
}

export type PrintBlockHeight = {
  readonly index: number
  readonly heightPt: number | null
}

export type PrintMeasure = {
  readonly pages: number
  readonly pdfBase64: string
  readonly fontMs: number
  readonly layoutMs: number
  readonly pdfMs: number
  readonly pngMs: number
  readonly pngBytes: number
  readonly memoryBefore: number
  readonly memoryAfter: number
  readonly memory: PrintMemory
  readonly faceBytes: number
  readonly faces: readonly PrintSessionFace[]
  readonly bootMs: number
  readonly fonts: readonly PrintFontInfo[]
}

export type PrintMemory = {
  readonly baseline: number
  readonly afterFonts: number
  readonly afterLayout: number
  readonly afterPdf: number
  readonly afterPng: number
}

export type PrintSessionFace = {
  readonly family: string
  readonly bytes: number
}

/** One open document: its editor, the host the editor draws into, and the file behind it. */
type Tab = {
  readonly id: number
  readonly editor: Editor
  readonly host: HTMLElement
  file: OpenFile
  /** A new document nobody has opened a file into yet. Open fills it in place instead of adding a tab while it is clean. */
  blank: boolean
}

export async function boot(host: HTMLElement): Promise<WordApp> {
  document.body.classList.add('shell')
  const engine = await loadEngine()
  const platform = detectPlatform(navigator.platform, navigator.userAgent)
  const tabs: Tab[] = []
  let active: Tab | null = null
  let nextTabId = 1
  let titleShown = ''
  let tail: Promise<void> = Promise.resolve()
  let resolveAsk: ((choice: UnsavedChoice) => void) | null = null
  let closingWindow = false
  let lastPageSetupMs = 0
  // listMarkers already caches on the document. read.ts keeps its outline only while the map
  // object is the same, so this map is what `markers()` hands out for a given Doc.
  const markersByDoc = new WeakMap<Doc, ReadonlyMap<ParaId, string>>()
  const changeListeners = new Set<() => void>()
  let findBar: FindBar | null = null
  let watchedEditor: Editor | null = null
  let unwatchEditor: (() => void) | null = null

  const toolbarRoot = required('#toolbar')
  const statusRoot = required('#status')
  const tabStrip = createTabs(required('#tabs'), {
    onSelect(id) {
      const tab = tabs.find(item => item.id === id)
      if (tab) activate(tab)
    },
    onClose(id) {
      const tab = tabs.find(item => item.id === id)
      if (tab) void closeTab(tab)
    },
    onNew() {
      newDocument()
    },
  })
  const toolbar = createToolbar(toolbarRoot, platform, {
    onNew() {
      newDocument()
    },
    onOpen() {
      void openWithPicker()
    },
    onSave() {
      void saveFromToolbar(false)
    },
    onSaveAs() {
      void saveFromToolbar(true)
    },
    onUndo() {
      edit(() => { current().undo() })
    },
    onRedo() {
      edit(() => { current().redo() })
    },
    onMark(mark: ToggleMark) {
      edit(() => { current().run(toggleMark(mark)) })
    },
    onFont(family: string) {
      edit(() => { current().run(setMark({ mark: 'font', value: family })) })
    },
    onSize(sizePt: number) {
      edit(() => { current().run(setMark({ mark: 'sizePt', value: sizePt as Pt })) })
    },
    onStyle(id: string) {
      edit(() => { current().run(setParaStyle(id as StyleId)) })
    },
    onAlign(align) {
      edit(() => { current().run(setAlign(align)) })
    },
    onList(kind) {
      edit(() => { current().run(toggleList(kind)) })
    },
    onIndent(delta) {
      edit(() => { current().run(changeIndentLevel(delta)) })
    },
    onColor(hex) {
      edit(() => { current().run(setMark({ mark: 'color', value: hex === undefined ? undefined : hex as Hex })) })
    },
    onHighlight(name: HighlightName | undefined) {
      edit(() => { current().run(setMark({ mark: 'highlight', value: name })) })
    },
    onVertAlign(which) {
      edit(() => { current().run(toggleVertAlign(which)) })
    },
    onClearMarks() {
      edit(() => { current().run(clearMarks) })
    },
    onLineSpacing(multiple) {
      edit(() => { current().run(setLineSpacing({ rule: 'auto', multiple })) })
    },
    onParaSpacing(edge, points) {
      const spacing = edge === 'before' ? { before: points as Pt } : { after: points as Pt }
      edit(() => { current().run(setSpacing(spacing)) })
    },
    onParagraph() {
      const editorNow = current()
      if (isDialogOpen()) return
      const paras = touchedParagraphs(editorNow.state.doc, editorNow.state.sel)
      const page = editorNow.state.doc.sections[editorNow.status().section]?.page
      if (!page || paras.length === 0) {
        editorNow.focus()
        return
      }
      const looks = paras.map((para) => editorNow.paraLook(para))
      void openParagraph(looks, page).then((format) => {
        if (!format) {
          current().focus()
          return
        }
        edit(() => { current().run(setParaFormat(format)) })
      })
    },
    onPageSetup() {
      const editorNow = current()
      const section = editorNow.status().section
      const page = editorNow.state.doc.sections[section]?.page
      if (!page || isDialogOpen()) return
      void openPageSetup(page).then((next) => {
        if (!next) {
          current().focus()
          return
        }
        const started = performance.now()
        edit(() => { current().run(setPageSetup(section, next)) })
        lastPageSetupMs = performance.now() - started
      })
    },
    onPrint() {
      void openPreview()
    },
    onFind() {
      findBar?.command('open')
    },
    onInsertTable() {
      if (isDialogOpen() || current().toolbar().inTable) return
      void openInsertTable().then((size) => {
        if (!size || current().toolbar().inTable) {
          current().focus()
          return
        }
        edit(() => { current().run(insertTable(size.rows, size.cols)) })
      })
    },
    onInsertImage() {
      void insertPicture()
    },
    onTableRow(where) {
      edit(() => { current().run(insertRow(where)) })
    },
    onDeleteRows() {
      edit(() => { current().run(deleteRows()) })
    },
    onTableColumn(where) {
      if (!current().toolbar().columnTools) return
      edit(() => { current().run(insertColumn(where)) })
    },
    onDeleteColumns() {
      if (!current().toolbar().columnTools) return
      edit(() => { current().run(deleteColumns()) })
    },
    onDeleteTable() {
      edit(() => { current().run(deleteEnclosingTable()) })
    },
  }, () => scroller())
  const status = createStatus(statusRoot, (zoom) => {
    edit(() => { current().setZoom(zoom) })
  })
  const dialog = createUnsavedDialog((choice) => {
    const resolve = resolveAsk
    resolveAsk = null
    resolve?.(choice)
  })
  const pages = host
  const machine = new PreviewMachine<PrintLayout>()
  let heldScroll = { top: 0, left: 0 }
  const preview = mountPreview(pages, machine, {
    onCancel() {
      requestPreviewClose()
    },
    onClose() {
      requestPreviewClose()
    },
    onExport() {
      void exportPdf()
    },
    onRenderError(error) {
      failPreview(error)
    },
  })
  // The dialog's own listener is capture-phase and stops Escape while it is open.
  document.addEventListener('keydown', (event) => {
    if (event.key !== 'Escape' || isDialogOpen()) return
    const phase = machine.phase
    if (phase === 'closed' || phase === 'unusable' || phase === 'failed') return
    event.preventDefault()
    requestPreviewClose()
  })

  function current(): Editor {
    if (!active) throw new Error('editor is not ready')
    return active.editor
  }

  function activeTab(): Tab {
    if (!active) throw new Error('editor is not ready')
    return active
  }

  /** The active tab's page scroller. */
  function scroller(): HTMLElement {
    return active?.host ?? pages
  }

  // Each tab has its own Editor, so current() changes when another tab is chosen. The automation
  // change subscription follows it: activate() rebinds and tells the listeners.
  function rebindWatch(): void {
    const next = current()
    if (next === watchedEditor) return
    unwatchEditor?.()
    watchedEditor = next
    unwatchEditor = next.on('change', () => {
      for (const listener of [...changeListeners]) listener()
    })
  }

  function editorNow(): Editor {
    const next = current()
    if (changeListeners.size > 0 && next !== watchedEditor) rebindWatch()
    return next
  }

  function renderChrome(): void {
    const tab = activeTab()
    const currentEditor = tab.editor
    const file = tab.file
    const name = file.name
    const title = formatTitle(name, isDirty(currentEditor.state.doc, file), t('title'))
    if (title !== titleShown) {
      titleShown = title
      document.title = title
      appWindow.setTitle(title)
    }
    toolbar.render(
      currentEditor.toolbar(),
      currentEditor.state.doc,
      currentEditor.source.fontsUsed(currentEditor.state.doc),
      canSave(currentEditor.state.doc, file),
    )
    status.render(currentEditor.status())
    tabStrip.render(
      tabs.map(item => ({ id: item.id, name: item.file.name, dirty: isDirty(item.editor.state.doc, item.file) })),
      tab.id,
    )
  }

  function edit(run: () => void): void {
    run()
    renderChrome()
    current().focus()
  }

  /** Mounts a new editor in its own host at the end of the strip. The caller activates it. */
  function createTab(opened: Opened, handle: FileSystemFileHandle | null, name: string, blank: boolean): Tab {
    const tabHost = document.createElement('div')
    tabHost.className = 'doc-host background'
    pages.append(tabHost)
    const tabEditor = Editor.mount(tabHost, opened, { locale: locale === 'zh' ? 'zh' : 'en' })
    const tab: Tab = { id: nextTabId++, editor: tabEditor, host: tabHost, file: { handle, name, saved: tabEditor.state.doc }, blank }
    tabEditor.on('change', () => { if (active) renderChrome() })
    tabEditor.onFile((op) => { fileOp(op) })
    tabEditor.onPictureError(() => showBand(t('image.error')))
    tabEditor.on('change', () => { if (tab === active) findBar?.noteDocument() })
    tabEditor.onFind((op) => { if (tab === active) findBar?.command(op) })
    tabs.push(tab)
    return tab
  }

  /** Shows `tab` and hides the others. Each host keeps its own scroll position while it is hidden. */
  function activate(tab: Tab): void {
    if (!tabs.includes(tab)) return
    const changed = active !== tab
    // The bar's matches and tints belong to the editor it searched. Close it on the old tab before switching.
    if (changed && active) findBar?.close(false)
    active = tab
    for (const item of tabs) {
      item.host.classList.toggle('active', item === tab)
      item.host.classList.toggle('background', item !== tab)
    }
    if (changed) {
      clearError()
      if (changeListeners.size > 0) {
        rebindWatch()
        for (const listener of [...changeListeners]) listener()
      }
    }
    renderChrome()
    tab.editor.focus()
  }

  function fileOp(op: FileOp): void {
    if (isDialogOpen() || previewCovers()) return
    switch (op) {
      case 'save':
      case 'saveAs':
        void saveFromToolbar(op === 'saveAs')
        return
      case 'new':
        newDocument()
        return
      case 'open':
        void openWithPicker()
        return
      case 'closeTab':
        void closeTab(activeTab())
        return
      case 'nextTab':
      case 'prevTab':
        activate(cycleTab(tabs, activeTab(), op === 'nextTab' ? 1 : -1))
        return
    }
  }

  function previewCovers(): boolean {
    const phase = machine.phase
    return phase === 'preparing' || phase === 'cancelling' || phase === 'open' || phase === 'exporting'
  }

  function adopt(tab: Tab, opened: Opened, handle: FileSystemFileHandle | null, name: string): void {
    tab.editor.load(opened)
    tab.file = { handle, name, saved: tab.editor.state.doc }
    tab.blank = false
    clearError()
    renderChrome()
  }

  function newDocument(): void {
    activate(createTab(engine.blank(), null, nextUntitled(tabs.map(item => item.file.name), t), true))
  }

  /** Where an opened file goes: the active tab when it is a clean new document, else a new tab. */
  function openInto(opened: Opened, handle: FileSystemFileHandle | null, name: string): void {
    const tab = activeTab()
    if (tab.blank && !isDirty(tab.editor.state.doc, tab.file)) {
      adopt(tab, opened, handle, name)
      return
    }
    activate(createTab(opened, handle, name, false))
  }

  function tabDirty(tab: Tab): boolean {
    return isDirty(tab.editor.state.doc, tab.file)
  }

  function askUnsaved(tab: Tab): Promise<UnsavedChoice> {
    return new Promise((resolve) => {
      resolveAsk = resolve
      dialog.ask(tab.file.name)
    })
  }

  /** Saves `tab` to its file and waits for every queued save. True when it is clean afterwards. */
  async function saveTab(tab: Tab): Promise<boolean> {
    const ok = await enqueueSave(tab, false)
    await drain()
    return ok && !tabDirty(tab)
  }

  /** Closes one tab, asking first when it has unsaved changes. The last tab gives way to a new blank document. */
  async function closeTab(tab: Tab): Promise<void> {
    if (dialog.open || isDialogOpen() || previewCovers() || closingWindow) return
    if (tabDirty(tab)) {
      activate(tab)
      const choice = await askUnsaved(tab)
      if (choice === 'cancel') {
        current().focus()
        return
      }
      if (choice === 'save' && !(await saveTab(tab))) {
        current().focus()
        return
      }
    }
    removeTab(tab)
  }

  function removeTab(tab: Tab): void {
    const at = tabs.indexOf(tab)
    if (at < 0) return
    const next = afterClose(tabs, tab, activeTab())
    if (next === null) {
      const fresh = createTab(engine.blank(), null, nextUntitled([], t), true)
      tabs.splice(tabs.indexOf(tab), 1)
      activate(fresh)
    } else {
      tabs.splice(at, 1)
      activate(next)
    }
    tab.editor.destroy()
    tab.host.remove()
    renderChrome()
  }

  // The window closes when no tab is dirty. Otherwise each dirty tab comes forward and asks in turn;
  // Cancel, or a save that does not finish, keeps the window. close() does not ask again.
  function requestClose(): boolean {
    if (dialog.open || closingWindow) return false
    if (!tabs.some(tabDirty)) return true
    void closeWindow()
    return false
  }

  async function closeWindow(): Promise<void> {
    closingWindow = true
    try {
      for (const tab of [...tabs]) {
        if (!tabs.includes(tab) || !tabDirty(tab)) continue
        activate(tab)
        const choice = await askUnsaved(tab)
        if (choice === 'cancel') {
          current().focus()
          return
        }
        if (choice === 'save' && !(await saveTab(tab))) {
          current().focus()
          return
        }
      }
    } finally {
      closingWindow = false
    }
    appWindow.close()
  }

  function enqueueSave(tab: Tab, asNew: boolean): Promise<boolean> {
    const run = tail.then(() => performSave(tab, asNew))
    tail = run.then(() => undefined, () => undefined)
    return run
  }

  async function drain(): Promise<void> {
    let seen = tail
    for (;;) {
      await seen
      if (tail === seen) return
      seen = tail
    }
  }

  async function saveFromToolbar(asNew: boolean): Promise<void> {
    const tab = activeTab()
    const ok = await enqueueSave(tab, asNew)
    if (ok && tab === active) current().focus()
  }

  // The tab is fixed when the save is queued, so choosing another tab while the picker is up still saves this one.
  async function performSave(tab: Tab, asNew: boolean): Promise<boolean> {
    if (!tabs.includes(tab)) return false
    const currentEditor = tab.editor
    const doc = currentEditor.state.doc
    const open = tab.file
    if (!asNew && open.handle && doc === open.saved) return true
    let bytes: Uint8Array
    try {
      bytes = currentEditor.source.save(doc)
    } catch (error) {
      showSaveError(error)
      renderChrome()
      return false
    }
    const name = open.name
    let handle = !asNew && open.handle ? open.handle : null
    if (handle) {
      const wrote = await tryWrite(handle, bytes)
      if (!wrote) handle = null
    }
    if (!handle) {
      let picked: FileSystemFileHandle | null
      try {
        picked = await pickSave(name)
      } catch (error) {
        showSaveError(error)
        renderChrome()
        return false
      }
      if (!picked) return false
      try {
        await writeBytes(picked, bytes)
      } catch (error) {
        if (isAbort(error)) return false
        showSaveError(error)
        renderChrome()
        return false
      }
      handle = picked
    }
    const writtenName = handle.name || name
    tab.file = { handle, name: writtenName, saved: doc }
    clearError()
    renderChrome()
    return true
  }

  async function openHandle(handle: FileSystemFileHandle): Promise<void> {
    // A file that is already open comes forward instead of opening twice.
    for (const tab of tabs) {
      const open = tab.file.handle
      if (open && (open === handle || await sameEntry(open, handle))) {
        activate(tab)
        return
      }
    }
    let bytes: Uint8Array
    let name: string
    try {
      const blob = await handle.getFile()
      bytes = new Uint8Array(await blob.arrayBuffer())
      name = blob.name || handle.name
    } catch {
      // The system can refuse the read after the pick, for example a folder the runtime has no permission for.
      showOpenError('unreadable')
      return
    }
    try {
      openInto(engine.open(bytes), handle, name)
    } catch (error) {
      if (error instanceof OpenError) {
        showOpenError(error.reason)
        return
      }
      throw error
    }
  }

  async function insertPicture(): Promise<void> {
    try {
      const picked = await window.showOpenFilePicker({
        multiple: false,
        types: [{
          description: t('file.image'),
          accept: { 'image/png': ['.png'], 'image/jpeg': ['.jpg', '.jpeg'] },
        }],
      })
      const handle = picked[0]
      if (!handle) return
      const blob = await handle.getFile()
      const target = current()
      const media = await target.source.addMedia(blob)
      if (target !== current()) return
      edit(() => {
        current().run(insertImage({
          key: media.key,
          contentType: media.contentType,
          widthPx: media.widthPt / 0.75,
          heightPx: media.heightPt / 0.75,
        }))
      })
      clearError()
    } catch (error) {
      if (isAbort(error)) return
      showBand(t('image.error'))
    } finally {
      current().focus()
    }
  }

  async function openWithPicker(): Promise<void> {
    try {
      const picked = await window.showOpenFilePicker({
        multiple: false,
        types: [{ description: t('file.docx'), accept: { [DOCX_MIME]: ['.docx'] } }],
      })
      const handle = picked[0]
      if (!handle) return
      await openHandle(handle)
    } catch (error) {
      if (!isAbort(error)) throw error
    } finally {
      current().focus()
    }
  }

  function openBytes(bytes: Uint8Array, name?: string): Promise<void> {
    // The test hook replaces the active tab's document at once, without asking.
    return new Promise((resolve, reject) => {
      try {
        adopt(activeTab(), engine.open(bytes), null, name && name.length > 0 ? name : t('openedName'))
      } catch (error) {
        if (error instanceof OpenError) showOpenError(error.reason)
        reject(error)
        return
      }
      void whenLaidOut(current()).then(resolve, reject)
    })
  }

  function timedSave(): { base64: string; ms: number } {
    const doc = current().state.doc
    const started = performance.now()
    const bytes = current().source.save(doc)
    const ms = performance.now() - started
    return { base64: encodeBase64(bytes), ms }
  }

  function failSave(): string {
    try {
      saveTampered(current().source, current().state.doc)
    } catch (error) {
      showSaveError(error)
      renderChrome()
      return saveDetail(error)
    }
    return ''
  }

  activate(createTab(engine.blank(), null, nextUntitled([], t), true))
  findBar = createFindBar(() => current(), pages)
  // A file the user opened with this app arrives as a system request. The runtime holds it until a handler exists.
  appWindow.on('request', request => {
    const handle = fileHandleFromRequest(request)
    if (!handle) return
    void openHandle(handle).then(() => current().focus())
  })
  // Return false at once. The runtime closes a dirty window if beforeclose is still pending after 3 seconds.
  // close() does not ask again, so Save and Don't save close the window themselves.
  appWindow.on('beforeclose', () => requestClose())

  function concealEditing(): void {
    toolbar.setCovered(true)
    tabStrip.setCovered(true)
    pages.hidden = true
  }

  function revealEditing(scroll: { top: number; left: number }): void {
    preview.hide()
    toolbar.setCovered(false)
    tabStrip.setCovered(false)
    pages.hidden = false
    const view = scroller()
    const restore = (): void => {
      const phase = machine.phase
      if (phase !== 'closed' && phase !== 'unusable' && phase !== 'failed') return
      view.scrollTop = scroll.top
      view.scrollLeft = scroll.left
    }
    restore()
    // The page view may relayout on the frame it becomes visible and clamp the position.
    requestAnimationFrame(() => {
      restore()
      requestAnimationFrame(restore)
    })
    current().focus()
  }

  function requestPreviewClose(): void {
    const phase = machine.phase
    if (phase === 'preparing') {
      machine.cancel()
      preview.markCancelling()
      return
    }
    if (phase !== 'open' && phase !== 'exporting') return
    const scroll = heldScroll
    try {
      preview.hide()
      if (machine.close() === 'unusable') showBand(t('print.reopen'))
    } catch (error) {
      const failure = describePrintFailure(error, t)
      if (failure.abort) machine.abort()
      showBand(failure.message)
    } finally {
      revealEditing(scroll)
    }
  }

  function failPreview(error: unknown): void {
    const failure = describePrintFailure(error, t)
    preview.hide()
    if (failure.abort) {
      machine.abort()
      revealEditing(heldScroll)
      showBand(failure.message)
      return
    }
    if (machine.phase === 'cancelling') {
      const phase = machine.finishCancel()
      revealEditing(heldScroll)
      if (phase === 'unusable') showBand(t('print.reopen'))
      return
    }
    if (machine.phase !== 'closed' && machine.phase !== 'unusable') {
      const phase = machine.fail()
      if (phase === 'unusable') {
        revealEditing(heldScroll)
        showBand(t('print.reopen'))
        return
      }
      machine.dismiss()
    }
    revealEditing(heldScroll)
    showBand(failure.message)
  }

  async function openPreview(): Promise<void> {
    if (isDialogOpen()) return
    if (machine.phase === 'unusable') {
      showBand(t('print.reopen'))
      return
    }
    if (!machine.begin()) return
    findBar?.close(false)
    const editorNow = current()
    const doc = editorNow.state.doc
    heldScroll = { top: scroller().scrollTop, left: scroller().scrollLeft }
    concealEditing()
    preview.showPreparing()
    clearError()
    try {
      // The status line has to paint before printLayout blocks the thread.
      await paintFrame()
      await paintFrame()
      if (machine.phase !== 'preparing') {
        machine.finishCancel()
        revealEditing(heldScroll)
        return
      }
      await ensurePrintFonts(editorNow.source, editorNow.source.fontsUsed(doc), documentHasHan(doc))
      if (machine.phase !== 'preparing') {
        machine.finishCancel()
        revealEditing(heldScroll)
        return
      }
      const layout = editorNow.source.printLayout(doc)
      // The click is queued during the blocking call. This timer is queued after it, so settle sees the cancel.
      await new Promise<void>((resolve) => { setTimeout(resolve, 0) })
      const settled = machine.settle(layout)
      if (settled !== 'open') {
        revealEditing(heldScroll)
        if (settled === 'unusable') showBand(t('print.reopen'))
        return
      }
      preview.showPages(layout, previewDpi(window.devicePixelRatio), {
        pages: layout.pageCount(),
        fonts: layout.fonts(),
      })
    } catch (error) {
      failPreview(error)
    }
  }

  async function exportPdf(): Promise<void> {
    if (!machine.startExport()) return
    preview.setExporting(true)
    const layout = machine.held()
    if (layout === null) {
      machine.finishExport()
      preview.setExporting(false)
      return
    }
    let bytes: Uint8Array
    try {
      bytes = layout.pdf()
    } catch (error) {
      failPreview(error)
      return
    }
    const suggested = pdfFileName(activeTab().file.name)
    try {
      const picked = await pickPdf(suggested)
      if (machine.phase !== 'exporting') return
      if (!picked) {
        machine.finishExport()
        preview.setExporting(false)
        return
      }
      await writeBytes(picked, bytes)
      if (machine.phase !== 'exporting') return
      machine.finishExport()
      preview.setExporting(false)
    } catch (error) {
      if (machine.phase === 'exporting') {
        machine.finishExport()
        preview.setExporting(false)
      }
      if (isAbort(error)) return
      const detail = error instanceof Error ? error.message : String(error)
      showBand(t('print.error.write', { detail }))
    }
  }

  async function measurePrint(): Promise<PrintMeasure> {
    const editorNow = current()
    const doc = editorNow.state.doc
    const baseline = wasmMemoryBytes()
    const fontStarted = performance.now()
    await ensurePrintFonts(editorNow.source, editorNow.source.fontsUsed(doc), documentHasHan(doc))
    const fontMs = performance.now() - fontStarted
    const afterFonts = wasmMemoryBytes()
    const faces = sessionFaces(editorNow.source)
    let faceBytes = 0
    for (const face of faces) faceBytes += face.bytes
    const layoutStarted = performance.now()
    const layout = editorNow.source.printLayout(doc)
    const layoutMs = performance.now() - layoutStarted
    const afterLayout = wasmMemoryBytes()
    try {
      const pdfStarted = performance.now()
      const pdf = layout.pdf()
      const pdfMs = performance.now() - pdfStarted
      const afterPdf = wasmMemoryBytes()
      const pngStarted = performance.now()
      const png = layout.pagePng(0, 144)
      const pngMs = performance.now() - pngStarted
      const afterPng = wasmMemoryBytes()
      return {
        pages: layout.pageCount(),
        pdfBase64: encodeBase64(pdf),
        fontMs,
        layoutMs,
        pdfMs,
        pngMs,
        pngBytes: png.byteLength,
        memoryBefore: baseline,
        memoryAfter: afterPng,
        memory: { baseline, afterFonts, afterLayout, afterPdf, afterPng },
        faceBytes,
        faces,
        bootMs: engineBootMs(),
        fonts: layout.fonts(),
      }
    } finally {
      layout.free()
    }
  }

  function measuredBlockHeights(): readonly BlockHeight[] {
    return current().blockHeights().map((row) => ({
      index: row.index,
      kind: heightKind(row.kind),
      heightPt: row.heightPt,
    }))
  }

  async function measurePrintHeights(): Promise<readonly PrintBlockHeight[]> {
    await whenLaidOut(current())
    const editorNow = current()
    const doc = editorNow.state.doc
    await ensurePrintFonts(editorNow.source, editorNow.source.fontsUsed(doc), documentHasHan(doc))
    const layout = editorNow.source.printLayout(doc)
    try {
      return layout.blockBoxes().map((box) => {
        let heightPt: number | null = null
        if (box.fragments.length > 0) {
          heightPt = 0
          for (const fragment of box.fragments) heightPt += fragment.height
        }
        return { index: box.index, heightPt }
      })
    } finally {
      layout.free()
    }
  }

  async function measureDrift(): Promise<DriftReport> {
    await whenLaidOut(current())
    const editorNow = current()
    const doc = editorNow.state.doc
    await ensurePrintFonts(editorNow.source, editorNow.source.fontsUsed(doc), documentHasHan(doc))
    const layout = editorNow.source.printLayout(doc)
    try {
      // `blocks` publishes a 1-based page. `pageOf` is the 0-based index the
      // print layout uses, so a matching block counts as the same page.
      const edit = doc.body.map((block) => editorNow.pageOf(block.id))
      const report = drift(edit, layout.blockPages())
      return {
        ...report,
        editPages: editorNow.status().pages,
        printPages: layout.pageCount(),
      }
    } finally {
      layout.free()
    }
  }

  const wordHost: Host = {
    state() {
      return editorNow().state
    },
    run(command) {
      return editorNow().run(command)
    },
    undo() {
      return editorNow().undo()
    },
    redo() {
      return editorNow().redo()
    },
    pageOf(id) {
      return editorNow().pageOf(id)
    },
    status() {
      return editorNow().status()
    },
    busy() {
      // A dialog covers the editor. A preview that is preparing, shown, exporting, or
      // cancelling does too. failed and unusable have put the editor back on screen.
      if (isDialogOpen()) return 'dialog'
      const phase = machine.phase
      if (phase === 'preparing' || phase === 'cancelling' || phase === 'open' || phase === 'exporting') return 'preview'
      if (editorNow().composing) return 'composition'
      return null
    },
    async addMedia(blob) {
      // The rejection reaches word.call, which answers refused. Do not show the error band.
      const media = await editorNow().source.addMedia(blob)
      return {
        key: media.key,
        contentType: media.contentType,
        widthPx: media.widthPt / 0.75,
        heightPx: media.heightPt / 0.75,
      }
    },
    hitTest(clientX, clientY) {
      return editorNow().positionAtPoint(clientX, clientY)
    },
    file() {
      const open = activeTab().file
      const live = editorNow()
      return {
        name: open.name,
        dirty: isDirty(live.state.doc, open),
        session: live.source,
      }
    },
    markers() {
      const doc = editorNow().state.doc
      const cached = markersByDoc.get(doc)
      if (cached) return cached
      const listed = listMarkers(doc)
      markersByDoc.set(doc, listed)
      return listed
    },
    watch(listener) {
      changeListeners.add(listener)
      rebindWatch()
      return () => {
        changeListeners.delete(listener)
        if (changeListeners.size === 0) {
          unwatchEditor?.()
          unwatchEditor = null
          watchedEditor = null
        }
      }
    },
  }
  const word = createInterface(wordHost)

  return {
    get editor() {
      return current()
    },
    word,
    openBytes,
    save: timedSave,
    saveTampered: failSave,
    requestClose,
    pageSetupMs: () => lastPageSetupMs,
    memoryBytes: () => wasmMemoryBytes(),
    printLayout: measurePrint,
    drift: measureDrift,
    blockHeights: measuredBlockHeights,
    printHeights: measurePrintHeights,
  }
}

function required(selector: string): HTMLElement {
  const el = document.querySelector(selector)
  if (!(el instanceof HTMLElement)) throw new Error(`missing ${selector}`)
  return el
}

/** The two handles name one file. A runtime without isSameEntry treats different handle objects as different files. */
async function sameEntry(a: FileSystemFileHandle, b: FileSystemFileHandle): Promise<boolean> {
  if (typeof a.isSameEntry !== 'function') return false
  try {
    return await a.isSameEntry(b)
  } catch {
    return false
  }
}

async function tryWrite(handle: FileSystemFileHandle, bytes: Uint8Array): Promise<boolean> {
  try {
    await writeBytes(handle, bytes)
    return true
  } catch {
    return false
  }
}

async function writeBytes(handle: FileSystemFileHandle, bytes: Uint8Array): Promise<void> {
  const writable = await handle.createWritable()
  try {
    const copy = new Uint8Array(bytes.byteLength)
    copy.set(bytes)
    await writable.write(copy)
    await writable.close()
  } catch (error) {
    await writable.abort().catch(() => undefined)
    throw error
  }
}

function paintFrame(): Promise<void> {
  return new Promise((resolve) => { requestAnimationFrame(() => resolve()) })
}

async function pickPdf(name: string): Promise<FileSystemFileHandle | null> {
  try {
    return await window.showSaveFilePicker({
      suggestedName: name,
      types: [{ description: t('file.pdf'), accept: { [PDF_MIME]: ['.pdf'] } }],
    })
  } catch (error) {
    if (isAbort(error)) return null
    throw error
  }
}

async function pickSave(name: string): Promise<FileSystemFileHandle | null> {
  try {
    return await window.showSaveFilePicker({
      suggestedName: name,
      types: [{ description: t('file.docx'), accept: { [DOCX_MIME]: ['.docx'] } }],
    })
  } catch (error) {
    if (isAbort(error)) return null
    throw error
  }
}

function saveDetail(error: unknown): string {
  if (error instanceof SaveVerifyError) return error.firstDifference
  if (error instanceof SaveWriteError) return error.detail
  if (error instanceof Error) return error.message
  return String(error)
}

function showSaveError(error: unknown): void {
  showBand(t('save.error', { detail: saveDetail(error) }))
}

function showOpenError(reason: OpenError['reason'] | 'unreadable'): void {
  showBand(t(`open.error.${reason}`))
}

function showBand(message: string): void {
  const el = document.querySelector('#open-error')
  if (!(el instanceof HTMLElement)) return
  if (el.textContent !== message) el.textContent = message
  if (el.hidden) el.hidden = false
}

function clearError(): void {
  const el = document.querySelector('#open-error')
  if (!(el instanceof HTMLElement)) return
  if (el.textContent) el.textContent = ''
  if (!el.hidden) el.hidden = true
}

function encodeBase64(bytes: Uint8Array): string {
  let binary = ''
  const chunk = 0x2000
  for (let i = 0; i < bytes.length; i += chunk) {
    const piece = bytes.subarray(i, i + chunk)
    binary += String.fromCharCode(...piece)
  }
  return btoa(binary)
}

function whenLaidOut(editor: Editor): Promise<void> {
  return new Promise(resolve => {
    const tick = (): void => {
      if (editor.status().pagesFinal) resolve()
      else requestAnimationFrame(tick)
    }
    requestAnimationFrame(tick)
  })
}
