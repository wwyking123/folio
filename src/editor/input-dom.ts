// DOM wiring for the input layer. The decisions it applies are pure functions in input.ts, so Node can test them
// without a document.
import { comparePos, flatText, indexOf } from '../model.ts'
import type { ParaId, Pos } from '../model.ts'
import { t } from '../i18n.ts'
import { formatMargin, pageUnit } from '../ui/page-setup.ts'
import { compositionDelete, deleteSelection, findImage, imageDropTarget, imageMaxWidth, insertText, moveAtom, resizeImage, selectedText } from './edit.ts'
import type { Command, EditorState } from './edit.ts'
import { MIN_IMAGE_PT, MOVE_THRESHOLD_PX, resizeBox } from './image-geom.ts'
import type { ImageGesture, PageBox, ResizeHandle } from './image-geom.ts'
import { PT_TO_PX } from './render.ts'
import {
  chordOf,
  defaultKeymap,
  dragEndpoints,
  dragUnit,
  IDLE,
  isPictureFile,
  isVerticalMotion,
  pastePlan,
  PICTURE_TEXT,
  planBeforeInput,
  trackGoal,
  transition,
} from './input.ts'
import type { Action, DragUnit, FileOp, FindOp, InputMode, Platform } from './input.ts'
import type { Decoration } from './render.ts'

/** The slice of the facade the DOM layer is allowed to see. Page geometry stays behind these names. */
export interface EditingSurface {
  readonly state: EditorState
  readonly platform: Platform
  run(command: Command): boolean
  undo(): boolean
  redo(): boolean
  setPreedit(deco: Decoration | null): void
  /** Puts the caret at the end of the runtime's composition span. No-op when that span is absent. */
  syncCompositionCaret(): void
  focus(): void
  positionAtPoint(clientX: number, clientY: number): Pos | null
  caretColumn(): number | null
  moveLine(direction: 'up' | 'down', extend: boolean, goalX: number | null): number | null
  movePage(direction: 'up' | 'down', extend: boolean, goalX: number | null): number | null
  moveToLineEdge(edge: 'start' | 'end', extend: boolean): void
  setTextSelection(anchor: Pos, head: Pos): void
  fileCommand(op: FileOp): void
  findCommand(op: FindOp): void
  zoom(): number
  imageBox(pos: Pos): PageBox | null
  previewImage(preview: ImageGesture | null): void
  setNotice(text: string | null): void
  selectedPicture(): PictureInput | null
  insertPictures(pictures: readonly PictureInput[], at: Pos | null): void
  pictureError(): void
}

/** Picture bytes on their way into the document. A size is given when a copied picture keeps the size it had. */
export interface PictureInput {
  readonly blob: Blob
  readonly widthPt?: number
  readonly heightPt?: number
}

/**
 * The last picture this app copied. The runtime's clipboard carries text only, so the bytes wait here and the
 * clipboard holds a mark. It outlives a document, so a picture can be pasted into the next one.
 */
let heldPicture: (PictureInput & { readonly text: string }) | null = null

/** Sets what a paste of the app's own mark brings back. Returns the plain text for the clipboard. */
function holdClipboard(editor: EditingSurface): string {
  const picture = editor.selectedPicture()
  heldPicture = picture ? { ...picture, text: PICTURE_TEXT } : null
  return picture ? PICTURE_TEXT : selectedText(editor.state)
}

function pictureFilesOf(data: DataTransfer | null): File[] {
  return data ? [...data.files].filter(isPictureFile) : []
}

function carriesFiles(data: DataTransfer | null): boolean {
  return data !== null && [...data.types].includes('Files')
}

/** Pictures on the system clipboard, where the runtime can read clipboard items. Empty when it cannot. */
async function clipboardPictures(): Promise<Blob[]> {
  const clipboard = navigator.clipboard
  if (!clipboard || typeof clipboard.read !== 'function' || typeof ClipboardItem !== 'function') return []
  try {
    const out: Blob[] = []
    for (const item of await clipboard.read()) {
      const type = item.types.find((name) => name === 'image/png' || name === 'image/jpeg')
      if (type) out.push(await item.getType(type))
    }
    return out
  } catch {
    return []
  }
}

export interface InputBinding {
  /** True from compositionstart until compositionend, blur, or reset. */
  readonly composing: boolean
  /** True during a text drag or a picture move or resize. */
  readonly dragging: boolean
  detach(): void
  reset(): void
}

const EDGE = 28
const STEP = 32
const HANDLES: ReadonlySet<string> = new Set(['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'])

type PictureDrag =
  | { kind: 'pending', pos: Pos, x: number, y: number }
  | { kind: 'resize', pos: Pos, handle: ResizeHandle, x: number, y: number, box: PageBox, widthPt: number, heightPt: number }
  | { kind: 'move', from: Pos, x: number, y: number, active: boolean, file: boolean }

function handleOf(target: EventTarget | null): ResizeHandle | null {
  const element = target instanceof Element ? target : null
  const node = element?.closest('[data-handle]')
  if (!(node instanceof HTMLElement)) return null
  const name = node.dataset.handle
  if (name && HANDLES.has(name)) return name as ResizeHandle
  return null
}

function imageOf(target: EventTarget | null): Pos | null {
  const element = target instanceof Element ? target : null
  const img = element?.closest('img.atom[data-atom="image"]')
  if (!(img instanceof HTMLElement)) return null
  const block = img.closest('.blk.p')
  const id = block instanceof HTMLElement ? block.dataset.id : undefined
  const offset = Number(img.dataset.o)
  if (!id || !Number.isInteger(offset)) return null
  return { p: id as ParaId, o: offset }
}

/**
 * Wires one textarea to the editor. The textarea is the only focus target. Returns the detach function and a reset
 * for a new document, which abandons a composition or a drag without committing it.
 */
export function attachInput(editor: EditingSurface, sink: HTMLTextAreaElement, host: HTMLElement): InputBinding {
  const keymap = defaultKeymap(editor.platform)
  const controller = new AbortController()
  const signal = controller.signal
  let mode: InputMode = IDLE
  let goal: number | null = null
  let fromKey = false
  let skipInsert: string | null = null
  let pastePending = false
  let drag: { move: (event: MouseEvent) => void; up: (event: MouseEvent) => void } | null = null
  let picture: PictureDrag | null = null
  let pictureMove: ((event: MouseEvent) => void) | null = null
  let pictureUp: ((event: MouseEvent) => void) | null = null
  let scrollFrame = 0
  let pointerX = 0
  let pointerY = 0

  sink.setAttribute('aria-label', t('inputLabel'))

  function arm(flag: 'key' | 'insert', value: string | null): void {
    if (flag === 'key') fromKey = true
    else skipInsert = value
    queueMicrotask(() => {
      if (flag === 'key') fromKey = false
      else if (skipInsert === value) skipInsert = null
    })
  }

  function endDrag(): void {
    if (!drag) return
    document.removeEventListener('mousemove', drag.move, true)
    document.removeEventListener('mouseup', drag.up, true)
    drag = null
  }

  function applyAction(action: Action, column: number | null): void {
    switch (action.kind) {
      case 'command':
        editor.run(action.command)
        return
      case 'history':
        if (action.step === 'undo') editor.undo()
        else editor.redo()
        return
      case 'clipboard':
        applyClipboard(action.op)
        return
      case 'file':
        editor.fileCommand(action.op)
        return
      case 'find':
        editor.findCommand(action.op)
        return
      case 'visual':
        applyVisual(action, column)
        return
    }
  }

  function applyVisual(action: Extract<Action, { kind: 'visual' }>, column: number | null): void {
    const extend = action.extend
    switch (action.motion) {
      case 'lineUp':
        goal = editor.moveLine('up', extend, column)
        return
      case 'lineDown':
        goal = editor.moveLine('down', extend, column)
        return
      case 'pageUp':
        goal = editor.movePage('up', extend, column)
        return
      case 'pageDown':
        goal = editor.movePage('down', extend, column)
        return
      case 'lineStart':
        editor.moveToLineEdge('start', extend)
        return
      case 'lineEnd':
        editor.moveToLineEdge('end', extend)
        return
    }
  }

  function copySelection(): void {
    const text = holdClipboard(editor)
    try {
      // Fires a copy event while the key gesture is current, so onCopy can fill clipboardData.
      document.execCommand('copy')
    } catch {
      // execCommand is missing in some runtimes; writeText below is the other path.
    }
    const clipboard = navigator.clipboard
    if (!clipboard) return
    const blob = heldPicture?.blob
    if (blob && blob.type === 'image/png' && typeof ClipboardItem === 'function' && typeof clipboard.write === 'function') {
      // Other apps get the picture where the runtime can write clipboard items. The mark rides along for this app.
      const item = new ClipboardItem({ 'image/png': blob, 'text/plain': new Blob([text], { type: 'text/plain' }) })
      void clipboard.write([item]).catch(() => clipboard.writeText(text).catch(() => {}))
      return
    }
    void clipboard.writeText(text).catch(() => {})
  }

  function pasteFrom(text: string, files: readonly Blob[]): void {
    const plan = pastePlan(text, files.length, heldPicture)
    if (plan === 'none') return
    goal = null
    if (plan === 'files') editor.insertPictures(files.map((blob) => ({ blob })), null)
    else if (plan === 'held' && heldPicture) editor.insertPictures([heldPicture], null)
    else editor.run(insertText(text))
  }

  function applyClipboard(op: 'copy' | 'cut' | 'paste'): void {
    if (op === 'paste') {
      pastePending = true
      const clipboard = navigator.clipboard
      if (!clipboard) return
      void clipboardPictures()
        .then(async (files) => {
          const text = files.length > 0 ? '' : await clipboard.readText()
          if (!pastePending) return
          pastePending = false
          pasteFrom(text, files)
        })
        .catch(() => {
          pastePending = false
        })
      return
    }
    copySelection()
    if (op === 'cut') {
      goal = null
      editor.run(deleteSelection)
    }
  }

  function onKeyDown(event: KeyboardEvent): void {
    if (picture && event.key === 'Escape') {
      event.preventDefault()
      cancelPicture()
      return
    }
    if (picture) cancelPicture()
    if (event.isComposing || mode.kind === 'composing') return
    const action = keymap.get(chordOf(event, editor.platform))
    if (!action) return
    event.preventDefault()
    arm('key', null)
    const vertical = action.kind === 'visual' && isVerticalMotion(action.motion)
    const column = trackGoal(goal, action, vertical && goal === null ? editor.caretColumn() : goal)
    goal = column
    applyAction(action, column)
  }

  function onBeforeInput(event: InputEvent): void {
    if (picture) cancelPicture()
    if (fromKey) {
      event.preventDefault()
      return
    }
    const data = event.data
    if (
      skipInsert !== null &&
      data === skipInsert &&
      (event.inputType === 'insertText' || event.inputType === 'insertFromComposition')
    ) {
      event.preventDefault()
      skipInsert = null
      return
    }
    const plan = planBeforeInput(event.inputType, data, mode.kind === 'composing' || event.isComposing)
    if (plan.kind !== 'run') return
    event.preventDefault()
    goal = null
    editor.run(plan.command)
  }

  function onCompositionStart(): void {
    cancelPicture()
    endDrag()
    const sel = editor.state.sel
    if (sel.kind === 'text' && (sel.anchor.p !== sel.head.p || sel.anchor.o !== sel.head.o)) {
      editor.run(compositionDelete())
    }
    const next = editor.state.sel
    if (next.kind !== 'text') return
    mode = transition(mode, { kind: 'compositionStart', at: next.head })
    goal = null
  }

  function onCompositionUpdate(event: CompositionEvent): void {
    if (mode.kind !== 'composing') return
    // The runtime draws this string at the textarea, after this event returns. A paragraph preedit would show it twice.
    if (event.data.length === 0) {
      queueMicrotask(() => {
        if (mode.kind === 'composing') editor.setPreedit(null)
      })
      return
    }
    queueMicrotask(() => {
      if (mode.kind === 'composing') editor.syncCompositionCaret()
    })
  }

  function onCompositionEnd(event: CompositionEvent): void {
    const wasComposing = mode.kind === 'composing'
    const data = event.data.length > 0 ? event.data : sink.value
    mode = transition(mode, { kind: 'compositionEnd' })
    goal = null
    editor.setPreedit(null)
    sink.value = ''
    // A blur resets the mode first and abandons the preedit. A commit that still owns the mode inserts once.
    if (!wasComposing || data.length === 0) return
    arm('insert', data)
    editor.run(insertText(data))
  }

  function onCopy(event: ClipboardEvent): void {
    event.preventDefault()
    event.clipboardData?.setData('text/plain', holdClipboard(editor))
  }

  function onCut(event: ClipboardEvent): void {
    event.preventDefault()
    if (fromKey) return
    const text = holdClipboard(editor)
    event.clipboardData?.setData('text/plain', text)
    goal = null
    if (text.length > 0) editor.run(deleteSelection)
  }

  function onPaste(event: ClipboardEvent): void {
    event.preventDefault()
    const files = pictureFilesOf(event.clipboardData)
    const text = event.clipboardData?.getData('text/plain') ?? ''
    if (files.length === 0 && text.length === 0) return
    pastePending = false
    pasteFrom(text, files)
  }

  // Files dragged in from the system. The caret follows the pointer, so the drop lands where it shows.
  function onDragOver(event: DragEvent): void {
    if (!carriesFiles(event.dataTransfer) || mode.kind === 'composing') return
    event.preventDefault()
    if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy'
    const pos = editor.positionAtPoint(event.clientX, event.clientY)
    const sel = editor.state.sel
    if (!pos) return
    const there = sel.kind === 'text' && sel.anchor.p === pos.p && sel.anchor.o === pos.o
      && sel.head.p === pos.p && sel.head.o === pos.o
    if (!there) editor.setTextSelection(pos, pos)
  }

  function onDrop(event: DragEvent): void {
    if (!carriesFiles(event.dataTransfer) || mode.kind === 'composing') return
    event.preventDefault()
    cancelPicture()
    endDrag()
    const files = pictureFilesOf(event.dataTransfer)
    const dropped = event.dataTransfer?.files.length ?? 0
    editor.focus()
    goal = null
    if (files.length > 0) editor.insertPictures(files.map((blob) => ({ blob })), editor.positionAtPoint(event.clientX, event.clientY))
    if (files.length < dropped || dropped === 0) editor.pictureError()
  }

  function selectGesture(anchor: Pos, head: Pos, unit: DragUnit): void {
    try {
      const index = indexOf(editor.state.doc)
      const anchorLoc = index.get(anchor.p)
      const headLoc = index.get(head.p)
      if (anchorLoc.story !== 'body' || headLoc.story !== 'body') return
      const same = anchor.p === head.p
      const headBefore = !same && comparePos(index, head, anchor) < 0
      const offsets = dragEndpoints(
        flatText(anchorLoc.para),
        anchor.o,
        flatText(headLoc.para),
        head.o,
        same,
        headBefore,
        unit,
      )
      editor.setTextSelection({ p: anchor.p, o: offsets.anchor }, { p: head.p, o: offsets.head })
    } catch {
      return
    }
  }

  function autoscroll(clientX: number, clientY: number): void {
    const box = host.getBoundingClientRect()
    if (clientY < box.top + EDGE) host.scrollTop -= STEP
    else if (clientY > box.bottom - EDGE) host.scrollTop += STEP
    if (clientX < box.left + EDGE) host.scrollLeft -= STEP
    else if (clientX > box.right - EDGE) host.scrollLeft += STEP
  }

  function stopScroll(): void {
    if (scrollFrame) cancelAnimationFrame(scrollFrame)
    scrollFrame = 0
  }

  function armScroll(): void {
    if (scrollFrame) return
    const tick = (): void => {
      scrollFrame = 0
      if (!picture || picture.kind !== 'move' || !picture.active) return
      autoscroll(pointerX, pointerY)
      const box = host.getBoundingClientRect()
      const near = pointerY < box.top + EDGE || pointerY > box.bottom - EDGE
      if (near) scrollFrame = requestAnimationFrame(tick)
    }
    scrollFrame = requestAnimationFrame(tick)
  }

  function cancelPicture(): void {
    if (pictureMove) document.removeEventListener('mousemove', pictureMove, true)
    if (pictureUp) document.removeEventListener('mouseup', pictureUp, true)
    window.removeEventListener('blur', cancelPicture)
    document.removeEventListener('mouseout', onPictureLeave, true)
    pictureMove = null
    pictureUp = null
    picture = null
    stopScroll()
    editor.previewImage(null)
    editor.setNotice(null)
  }

  function onPictureLeave(event: MouseEvent): void {
    if (!picture || event.relatedTarget) return
    cancelPicture()
  }

  function listenPicture(): void {
    const move = (event: MouseEvent): void => onPictureMove(event)
    const up = (event: MouseEvent): void => onPictureUp(event)
    pictureMove = move
    pictureUp = up
    document.addEventListener('mousemove', move, true)
    document.addEventListener('mouseup', up, true)
    window.addEventListener('blur', cancelPicture)
    document.addEventListener('mouseout', onPictureLeave, true)
  }

  function sizeLabel(widthPt: number, heightPt: number): string {
    const unit = pageUnit()
    return `${formatMargin(widthPt, unit)} × ${formatMargin(heightPt, unit)} ${t('page.unit')}`
  }

  function resizeLimits(pos: Pos): { min: number, maxWidth: number } {
    return {
      min: MIN_IMAGE_PT * PT_TO_PX,
      maxWidth: imageMaxWidth(editor.state.doc, pos.p) * PT_TO_PX,
    }
  }

  function startResize(handle: ResizeHandle, event: MouseEvent): void {
    const sel = editor.state.sel
    if (sel.kind !== 'text') return
    const from = sel.anchor.o <= sel.head.o && sel.anchor.p === sel.head.p ? sel.anchor : sel.head
    const atom = findImage(editor.state.doc, from)
    const box = editor.imageBox(from)
    if (!atom || atom.atom.type !== 'image' || !box) return
    picture = {
      kind: 'resize',
      pos: from,
      handle,
      x: event.clientX,
      y: event.clientY,
      box,
      widthPt: atom.atom.widthPt,
      heightPt: atom.atom.heightPt,
    }
    listenPicture()
  }

  function startPicturePress(pos: Pos, event: MouseEvent): void {
    const sel = editor.state.sel
    const covered = sel.kind === 'text'
      && ((sel.anchor.p === pos.p && sel.anchor.o === pos.o && sel.head.p === pos.p && sel.head.o === pos.o + 1)
        || (sel.head.p === pos.p && sel.head.o === pos.o && sel.anchor.p === pos.p && sel.anchor.o === pos.o + 1))
    if (!covered) editor.setTextSelection(pos, { p: pos.p, o: pos.o + 1 })
    picture = { kind: 'pending', pos, x: event.clientX, y: event.clientY }
    listenPicture()
  }

  function onPictureMove(event: MouseEvent): void {
    if (!picture) return
    event.preventDefault()
    pointerX = event.clientX
    pointerY = event.clientY
    if (picture.kind === 'pending') {
      const dx = event.clientX - picture.x
      const dy = event.clientY - picture.y
      if (dx * dx + dy * dy <= MOVE_THRESHOLD_PX * MOVE_THRESHOLD_PX) return
      const atom = findImage(editor.state.doc, picture.pos)
      const file = atom?.src != null
      if (file) editor.setNotice(t('image.move.kept'))
      picture = { kind: 'move', from: picture.pos, x: picture.x, y: picture.y, active: true, file }
    }
    if (picture.kind === 'resize') {
      const next = resizeBox(
        picture.box,
        picture.handle,
        event.clientX - picture.x,
        event.clientY - picture.y,
        editor.zoom(),
        resizeLimits(picture.pos),
      )
      editor.previewImage({
        kind: 'resize',
        box: next,
        label: sizeLabel(next.width / PT_TO_PX, next.height / PT_TO_PX),
      })
      return
    }
    if (picture.kind === 'move') {
      const point = picture.file ? null : editor.positionAtPoint(event.clientX, event.clientY)
      const target = point && imageDropTarget(editor.state.doc, picture.from, point) ? point : null
      editor.previewImage({ kind: 'move', from: picture.from, target })
      armScroll()
    }
  }

  function onPictureUp(event: MouseEvent): void {
    const current = picture
    if (!current) return
    if (current.kind === 'resize') {
      const next = resizeBox(
        current.box,
        current.handle,
        event.clientX - current.x,
        event.clientY - current.y,
        editor.zoom(),
        resizeLimits(current.pos),
      )
      const widthPt = next.width / PT_TO_PX
      const heightPt = next.height / PT_TO_PX
      cancelPicture()
      const moved = Math.abs(next.width - current.box.width) > 0.05 * PT_TO_PX
        || Math.abs(next.height - current.box.height) > 0.05 * PT_TO_PX
      if (moved) editor.run(resizeImage(current.pos, widthPt, heightPt))
      return
    }
    if (current.kind === 'move' && current.active && !current.file) {
      const point = editor.positionAtPoint(event.clientX, event.clientY)
      const target = point && imageDropTarget(editor.state.doc, current.from, point) ? point : null
      cancelPicture()
      if (target) editor.run(moveAtom(current.from, target))
      return
    }
    cancelPicture()
  }

  function onMouseDown(event: MouseEvent): void {
    if (event.button !== 0 || mode.kind === 'composing') return
    if (event.metaKey || event.ctrlKey || event.altKey) return
    if (event.target === sink) return
    const target = event.target
    const element = target instanceof Element ? target : target instanceof Node ? target.parentElement : null
    if (!element?.closest('.page')) return
    const handle = handleOf(event.target)
    const image = event.shiftKey ? null : imageOf(event.target)
    if (handle || image) {
      event.preventDefault()
      editor.focus()
      endDrag()
      cancelPicture()
      goal = null
      if (handle) startResize(handle, event)
      else if (image) startPicturePress(image, event)
      return
    }
    const pos = editor.positionAtPoint(event.clientX, event.clientY)
    if (!pos) return
    event.preventDefault()
    editor.focus()
    const unit = dragUnit(event.detail)
    const sel = editor.state.sel
    const anchor = event.shiftKey && sel.kind === 'text' ? sel.anchor : pos
    endDrag()
    mode = transition(mode, { kind: 'pointerDown', at: anchor, unit })
    goal = null
    selectGesture(anchor, pos, unit)
    const move = (ev: MouseEvent): void => {
      if (mode.kind !== 'dragging') return
      ev.preventDefault()
      autoscroll(ev.clientX, ev.clientY)
      const next = editor.positionAtPoint(ev.clientX, ev.clientY)
      if (next) selectGesture(anchor, next, unit)
    }
    const up = (): void => {
      endDrag()
      mode = transition(mode, { kind: 'pointerUp' })
    }
    drag = { move, up }
    document.addEventListener('mousemove', move, true)
    document.addEventListener('mouseup', up, true)
  }

  function onBlur(): void {
    const composing = mode.kind === 'composing'
    endDrag()
    mode = transition(mode, { kind: 'reset' })
    goal = null
    sink.value = ''
    if (composing) editor.setPreedit(null)
  }

  function onScroll(): void {
    if (mode.kind === 'composing') editor.syncCompositionCaret()
  }

  sink.addEventListener('keydown', onKeyDown, { signal })
  sink.addEventListener('beforeinput', onBeforeInput, { signal })
  sink.addEventListener('compositionstart', onCompositionStart, { signal })
  sink.addEventListener('compositionupdate', onCompositionUpdate, { signal })
  sink.addEventListener('compositionend', onCompositionEnd, { signal })
  sink.addEventListener('copy', onCopy, { signal })
  sink.addEventListener('cut', onCut, { signal })
  sink.addEventListener('paste', onPaste, { signal })
  sink.addEventListener('blur', onBlur, { signal })
  host.addEventListener('mousedown', onMouseDown, { capture: true, signal })
  host.addEventListener('scroll', onScroll, { passive: true, signal })
  host.addEventListener('dragover', onDragOver, { signal })
  host.addEventListener('drop', onDrop, { signal })

  return {
    get composing() {
      return mode.kind === 'composing'
    },
    get dragging() {
      return mode.kind === 'dragging' || picture !== null
    },
    detach() {
      cancelPicture()
      endDrag()
      controller.abort()
      mode = IDLE
      goal = null
      pastePending = false
    },
    reset() {
      const composing = mode.kind === 'composing'
      cancelPicture()
      endDrag()
      mode = IDLE
      goal = null
      pastePending = false
      sink.value = ''
      if (composing) editor.setPreedit(null)
    },
  }
}
