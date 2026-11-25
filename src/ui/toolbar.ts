// Toolbar controls. The buttons are created once; a later change only writes an attribute when its value differs.
import type { ToggleMark, ToolbarState } from '../editor/edit.ts'
import type { Platform, ShortcutName } from '../editor/input.ts'
import { shortcutChord } from '../editor/input.ts'
import { installedFamilies } from '../fonts.ts'
import { t } from '../i18n.ts'
import type { Doc, HighlightName } from '../model.ts'
import { createFormatChrome } from './format-menu.ts'
import { icon } from './icons.ts'
import type { IconName } from './icons.ts'
import { createFileMenu } from './menu.ts'
import type { FileMenuCommand, FileMenuId } from './menu.ts'
import { commandTitle, fontOptions, sizeOptions, styleOptions } from './text.ts'
import type { StyleOption } from './text.ts'

export type Align = 'left' | 'center' | 'right' | 'justify'

export interface ToolbarHandlers {
  onNew(): void
  onOpen(): void
  onSave(): void
  onSaveAs(): void
  onUndo(): void
  onRedo(): void
  onMark(mark: ToggleMark): void
  onFont(family: string): void
  onSize(sizePt: number): void
  onStyle(id: string): void
  onAlign(align: Align): void
  onList(kind: 'bullet' | 'decimal'): void
  onIndent(delta: 1 | -1): void
  onColor(hex: string | undefined): void
  onHighlight(name: HighlightName | undefined): void
  onVertAlign(which: 'super' | 'sub'): void
  onClearMarks(): void
  onLineSpacing(multiple: number): void
  onParaSpacing(edge: 'before' | 'after', points: number): void
  onParagraph(): void
  onPageSetup(): void
  onPrint(): void
  onFind(): void
  onInsertTable(): void
  onInsertImage(): void
  onTableRow(where: 'above' | 'below'): void
  onDeleteRows(): void
  onTableColumn(where: 'left' | 'right'): void
  onDeleteColumns(): void
  onDeleteTable(): void
}

export interface ToolbarView {
  render(state: ToolbarState, doc: Doc, usedFonts: readonly string[], canSave: boolean): void
  /** Preview covers the editing chrome. The table strip hides without moving the document. */
  setCovered(covered: boolean): void
}

const ALIGNS: readonly Align[] = ['left', 'center', 'right', 'justify']

export function createToolbar(
  root: HTMLElement,
  platform: Platform,
  handlers: ToolbarHandlers,
  /** The active tab's page scroller, which the table strip keeps still when it shows or hides. */
  scrollerNow: () => HTMLElement,
): ToolbarView {
  root.setAttribute('aria-label', t('toolbar'))
  let painting = false
  let covered = false
  let tableOpen = false
  let styleKey = ''
  let fontKey = ''
  let sizeKey = ''
  let seenStyles: object | null = null
  let cachedOptions: readonly StyleOption[] = []
  let installed: readonly string[] = []
  let usedFonts: readonly string[] = []
  let fontNow: string | 'mixed' = 'mixed'
  let fontReady = false
  let fontListToken = 0
  const styleGuard = { n: 0, ignore: false }
  const fontGuard = { n: 0, ignore: false }
  const sizeGuard = { n: 0, ignore: false }

  const fileMenu = createFileMenu(root, platform, fileCommands(handlers))
  const saveButton = iconButton(t('save'), 'save', platform, () => handlers.onSave(), 'save')
  group(root).append(fileMenu.button, saveButton)

  const undo = iconButton(t('undo'), 'undo', platform, () => handlers.onUndo(), 'undo')
  const redo = iconButton(t('redo'), 'redo', platform, () => handlers.onRedo(), 'redo')
  undo.disabled = true
  redo.disabled = true
  group(root).append(undo, redo)

  const fontSelect = labeledSelect(t('font'), 'font')
  watchSelect(fontSelect, () => {
    if (painting || fontGuard.ignore) return
    const value = fontSelect.value
    if (value.length === 0) return
    handlers.onFont(value)
  })
  const sizeSelect = labeledSelect(t('size'), 'size')
  watchSelect(sizeSelect, () => {
    if (painting || sizeGuard.ignore) return
    const value = Number(sizeSelect.value)
    if (!Number.isFinite(value) || value <= 0) return
    handlers.onSize(value)
  })
  group(root).append(fontSelect, sizeSelect)

  const marks = {
    bold: toggleButton(t('bold'), 'bold', platform, () => handlers.onMark('bold'), 'bold'),
    italic: toggleButton(t('italic'), 'italic', platform, () => handlers.onMark('italic'), 'italic'),
    underline: toggleButton(t('underline'), 'underline', platform, () => handlers.onMark('underline'), 'underline'),
    strike: toggleButton(t('strike'), null, platform, () => handlers.onMark('strike'), 'strike'),
  }
  const format = createFormatChrome(platform, {
    onColor: (hex) => handlers.onColor(hex),
    onHighlight: (name) => handlers.onHighlight(name),
    onVertAlign: (which) => handlers.onVertAlign(which),
    onClear: () => handlers.onClearMarks(),
    onLineSpacing: (multiple) => handlers.onLineSpacing(multiple),
    onParaSpacing: (edge, points) => handlers.onParaSpacing(edge, points),
    onParagraph: () => handlers.onParagraph(),
  }, () => fileMenu.close(false))
  group(root).append(marks.bold, marks.italic, marks.underline, marks.strike, format.color, format.highlight, format.more)

  const bullet = toggleButton(t('list.bullet'), null, platform, () => handlers.onList('bullet'), 'list-bullet')
  const numbered = toggleButton(t('list.number'), null, platform, () => handlers.onList('decimal'), 'list-number')
  const decrease = iconButton(t('list.decrease'), null, platform, () => handlers.onIndent(-1), 'indent-decrease')
  const increase = iconButton(t('list.increase'), null, platform, () => handlers.onIndent(1), 'indent-increase')
  group(root).append(bullet, numbered, decrease, increase)

  const alignButtons = {
    left: toggleButton(t('align.left'), null, platform, () => handlers.onAlign('left'), 'align-left'),
    center: toggleButton(t('align.center'), null, platform, () => handlers.onAlign('center'), 'align-center'),
    right: toggleButton(t('align.right'), null, platform, () => handlers.onAlign('right'), 'align-right'),
    justify: toggleButton(t('align.justify'), null, platform, () => handlers.onAlign('justify'), 'align-justify'),
  }
  group(root).append(
    alignButtons.left,
    alignButtons.center,
    alignButtons.right,
    alignButtons.justify,
    format.spacing,
  )

  const styleSelect = labeledSelect(t('style'), 'style')
  watchSelect(styleSelect, () => {
    if (painting || styleGuard.ignore) return
    const value = styleSelect.value
    if (value.length === 0) return
    handlers.onStyle(value)
  })
  group(root).append(styleSelect)

  function refreshInstalled(): void {
    if (installed.length > 0) return
    const token = ++fontListToken
    void installedFamilies().then((names) => {
      if (token !== fontListToken) return
      installed = names
      if (!fontReady) return
      paintFont(usedFonts, fontNow)
    })
  }

  // Opening the menu reads the index again. A failed read is not kept; a success is already cached.
  refreshInstalled()
  fontSelect.addEventListener('mousedown', refreshInstalled)
  fontSelect.addEventListener('keydown', (event) => {
    if (event.repeat) return
    if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp' && event.key !== 'Enter' && event.key !== ' ') return
    refreshInstalled()
  })

  const insertTable = iconButton(t('table.insert'), null, platform, () => handlers.onInsertTable(), 'table')
  const insertPicture = iconButton(t('image.insert'), null, platform, () => handlers.onInsertImage(), 'image')
  group(root).append(insertTable, insertPicture)
  group(root, true).append(iconButton(t('find'), 'find', platform, () => handlers.onFind(), 'search'))

  const strip = document.createElement('div')
  strip.id = 'tablebar'
  strip.className = 'tablebar'
  strip.hidden = true
  strip.setAttribute('role', 'toolbar')
  strip.setAttribute('aria-label', t('table.bar'))
  const caption = document.createElement('span')
  caption.className = 'caption'
  caption.setAttribute('aria-hidden', 'true')
  caption.textContent = t('table.bar')
  strip.append(caption)
  const rowAbove = toolButton(t('table.rowAbove'), platform, () => handlers.onTableRow('above'), 'table-row-above')
  const rowBelow = toolButton(t('table.rowBelow'), platform, () => handlers.onTableRow('below'), 'table-row-below')
  const deleteRow = toolButton(t('table.deleteRow'), platform, () => handlers.onDeleteRows(), 'table-row-delete')
  const columnLeft = toolButton(t('table.columnLeft'), platform, () => handlers.onTableColumn('left'), 'table-col-left')
  const columnRight = toolButton(t('table.columnRight'), platform, () => handlers.onTableColumn('right'), 'table-col-right')
  const deleteColumn = toolButton(t('table.deleteColumn'), platform, () => handlers.onDeleteColumns(), 'table-col-delete')
  const deleteTableButton = toolButton(t('table.delete'), platform, () => handlers.onDeleteTable(), 'table-delete')
  cluster(strip, false).append(rowAbove, rowBelow, deleteRow)
  cluster(strip, true).append(columnLeft, columnRight, deleteColumn)
  cluster(strip, true).append(deleteTableButton)
  for (const button of [rowAbove, rowBelow, deleteRow, columnLeft, columnRight, deleteColumn, deleteTableButton]) {
    button.disabled = true
  }
  root.insertAdjacentElement('afterend', strip)

  function optionsFor(doc: Doc): readonly StyleOption[] {
    const styles = doc.styles.paragraph
    if (styles === seenStyles) return cachedOptions
    seenStyles = styles
    cachedOptions = styleOptions(styles, role => t(`style.${role}`))
    return cachedOptions
  }

  function paintStyles(doc: Doc, current: string | 'mixed'): void {
    let options = optionsFor(doc)
    if (current !== 'mixed' && !options.some(option => option.id === current)) {
      options = options.concat({ id: current, label: current })
    }
    const key = options.map(option => `${option.id}\t${option.label}`).join('\n')
    if (key !== styleKey) {
      styleKey = key
      replaceOptions(styleSelect, options.map((option) => ({ value: option.id, label: option.label })), styleGuard)
    }
    const value = current === 'mixed' ? '' : current
    if (styleSelect.value !== value) styleSelect.value = value
  }

  function paintFont(used: readonly string[], current: string | 'mixed'): void {
    usedFonts = used
    fontNow = current
    let options = fontOptions(used, installed, current)
    // The helper folds case. The select value has to be the toolbar's spelling, so keep that string too.
    if (current !== 'mixed' && current.trim() !== '' && !options.includes(current)) options = options.concat(current)
    const key = JSON.stringify(options)
    if (key !== fontKey) {
      fontKey = key
      replaceOptions(fontSelect, options.map((name) => ({ value: name, label: name })), fontGuard)
    }
    const value = current === 'mixed' ? '' : current
    if (fontSelect.value !== value) fontSelect.value = value
  }

  function paintSize(current: ToolbarState['sizePt']): void {
    const options = sizeOptions(current)
    const key = JSON.stringify(options)
    if (key !== sizeKey) {
      sizeKey = key
      replaceOptions(
        sizeSelect,
        options.map((size) => ({ value: String(size), label: String(size) })),
        sizeGuard,
      )
    }
    const value = current === 'mixed' ? '' : String(current)
    if (sizeSelect.value !== value) sizeSelect.value = value
  }

  // The strip pushes #pages down. Match scrollTop to that shift in the same turn, or the
  // text moves under the pointer. PageView watches scrollTop to mount pages; it does not
  // write scrollTop back. overflow-anchor is off only for the write so it cannot add a second shift.
  function syncStrip(inTable: boolean, shift: boolean): void {
    const show = inTable && !covered
    const showing = !strip.hidden
    tableOpen = inTable
    if (show === showing) {
      root.classList.toggle('with-table', show)
      return
    }
    const scroller = scrollerNow()
    const previous = scroller.style.overflowAnchor
    if (shift) scroller.style.overflowAnchor = 'none'
    try {
      const before = shift ? scroller.getBoundingClientRect().top : 0
      strip.hidden = !show
      root.classList.toggle('with-table', show)
      if (!shift) return
      const delta = scroller.getBoundingClientRect().top - before
      if (Math.abs(delta) >= 0.5) scroller.scrollTop += delta
    } finally {
      if (shift) scroller.style.overflowAnchor = previous
    }
  }

  return {
    setCovered(next) {
      if (covered === next) return
      covered = next
      root.hidden = next
      fileMenu.close(false)
      format.close(false)
      syncStrip(tableOpen, false)
    },
    render(state, doc, used, canSave) {
      painting = true
      fontReady = true
      try {
        fileMenu.setDisabled('save', !canSave)
        if (saveButton.disabled !== !canSave) saveButton.disabled = !canSave
        if (undo.disabled !== !state.canUndo) undo.disabled = !state.canUndo
        if (redo.disabled !== !state.canRedo) redo.disabled = !state.canRedo
        // A table nested in a cell is refused on save, so insert stays off while the caret is in one.
        setDisabled(insertTable, state.inTable)
        setDisabled(rowAbove, !state.rowTools)
        setDisabled(rowBelow, !state.rowTools)
        setDisabled(deleteRow, !state.rowTools)
        setDisabled(columnLeft, !state.columnTools)
        setDisabled(columnRight, !state.columnTools)
        setDisabled(deleteColumn, !state.columnTools)
        setDisabled(deleteTableButton, !state.inTable)
        setPressed(marks.bold, pressedOf(state.bold))
        setPressed(marks.italic, pressedOf(state.italic))
        setPressed(marks.underline, pressedOf(state.underline))
        setPressed(marks.strike, pressedOf(state.strike))
        for (const align of ALIGNS) {
          const pressed = state.align === 'mixed' ? 'mixed' : state.align === align ? 'true' : 'false'
          setPressed(alignButtons[align], pressed)
        }
        setPressed(bullet, state.list === 'mixed' ? 'mixed' : state.list === 'bullet' ? 'true' : 'false')
        setPressed(numbered, state.list === 'mixed' ? 'mixed' : state.list === 'decimal' ? 'true' : 'false')
        paintStyles(doc, state.style)
        paintFont(used, state.font)
        paintSize(state.sizePt)
        format.sync(state)
        syncStrip(state.inTable, !covered)
      } finally {
        painting = false
      }
    },
  }
}

function fileCommands(handlers: ToolbarHandlers): Readonly<Record<FileMenuId, FileMenuCommand>> {
  // shortcutChord is the only source of the chord a menu row shows.
  return {
    new: { label: t('new'), icon: 'doc', shortcut: 'new', showTitle: false, run: () => handlers.onNew() },
    open: { label: t('open'), icon: 'folder', shortcut: 'open', showTitle: false, run: () => handlers.onOpen() },
    save: { label: t('save'), icon: 'save', shortcut: 'save', showTitle: true, run: () => handlers.onSave() },
    saveAs: { label: t('saveAs'), icon: 'save-as', shortcut: 'saveAs', showTitle: true, run: () => handlers.onSaveAs() },
    pageSetup: { label: t('pageSetup'), icon: 'page-setup', shortcut: null, showTitle: true, run: () => handlers.onPageSetup() },
    print: { label: t('print.preview'), icon: 'print', shortcut: null, showTitle: false, run: () => handlers.onPrint() },
  }
}

function group(root: HTMLElement, end = false): HTMLElement {
  const el = document.createElement('div')
  el.className = end ? 'group end' : 'group'
  const divided = !end && root.querySelector('.group') !== null
  root.append(el)
  if (divided) el.append(divider())
  return el
}

function cluster(parent: HTMLElement, divided: boolean): HTMLElement {
  const el = document.createElement('div')
  el.className = 'group'
  if (divided) el.append(divider())
  parent.append(el)
  return el
}

function divider(): HTMLElement {
  const el = document.createElement('div')
  el.className = 'divider'
  el.setAttribute('aria-hidden', 'true')
  return el
}

function iconButton(
  label: string,
  shortcut: ShortcutName | null,
  platform: Platform,
  onClick: () => void,
  graphic: IconName,
): HTMLButtonElement {
  return commandButton(label, shortcut, platform, onClick, graphic, false)
}

function toolButton(
  label: string,
  platform: Platform,
  onClick: () => void,
  graphic: IconName,
): HTMLButtonElement {
  return commandButton(label, null, platform, onClick, graphic, true)
}

function commandButton(
  label: string,
  shortcut: ShortcutName | null,
  platform: Platform,
  onClick: () => void,
  graphic: IconName,
  withLabel: boolean,
): HTMLButtonElement {
  const el = document.createElement('button')
  el.type = 'button'
  el.setAttribute('aria-label', label)
  el.append(icon(graphic))
  if (withLabel) {
    const text = document.createElement('span')
    text.textContent = label
    el.append(text)
  }
  const chord = shortcut === null ? null : shortcutChord(platform, shortcut)
  el.title = commandTitle(label, chord, platform)
  el.addEventListener('click', onClick)
  return el
}

function toggleButton(
  label: string,
  shortcut: ShortcutName | null,
  platform: Platform,
  onClick: () => void,
  graphic: IconName,
): HTMLButtonElement {
  const el = iconButton(label, shortcut, platform, onClick, graphic)
  el.setAttribute('aria-pressed', 'false')
  return el
}

function labeledSelect(label: string, className: string): HTMLSelectElement {
  const el = document.createElement('select')
  el.className = className
  el.setAttribute('aria-label', label)
  return el
}

function replaceOptions(
  select: HTMLSelectElement,
  entries: readonly { value: string, label: string }[],
  guard: { n: number, ignore: boolean },
): void {
  const fragment = document.createDocumentFragment()
  const mixed = document.createElement('option')
  mixed.value = ''
  mixed.textContent = t('style.mixed')
  mixed.hidden = true
  fragment.append(mixed)
  for (const entry of entries) {
    const option = document.createElement('option')
    option.value = entry.value
    option.textContent = entry.label
    fragment.append(option)
  }
  // Rebuilding the options can emit change. Ignore it until the next frame so a rebuild does not edit the document.
  guard.n += 1
  const token = guard.n
  guard.ignore = true
  select.replaceChildren(fragment)
  requestAnimationFrame(() => {
    if (guard.n === token) guard.ignore = false
  })
}

function watchSelect(select: HTMLSelectElement, apply: () => void): void {
  select.addEventListener('change', apply)
  select.addEventListener('input', apply)
}

function pressedOf(value: boolean | 'mixed'): 'true' | 'false' | 'mixed' {
  if (value === 'mixed') return 'mixed'
  return value ? 'true' : 'false'
}

function setPressed(button: HTMLButtonElement, pressed: 'true' | 'false' | 'mixed'): void {
  if (button.getAttribute('aria-pressed') !== pressed) button.setAttribute('aria-pressed', pressed)
}

function setDisabled(button: HTMLButtonElement, disabled: boolean): void {
  if (button.disabled !== disabled) button.disabled = disabled
}
