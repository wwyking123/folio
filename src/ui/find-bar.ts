// A floating find and replace bar. It is not a dialog: the document stays editable, and isDialogOpen() stays false.
import { replaceAllText, replaceCounts, replaceMatch, setSelection } from '../editor/edit.ts'
import type { Editor } from '../editor/editor.ts'
import { findMatches, followingMatch, matchFrom, matchUntil } from '../editor/find.ts'
import type { FindOptions, TextMatch } from '../editor/find.ts'
import { chordOf } from '../editor/input.ts'
import type { FindOp } from '../editor/input.ts'
import { t } from '../i18n.ts'
import { comparePos, flatText, indexOf } from '../model.ts'
import type { Pos, Selection } from '../model.ts'
import { isDialogOpen } from './dialog.ts'
import { icon } from './icons.ts'

export interface FindBar {
  command(op: FindOp): void
  noteDocument(): void
  close(restoreFocus: boolean): void
  readonly shown: boolean
}

const ATOM = '\uFFFC'

/** `editor` reads the active tab's editor at each use, so the bar follows the tab on screen. */
export function createFindBar(editor: () => Editor, pages: HTMLElement): FindBar {
  const root = document.createElement('div')
  root.className = 'findbar'
  root.hidden = true
  root.setAttribute('role', 'search')
  root.setAttribute('aria-label', t('find'))

  const queryInput = field(t('find'), 'find-query')
  const replaceInput = field(t('find.with'), 'find-replace-text')
  const toggle = iconControl(t('find.replace'), 'find-toggle', 'chevron')
  toggle.setAttribute('aria-expanded', 'false')
  const previous = iconControl(t('find.previous'), 'find-prev', 'chevron')
  previous.classList.add('prev')
  const next = iconControl(t('find.next'), 'find-next', 'chevron')
  const matchCase = letterControl(t('find.case'), 'find-case', 'Aa', 'case')
  const wholeWord = letterControl(t('find.word'), 'find-word', 'ab', 'word')
  const closeButton = iconControl(t('find.close'), 'find-close', 'close')
  const replaceButton = textControl(t('find.replace'), 'find-replace')
  const replaceAllButton = textControl(t('find.all'), 'find-all')

  const count = document.createElement('span')
  count.className = 'count'
  count.setAttribute('role', 'status')

  const tools = document.createElement('div')
  tools.className = 'find-tools'
  tools.append(count, previous, next, matchCase, wholeWord, closeButton)
  const actions = document.createElement('div')
  actions.className = 'find-actions'
  actions.append(replaceButton, replaceAllButton)
  actions.hidden = true
  replaceInput.hidden = true

  root.append(queryInput, replaceInput, toggle, tools, actions)
  document.body.append(root)

  let shown = false
  let expanded = false
  let caseOn = false
  let wordOn = false
  let matches: TextMatch[] = []
  let current = -1
  let held: { replaced: number, skipped: number } | null = null
  let seenQuery = ''
  let seenCase = false
  let seenWord = false
  let seenDoc: object | null = null
  let relocate = false
  let frame = 0

  const observer = new ResizeObserver(() => place())
  observer.observe(pages)

  function place(): void {
    root.style.top = `${pages.offsetTop + 8}px`
  }

  function options(): FindOptions {
    return { matchCase: caseOn, wholeWord: wordOn }
  }

  function poke(): void {
    if (frame !== 0) return
    frame = requestAnimationFrame(() => {
      frame = 0
      refresh()
    })
  }

  function flush(): void {
    if (frame !== 0) {
      cancelAnimationFrame(frame)
      frame = 0
    }
    refresh()
  }

  function refresh(): void {
    if (!shown) return
    place()
    const query = queryInput.value
    const doc = editor().state.doc
    const queryChanged = query !== seenQuery || caseOn !== seenCase || wordOn !== seenWord
    const docChanged = doc !== seenDoc
    const jump = queryChanged || relocate
    relocate = false
    if (!queryChanged && !docChanged && !jump) return
    if (queryChanged) held = null
    seenQuery = query
    seenCase = caseOn
    seenWord = wordOn
    seenDoc = doc
    if (query.length === 0) {
      matches = []
      current = -1
      editor().setFindHits([], -1)
      paint()
      return
    }
    if (queryChanged || docChanged) matches = findMatches(doc, query, options())
    if (jump) land()
    else follow()
    paint()
  }

  function land(): void {
    if (matches.length === 0) {
      current = -1
      editor().setFindHits([], -1)
      return
    }
    const at = matchFrom(editor().state.doc, matches, earlier(editor().state.sel))
    choose(at < 0 ? 0 : at)
  }

  function follow(): void {
    const equal = selectionIndex()
    if (equal >= 0) current = equal
    else if (matches.length === 0) current = -1
    else {
      const sel = editor().state.sel
      const pos = sel.kind === 'text' ? sel.head : earlier(sel)
      const at = matchFrom(editor().state.doc, matches, pos)
      current = at < 0 ? 0 : at
    }
    editor().setFindHits(matches, current)
  }

  function choose(index: number): void {
    const match = matches[index]
    if (!match) return
    current = index
    editor().setFindHits(matches, current)
    editor().run(setSelection({
      kind: 'text',
      anchor: { p: match.para, o: match.start },
      head: { p: match.para, o: match.end },
    }))
  }

  function move(direction: 1 | -1): void {
    flush()
    if (matches.length === 0) return
    const sel = editor().state.sel
    let index = -1
    if (sel.kind === 'text') {
      const doc = editor().state.doc
      const start = earlier(sel)
      const end = later(sel)
      index = direction > 0 ? matchFrom(doc, matches, end) : matchUntil(doc, matches, start)
    } else if (direction > 0) {
      index = current < 0 ? 0 : (current + 1) % matches.length
    } else {
      index = current <= 0 ? matches.length - 1 : current - 1
    }
    if (index < 0) index = direction > 0 ? 0 : matches.length - 1
    choose(index)
    paint()
  }

  function selectionIs(match: TextMatch): boolean {
    const sel = editor().state.sel
    if (sel.kind !== 'text' || sel.anchor.p !== match.para || sel.head.p !== match.para) return false
    const start = Math.min(sel.anchor.o, sel.head.o)
    const end = Math.max(sel.anchor.o, sel.head.o)
    return start === match.start && end === match.end
  }

  function selectionIndex(): number {
    const sel = editor().state.sel
    if (sel.kind !== 'text') return -1
    for (let i = 0; i < matches.length; i++) {
      const match = matches[i]
      if (match && selectionIs(match)) return i
    }
    return -1
  }

  function replaceOne(): void {
    flush()
    const match = matches[current]
    if (current < 0 || !match) return
    if (!selectionIs(match)) {
      move(1)
      return
    }
    const replacement = replaceInput.value
    editor().run(replaceMatch(match, replacement, followingMatch(matches, current, replacement)))
  }

  function replaceEvery(): void {
    flush()
    if (matches.length === 0) return
    const replacement = replaceInput.value
    const counts = replaceCounts(editor().state.doc, matches, replacement)
    if (counts.replaced === 0 && counts.skipped === 0) return
    if (counts.replaced > 0) {
      editor().run(replaceAllText(queryInput.value, replacement, options()))
    }
    held = counts
    paint()
  }

  function reveal(expandReplace: boolean): void {
    const was = shown
    shown = true
    root.hidden = false
    if (expandReplace) setExpanded(true)
    if (!was) {
      const seed = seedText()
      if (seed !== null) queryInput.value = seed
      relocate = true
    }
    place()
    focusQuery()
    flush()
  }

  function close(restoreFocus: boolean): void {
    if (!shown) return
    shown = false
    root.hidden = true
    if (frame !== 0) {
      cancelAnimationFrame(frame)
      frame = 0
    }
    editor().setFindHits([], -1)
    if (restoreFocus) editor().focus()
  }

  function setExpanded(value: boolean): void {
    expanded = value
    toggle.setAttribute('aria-expanded', value ? 'true' : 'false')
    replaceInput.hidden = !value
    actions.hidden = !value
  }

  function focusQuery(): void {
    queryInput.focus()
    queryInput.select()
  }

  function paint(): void {
    const idle = matches.length === 0
    previous.disabled = idle
    next.disabled = idle
    replaceButton.disabled = idle
    replaceAllButton.disabled = idle
    if (held) {
      count.textContent = held.skipped > 0
        ? t('find.replacedSkipped', { count: held.replaced, skipped: held.skipped })
        : t('find.replaced', { count: held.replaced })
      count.classList.remove('muted')
      queryInput.removeAttribute('aria-invalid')
      return
    }
    if (queryInput.value.length === 0) {
      count.textContent = ''
      count.classList.remove('muted')
      queryInput.removeAttribute('aria-invalid')
      return
    }
    if (matches.length === 0) {
      count.textContent = t('find.none')
      count.classList.add('muted')
      queryInput.setAttribute('aria-invalid', 'true')
      return
    }
    count.classList.remove('muted')
    queryInput.removeAttribute('aria-invalid')
    const shownIndex = current >= 0 ? current + 1 : 1
    count.textContent = `${cap(shownIndex)}/${cap(matches.length)}`
  }

  function earlier(sel: Selection): Pos {
    if (sel.kind !== 'text') return documentStart()
    const index = indexOf(editor().state.doc)
    return comparePos(index, sel.anchor, sel.head) <= 0 ? sel.anchor : sel.head
  }

  function later(sel: Selection): Pos {
    if (sel.kind !== 'text') return documentStart()
    const index = indexOf(editor().state.doc)
    return comparePos(index, sel.anchor, sel.head) <= 0 ? sel.head : sel.anchor
  }

  function documentStart(): Pos {
    const id = indexOf(editor().state.doc).ordered[0]
    if (id) return { p: id, o: 0 }
    const sel = editor().state.sel
    if (sel.kind === 'text') return sel.head
    return { p: '' as Pos['p'], o: 0 }
  }

  function seedText(): string | null {
    const sel = editor().state.sel
    if (sel.kind !== 'text' || sel.anchor.p !== sel.head.p) return null
    let loc
    try {
      loc = indexOf(editor().state.doc).get(sel.anchor.p)
    } catch {
      return null
    }
    if (loc.story !== 'body') return null
    const start = Math.min(sel.anchor.o, sel.head.o)
    const end = Math.max(sel.anchor.o, sel.head.o)
    if (end === start || end - start > 200) return null
    const slice = flatText(loc.para).slice(start, end)
    if (slice.includes(ATOM)) return null
    return slice
  }

  function onKey(event: KeyboardEvent): void {
    if (event.key === 'Enter' && event.target === queryInput) {
      event.preventDefault()
      move(event.shiftKey ? -1 : 1)
      return
    }
    if (event.key === 'Enter' && event.target === replaceInput) {
      event.preventDefault()
      replaceOne()
      return
    }
    const chord = chordOf(event, editor().platform)
    if (chord === 'f3' || chord === 'mod+g') {
      event.preventDefault()
      move(1)
      return
    }
    if (chord === 'shift+f3' || chord === 'mod+shift+g') {
      event.preventDefault()
      move(-1)
      return
    }
    if (chord === 'mod+f') {
      event.preventDefault()
      focusQuery()
      return
    }
    const replaceChord = editor().platform === 'mac' ? 'ctrl+h' : 'mod+h'
    if (chord === replaceChord) {
      event.preventDefault()
      setExpanded(true)
      focusQuery()
    }
  }

  function onEscape(event: KeyboardEvent): void {
    if (!shown || event.key !== 'Escape') return
    if (isDialogOpen() || document.querySelector('.file-menu') || editor().dragging()) return
    event.preventDefault()
    event.stopPropagation()
    close(true)
  }

  toggle.addEventListener('click', () => setExpanded(!expanded))
  previous.addEventListener('click', () => move(-1))
  next.addEventListener('click', () => move(1))
  matchCase.addEventListener('click', () => {
    caseOn = !caseOn
    matchCase.setAttribute('aria-pressed', caseOn ? 'true' : 'false')
    poke()
  })
  wholeWord.addEventListener('click', () => {
    wordOn = !wordOn
    wholeWord.setAttribute('aria-pressed', wordOn ? 'true' : 'false')
    poke()
  })
  closeButton.addEventListener('click', () => close(true))
  replaceButton.addEventListener('click', () => replaceOne())
  replaceAllButton.addEventListener('click', () => replaceEvery())
  queryInput.addEventListener('input', () => poke())
  root.addEventListener('keydown', onKey)
  root.addEventListener('mousedown', (event) => {
    if (event.target instanceof HTMLInputElement) return
    event.preventDefault()
  })
  document.addEventListener('keydown', onEscape, true)

  return {
    command(op) {
      if (op === 'open') {
        reveal(false)
        return
      }
      if (op === 'replace') {
        reveal(true)
        return
      }
      if (!shown) {
        reveal(false)
        return
      }
      move(op === 'next' ? 1 : -1)
    },
    noteDocument() {
      if (shown) poke()
    },
    close,
    get shown() {
      return shown
    },
  }
}

function field(label: string, className: string): HTMLInputElement {
  const input = document.createElement('input')
  input.type = 'text'
  input.className = className
  input.autocomplete = 'off'
  input.spellcheck = false
  input.setAttribute('aria-label', label)
  input.placeholder = label
  return input
}

function iconControl(label: string, className: string, graphic: 'chevron' | 'close'): HTMLButtonElement {
  const button = document.createElement('button')
  button.type = 'button'
  button.className = `icon ${className}`
  button.setAttribute('aria-label', label)
  button.title = label
  button.append(icon(graphic))
  return button
}

function letterControl(label: string, className: string, letters: string, kind: 'case' | 'word'): HTMLButtonElement {
  const button = document.createElement('button')
  button.type = 'button'
  button.className = `icon ${className}`
  button.setAttribute('aria-label', label)
  button.setAttribute('aria-pressed', 'false')
  button.title = label
  const glyph = document.createElement('span')
  glyph.className = `glyph ${kind}`
  glyph.setAttribute('aria-hidden', 'true')
  glyph.textContent = letters
  button.append(glyph)
  return button
}

function textControl(label: string, className: string): HTMLButtonElement {
  const button = document.createElement('button')
  button.type = 'button'
  button.className = `text ${className}`
  button.textContent = label
  return button
}

function cap(n: number): string {
  return n > 999 ? '999+' : String(n)
}
