import assert from 'node:assert/strict'
import { test } from 'node:test'
import { checkDoc, flatText } from '../model.ts'
import type { Paragraph, Pos } from '../model.ts'
import { apply, compositionDelete, initialState, insertTable, insertText, selectAll, selectedText, setSelection, splitParagraph, toggleList, undo } from './edit.ts'
import type { EditorState } from './edit.ts'
import {
  buildKeymap,
  caretAtPreeditEnd,
  chordOf,
  defaultKeymap,
  hostPreeditBox,
  isHostPreeditStyle,
  detectPlatform,
  dragEndpoints,
  dragUnit,
  gestureOffsets,
  IDLE,
  isPictureFile,
  paragraphRange,
  pastePlan,
  PICTURE_TEXT,
  planBeforeInput,
  shortcutChord,
  trackGoal,
  transition,
  wordRange,
} from './input.ts'
import type { Action, DragUnit, InputMode, KeyFields, ModeEvent, Platform } from './input.ts'
import { makeDoc, paragraph, paraId, text } from './testkit.ts'

const PLATFORMS: readonly Platform[] = ['mac', 'other']
const keymaps = { mac: defaultKeymap('mac'), other: defaultKeymap('other') }

const ctrl = { ctrlKey: true }
const cmd = { metaKey: true }
const alt = { altKey: true }
const shift = { shiftKey: true }

function keyEvent(key: string, held: Partial<Omit<KeyFields, 'key'>> = {}): KeyFields {
  return { key, ctrlKey: false, metaKey: false, altKey: false, shiftKey: false, ...held }
}

const first = paragraph('a', [text('hello world', 1)], {}, { src: 1 })
const second = paragraph('b', [text('second line', 2)], {}, { src: 2 })
const third = paragraph('c', [text('third', 3)], {}, { src: 3 })

function stateAt(para: Paragraph, o: number): EditorState {
  const state = initialState(makeDoc([first, second, third]))
  const caret = { p: para.id, o }
  const sel = { kind: 'text', anchor: caret, head: caret } as const
  const tr = setSelection(sel)(state)
  if (!tr) {
    assert.deepEqual(state.sel, sel, 'the caret is already there')
    return state
  }
  return apply(state, tr, 0)
}

function bound(platform: Platform, event: KeyFields): Action | undefined {
  return keymaps[platform].get(chordOf(event, platform))
}

function press(state: EditorState, platform: Platform, event: KeyFields): EditorState {
  const chord = chordOf(event, platform)
  const action = keymaps[platform].get(chord)
  assert.ok(action && action.kind === 'command', `${chord} runs a command on ${platform}`)
  const tr = action.command(state)
  assert.ok(tr, `${chord} applies`)
  const next = apply(state, tr, 0)
  checkDoc(next.doc)
  return next
}

function heldFor(chord: string, platform: Platform): KeyFields {
  const parts = chord.split('+')
  const key = parts.pop() ?? ''
  const held = { ctrlKey: false, metaKey: false, altKey: false, shiftKey: false }
  for (const part of parts) {
    if (part === 'mod') held[platform === 'mac' ? 'metaKey' : 'ctrlKey'] = true
    else if (part === 'ctrl') held.ctrlKey = true
    else if (part === 'meta') held.metaKey = true
    else if (part === 'alt') held.altKey = true
    else if (part === 'shift') held.shiftKey = true
    else assert.fail(`${chord} has an unknown modifier ${part}`)
  }
  return { key, ...held }
}

function caretOf(state: EditorState): Pos {
  assert.ok(state.sel.kind === 'text', 'the selection is in text')
  assert.deepEqual(state.sel.anchor, state.sel.head, 'the selection is collapsed')
  return state.sel.head
}

test('chordOf spells the platform primary modifier as mod and keeps the other one by name', () => {
  assert.equal(chordOf(keyEvent('b', ctrl), 'other'), 'mod+b', 'Ctrl is the primary modifier outside macOS')
  assert.equal(chordOf(keyEvent('b', cmd), 'mac'), 'mod+b', 'Cmd is the primary modifier on macOS')
  assert.equal(chordOf(keyEvent('b', cmd), 'other'), 'meta+b', 'the Windows key stays meta')
  assert.equal(chordOf(keyEvent('b', ctrl), 'mac'), 'ctrl+b', 'Ctrl on macOS stays ctrl')
})

test('chordOf lower-cases the key and puts shift after the other modifiers', () => {
  assert.equal(chordOf(keyEvent('B', { ctrlKey: true, shiftKey: true }), 'other'), 'mod+shift+b')
  assert.equal(chordOf(keyEvent('Z', { metaKey: true, shiftKey: true }), 'mac'), 'mod+shift+z')
  assert.equal(chordOf(keyEvent('ArrowDown', shift), 'mac'), 'shift+arrowdown')
  assert.equal(chordOf(keyEvent('ArrowDown', shift), 'other'), 'shift+arrowdown')
  assert.equal(chordOf(keyEvent('Enter'), 'other'), 'enter')
  assert.equal(chordOf(keyEvent('a'), 'mac'), 'a', 'a plain key has no modifiers')
})

test('chordOf gives every modifier combination one canonical name', () => {
  const all = { ctrlKey: true, metaKey: true, altKey: true, shiftKey: true }
  assert.equal(chordOf(keyEvent('a', all), 'other'), 'mod+meta+alt+shift+a')
  assert.equal(chordOf(keyEvent('a', all), 'mac'), 'mod+ctrl+alt+shift+a')
})

test('chordOf names space and plus so the chord can be split on "+"', () => {
  assert.equal(chordOf(keyEvent(' '), 'other'), 'space')
  assert.equal(chordOf(keyEvent('+', ctrl), 'other'), 'mod+plus')
})

test('a chord bound twice is refused, including one that collides with a default', () => {
  const noop: Action = { kind: 'history', step: 'undo' }
  assert.throws(() => buildKeymap([['mod+q', noop], ['mod+q', noop]]), RangeError)
  assert.throws(() => buildKeymap([...defaultKeymap('other'), ['mod+b', noop]]), RangeError, 'an app binding cannot shadow bold')
  assert.throws(() => buildKeymap([...defaultKeymap('other'), ['mod+s', noop]]), RangeError, 'an app binding cannot shadow save')
  const extended = buildKeymap([...defaultKeymap('other'), ['mod+q', noop]])
  assert.equal(extended.size, defaultKeymap('other').size + 1, 'a free chord is added')
})

test('the default keymap binds each chord once on both platforms', () => {
  for (const platform of PLATFORMS) {
    assert.doesNotThrow(() => defaultKeymap(platform), platform)
  }
})

test('every default chord is one that chordOf produces for a real key press', () => {
  for (const platform of PLATFORMS) {
    for (const chord of keymaps[platform].keys()) {
      assert.equal(chordOf(heldFor(chord, platform), platform), chord, `${chord} is canonical on ${platform}`)
    }
  }
})

test('every command in the default keymap applies on a mid-document caret and keeps the document valid', () => {
  for (const platform of PLATFORMS) {
    for (const [chord, action] of keymaps[platform]) {
      if (action.kind !== 'command') continue
      const state = stateAt(second, 6)
      if (chord === 'shift+tab') {
        assert.equal(action.command(state), null, `${chord} does not apply in plain text on ${platform}`)
        continue
      }
      const tr = action.command(state)
      assert.ok(tr, `${chord} applies on ${platform}`)
      checkDoc(apply(state, tr, 0).doc)
    }
  }
})

test('the keymap reaches every visual motion, with and without Shift, plus undo, redo and each clipboard operation', () => {
  for (const platform of PLATFORMS) {
    const actions = [...keymaps[platform].values()]
    const motions = actions.flatMap(action => action.kind === 'visual' ? [`${action.motion}:${action.extend}`] : [])
    for (const motion of ['lineUp', 'lineDown', 'pageUp', 'pageDown', 'lineStart', 'lineEnd']) {
      assert.ok(motions.includes(`${motion}:false`), `${motion} moves on ${platform}`)
      assert.ok(motions.includes(`${motion}:true`), `${motion} extends on ${platform}`)
    }
    const steps = actions.flatMap(action => action.kind === 'history' ? [action.step] : [])
    assert.deepEqual([...new Set(steps)].sort(), ['redo', 'undo'], `history on ${platform}`)
    const ops = actions.flatMap(action => action.kind === 'clipboard' ? [action.op] : [])
    assert.deepEqual([...new Set(ops)].sort(), ['copy', 'cut', 'paste'], `clipboard on ${platform}`)
    const files = actions.flatMap(action => action.kind === 'file' ? [action.op] : [])
    assert.deepEqual(
      [...new Set(files)].sort(),
      ['closeTab', 'new', 'nextTab', 'open', 'prevTab', 'save', 'saveAs'],
      `file and tab chords on ${platform}`,
    )
  }
})

test('tab chords sit on Ctrl on both platforms, and new, open and close on the primary modifier', () => {
  const op = (platform: Platform, chord: string): string | null => {
    const action = keymaps[platform].get(chord)
    return action?.kind === 'file' ? action.op : null
  }
  assert.equal(op('mac', 'ctrl+tab'), 'nextTab')
  assert.equal(op('mac', 'ctrl+shift+tab'), 'prevTab')
  assert.equal(op('other', 'mod+tab'), 'nextTab', 'Ctrl reads as mod off macOS')
  assert.equal(op('other', 'mod+shift+tab'), 'prevTab')
  for (const platform of PLATFORMS) {
    assert.equal(op(platform, 'mod+n'), 'new', platform)
    assert.equal(op(platform, 'mod+o'), 'open', platform)
    assert.equal(op(platform, 'mod+w'), 'closeTab', platform)
    assert.equal(shortcutChord(platform, 'new'), 'mod+n', platform)
    assert.equal(shortcutChord(platform, 'open'), 'mod+o', platform)
  }
})

test('shortcutChord reads the chord the keymap bound', () => {
  for (const platform of PLATFORMS) {
    assert.equal(shortcutChord(platform, 'bold'), 'mod+b', platform)
    assert.equal(shortcutChord(platform, 'italic'), 'mod+i', platform)
    assert.equal(shortcutChord(platform, 'underline'), 'mod+u', platform)
    assert.equal(shortcutChord(platform, 'undo'), 'mod+z', platform)
    assert.equal(shortcutChord(platform, 'redo'), 'mod+y', platform)
    assert.equal(shortcutChord(platform, 'save'), 'mod+s', platform)
    assert.equal(shortcutChord(platform, 'saveAs'), 'mod+shift+s', platform)
    assert.equal(shortcutChord(platform, 'superscript'), 'mod+shift+plus', platform)
    assert.equal(shortcutChord(platform, 'subscript'), 'mod+=', platform)
  }
})

test('find chords open, step, and leave ⌘H unbound on macOS', () => {
  assert.equal(shortcutChord('other', 'find'), 'mod+f')
  assert.equal(shortcutChord('mac', 'find'), 'mod+f')
  assert.equal(shortcutChord('other', 'replace'), 'mod+h')
  assert.equal(shortcutChord('mac', 'replace'), 'ctrl+h')
  assert.equal(shortcutChord('other', 'findNext'), 'mod+g')
  assert.equal(shortcutChord('mac', 'findPrevious'), 'mod+shift+g')
  assert.deepEqual(bound('other', keyEvent('f3')), { kind: 'find', op: 'next' })
  assert.deepEqual(bound('other', keyEvent('f3', shift)), { kind: 'find', op: 'previous' })
  assert.deepEqual(bound('mac', keyEvent('h', ctrl)), { kind: 'find', op: 'replace' })
  assert.equal(bound('mac', keyEvent('h', cmd)), undefined)
  assert.deepEqual(bound('other', keyEvent('g', { ...ctrl, ...shift })), { kind: 'find', op: 'previous' })
})

test('superscript is the plus key Shift delivers, and subscript stays on mod+=', () => {
  for (const [platform, held] of [['mac', cmd], ['other', ctrl]] as const) {
    assert.equal(chordOf(keyEvent('+', { ...held, ...shift }), platform), 'mod+shift+plus')
    const raised = press(stateAt(second, 6), platform, keyEvent('+', { ...held, ...shift }))
    assert.equal(raised.stored?.vertAlign, 'super')
    const lowered = press(stateAt(second, 6), platform, keyEvent('=', held))
    assert.equal(lowered.stored?.vertAlign, 'sub')
  }
})

test('undo, redo and the clipboard sit on mod on both platforms, with redo on Shift+Z and on Y', () => {
  for (const [platform, held] of [['mac', cmd], ['other', ctrl]] as const) {
    assert.deepEqual(bound(platform, keyEvent('z', held)), { kind: 'history', step: 'undo' })
    assert.deepEqual(bound(platform, keyEvent('Z', { ...held, ...shift })), { kind: 'history', step: 'redo' })
    assert.deepEqual(bound(platform, keyEvent('y', held)), { kind: 'history', step: 'redo' })
    assert.deepEqual(bound(platform, keyEvent('c', held)), { kind: 'clipboard', op: 'copy' })
    assert.deepEqual(bound(platform, keyEvent('x', held)), { kind: 'clipboard', op: 'cut' })
    assert.deepEqual(bound(platform, keyEvent('v', held)), { kind: 'clipboard', op: 'paste' })
    assert.deepEqual(bound(platform, keyEvent('s', held)), { kind: 'file', op: 'save' })
    assert.deepEqual(bound(platform, keyEvent('S', { ...held, ...shift })), { kind: 'file', op: 'saveAs' })
  }
})

test('Home, End, Page Up, Page Down and the vertical arrows are visual motions that extend with Shift', () => {
  for (const platform of PLATFORMS) {
    assert.deepEqual(bound(platform, keyEvent('Home')), { kind: 'visual', motion: 'lineStart', extend: false })
    assert.deepEqual(bound(platform, keyEvent('End', shift)), { kind: 'visual', motion: 'lineEnd', extend: true })
    assert.deepEqual(bound(platform, keyEvent('PageUp')), { kind: 'visual', motion: 'pageUp', extend: false })
    assert.deepEqual(bound(platform, keyEvent('PageDown', shift)), { kind: 'visual', motion: 'pageDown', extend: true })
    assert.deepEqual(bound(platform, keyEvent('ArrowUp')), { kind: 'visual', motion: 'lineUp', extend: false })
    assert.deepEqual(bound(platform, keyEvent('ArrowDown', shift)), { kind: 'visual', motion: 'lineDown', extend: true })
  }
})

test('Enter splits the paragraph, Shift+Enter breaks the line inside it, and Mod+Enter inserts a page break', () => {
  for (const [platform, held] of [['mac', cmd], ['other', ctrl]] as const) {
    const start = stateAt(first, 5)
    const split = press(start, platform, keyEvent('Enter'))
    assert.deepEqual(split.doc.body.map(block => block.kind === 'p' ? flatText(block) : ''), ['hello', ' world', 'second line', 'third'], 'enter splits at the caret')
    const broken = press(start, platform, keyEvent('Enter', shift))
    assert.equal(broken.doc.body.length, 3, 'a line break adds no paragraph')
    const para = broken.doc.body[0]
    assert.ok(para && para.kind === 'p')
    assert.equal(flatText(para), 'hello\n world', 'the break is inside the paragraph')
    const paged = press(start, platform, keyEvent('Enter', held))
    const marked = paged.doc.body[0]
    assert.ok(marked && marked.kind === 'p')
    assert.equal(flatText(marked), 'hello￼ world', 'the page break is one atom at the caret')
  }
})

test('Backspace at a paragraph start and Delete at a paragraph end merge across the boundary', () => {
  for (const platform of PLATFORMS) {
    const back = press(stateAt(second, 0), platform, keyEvent('Backspace'))
    assert.deepEqual(back.doc.body.map(block => block.kind === 'p' ? flatText(block) : ''), ['hello worldsecond line', 'third'])
    assert.deepEqual(caretOf(back), { p: first.id, o: 11 }, 'the caret sits at the join')
    const forward = press(stateAt(first, 11), platform, keyEvent('Delete'))
    assert.deepEqual(forward.doc.body.map(block => block.kind === 'p' ? flatText(block) : ''), ['hello worldsecond line', 'third'])
  }
})

test('Backspace and Delete remove one character inside a paragraph', () => {
  const back = press(stateAt(first, 5), 'other', keyEvent('Backspace'))
  const afterBack = back.doc.body[0]
  assert.ok(afterBack && afterBack.kind === 'p')
  assert.equal(flatText(afterBack), 'hell world')
  const forward = press(stateAt(first, 5), 'other', keyEvent('Delete'))
  const afterForward = forward.doc.body[0]
  assert.ok(afterForward && afterForward.kind === 'p')
  assert.equal(flatText(afterForward), 'helloworld')
})

test('word deletion is Ctrl+Backspace outside macOS and Alt+Backspace on macOS, and neither platform answers the other chord', () => {
  const end = stateAt(first, 11)
  const other = press(end, 'other', keyEvent('Backspace', ctrl))
  const mac = press(end, 'mac', keyEvent('Backspace', alt))
  for (const state of [other, mac]) {
    const para = state.doc.body[0]
    assert.ok(para && para.kind === 'p')
    assert.equal(flatText(para), 'hello ', 'the last word is gone')
  }
  assert.equal(bound('mac', keyEvent('Backspace', ctrl)), undefined, 'Ctrl+Backspace is not a word delete on macOS')
  assert.equal(bound('other', keyEvent('Backspace', alt)), undefined, 'Alt+Backspace is not a word delete elsewhere')
})

test('arrows move one character and Shift extends the selection from the original anchor', () => {
  for (const platform of PLATFORMS) {
    const right = press(stateAt(first, 3), platform, keyEvent('ArrowRight'))
    assert.deepEqual(caretOf(right), { p: first.id, o: 4 })
    const extended = press(press(stateAt(first, 3), platform, keyEvent('ArrowRight', shift)), platform, keyEvent('ArrowRight', shift))
    assert.deepEqual(extended.sel, { kind: 'text', anchor: { p: first.id, o: 3 }, head: { p: first.id, o: 5 } })
    const collapsed = press(extended, platform, keyEvent('ArrowLeft'))
    assert.deepEqual(caretOf(collapsed), { p: first.id, o: 3 }, 'a plain arrow collapses to the start of the selection')
  }
})

test('an arrow at a paragraph edge crosses into the neighbor', () => {
  const next = press(stateAt(first, 11), 'other', keyEvent('ArrowRight'))
  assert.deepEqual(caretOf(next), { p: second.id, o: 0 })
  const previous = press(stateAt(second, 0), 'other', keyEvent('ArrowLeft'))
  assert.deepEqual(caretOf(previous), { p: first.id, o: 11 })
})

test('word motion is Ctrl+arrow outside macOS and Alt+arrow on macOS', () => {
  const start = stateAt(first, 0)
  const other = press(start, 'other', keyEvent('ArrowRight', ctrl))
  const mac = press(start, 'mac', keyEvent('ArrowRight', alt))
  assert.deepEqual(caretOf(other), { p: first.id, o: 6 }, 'Ctrl+Right lands at the next word start')
  assert.deepEqual(caretOf(mac), { p: first.id, o: 6 }, 'Alt+Right lands at the next word start')
  assert.equal(bound('mac', keyEvent('ArrowRight', ctrl)), undefined, 'Ctrl+arrow belongs to the OS on macOS')
  const selecting = press(start, 'other', keyEvent('ArrowRight', { ...ctrl, ...shift }))
  assert.deepEqual(selecting.sel, { kind: 'text', anchor: { p: first.id, o: 0 }, head: { p: first.id, o: 6 } })
})

test('the document edges are Ctrl+Home and Ctrl+End outside macOS and Cmd+Up and Cmd+Down on macOS', () => {
  const middle = stateAt(second, 3)
  for (const [platform, event] of [['other', keyEvent('Home', ctrl)], ['mac', keyEvent('ArrowUp', cmd)]] as const) {
    assert.deepEqual(caretOf(press(middle, platform, event)), { p: first.id, o: 0 }, `document start on ${platform}`)
  }
  for (const [platform, event] of [['other', keyEvent('End', ctrl)], ['mac', keyEvent('ArrowDown', cmd)]] as const) {
    assert.deepEqual(caretOf(press(middle, platform, event)), { p: third.id, o: 5 }, `document end on ${platform}`)
  }
})

test('Cmd+Left and Cmd+Right are line edges on macOS', () => {
  assert.deepEqual(bound('mac', keyEvent('ArrowLeft', cmd)), { kind: 'visual', motion: 'lineStart', extend: false })
  assert.deepEqual(bound('mac', keyEvent('ArrowRight', { ...cmd, ...shift })), { kind: 'visual', motion: 'lineEnd', extend: true })
})

test('Alt+Up and Alt+Down move by paragraph on macOS', () => {
  const middle = stateAt(second, 3)
  assert.deepEqual(caretOf(press(middle, 'mac', keyEvent('ArrowUp', alt))), { p: second.id, o: 0 })
  assert.deepEqual(caretOf(press(middle, 'mac', keyEvent('ArrowDown', alt))), { p: second.id, o: 11 })
})

test('Mod+A selects the whole document', () => {
  for (const [platform, held] of [['mac', cmd], ['other', ctrl]] as const) {
    const state = press(stateAt(second, 3), platform, keyEvent('a', held))
    assert.deepEqual(state.sel, { kind: 'text', anchor: { p: first.id, o: 0 }, head: { p: third.id, o: 5 } })
  }
})

test('Mod+B, Mod+I and Mod+U mark the next typed text and the other platform primary key does not', () => {
  for (const [platform, held, foreign] of [['mac', cmd, ctrl], ['other', ctrl, cmd]] as const) {
    for (const [letter, mark] of [['b', 'bold'], ['i', 'italic'], ['u', 'underline']] as const) {
      const marked = press(stateAt(first, 5), platform, keyEvent(letter, held))
      assert.equal(marked.stored?.[mark], true, `${letter} stores ${mark} on ${platform}`)
      assert.equal(bound(platform, keyEvent(letter, foreign)), undefined, `${letter} with the other primary key is unbound on ${platform}`)
    }
  }
  const bold = press(stateAt(first, 5), 'other', keyEvent('b', ctrl))
  const typed = apply(bold, insertText('!')(bold) ?? assert.fail('typing applies'), 10)
  const para = typed.doc.body[0]
  assert.ok(para && para.kind === 'p')
  const runs = para.inlines.flatMap(inline => inline.kind === 'text' ? [[inline.text, inline.marks.bold]] : [])
  assert.deepEqual(runs, [['hello', undefined], ['!', true], [' world', undefined]], 'only the typed character is bold')
})

const at: Pos = { p: paraId('a'), o: 3 }
const elsewhere: Pos = { p: paraId('b'), o: 1 }
const composing: InputMode = { kind: 'composing', at }
const dragging: InputMode = { kind: 'dragging', anchor: at, unit: 'char' }

const compositionStart: ModeEvent = { kind: 'compositionStart', at }
const compositionEnd: ModeEvent = { kind: 'compositionEnd' }
const pointerDown = (unit: DragUnit = 'char'): ModeEvent => ({ kind: 'pointerDown', at: elsewhere, unit })
const pointerUp: ModeEvent = { kind: 'pointerUp' }
const reset: ModeEvent = { kind: 'reset' }

test('the runtime composition span is recognized by the style it is given', () => {
  const host = 'position:fixed;left:10px;top:20px;height:16px;line-height:16px;z-index:2147483647;text-decoration:underline;background:Canvas;color:CanvasText;pointer-events:none;white-space:pre'
  const cssom = 'position: fixed; left: 10px; top: 20px; height: 16px; line-height: 16px; z-index: 2147483647; text-decoration: underline; background: Canvas; color: CanvasText; pointer-events: none; white-space: pre;'
  assert.equal(isHostPreeditStyle(host), true)
  assert.equal(isHostPreeditStyle(cssom), true)
  assert.equal(isHostPreeditStyle('position:fixed;z-index:2147483647;border-radius:12px'), false)
})

test('the runtime preedit shifts onto the paragraph baseline at the active zoom', () => {
  assert.deepEqual(
    hostPreeditBox({ left: 96, top: 37 }, { deltaTop: 2.25, height: 15 }, 1),
    { left: 96, top: 39.25, height: 15 },
  )
  assert.deepEqual(
    hostPreeditBox({ left: 192, top: 74 }, { deltaTop: 2.25, height: 15 }, 2),
    { left: 192, top: 78.5, height: 30 },
  )
})

test('the caret sits at the end of the runtime preedit, in overlay pixels', () => {
  assert.deepEqual(
    caretAtPreeditEnd({ right: 140, top: 80, height: 20 }, { left: 40, top: 60, width: 200, offsetWidth: 200 }),
    { left: 100, top: 20, height: 20 },
  )
  assert.deepEqual(
    caretAtPreeditEnd({ right: 240, top: 160, height: 40 }, { left: 80, top: 120, width: 400, offsetWidth: 200 }),
    { left: 80, top: 20, height: 20 },
  )
})

test('a composition starts at the remembered caret and ends back at idle', () => {
  const started = transition(IDLE, compositionStart)
  assert.deepEqual(started, { kind: 'composing', at })
  assert.deepEqual(transition(started, compositionEnd), IDLE)
})

test('a drag remembers its anchor and unit until the pointer is released', () => {
  const started = transition(IDLE, { kind: 'pointerDown', at, unit: 'word' })
  assert.deepEqual(started, { kind: 'dragging', anchor: at, unit: 'word' }, 'a double click drags by word')
  assert.deepEqual(transition(started, pointerUp), IDLE)
})

test('events that mean nothing in a mode return that very mode', () => {
  const ignored: readonly [string, InputMode, ModeEvent][] = [
    ['idle ignores a composition end', IDLE, compositionEnd],
    ['idle ignores a pointer release', IDLE, pointerUp],
    ['idle ignores a reset', IDLE, reset],
    ['composing ignores a second composition start', composing, { kind: 'compositionStart', at: elsewhere }],
    ['composing ignores a pointer press', composing, pointerDown()],
    ['composing ignores a pointer release', composing, pointerUp],
    ['dragging ignores a composition end', dragging, compositionEnd],
  ]
  for (const [name, mode, event] of ignored) {
    assert.equal(transition(mode, event), mode, name)
  }
})

test('a click during composition does not start a drag, and the commit still finds the original caret', () => {
  const clicked = transition(composing, pointerDown())
  assert.deepEqual(clicked, { kind: 'composing', at }, 'the remembered caret is kept')
  assert.deepEqual(transition(clicked, compositionEnd), IDLE)
})

test('a second composition start keeps the first caret', () => {
  assert.deepEqual(transition(composing, { kind: 'compositionStart', at: elsewhere }), { kind: 'composing', at })
})

test('a press during a drag restarts it at the new anchor and unit', () => {
  assert.deepEqual(transition(dragging, pointerDown('paragraph')), { kind: 'dragging', anchor: elsewhere, unit: 'paragraph' })
})

test('a composition that starts during a drag takes over from the drag', () => {
  const taken = transition(dragging, { kind: 'compositionStart', at: elsewhere })
  assert.deepEqual(taken, { kind: 'composing', at: elsewhere })
  assert.equal(transition(taken, pointerUp), taken, 'the late mouseup does not disturb the composition')
})

test('reset abandons a composition or a drag', () => {
  assert.deepEqual(transition(composing, reset), IDLE)
  assert.deepEqual(transition(dragging, reset), IDLE)
})

test('wordRange selects the word under the offset and the word before a space', () => {
  assert.deepEqual(wordRange('hello world', 1), { start: 0, end: 5 })
  assert.deepEqual(wordRange('hello world', 5), { start: 0, end: 5 }, 'a space belongs to the word before it')
  assert.deepEqual(wordRange('hello world', 6), { start: 6, end: 11 })
  assert.deepEqual(wordRange('hello world', 11), { start: 6, end: 11 }, 'the end of the text is the last word')
  assert.deepEqual(wordRange('', 0), { start: 0, end: 0 })
  const cjk = wordRange('你好世界', 1)
  assert.ok(cjk.start <= 1 && cjk.end > 1 && cjk.end <= 4, `cjk word ${cjk.start}-${cjk.end}`)
})

test('paragraphRange is the whole paragraph, including an empty one', () => {
  assert.deepEqual(paragraphRange('hello'), { start: 0, end: 5 })
  assert.deepEqual(paragraphRange(''), { start: 0, end: 0 })
})

test('dragUnit follows the click count', () => {
  assert.equal(dragUnit(1), 'char')
  assert.equal(dragUnit(0), 'char')
  assert.equal(dragUnit(2), 'word')
  assert.equal(dragUnit(3), 'paragraph')
  assert.equal(dragUnit(4), 'paragraph')
})

test('a double click selects the word and a drag grows by that word', () => {
  assert.deepEqual(gestureOffsets('hello world', 1, 1, 'word'), { anchor: 0, head: 5 })
  assert.deepEqual(gestureOffsets('hello world', 1, 8, 'word'), { anchor: 0, head: 11 }, 'dragging forward covers both words')
  assert.deepEqual(gestureOffsets('hello world', 8, 1, 'word'), { anchor: 11, head: 0 }, 'dragging backward keeps the press as the anchor side')
  assert.deepEqual(gestureOffsets('hello world', 2, 5, 'char'), { anchor: 2, head: 5 })
  assert.deepEqual(gestureOffsets('hello world', 5, 2, 'char'), { anchor: 5, head: 2 })
  assert.deepEqual(gestureOffsets('abcdef', 4, 1, 'paragraph'), { anchor: 6, head: 0 })
  assert.deepEqual(gestureOffsets('abcdef', 1, 4, 'paragraph'), { anchor: 0, head: 6 })
})

test('a drag across paragraphs expands each end to its unit', () => {
  assert.deepEqual(dragEndpoints('hello', 1, 'world', 1, false, false, 'word'), { anchor: 0, head: 5 })
  assert.deepEqual(dragEndpoints('hello', 1, 'world', 1, false, true, 'word'), { anchor: 5, head: 0 })
  assert.deepEqual(dragEndpoints('hello', 1, 'world', 3, false, false, 'char'), { anchor: 1, head: 3 })
  assert.deepEqual(dragEndpoints('hello', 2, 'world', 1, false, false, 'paragraph'), { anchor: 0, head: 5 })
  assert.deepEqual(dragEndpoints('hello', 2, 'world', 1, true, false, 'word'), gestureOffsets('hello', 2, 1, 'word'))
})

test('the goal column sticks across vertical moves and clears on any other action', () => {
  const down: Action = { kind: 'visual', motion: 'lineDown', extend: false }
  const downExtend: Action = { kind: 'visual', motion: 'lineDown', extend: true }
  const up: Action = { kind: 'visual', motion: 'lineUp', extend: false }
  const page: Action = { kind: 'visual', motion: 'pageDown', extend: false }
  assert.equal(trackGoal(null, down, 40), 40, 'the first vertical move measures the caret')
  assert.equal(trackGoal(40, down, 12), 40, 'the next vertical move keeps the column')
  assert.equal(trackGoal(40, downExtend, 3), 40, 'shift does not start a new column')
  assert.equal(trackGoal(40, up, 90), 40)
  assert.equal(trackGoal(40, page, 1), 40)
  assert.equal(trackGoal(40, { kind: 'visual', motion: 'lineStart', extend: false }, 40), null)
  assert.equal(trackGoal(40, { kind: 'visual', motion: 'lineEnd', extend: true }, 40), null)
  assert.equal(trackGoal(40, { kind: 'history', step: 'undo' }, 40), null)
  assert.equal(trackGoal(40, { kind: 'clipboard', op: 'paste' }, 40), null)
  assert.equal(trackGoal(40, { kind: 'command', command: splitParagraph }, 40), null)
  assert.equal(trackGoal(null, down, null), null)
})

test('beforeinput maps text, breaks and deletes, and leaves composition to the composition events', () => {
  const start = stateAt(first, 5)
  const inserted = planBeforeInput('insertText', 'XY', false)
  assert.equal(inserted.kind, 'run')
  if (inserted.kind === 'run') {
    const tr = inserted.command(start)
    assert.ok(tr)
    const para = apply(start, tr, 0).doc.body[0]
    assert.ok(para && para.kind === 'p')
    assert.equal(flatText(para), 'helloXY world')
  }
  assert.equal(planBeforeInput('insertText', '', false).kind, 'ignore')
  assert.equal(planBeforeInput('insertText', null, false).kind, 'ignore')
  const line = planBeforeInput('insertLineBreak', null, false)
  assert.equal(line.kind, 'run')
  if (line.kind === 'run') {
    const tr = line.command(start)
    assert.ok(tr)
    const para = apply(start, tr, 0).doc.body[0]
    assert.ok(para && para.kind === 'p')
    assert.equal(flatText(para), 'hello\n world')
  }
  const split = planBeforeInput('insertParagraph', null, false)
  assert.equal(split.kind, 'run')
  if (split.kind === 'run') assert.equal(split.command, splitParagraph)
  const back = planBeforeInput('deleteContentBackward', null, false)
  assert.equal(back.kind, 'run')
  if (back.kind === 'run') {
    const tr = back.command(start)
    assert.ok(tr)
    const para = apply(start, tr, 0).doc.body[0]
    assert.ok(para && para.kind === 'p')
    assert.equal(flatText(para), 'hell world')
  }
  const forward = planBeforeInput('deleteContentForward', null, false)
  assert.equal(forward.kind, 'run')
  if (forward.kind === 'run') {
    const tr = forward.command(start)
    assert.ok(tr)
    const para = apply(start, tr, 0).doc.body[0]
    assert.ok(para && para.kind === 'p')
    assert.equal(flatText(para), 'helloworld')
  }
  assert.equal(planBeforeInput('insertText', 'a', true).kind, 'compose', 'typing during composition stays out of the model')
  for (const type of ['insertCompositionText', 'deleteCompositionText', 'insertFromComposition', 'deleteByComposition']) {
    assert.equal(planBeforeInput(type, 'a', false).kind, 'compose', type)
  }
  assert.equal(planBeforeInput('insertFromPaste', 'pasted', false).kind, 'ignore', 'paste arrives as a clipboard event')
})

test('replacing a selection with a composition is one typing entry', () => {
  const start = stateAt(first, 2)
  const selectedTr = setSelection({ kind: 'text', anchor: { p: first.id, o: 2 }, head: { p: first.id, o: 5 } })(start)
  assert.ok(selectedTr)
  const selected = apply(start, selectedTr, 0)
  assert.equal(compositionDelete()(stateAt(first, 3)), null, 'a collapsed caret deletes nothing')
  const cleared = compositionDelete()(selected)
  assert.ok(cleared)
  const mid = apply(selected, cleared, 0)
  const typed = insertText('XY')(mid)
  assert.ok(typed)
  const done = apply(mid, typed, 20)
  assert.equal(done.history.past.length, 1, 'the deletion and the commit coalesce')
  const para = done.doc.body[0]
  assert.ok(para && para.kind === 'p')
  assert.equal(flatText(para), 'heXY world')
  const undone = undo(done)
  assert.ok(undone)
  assert.equal(selectedText(undone), 'llo')
})

test('tab indents a list, moves through table cells, appends a row, and inserts a tab in plain text', () => {
  for (const platform of PLATFORMS) {
    const tab = keyEvent('Tab')
    const untab = keyEvent('Tab', shift)
    assert.equal(chordOf(tab, platform), 'tab')
    assert.equal(chordOf(untab, platform), 'shift+tab')

    const plain = stateAt(second, 6)
    const tabbed = press(plain, platform, tab)
    const written = tabbed.doc.body[1]
    assert.ok(written && written.kind === 'p')
    assert.equal(flatText(written), 'second\t line', 'tab in plain text inserts a tab character')
    assert.deepEqual(caretOf(tabbed), { p: second.id, o: 7 })
    const shiftPlain = bound(platform, untab)
    assert.ok(shiftPlain && shiftPlain.kind === 'command')
    assert.equal(shiftPlain.command(plain), null, 'shift+tab in plain text does not change the document')

    const items = [paragraph('a', [text('one')]), paragraph('b', [text('two')]), paragraph('c', [text('three')])]
    let listed = initialState(makeDoc(items))
    const all = selectAll(listed)
    assert.ok(all)
    listed = apply(listed, all, 0)
    const numbered = toggleList('decimal')(listed)
    assert.ok(numbered)
    listed = apply(listed, numbered, 0)
    const secondItem = items[1]
    assert.ok(secondItem)
    const atSecond = setSelection({ kind: 'text', anchor: { p: secondItem.id, o: 0 }, head: { p: secondItem.id, o: 0 } })(listed)
    assert.ok(atSecond)
    listed = apply(listed, atSecond, 0)
    listed = press(listed, platform, tab)
    const indented = listed.doc.body[1]
    assert.ok(indented && indented.kind === 'p')
    assert.equal(indented.props.numbering?.level, 1, 'tab in a list raises the level')
    listed = press(listed, platform, untab)
    const restored = listed.doc.body[1]
    assert.ok(restored && restored.kind === 'p')
    assert.equal(restored.props.numbering?.level, 0, 'shift+tab lowers the level')

    let cells = initialState(makeDoc([paragraph('host', [text('ab')])]))
    const made = insertTable(2, 2)(cells)
    assert.ok(made)
    cells = apply(cells, made, 0)
    const table = cells.doc.body.find((block) => block.kind === 'table')
    assert.ok(table && table.kind === 'table')
    const cellId = (row: number, cell: number) => {
      const block = table.rows[row]?.cells[cell]?.blocks[0]
      assert.ok(block && block.kind === 'p')
      return block.id
    }
    assert.equal(caretOf(cells).p, cellId(0, 0))
    cells = press(cells, platform, tab)
    assert.equal(caretOf(cells).p, cellId(0, 1), 'tab moves to the next cell')
    cells = press(cells, platform, tab)
    assert.equal(caretOf(cells).p, cellId(1, 0), 'tab wraps to the next row')
    cells = press(cells, platform, untab)
    assert.equal(caretOf(cells).p, cellId(0, 1), 'shift+tab moves to the previous cell')
    const last = cellId(1, 1)
    const atLast = setSelection({
      kind: 'text',
      anchor: { p: last, o: 0 },
      head: { p: last, o: 0 },
    })(cells)
    assert.ok(atLast)
    cells = apply(cells, atLast, 0)
    const grown = press(cells, platform, tab)
    const grownTable = grown.doc.body.find((block) => block.kind === 'table')
    assert.ok(grownTable && grownTable.kind === 'table')
    assert.equal(grownTable.rows.length, 3, 'tab in the last cell appends a row')
    const fresh = grownTable.rows[2]?.cells[0]?.blocks[0]
    assert.ok(fresh && fresh.kind === 'p')
    assert.deepEqual(caretOf(grown), { p: fresh.id, o: 0 }, 'the caret moves into the new row')

    let inside = initialState(makeDoc([paragraph('host', [text('ab')])]))
    const nested = insertTable(2, 2)(inside)
    assert.ok(nested)
    inside = apply(inside, nested, 0)
    const asList = toggleList('decimal')(inside)
    assert.ok(asList)
    inside = apply(inside, asList, 0)
    const host = inside.doc.body.find((block) => block.kind === 'table')
    assert.ok(host && host.kind === 'table')
    const origin = host.rows[0]?.cells[0]?.blocks[0]
    const neighbor = host.rows[0]?.cells[1]?.blocks[0]
    assert.ok(origin && origin.kind === 'p' && neighbor && neighbor.kind === 'p')
    assert.equal(origin.props.numbering?.level, 0)
    inside = press(inside, platform, tab)
    assert.equal(caretOf(inside).p, neighbor.id, 'tab in a list inside a cell still moves to the next cell')
    const kept = inside.doc.body.find((block) => block.kind === 'table')
    const level = kept && kept.kind === 'table' ? kept.rows[0]?.cells[0]?.blocks[0] : undefined
    assert.ok(level && level.kind === 'p')
    assert.equal(level.props.numbering?.level, 0, 'the cell list level stays')
  }
})

test('detectPlatform treats Macintosh as mac and Linux, including Dopame, as other', () => {
  assert.equal(detectPlatform('MacIntel'), 'mac')
  assert.equal(detectPlatform('', 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)'), 'mac')
  assert.equal(detectPlatform('Linux', 'Mozilla/5.0 (X11; Linux x86_64)'), 'other')
  assert.equal(detectPlatform('Win32'), 'other')
})

test('every mode and event pair returns a mode without throwing', () => {
  const events: readonly ModeEvent[] = [compositionStart, compositionEnd, pointerDown('word'), pointerUp, reset]
  for (const mode of [IDLE, composing, dragging]) {
    for (const event of events) {
      const next = transition(mode, event)
      assert.ok(['idle', 'composing', 'dragging'].includes(next.kind), `${mode.kind} then ${event.kind}`)
    }
  }
})

test('a paste takes picture files first, then the held picture for its own mark, then text', () => {
  const held = { text: PICTURE_TEXT }
  assert.equal(pastePlan('words', 1, held), 'files', 'a picture on the clipboard wins over its text')
  assert.equal(pastePlan(PICTURE_TEXT, 0, held), 'held', 'the mark this app wrote brings the picture back')
  assert.equal(pastePlan(PICTURE_TEXT, 0, null), 'text', 'the mark with nothing held is only a character')
  assert.equal(pastePlan('words', 0, held), 'text', 'text copied elsewhere since is typed, not the old picture')
  assert.equal(pastePlan('', 0, held), 'none', 'an empty clipboard does not bring the old picture back')
  assert.equal(pastePlan('', 0, null), 'none')
})

test('isPictureFile accepts PNG and JPEG by type, and by name only when the type is empty', () => {
  assert.equal(isPictureFile({ type: 'image/png', name: 'a.png' }), true)
  assert.equal(isPictureFile({ type: 'image/jpeg' }), true)
  assert.equal(isPictureFile({ type: '', name: 'shot.JPG' }), true)
  assert.equal(isPictureFile({ type: '', name: 'shot.jpeg' }), true)
  assert.equal(isPictureFile({ type: '', name: 'notes.txt' }), false)
  assert.equal(isPictureFile({ type: 'image/gif', name: 'a.gif' }), false)
  assert.equal(isPictureFile({ type: 'application/pdf', name: 'a.png' }), false)
})
