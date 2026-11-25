import assert from 'node:assert/strict'
import { test } from 'node:test'
import {
  afterClose,
  canSave,
  cycleTab,
  commandTitle,
  fileHandleFromRequest,
  fontOptions,
  formatChord,
  formatTitle,
  isAbort,
  isDirty,
  nextUntitled,
  sizeOptions,
  statusText,
  stepZoom,
  styleOptions,
  tabLabel,
  zoomAtEnd,
  ZOOM_PERCENTS,
} from './text.ts'
import type { Translate } from './text.ts'

const translate: Translate = (key, values = {}) => {
  const templates: Record<string, string> = {
    'status.page': 'Page {page} of {pages}',
    'status.words': '{words} words',
    untitled: 'Untitled.docx',
    'untitled.n': 'Untitled {n}.docx',
  }
  const text = templates[key] ?? key
  return text.replace(/\{(\w+)\}/g, (all, name: string) => {
    const value = values[name]
    return value === undefined ? all : String(value)
  })
}

test('isDirty compares the document object with the saved snapshot', () => {
  const saved = { body: [] }
  const edited = { body: [] }
  assert.equal(isDirty(saved, null), false, 'no open file')
  assert.equal(isDirty(saved, { saved }), false, 'the snapshot is the document')
  assert.equal(isDirty(edited, { saved }), true, 'a later document object is dirty')
})

test('canSave is off only for a clean document that already has a file', () => {
  const saved = { body: [] }
  const edited = { body: [] }
  const handle = { kind: 'file' }
  assert.equal(canSave(saved, null), true, 'no open file still offers Save')
  assert.equal(canSave(saved, { handle: null, saved }), true, 'no handle falls through to Save As')
  assert.equal(canSave(saved, { handle, saved }), false, 'a clean file has nothing to write')
  assert.equal(canSave(edited, { handle, saved }), true, 'edits can be written')
})

test('formatTitle puts the dirty mark before the file name', () => {
  assert.equal(formatTitle('basic.docx', false, 'Word'), 'basic.docx - Word')
  assert.equal(formatTitle('basic.docx', true, 'Word'), '• basic.docx - Word')
  assert.equal(formatTitle('未命名.docx', false, 'Word'), '未命名.docx - Word')
})

test('statusText shows the page and the word count', () => {
  const final = statusText({ page: 1, pages: 3, pagesFinal: true, words: 12 }, translate)
  assert.equal(final, 'Page 1 of 3 · 12 words')
  const partial = statusText({ page: 2, pages: 4, pagesFinal: false, words: 0 }, translate)
  assert.equal(partial, 'Page 2 of ≥4 · 0 words')
  const half = statusText({ page: 1, pages: 1, pagesFinal: true, words: 1 }, translate)
  assert.equal(half, 'Page 1 of 1 · 1 words')
})

test('stepZoom walks the zoom list and stays put at either end', () => {
  assert.deepEqual(ZOOM_PERCENTS, [50, 75, 100, 125, 150, 200])
  assert.equal(stepZoom(100, 1), 125)
  assert.equal(stepZoom(100, -1), 75)
  assert.equal(stepZoom(50, -1), 50)
  assert.equal(stepZoom(50, 1), 75)
  assert.equal(stepZoom(200, 1), 200)
  assert.equal(stepZoom(200, -1), 150)
  assert.equal(stepZoom(60, -1), 50)
  assert.equal(stepZoom(60, 1), 75)
  assert.deepEqual(zoomAtEnd(50), { out: true, in: false })
  assert.deepEqual(zoomAtEnd(75), { out: false, in: false })
  assert.deepEqual(zoomAtEnd(200), { out: false, in: true })
})

test('formatChord spells the platform modifier', () => {
  assert.equal(formatChord('mod+b', 'other'), 'Ctrl+B')
  assert.equal(formatChord('mod+b', 'mac'), '⌘B')
  assert.equal(formatChord('mod+shift+z', 'other'), 'Ctrl+Shift+Z')
  assert.equal(formatChord('mod+shift+z', 'mac'), '⌘⇧Z')
  assert.equal(formatChord('mod+y', 'mac'), '⌘Y')
  assert.equal(formatChord('mod+space', 'other'), 'Ctrl+Space')
  assert.equal(formatChord('mod+=', 'mac'), '⌘=')
  assert.equal(formatChord('mod+shift+plus', 'mac'), '⌘⇧+')
  assert.equal(formatChord('mod+shift+plus', 'other'), 'Ctrl+Shift++')
})

test('commandTitle keeps the label when the keymap has no chord', () => {
  assert.equal(commandTitle('Bold', 'mod+b', 'other'), 'Bold (Ctrl+B)')
  assert.equal(commandTitle('加粗', 'mod+b', 'mac'), '加粗 (⌘B)')
  assert.equal(commandTitle('Strikethrough', null, 'other'), 'Strikethrough')
  assert.equal(commandTitle('Strikethrough', '', 'mac'), 'Strikethrough')
})

test('fileHandleFromRequest accepts only an open-file request for a file', () => {
  const file = { kind: 'file', name: 'basic.docx' }
  const directory = { kind: 'directory', name: 'docs' }
  assert.equal(fileHandleFromRequest({ kind: 'open-file', handle: file }), file)
  assert.equal(fileHandleFromRequest({ kind: 'open-file', handle: directory }), null)
  assert.equal(fileHandleFromRequest({ kind: 'open-file' }), null)
  assert.equal(fileHandleFromRequest({ kind: 'open-link', handle: file }), null)
})

test('isAbort recognizes a cancelled picker', () => {
  assert.equal(isAbort({ name: 'AbortError' }), true)
  assert.equal(isAbort({ name: 'NotFoundError' }), false)
  assert.equal(isAbort(null), false)
  assert.equal(isAbort('AbortError'), false)
})

test('styleOptions lists the five roles first and localizes only those', () => {
  const styles = new Map<string, { name: string; role: string | null }>([
    ['Quote', { name: 'Quote', role: null }],
    ['Heading2', { name: 'heading 2', role: 'heading2' }],
    ['Normal', { name: 'Normal', role: 'normal' }],
    ['Alpha', { name: 'Alpha', role: null }],
    ['Title', { name: 'Title', role: 'title' }],
    ['Caption', { name: 'caption', role: 'caption' }],
  ])
  const options = styleOptions(styles, role => `role:${role}`)
  assert.deepEqual(options.map(option => option.id), ['Normal', 'Title', 'Heading2', 'Caption', 'Alpha', 'Quote'])
  assert.deepEqual(options.map(option => option.label), [
    'role:normal',
    'role:title',
    'role:heading2',
    'caption',
    'Alpha',
    'Quote',
  ])
})

const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]

test('fontOptions lists document families, then installed ones, once each', () => {
  const options = fontOptions(
    ['宋体', 'Calibri', 'calibri', ''],
    ['Arial', 'calibri', '宋体', 'Times New Roman', 'arial'],
    'mixed',
  )
  assert.deepEqual(options, ['宋体', 'Calibri', 'Arial', 'Times New Roman'])
  assert.ok(!options.includes('mixed'), 'mixed is not a family')
})

test('fontOptions appends a current family that is in neither list', () => {
  const options = fontOptions(['Calibri'], ['Arial'], 'Consolas')
  assert.deepEqual(options, ['Calibri', 'Arial', 'Consolas'])
  const again = fontOptions(['Calibri'], ['Arial'], 'arial')
  assert.deepEqual(again, ['Calibri', 'Arial'], 'a different case of a listed family is not added')
})

test('sizeOptions keeps the fixed list, and inserts another size in order', () => {
  assert.deepEqual(sizeOptions('mixed'), FONT_SIZES)
  assert.deepEqual(sizeOptions(11), FONT_SIZES)
  assert.deepEqual(sizeOptions(10.5), FONT_SIZES)
  assert.deepEqual(sizeOptions(13), [8, 9, 10, 10.5, 11, 12, 13, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72])
  assert.deepEqual(sizeOptions(7), [7, ...FONT_SIZES])
  assert.deepEqual(sizeOptions(80), [...FONT_SIZES, 80])
})

test('tabLabel drops the .docx extension and keeps any other name', () => {
  assert.equal(tabLabel('Report.docx'), 'Report')
  assert.equal(tabLabel('REPORT.DOCX'), 'REPORT')
  assert.equal(tabLabel('notes.txt'), 'notes.txt')
  assert.equal(tabLabel('.docx'), '.docx', 'a name that is only the extension stays whole')
})

test('nextUntitled takes the plain name first, then the lowest free number', () => {
  assert.equal(nextUntitled([], translate), 'Untitled.docx')
  assert.equal(nextUntitled(['Untitled.docx'], translate), 'Untitled 2.docx')
  assert.equal(nextUntitled(['Untitled.docx', 'Untitled 2.docx', 'Untitled 4.docx'], translate), 'Untitled 3.docx')
  assert.equal(nextUntitled(['Untitled 2.docx', 'a.docx'], translate), 'Untitled.docx', 'a closed first tab frees the plain name')
})

test('afterClose keeps the active tab, or moves to the right neighbor of the closing one', () => {
  assert.equal(afterClose(['a', 'b', 'c'], 'b', 'b'), 'c')
  assert.equal(afterClose(['a', 'b', 'c'], 'c', 'c'), 'b', 'the last tab falls back to its left neighbor')
  assert.equal(afterClose(['a', 'b', 'c'], 'a', 'c'), 'c', 'closing a background tab keeps the active one')
  assert.equal(afterClose(['a'], 'a', 'a'), null, 'nothing is left')
})

test('cycleTab wraps at both ends', () => {
  assert.equal(cycleTab(['a', 'b', 'c'], 'c', 1), 'a')
  assert.equal(cycleTab(['a', 'b', 'c'], 'a', -1), 'c')
  assert.equal(cycleTab(['a', 'b', 'c'], 'b', 1), 'c')
  assert.equal(cycleTab(['a'], 'a', 1), 'a')
})
