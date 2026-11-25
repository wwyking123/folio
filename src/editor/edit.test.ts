import assert from 'node:assert/strict'
import { test } from 'node:test'
import { checkDoc, flatText, indexOf } from '../model.ts'
import type { Block, Doc, Marks, Paragraph, Row, RunLook, Table, TextRun } from '../model.ts'
import {
  apply,
  clearMarks,
  changeIndentLevel,
  COALESCE_MS,
  deleteBackward,
  deleteColumns,
  deleteEnclosingTable,
  deleteForward,
  deleteRows,
  deleteSelection,
  deleteTable,
  HISTORY_LIMIT,
  INSERTED_CELL_MARGINS,
  INSERTED_TABLE_BORDERS,
  initialState,
  insertColumn,
  imageChrome,
  imageMoveIssue,
  insertImage,
  insertLineBreak,
  insertPageBreak,
  insertRow,
  insertTable,
  insertText,
  move,
  replaceAllText,
  replaceCounts,
  replaceMatch,
  moveAtom,
  redo,
  resizeImage,
  selectAll,
  selectAround,
  selectedText,
  setAlign,
  setIndent,
  setLineSpacing,
  setMark,
  setPageSetup,
  setParaFormat,
  setParaStyle,
  setSpacing,
  setSelection,
  splitParagraph,
  tableEditTools,
  toggleList,
  toggleMark,
  toggleVertAlign,
  toolbarState,
  undo,
  wordCount,
} from './edit.ts'
import type { Command, EditorState, Transaction } from './edit.ts'
import {
  cell,
  fieldAtom,
  HEADING,
  hex,
  imageAtom,
  makeDoc,
  mediaKey,
  mulberry32,
  NORMAL,
  numId,
  opaqueAtom,
  opaqueBlock,
  page,
  paragraph,
  paraLook,
  pt,
  randomCommand,
  randomDoc,
  row,
  section,
  storyKey,
  styleId,
  stylesheet,
  table,
  text,
} from './testkit.ts'

function run(state: EditorState, command: Command, now = 0): EditorState {
  const tr = command(state)
  assert.ok(tr, 'command applies')
  const next = apply(state, tr, now)
  checkDoc(next.doc)
  assertNoLoneSurrogate(next.doc)
  return next
}

function refused(state: EditorState, command: Command): void {
  assert.equal(command(state), null, 'command does not apply')
}

function caretAt(state: EditorState, p: Paragraph['id'], o: number): EditorState {
  const tr = setSelection({ kind: 'text', anchor: { p, o }, head: { p, o } })(state)
  if (!tr) return state
  return apply(state, tr, 0)
}

function rangeAt(state: EditorState, anchor: { p: Paragraph['id'], o: number }, head: { p: Paragraph['id'], o: number }): EditorState {
  const tr = setSelection({ kind: 'text', anchor, head })(state)
  assert.ok(tr, 'the selection moves')
  return apply(state, tr, 0)
}

function asParagraph(block: Block | undefined): Paragraph {
  assert.ok(block && block.kind === 'p', 'block is a paragraph')
  return block
}

function asTable(block: Block | undefined): Table {
  assert.ok(block && block.kind === 'table', 'block is a table')
  return block
}

function textRuns(para: Paragraph): TextRun[] {
  return para.inlines.filter((inline): inline is TextRun => inline.kind === 'text')
}

function look(para: Paragraph, marks: Marks): RunLook {
  const heading = para.props.style === HEADING
  return {
    fonts: { ascii: marks.font ?? 'Calibri', hAnsi: 'Calibri', eastAsia: marks.fontEastAsia ?? '宋体', cs: 'Calibri' },
    sizePt: marks.sizePt ?? pt(heading ? 16 : 12),
    bold: marks.bold === undefined ? heading : marks.bold,
    italic: marks.italic === true,
    underline: marks.underline === true,
    strike: marks.strike === true,
    color: marks.color ?? 'auto',
    highlight: marks.highlight,
    caps: false,
    charSpacingPt: pt(0),
  }
}

test('typing extends the run at the caret and keeps its src', () => {
  const para = paragraph('p', [text('ab', 7)], {}, { src: 1 })
  const other = paragraph('q', [text('stay', 8)], {}, { src: 2 })
  const doc = makeDoc([para, other])
  const state = run(caretAt(initialState(doc), para.id, 2), insertText('c'))
  const edited = asParagraph(state.doc.body[0])
  assert.equal(textRuns(edited).length, 1, 'the typed letter joins the run')
  assert.equal(textRuns(edited)[0]?.text, 'abc', 'the letter is appended')
  assert.equal(textRuns(edited)[0]?.src, para.inlines[0] && para.inlines[0].kind === 'text' ? para.inlines[0].src : null, 'the run src is reused')
  assert.equal(state.doc.body[1], other, 'the other paragraph is the same object')
})

test('stored marks style only the text typed after the caret', () => {
  const para = paragraph('p', [text('ac', 1)], {}, { src: 2 })
  let state = caretAt(initialState(makeDoc([para])), para.id, 1)
  const marked = toggleMark('bold')(state)
  assert.ok(marked, 'toggling at a caret records stored marks')
  assert.equal(marked.doc, state.doc, 'a collapsed toggle does not change the document')
  state = apply(state, marked, 0)
  assert.equal(state.stored?.bold, true, 'bold is stored for the next typed text')
  state = run(state, insertText('b'), 10)
  const runs = textRuns(asParagraph(state.doc.body[0]))
  assert.deepEqual(runs.map(item => [item.text, item.marks.bold]), [['a', undefined], ['b', true], ['c', undefined]], 'only the new letter is bold')
  const neighbors = runs.filter(item => item.text !== 'bB' && item.text !== 'b')
  state = run(state, insertText('B'), 20)
  const again = textRuns(asParagraph(state.doc.body[0]))
  assert.deepEqual(again.map(item => item.text), ['a', 'bB', 'c'], 'more typing extends the new run')
  assert.equal(again[0], neighbors[0], 'the left neighbor run is not rewritten')
  assert.equal(again[2], neighbors[1], 'the right neighbor run is not rewritten')
  assert.equal(state.stored, null, 'typing clears stored marks')
})

test('typing after an atom starts a run with a null src', () => {
  const atom = opaqueAtom('sym', '§', true, 2)
  const para = paragraph('p', [text('a', 1), atom], {}, { src: 3 })
  const state = run(caretAt(initialState(makeDoc([para])), para.id, 2), insertText('z'))
  const edited = asParagraph(state.doc.body[0])
  assert.equal(edited.inlines[1], atom, 'the atom object stays')
  const typed = edited.inlines[2]
  assert.ok(typed && typed.kind === 'text' && typed.text === 'z', 'the letter follows the atom')
  assert.equal(typed.src, null, 'the new run does not reuse the src from before the atom')
})

test('a line break stays inside the paragraph and a page break is an atom', () => {
  const para = paragraph('p', [text('ab', 1)], {}, { src: 2 })
  let state = run(caretAt(initialState(makeDoc([para])), para.id, 1), insertLineBreak)
  assert.equal(flatText(asParagraph(state.doc.body[0])), 'a\nb', 'shift-enter inserts a line break')
  assert.equal(state.doc.body.length, 1, 'a line break does not split the paragraph')
  state = run(state, insertPageBreak)
  const atom = asParagraph(state.doc.body[0]).inlines.find(item => item.kind === 'atom')
  assert.ok(atom && atom.kind === 'atom' && atom.atom.type === 'pageBreak', 'ctrl-enter inserts a page break')
  assert.equal(atom.src, null, 'a new page break has no source run')
})

test('enter splits a heading at the start, middle, and end', () => {
  const title = paragraph('h', [text('Title', 2)], { style: HEADING }, { src: 1 })
  const start = run(initialState(makeDoc([title])), splitParagraph)
  assert.equal(start.doc.body[1], title, 'enter at the start keeps the original paragraph')
  assert.equal(flatText(asParagraph(start.doc.body[0])), '', 'the new paragraph above is empty')
  assert.equal(asParagraph(start.doc.body[0]).props.style, HEADING, 'the new paragraph keeps the heading style')
  assert.equal(asParagraph(start.doc.body[0]).src, title.src, 'the new paragraph copies the paragraph src')

  const middle = run(caretAt(initialState(makeDoc([title])), title.id, 2), splitParagraph)
  assert.equal(asParagraph(middle.doc.body[0]).id, title.id, 'the head keeps the paragraph id')
  assert.equal(flatText(asParagraph(middle.doc.body[0])), 'Ti', 'the head keeps the text before the caret')
  assert.equal(flatText(asParagraph(middle.doc.body[1])), 'tle', 'the tail keeps the text after the caret')
  assert.equal(asParagraph(middle.doc.body[1]).props.style, HEADING, 'enter in the middle does not take the next style')
  assert.equal(asParagraph(middle.doc.body[0]).src, title.src, 'the head keeps the src')
  assert.equal(asParagraph(middle.doc.body[1]).src, title.src, 'the tail copies the src')

  const end = run(caretAt(initialState(makeDoc([title])), title.id, 5), splitParagraph)
  assert.equal(flatText(asParagraph(end.doc.body[0])), 'Title', 'the heading text stays put')
  assert.equal(asParagraph(end.doc.body[1]).props.style, NORMAL, 'enter at the end of a heading takes the next style')
  assert.equal(asParagraph(end.doc.body[1]).src, title.src, 'the following paragraph copies the src')

  const empty = paragraph('h', [], { style: HEADING }, { src: 1 })
  const opened = run(initialState(makeDoc([empty])), splitParagraph)
  assert.equal(opened.doc.body[0], empty, 'enter in an empty heading keeps that paragraph')
  assert.equal(asParagraph(opened.doc.body[1]).props.style, NORMAL, 'the following paragraph takes the next style')
})

test('a pasted newline does not apply the heading next style', () => {
  const title = paragraph('h', [text('Title', 2)], { style: HEADING }, { src: 1 })
  const state = run(caretAt(initialState(makeDoc([title])), title.id, 5), insertText('\n'))
  assert.equal(asParagraph(state.doc.body[1]).props.style, HEADING, 'pasted newlines keep the current style')
})

test('backspace and delete merge with the side mergeSide picks', () => {
  const heading = paragraph('h', [text('Hello', 2)], { style: HEADING, align: 'center' }, { src: 1 })
  const next = paragraph('n', [text('world', 4)], { style: NORMAL, align: 'right' }, { src: 3 })
  const keptFirst = run(caretAt(initialState(makeDoc([heading, next])), next.id, 0), deleteBackward('char'))
  const merged = asParagraph(keptFirst.doc.body[0])
  assert.equal(merged.id, heading.id, 'backspace into a non-empty paragraph keeps its id')
  assert.equal(flatText(merged), 'Helloworld', 'the texts join')
  assert.equal(merged.props.align, 'center', 'a deletion that starts later keeps the first paragraph')
  assert.equal(merged.src, heading.src, 'the surviving side keeps its src')

  const empty = paragraph('h', [], { style: HEADING, align: 'center' }, { src: 1 })
  const keptSecond = run(caretAt(initialState(makeDoc([empty, next])), next.id, 0), deleteBackward('char'))
  const survivor = asParagraph(keptSecond.doc.body[0])
  assert.equal(survivor.id, next.id, 'backspace through an empty paragraph keeps the next id')
  assert.equal(survivor.props.align, 'right', 'an empty first paragraph gives way to the second')
  assert.equal(flatText(survivor), 'world', 'the next paragraph text stays')

  const whole = rangeAt(initialState(makeDoc([heading, next])), { p: heading.id, o: 0 }, { p: next.id, o: 0 })
  const removed = run(whole, deleteSelection)
  const rest = asParagraph(removed.doc.body[0])
  assert.equal(rest.id, next.id, 'deleting a whole heading removes that paragraph')
  assert.equal(rest.props.align, 'right', 'deleting a whole heading keeps the next paragraph props')
  assert.equal(flatText(rest), 'world', 'the heading text is gone')

  const forward = run(caretAt(initialState(makeDoc([heading, next])), heading.id, 5), deleteForward('char'))
  const joined = asParagraph(forward.doc.body[0])
  assert.equal(joined.id, heading.id, 'delete at the end keeps the first paragraph')
  assert.equal(joined.props.align, 'center', 'delete at the end keeps the first props')
  assert.equal(flatText(joined), 'Helloworld', 'delete joins the next paragraph')

  const unsourced = paragraph('a', [text('Hello')], { align: 'center' })
  const sourced = paragraph('b', [text('world')], { align: 'right' }, { src: 4 })
  const fallback = run(caretAt(initialState(makeDoc([unsourced, sourced])), sourced.id, 0), deleteBackward('char'))
  assert.equal(asParagraph(fallback.doc.body[0]).src, sourced.src, 'a null surviving src falls back to the other side')
})

test('delete across paragraphs, emoji, and atoms', () => {
  const first = paragraph('a', [text('abcd', 1)], {}, { src: 2 })
  const second = paragraph('b', [text('efgh', 3)], {}, { src: 4 })
  const across = run(
    rangeAt(initialState(makeDoc([first, second])), { p: first.id, o: 2 }, { p: second.id, o: 2 }),
    deleteSelection,
  )
  assert.equal(flatText(asParagraph(across.doc.body[0])), 'abgh', 'a cross-paragraph selection keeps the head and the tail')
  assert.equal(asParagraph(across.doc.body[0]).props, first.props, 'the deletion starts inside the first paragraph')

  const emoji = paragraph('e', [text('a😀b')])
  const afterEmoji = run(caretAt(initialState(makeDoc([emoji])), emoji.id, 3), deleteBackward('char'))
  assert.equal(flatText(asParagraph(afterEmoji.doc.body[0])), 'ab', 'backspace removes both units of an emoji')
  const inside = run(caretAt(initialState(makeDoc([emoji])), emoji.id, 2), deleteBackward('char'))
  assert.equal(flatText(asParagraph(inside.doc.body[0])), 'ab', 'backspace inside an emoji removes the whole grapheme')

  const chinese = paragraph('c', [text('你好世界')])
  const word = run(caretAt(initialState(makeDoc([chinese])), chinese.id, 4), deleteBackward('word'))
  assert.equal(flatText(asParagraph(word.doc.body[0])), '你好', 'a word backspace removes the last Chinese word')
  const latin = paragraph('l', [text('hello world')])
  const latinWord = run(caretAt(initialState(makeDoc([latin])), latin.id, 11), deleteBackward('word'))
  assert.equal(flatText(asParagraph(latinWord.doc.body[0])), 'hello ', 'a word backspace removes the last Latin word')
  const forwardWord = run(initialState(makeDoc([latin])), deleteForward('word'))
  assert.equal(flatText(asParagraph(forwardWord.doc.body[0])), 'world', 'a word delete removes through the following space')

  const fixed = opaqueAtom('eq', '∫', false, 2)
  const host = paragraph('h', [text('ab', 1), fixed, text('cd', 3)], {}, { src: 4 })
  const tail = paragraph('t', [text('ef', 5)], {}, { src: 6 })
  const kept = run(
    rangeAt(initialState(makeDoc([host, tail])), { p: host.id, o: 1 }, { p: tail.id, o: 1 }),
    deleteSelection,
  )
  const pieces = asParagraph(kept.doc.body[0]).inlines
  assert.equal(pieces[1], fixed, 'a non-removable atom inside the selection stays the same object')
  assert.equal(flatText(asParagraph(kept.doc.body[0])), 'a\uFFFCf', 'text around the atom is deleted')
  const onlyAtom = rangeAt(initialState(makeDoc([host])), { p: host.id, o: 2 }, { p: host.id, o: 3 })
  refused(onlyAtom, deleteSelection)
  assert.equal(onlyAtom.doc.body[0], host, 'deleting only a non-removable atom leaves the document')

  const image = imageAtom('k', 10, 10, 2)
  const pictured = paragraph('p', [text('ab', 1), image, text('cd', 3)], {}, { src: 4 })
  const dropped = run(rangeAt(initialState(makeDoc([pictured])), { p: pictured.id, o: 2 }, { p: pictured.id, o: 3 }), deleteSelection)
  assert.equal(flatText(asParagraph(dropped.doc.body[0])), 'abcd', 'a removable atom is deleted')
})

test('locked paragraphs, section marks, and the document edges refuse the edit', () => {
  const locked = paragraph('p', [text('no', 1)], {}, { src: 2, lock: { reason: 'revisions' } })
  const lockedDoc = makeDoc([locked])
  const lockedState = initialState(lockedDoc)
  refused(lockedState, insertText('a'))
  refused(lockedState, deleteForward('char'))
  assert.equal(lockedState.doc, lockedDoc, 'a locked paragraph is never replaced')

  const ended = paragraph('a', [text('aa', 1)], {}, { src: 2, sectionEnd: 0 })
  const follow = paragraph('b', [text('bb', 3)], {}, { src: 4 })
  const edge = caretAt(initialState(makeDoc([ended, follow], { sections: [section(), section(page(500))] })), ended.id, 2)
  refused(edge, deleteForward('char'))
  assert.equal(edge.doc.body[0], ended, 'delete does not merge across a section mark')

  const only = paragraph('p', [text('ab')])
  refused(initialState(makeDoc([only])), deleteBackward('char'))
  refused(caretAt(initialState(makeDoc([only])), only.id, 2), deleteForward('char'))
})

test('lists toggle, join a neighbor, and backspace removes numbering before merging', () => {
  const item = paragraph('p', [text('item')])
  let state = run(initialState(makeDoc([item])), toggleList('bullet'))
  const numbering = asParagraph(state.doc.body[0]).props.numbering
  assert.ok(numbering, 'the paragraph joins a list')
  assert.equal(numbering.num, numId(-1), 'a new list gets the next negative num id')
  assert.equal(numbering.level, 0, 'a new list starts at level 0')
  const def = state.doc.lists.byNum.get(numbering.num)
  assert.equal(def?.fresh, true, 'the list definition is fresh')
  assert.equal(def?.levels.length, 9, 'the list has nine levels')
  assert.equal(def?.levels[0]?.text, '•', 'the first level is a bullet')
  assert.equal(def?.levels[0]?.indentLeft, pt(21), 'the first level indents 21pt')
  assert.equal(def?.levels[0]?.hanging, pt(21), 'the marker hangs back to the margin')
  state = run(state, toggleList('bullet'))
  assert.equal(asParagraph(state.doc.body[0]).props.numbering, null, 'toggling the same kind removes numbering')

  const left = paragraph('a', [text('one')])
  const right = paragraph('b', [text('two')])
  state = run(initialState(makeDoc([left, right])), toggleList('decimal'))
  const shared = asParagraph(state.doc.body[0]).props.numbering
  state = run(caretAt(state, right.id, 0), toggleList('decimal'))
  assert.deepEqual(asParagraph(state.doc.body[1]).props.numbering, shared, 'the next paragraph joins the neighbor list')
  state = run(caretAt(state, right.id, 0), deleteBackward('char'))
  assert.equal(asParagraph(state.doc.body[1]).props.numbering, null, 'backspace at offset 0 clears the list first')
  assert.equal(state.doc.body.length, 2, 'clearing the list does not merge')
  state = run(state, deleteBackward('char'))
  assert.equal(state.doc.body.length, 1, 'the next backspace merges the paragraphs')
  assert.equal(flatText(asParagraph(state.doc.body[0])), 'onetwo', 'the list item text is kept')

  const styles = stylesheet()
  const defs = new Map(styles.paragraph)
  const inherited = numId(4)
  defs.set(styleId('List'), { name: 'List', role: null, next: null, look: paraLook({ num: inherited, level: 0 }) })
  const lists = {
    byNum: new Map([[inherited, { fresh: false, levels: [{ format: 'decimal' as const, text: '%1.', start: 1, indentLeft: pt(36), hanging: pt(18) }] }]]),
  }
  const prev = paragraph('a', [text('prev')])
  const styled = paragraph('b', [text('item')], { style: styleId('List') })
  const styledDoc = makeDoc([prev, styled], { styles: { ...styles, paragraph: defs }, lists })
  const cleared = run(caretAt(initialState(styledDoc), styled.id, 0), deleteBackward('char'))
  assert.equal(asParagraph(cleared.doc.body[1]).props.numbering, null, 'backspace clears numbering inherited from the style')
  assert.equal(cleared.doc.body.length, 2, 'the styled item stays its own paragraph')

  const explicit = paragraph('b', [text('item')], { style: styleId('List'), numbering: null })
  const merged = run(caretAt(initialState(makeDoc([prev, explicit], { styles: { ...styles, paragraph: defs }, lists })), explicit.id, 0), deleteBackward('char'))
  assert.equal(merged.doc.body.length, 1, 'explicit numbering null is not a list, so backspace merges')
})

test('paragraph formatting updates style, alignment, spacing, and indent', () => {
  const para = paragraph('p', [text('ab')], { indent: { firstLineChars: 200 } })
  let state = run(initialState(makeDoc([para])), setParaStyle(HEADING))
  assert.equal(asParagraph(state.doc.body[0]).props.style, HEADING, 'the paragraph takes the heading style')
  state = run(state, setParaStyle(styleId('heading2')))
  assert.ok(state.doc.styles.added.includes('heading2'), 'a missing built-in role is recorded')
  assert.equal(state.doc.styles.paragraph.get(styleId('heading2'))?.next, NORMAL, 'a heading next style is the default paragraph')
  state = run(state, setAlign('center'))
  assert.equal(asParagraph(state.doc.body[0]).props.align, 'center', 'alignment is set')
  state = run(state, setLineSpacing({ rule: 'exact', pt: pt(16) }))
  assert.deepEqual(asParagraph(state.doc.body[0]).props.spacing?.line, { rule: 'exact', pt: pt(16) }, 'line spacing is set')
  state = run(state, setIndent({ left: pt(36) }))
  const indent = asParagraph(state.doc.body[0]).props.indent
  assert.equal(indent?.left, pt(36), 'left indent is set')
  assert.equal(indent?.firstLineChars, 200, 'an omitted firstLineChars value is kept')
  state = run(initialState(makeDoc([paragraph('p', [text('ab')])])), changeIndentLevel(1))
  assert.equal(asParagraph(state.doc.body[0]).props.indent?.left, pt(36), 'indent level adds 36pt')
  state = run(state, changeIndentLevel(1))
  assert.equal(asParagraph(state.doc.body[0]).props.indent?.left, pt(72), 'a second increase adds another 36pt')
  state = run(state, changeIndentLevel(-1))
  state = run(state, changeIndentLevel(-1))
  state = run(state, changeIndentLevel(-1))
  assert.equal(asParagraph(state.doc.body[0]).props.indent?.left, pt(0), 'indent does not go below zero')
  state = run(state, toggleList('decimal'))
  assert.equal(asParagraph(state.doc.body[0]).props.indent?.left, undefined, 'entering a list clears direct left indent so the list level indents')

  let listed = run(initialState(makeDoc([paragraph('p', [text('ab')])])), toggleList('bullet'))
  listed = run(listed, changeIndentLevel(1))
  assert.equal(asParagraph(listed.doc.body[0]).props.numbering?.level, 1, 'indent level raises the list level')
  listed = run(listed, changeIndentLevel(-1))
  listed = run(listed, changeIndentLevel(-1))
  assert.equal(asParagraph(listed.doc.body[0]).props.numbering, null, 'indent below level 0 clears the list')
})

test('range marks turn on everywhere when any covered text lacks the mark', () => {
  const para = paragraph('p', [text('a', 1, { bold: true }), text('b', 2)], {}, { src: 3 })
  let state = rangeAt(initialState(makeDoc([para])), { p: para.id, o: 0 }, { p: para.id, o: 2 })
  state = run(state, toggleMark('bold'))
  assert.deepEqual(textRuns(asParagraph(state.doc.body[0])).map(item => item.marks.bold), [true, true], 'a mixed selection sets the mark on every run')
  state = run(state, toggleMark('bold'))
  assert.deepEqual(textRuns(asParagraph(state.doc.body[0])).map(item => item.marks.bold), [false, false], 'turning the mark off writes an explicit false')
  const plain = paragraph('q', [text('abcd', 1)], {}, { src: 2 })
  state = rangeAt(initialState(makeDoc([plain])), { p: plain.id, o: 1 }, { p: plain.id, o: 3 })
  state = run(state, setMark({ mark: 'font', value: 'Songti' }))
  assert.deepEqual(textRuns(asParagraph(state.doc.body[0])).map(item => item.text), ['a', 'bc', 'd'], 'setMark splits the run at the selection')
  assert.equal(textRuns(asParagraph(state.doc.body[0]))[1]?.marks.font, 'Songti', 'the covered piece takes the font')
  assert.equal(textRuns(asParagraph(state.doc.body[0]))[1]?.src, runSrcOf(plain), 'the split pieces keep the run src')
  const collapsed = setMark({ mark: 'color', value: hex('FF0000') })(caretAt(state, plain.id, 0))
  assert.ok(collapsed, 'setMark at a caret stores the value')
  assert.equal(collapsed.doc, state.doc, 'setMark at a caret does not change the document')
  assert.equal(collapsed.stored?.color, hex('FF0000'), 'the color is stored')
})

test('clearMarks removes direct character formatting and keeps the link', () => {
  const link = 'https://example.com/' as Marks['link']
  const para = paragraph('p', [
    text('ab', 1, { bold: true, color: hex('C00000'), sizePt: pt(20), vertAlign: 'super', link, charStyle: HEADING }),
    text('c', 2, { italic: true }),
  ])
  const doc = makeDoc([para])
  let state = rangeAt(initialState(doc), { p: para.id, o: 0 }, { p: para.id, o: 2 })
  state = run(state, clearMarks)
  const runs = textRuns(asParagraph(state.doc.body[0]))
  assert.deepEqual(runs[0]?.marks, { link }, 'a range clear keeps only the link')
  assert.equal(runs[1]?.marks.italic, true, 'text outside the range stays')
  assert.equal(state.history.past.length, 1, 'a range clear is one undo entry')
  const undone = undo(state)
  assert.ok(undone)
  assert.equal(undone.doc, doc, 'undo restores the original document')
  checkDoc(state.doc)

  const caret = clearMarks(initialState(doc))
  assert.ok(caret)
  assert.equal(caret.doc, doc, 'a caret clear does not change the document')
  assert.deepEqual(caret.stored, {}, 'a caret clear stores empty marks')
  assert.equal(caret.intent, 'none', 'a caret clear is not an undo entry')
  assert.equal(clearMarks(apply(initialState(doc), caret, 0)), null, 'clearing empty stored marks changes nothing')

  const styled = paragraph('s', [text('ab', 1, { bold: true })], { style: HEADING, align: 'center', spacing: { before: pt(6) } })
  const styledDoc = makeDoc([styled])
  const cleared = run(rangeAt(initialState(styledDoc), { p: styled.id, o: 0 }, { p: styled.id, o: 2 }), clearMarks)
  assert.equal(asParagraph(cleared.doc.body[0]).props, styled.props, 'clear formatting does not touch paragraph properties')
  assert.equal(textRuns(asParagraph(cleared.doc.body[0]))[0]?.marks.bold, undefined)
})

test('setSpacing and setParaFormat are one undo entry each', () => {
  const para = paragraph('p', [text('ab')], {
    align: 'left',
    indent: { left: pt(12), firstLineChars: 200 },
    spacing: { before: pt(3), line: { rule: 'auto', multiple: 1 } },
  })
  const doc = makeDoc([para])
  let state = run(initialState(doc), setSpacing({ before: pt(12), after: pt(6) }))
  const spacing = asParagraph(state.doc.body[0]).props.spacing
  assert.equal(spacing?.before, pt(12))
  assert.equal(spacing?.after, pt(6))
  assert.deepEqual(spacing?.line, { rule: 'auto', multiple: 1 }, 'setSpacing leaves the line rule')
  assert.equal(asParagraph(state.doc.body[0]).props.indent?.firstLineChars, 200, 'setSpacing leaves the indent')
  assert.equal(state.history.past.length, 1)
  assert.equal(undo(state)?.doc, doc, 'undo restores the original document')
  assert.equal(setSpacing({})(initialState(doc)), null, 'an empty spacing patch changes nothing')

  state = run(initialState(doc), setParaFormat({
    align: 'justify',
    indent: { left: pt(18), firstLine: pt(24), hanging: null, firstLineChars: null },
    spacing: { before: pt(6), after: pt(6), line: { rule: 'exact', pt: pt(20) } },
  }))
  const props = asParagraph(state.doc.body[0]).props
  assert.equal(props.align, 'justify')
  assert.equal(props.indent?.left, pt(18))
  assert.equal(props.indent?.firstLine, pt(24))
  assert.equal(props.indent?.hanging, undefined)
  assert.equal(props.indent?.firstLineChars, undefined)
  assert.equal(props.spacing?.before, pt(6))
  assert.equal(props.spacing?.after, pt(6))
  assert.deepEqual(props.spacing?.line, { rule: 'exact', pt: pt(20) })
  assert.equal(state.history.past.length, 1, 'the dialog applies one undo entry')
  const restored = undo(state)
  assert.ok(restored)
  assert.equal(restored.doc, doc, 'undo restores the original document')
  checkDoc(state.doc)
  assert.equal(setParaFormat({})(initialState(doc)), null, 'an empty paragraph patch changes nothing')
})

test('superscript and subscript replace each other', () => {
  const para = paragraph('p', [text('abcd')])
  let state = rangeAt(initialState(makeDoc([para])), { p: para.id, o: 1 }, { p: para.id, o: 3 })
  state = run(state, toggleVertAlign('super'))
  assert.equal(textRuns(asParagraph(state.doc.body[0]))[1]?.marks.vertAlign, 'super')
  state = run(state, toggleVertAlign('sub'))
  assert.equal(textRuns(asParagraph(state.doc.body[0]))[1]?.marks.vertAlign, 'sub', 'subscript replaces superscript')
  state = run(state, toggleVertAlign('sub'))
  assert.equal(textRuns(asParagraph(state.doc.body[0]))[1]?.marks.vertAlign, undefined, 'choosing the active one clears it')
})

test('tables insert, grow, shrink, and clear cells without dropping structure', () => {
  const para = paragraph('p', [text('abcd', 1)], {}, { src: 2 })
  const middle = run(caretAt(initialState(makeDoc([para])), para.id, 2), insertTable(2, 2))
  assert.equal(flatText(asParagraph(middle.doc.body[0])), 'ab', 'a table in the middle keeps the head')
  const inserted = asTable(middle.doc.body[1])
  assert.deepEqual(inserted.props.borders, INSERTED_TABLE_BORDERS, 'an inserted table has the borders save writes')
  assert.deepEqual(inserted.props.margins, INSERTED_CELL_MARGINS, 'an inserted table keeps the old 3px 5px padding')
  assert.equal(inserted.rows.length, 2, 'the table has the requested rows')
  assert.equal(inserted.grid.length, 2, 'the table has the requested columns')
  const firstCell = inserted.rows[0]?.cells[0]?.blocks[0]
  assert.ok(firstCell && firstCell.kind === 'p', 'the first cell has a paragraph')
  assert.equal(middle.sel.kind, 'text', 'insert leaves a text caret')
  if (firstCell && firstCell.kind === 'p' && middle.sel.kind === 'text') {
    assert.equal(middle.sel.head.p, firstCell.id, 'the caret is in the first cell')
    assert.equal(middle.sel.head.o, 0, 'the caret is at the start of that cell')
    assert.equal(middle.sel.anchor.p, firstCell.id, 'the caret is collapsed')
    assert.equal(middle.sel.anchor.o, 0, 'the collapsed caret stays at the start')
  }
  assert.equal(inserted.grid[0], pt(234), 'columns share the content width')
  const gutter = page()
  const withGutter = { ...gutter, margin: { ...gutter.margin, gutter: pt(36) } }
  const guttered = run(caretAt(initialState(makeDoc([para], { sections: [section(withGutter)] })), para.id, 4), insertTable(1, 2))
  assert.equal(asTable(guttered.doc.body[1]).grid[0], pt(216), 'columns share the text width')
  assert.equal(flatText(asParagraph(middle.doc.body[2])), 'cd', 'the tail follows the table')
  assert.equal(asParagraph(middle.doc.body[0]).src, para.src, 'the head keeps the paragraph src')
  assert.equal(asParagraph(middle.doc.body[2]).src, para.src, 'the tail keeps the paragraph src')
  for (const entry of inserted.rows) for (const item of entry.cells) {
    assert.equal(item.blocks[item.blocks.length - 1]?.kind, 'p', 'every cell ends with a paragraph')
  }

  const atEnd = run(caretAt(initialState(makeDoc([para])), para.id, 4), insertTable(1, 1))
  assert.equal(atEnd.doc.body[0], para, 'a table after the paragraph keeps that paragraph object')
  const atStart = run(initialState(makeDoc([para])), insertTable(1, 1))
  assert.equal(atStart.doc.body[1], para, 'a table before the paragraph keeps that paragraph object')

  const left = paragraph('l', [text('hello', 2)], {}, { src: 1 })
  const right = paragraph('r', [text('world', 5)], {}, { src: 4 })
  const blank = paragraph('b', [], {}, { src: 3 })
  const grid = table('t', [row([cell([left]), cell([blank]), cell([right])])])
  const cleared = run(
    rangeAt(initialState(makeDoc([grid])), { p: left.id, o: 1 }, { p: right.id, o: 4 }),
    deleteSelection,
  )
  const kept = asTable(cleared.doc.body[0])
  assert.equal(kept.id, grid.id, 'clearing cells keeps the table')
  assert.equal(kept.rows.length, 1, 'clearing cells keeps the rows')
  assert.equal(kept.rows[0]?.cells.length, 3, 'clearing cells keeps the columns')
  assert.equal(flatText(asParagraph(kept.rows[0]?.cells[0]?.blocks[0])), '', 'a covered cell is emptied')
  assert.equal(flatText(asParagraph(kept.rows[0]?.cells[2]?.blocks[0])), '', 'the other covered cell is emptied')

  const partial = run(rangeAt(initialState(makeDoc([grid])), { p: left.id, o: 1 }, { p: left.id, o: 4 }), deleteSelection)
  assert.equal(flatText(asParagraph(asTable(partial.doc.body[0]).rows[0]?.cells[0]?.blocks[0])), 'ho', 'a selection inside one cell deletes text')
  assert.equal(asTable(partial.doc.body[0]).rows[0]?.cells[2], grid.rows[0]?.cells[2], 'an untouched cell is the same object')

  const lead = paragraph('a', [text('abcd', 2)], {}, { src: 1 })
  const inner = paragraph('c', [text('x', 4)], {}, { src: 3 })
  const boxed = table('box', [row([cell([inner])])])
  const trail = paragraph('b', [text('efgh', 6)], {}, { src: 5 })
  const gone = run(
    rangeAt(initialState(makeDoc([lead, boxed, trail])), { p: lead.id, o: 2 }, { p: trail.id, o: 2 }),
    deleteSelection,
  )
  assert.equal(gone.doc.body.length, 1, 'a table wholly inside the selection is removed')
  assert.equal(flatText(asParagraph(gone.doc.body[0])), 'abgh', 'the paragraphs on either side join')

  const top = row([cell([paragraph('t', [text('top')])])])
  const bottom = row([cell([paragraph('b', [text('bottom')])])])
  const rows = table('rows', [top, bottom], [72])
  const deletedRow = run(initialState(makeDoc([rows])), deleteRows())
  const remaining = asTable(deletedRow.doc.body[0])
  assert.equal(remaining.rows.length, 1, 'deleteRows removes the caret row')
  assert.equal(remaining.rows[0], bottom, 'the other row is the same object')
  const removed = run(deletedRow, deleteRows())
  assert.equal(removed.doc.body[0]?.kind, 'p', 'deleting the last row removes the table')

  const grown = run(initialState(makeDoc([rows])), insertRow('below'))
  assert.equal(asTable(grown.doc.body[0]).rows.length, 3, 'insertRow adds a row')
  assert.equal(asTable(grown.doc.body[0]).rows[0], top, 'insertRow keeps the other rows')
  assert.deepEqual(asTable(grown.doc.body[0]).rows[2]?.props, {}, 'a row inserted beside one with no height has none')

  const wide = table('wide', [row([cell([paragraph('a', [text('a')])]), cell([paragraph('b', [text('b')])])])], [72, 72])
  const added = run(initialState(makeDoc([wide])), insertColumn('right'))
  assert.equal(asTable(added.doc.body[0]).grid.length, 3, 'insertColumn adds a column')
  assert.equal(asTable(added.doc.body[0]).rows[0]?.cells[0], wide.rows[0]?.cells[0], 'the existing cell object stays')
  const narrowed = run(added, deleteColumns())
  assert.equal(asTable(narrowed.doc.body[0]).grid.length, 2, 'deleteColumns removes a column')
  const only = paragraph('only', [text('z')])
  const single = table('one', [row([cell([only])])], [72])
  const clearedCol = run(initialState(makeDoc([single])), deleteColumns())
  assert.equal(clearedCol.doc.body[0]?.kind, 'p', 'deleting the last column removes the table')
  const beside = run(caretAt(initialState(makeDoc([para, single])), only.id, 0), deleteColumns())
  assert.equal(beside.doc.body.length, 1, 'the last column takes the table with it')
  assert.equal(beside.doc.body[0], para, 'the paragraph outside the table stays')
  const exact: Row = { ...row([cell([paragraph('h', [text('h')])])]), props: { heightPt: pt(36), heightRule: 'exact' } }
  const copied = run(initialState(makeDoc([table('h', [exact])])), insertRow('below'))
  const copiedRows = asTable(copied.doc.body[0]).rows
  assert.equal(copiedRows[1]?.props.heightPt, pt(36), 'insertRow copies the neighbour height')
  assert.equal(copiedRows[1]?.props.heightRule, 'exact', 'insertRow copies the neighbour rule')
  assert.equal(copiedRows[1]?.props.header, undefined, 'insertRow does not copy the header flag')
  const least: Row = { ...row([cell([paragraph('m', [text('m')])])]), props: { heightPt: pt(20), heightRule: 'atLeast' } }
  const copiedLeast = run(initialState(makeDoc([table('m', [least])])), insertRow('above'))
  assert.equal(asTable(copiedLeast.doc.body[0]).rows[0]?.props.heightRule, 'atLeast')

  const dropped = run(initialState(makeDoc([para, wide])), deleteTable(wide.id))
  assert.equal(dropped.doc.body.length, 1, 'deleteTable removes the table')
  assert.equal(dropped.doc.body[0], para, 'the paragraph outside the table stays')
  const host = wide.rows[0]?.cells[0]?.blocks[0]
  assert.ok(host && host.kind === 'p')
  if (host && host.kind === 'p') {
    const droppedByCaret = run(caretAt(initialState(makeDoc([para, wide])), host.id, 0), deleteEnclosingTable())
    assert.equal(droppedByCaret.doc.body.length, 1, 'deleteEnclosingTable removes the table around the caret')
    assert.equal(droppedByCaret.doc.body[0], para, 'the paragraph outside the table stays')
  }
})

test('column edits rewrite span inside a merged cell, so those tools turn off; row edits and a vertical merge do not', () => {
  const widePara = paragraph('wide', [text('wide')])
  const side = paragraph('side', [text('side')])
  const narrow = paragraph('narrow', [text('narrow')])
  const mid = paragraph('mid', [text('mid')])
  const end = paragraph('end', [text('end')])
  const horizontal = table('h', [
    row([{ ...cell([widePara], 2), vMerge: 'restart' }, cell([side])]),
    row([cell([narrow]), cell([mid]), cell([end])]),
  ], [72, 72, 72])
  assert.deepEqual(tableEditTools(horizontal), { rows: true, columns: false }, 'a span above 1 turns column tools off for the whole table')
  const atMid = caretAt(initialState(makeDoc([horizontal])), mid.id, 0)
  const bar = toolbarState(atMid, look)
  assert.equal(bar.inTable, true, 'the caret is in the table')
  assert.equal(bar.rowTools, true, 'row tools stay available')
  assert.equal(bar.columnTools, false, 'column tools follow tableEditTools')

  const grown = run(atMid, insertColumn('left'))
  const grownCell = asTable(grown.doc.body[0]).rows[0]?.cells[0]
  assert.equal(grownCell?.span, 3, 'insertColumn inside a span rewrites that span')
  assert.equal(grownCell?.vMerge, 'restart', 'insertColumn leaves vMerge as it was')
  assert.equal(flatText(asParagraph(grownCell?.blocks[0])), 'wide', 'the rewritten cell is the merged one')

  const shrunk = run(atMid, deleteColumns())
  const shrunkCell = asTable(shrunk.doc.body[0]).rows[0]?.cells[0]
  assert.equal(shrunkCell?.span, 1, 'deleteColumns inside a span rewrites that span')
  assert.equal(shrunkCell?.vMerge, 'restart', 'deleteColumns leaves vMerge as it was')

  const added = run(caretAt(initialState(makeDoc([horizontal])), widePara.id, 0), insertRow('below'))
  const addedTable = asTable(added.doc.body[0])
  assert.equal(addedTable.rows[0], horizontal.rows[0], 'insertRow does not rewrite the existing row')
  assert.equal(addedTable.rows[1]?.cells[0]?.span, 2, 'the new row copies span')
  assert.equal(addedTable.rows[1]?.cells[0]?.vMerge, null, 'the new row clears vMerge')
  const removedRow = run(caretAt(initialState(makeDoc([horizontal])), widePara.id, 0), deleteRows())
  assert.equal(asTable(removedRow.doc.body[0]).rows[0], horizontal.rows[1], 'deleteRows keeps the other row object')

  const top = paragraph('top', [text('top')])
  const topRight = paragraph('tr', [text('tr')])
  const bottom = paragraph('bot', [text('bot')])
  const bottomRight = paragraph('br', [text('br')])
  const vertical = table('v', [
    row([{ ...cell([top]), vMerge: 'restart' }, cell([topRight])]),
    row([{ ...cell([bottom]), vMerge: 'continue' }, cell([bottomRight])]),
  ], [72, 72])
  assert.deepEqual(tableEditTools(vertical), { rows: true, columns: true }, 'a vertical merge does not turn column tools off')
  const verticalBar = toolbarState(initialState(makeDoc([vertical])), look)
  assert.equal(verticalBar.columnTools, true, 'column tools stay on for a vertical merge')
  assert.equal(verticalBar.rowTools, true, 'row tools stay on for a vertical merge')
  const extended = run(initialState(makeDoc([vertical])), insertColumn('right'))
  assert.equal(asTable(extended.doc.body[0]).rows[0]?.cells[0], vertical.rows[0]?.cells[0], 'insertColumn beside a vertical merge keeps that cell')
  const trimmed = run(caretAt(initialState(makeDoc([vertical])), topRight.id, 0), deleteColumns())
  assert.equal(asTable(trimmed.doc.body[0]).rows[0]?.cells[0], vertical.rows[0]?.cells[0], 'deleteColumns beside a vertical merge keeps that cell')
  assert.equal(asTable(trimmed.doc.body[0]).rows[0]?.cells[0]?.vMerge, 'restart', 'the kept cell keeps its vMerge')
})

test('images scale to the content width and page setup replaces one section', () => {
  const para = paragraph('p', [text('a')])
  const key = mediaKey('pic')
  const doc = makeDoc([para], { media: { byKey: new Map([[key, { contentType: 'image/jpeg', origin: 'pristine' }]]) } })
  const small = run(caretAt(initialState(doc), para.id, 1), insertImage({ key, contentType: 'image/jpeg', widthPx: 200, heightPx: 100 }))
  const image = asParagraph(small.doc.body[0]).inlines[1]
  assert.ok(image && image.kind === 'atom' && image.atom.type === 'image', 'the image is an atom')
  if (image && image.kind === 'atom' && image.atom.type === 'image') {
    assert.equal(image.atom.widthPt, pt(150), 'pixel size converts at 72/96')
    assert.equal(image.atom.heightPt, pt(75), 'height converts with the same scale')
  }
  assert.equal(small.doc.media.byKey.get(key)?.origin, 'pristine', 'an existing media key keeps its origin')
  const fresh = run(caretAt(initialState(makeDoc([para])), para.id, 1), insertImage({ key: mediaKey('new'), contentType: 'image/jpeg', widthPx: 1000, heightPx: 500 }))
  const scaled = asParagraph(fresh.doc.body[0]).inlines[1]
  assert.ok(scaled && scaled.kind === 'atom' && scaled.atom.type === 'image', 'a new image is inserted')
  if (scaled && scaled.kind === 'atom' && scaled.atom.type === 'image') {
    assert.equal(scaled.atom.widthPt, pt(468), 'a wide image scales down to the content width')
    assert.equal(scaled.atom.heightPt, pt(234), 'height scales by the same factor')
  }
  assert.equal(fresh.doc.media.byKey.get(mediaKey('new'))?.origin, 'added', 'a new key is registered as added')
  assert.equal(fresh.doc.media.byKey.get(mediaKey('new'))?.contentType, 'image/jpeg', 'a new image keeps the content type it was given')

  const first = section(page(612))
  const second = section(page(500))
  const setup = run(
    initialState(makeDoc([paragraph('a', [text('a')], {}, { sectionEnd: 0 }), paragraph('b', [text('b')])], { sections: [first, second] })),
    setPageSetup(1, page(700)),
  )
  assert.equal(setup.doc.sections[0], first, 'the other section object stays')
  assert.equal(setup.doc.sections[1]?.page.widthPt, pt(700), 'the chosen section takes the new page')
  assert.equal(setup.doc.sections[1]?.src, second.src, 'page setup keeps the section src')

  const inner = paragraph('c', [text('a')])
  const boxed = table('t', [row([cell([inner])])], [100])
  const inCell = run(caretAt(initialState(makeDoc([boxed])), inner.id, 0), insertImage({
    key: mediaKey('cell'),
    contentType: 'image/jpeg',
    widthPx: 1000,
    heightPx: 500,
  }))
  const cellImage = asParagraph(asTable(inCell.doc.body[0]).rows[0]?.cells[0]?.blocks[0]).inlines[0]
  assert.ok(cellImage && cellImage.kind === 'atom' && cellImage.atom.type === 'image', 'the cell image is an atom')
  if (cellImage && cellImage.kind === 'atom' && cellImage.atom.type === 'image') {
    assert.equal(cellImage.atom.widthPt, pt(100), 'a wide image scales to the cell width')
    assert.equal(cellImage.atom.heightPt, pt(50), 'the cell scale keeps the ratio')
  }
  assert.equal(inCell.doc.media.byKey.get(mediaKey('cell'))?.contentType, 'image/jpeg', 'the cell image keeps its content type')
})

test('a selected picture resizes and moves as one command, and a file picture cannot move', () => {
  const picture = imageAtom('pic', 80, 40)
  const sourced = imageAtom('file', 144, 72, 4)
  const para = paragraph('p', [text('ab', 1), picture, text('cd', 2)], {}, { src: 3 })
  const other = paragraph('q', [text('end', 5)], {}, { src: 6 })
  const locked = paragraph('locked', [picture], {}, { lock: { reason: 'revisions' } })
  const doc = makeDoc([para, other])
  const selected = rangeAt(initialState(doc), { p: para.id, o: 2 }, { p: para.id, o: 3 })
  assert.equal(imageChrome(selected.doc, selected.sel), 'handles', 'one picture shows handles')
  assert.equal(selectedText(selected), '', 'copy takes the range and the picture has no plain text')

  const replaced = run(selected, insertText('Z'))
  assert.equal(flatText(asParagraph(replaced.doc.body[0])), 'abZcd', 'typing replaces the picture')
  const removed = run(selected, deleteBackward('char'))
  assert.equal(flatText(asParagraph(removed.doc.body[0])), 'abcd', 'backspace removes the selected picture')
  const deleted = run(selected, deleteForward('char'))
  assert.equal(flatText(asParagraph(deleted.doc.body[0])), 'abcd', 'delete removes the selected picture')
  const after = run(selected, move('forward', 'char', false))
  assert.deepEqual(after.sel, { kind: 'text', anchor: { p: para.id, o: 3 }, head: { p: para.id, o: 3 } }, 'forward leaves the caret after the picture')
  const before = run(selected, move('backward', 'char', false))
  assert.deepEqual(before.sel, { kind: 'text', anchor: { p: para.id, o: 2 }, head: { p: para.id, o: 2 } }, 'backward leaves the caret before the picture')

  const resized = run(selected, resizeImage({ p: para.id, o: 2 }, 40, 20))
  const resizedAtom = asParagraph(resized.doc.body[0]).inlines[1]
  assert.ok(resizedAtom && resizedAtom.kind === 'atom' && resizedAtom.atom.type === 'image', 'the picture is still an image')
  if (resizedAtom && resizedAtom.kind === 'atom' && resizedAtom.atom.type === 'image') {
    assert.equal(resizedAtom.atom.widthPt, pt(40), 'the width is replaced')
    assert.equal(resizedAtom.atom.heightPt, pt(20), 'the height is replaced')
    assert.equal(resizedAtom.atom.media, picture.atom.type === 'image' ? picture.atom.media : picture.atom, 'the media stays')
    assert.equal(resizedAtom.src, null, 'a fresh picture keeps a null src')
  }
  assert.equal(resized.doc.body[1], other, 'the other paragraph stays')
  assert.deepEqual(resized.sel, { kind: 'text', anchor: { p: para.id, o: 2 }, head: { p: para.id, o: 3 } }, 'the picture stays selected')
  const restored = undo(resized)
  assert.ok(restored, 'undo applies')
  if (restored) {
    assert.equal(asParagraph(restored.doc.body[0]).inlines[1], picture, 'undo restores the original picture object')
    checkDoc(restored.doc)
  }
  const same = resizeImage({ p: para.id, o: 2 }, 80, 40)(selected)
  assert.equal(same?.intent, 'none', 'the starting size adds no undo entry')
  if (same) {
    const held = apply(selected, same, 0)
    assert.equal(held.history, selected.history, 'an unchanged size is not history')
    assert.equal(held.doc, selected.doc, 'an unchanged size keeps the document')
  }
  refused(selected, resizeImage({ p: para.id, o: 2 }, 4, 20))
  refused(selected, resizeImage({ p: para.id, o: 2 }, 500, 20))
  refused(selected, resizeImage({ p: para.id, o: 2 }, 40, 1.5))
  refused(selected, resizeImage({ p: para.id, o: 0 }, 40, 20))
  const short = imageAtom('short', 40, 1.5)
  const shortPara = paragraph('s', [short])
  const keptShort = run(initialState(makeDoc([shortPara])), resizeImage({ p: shortPara.id, o: 0 }, 50, 1.5))
  const keptAtom = asParagraph(keptShort.doc.body[0]).inlines[0]
  assert.ok(keptAtom && keptAtom.kind === 'atom' && keptAtom.atom.type === 'image' && keptAtom.atom.heightPt === pt(1.5), 'a side that was already under 8pt can stay')

  const moved = run(selected, moveAtom({ p: para.id, o: 2 }, { p: other.id, o: 3 }))
  const host = asParagraph(moved.doc.body[1])
  assert.equal(host.inlines[1], picture, 'the move keeps the picture object')
  assert.equal(flatText(asParagraph(moved.doc.body[0])), 'abcd', 'the picture leaves its paragraph')
  assert.deepEqual(moved.sel, { kind: 'text', anchor: { p: other.id, o: 3 }, head: { p: other.id, o: 4 } }, 'the picture is selected in its new place')
  const back = undo(moved)
  assert.ok(back, 'one undo restores the move')
  if (back) assert.equal(asParagraph(back.doc.body[0]).inlines[1], picture, 'undo puts the original picture back')
  assert.equal(imageMoveIssue(doc, { p: para.id, o: 2 }, { p: para.id, o: 2 }), 'same', 'the picture is not a drop target for itself')
  refused(selected, moveAtom({ p: para.id, o: 2 }, { p: para.id, o: 3 }))

  const inner = paragraph('cell', [text('x')])
  const boxed = table('t', [row([cell([inner])])], [40])
  const placed = run(initialState(makeDoc([para, boxed])), moveAtom({ p: para.id, o: 2 }, { p: inner.id, o: 0 }))
  const fitted = asParagraph(asTable(placed.doc.body[1]).rows[0]?.cells[0]?.blocks[0]).inlines[0]
  assert.ok(fitted && fitted.kind === 'atom' && fitted.atom.type === 'image', 'the picture lands in the cell')
  if (fitted && fitted.kind === 'atom' && fitted.atom.type === 'image') {
    assert.equal(fitted.atom.widthPt, pt(40), 'a narrower cell scales the width')
    assert.equal(fitted.atom.heightPt, pt(20), 'the cell scale keeps the ratio')
    assert.equal(fitted.atom.media, picture.atom.type === 'image' ? picture.atom.media : '', 'the scaled picture keeps its media')
    assert.equal(fitted.src, null, 'the scaled picture keeps its src')
  }
  const cellBack = undo(placed)
  assert.ok(cellBack, 'the scale shares the move undo entry')
  if (cellBack) assert.equal(asParagraph(cellBack.doc.body[0]).inlines[1], picture, 'undo restores the unscaled picture')

  const filed = paragraph('file', [text('ab', 1), sourced], {}, { src: 2 })
  const fileDoc = makeDoc([filed, other])
  const fileState = rangeAt(initialState(fileDoc), { p: filed.id, o: 2 }, { p: filed.id, o: 3 })
  const fileSized = run(fileState, resizeImage({ p: filed.id, o: 2 }, 120, 60))
  const fileAtom = asParagraph(fileSized.doc.body[0]).inlines[1]
  assert.ok(fileAtom && fileAtom.kind === 'atom' && fileAtom.atom.type === 'image' && fileAtom.src === sourced.src, 'resizing a file picture keeps its run')
  assert.equal(imageMoveIssue(fileDoc, { p: filed.id, o: 2 }, { p: other.id, o: 0 }), 'file', 'a file picture cannot move')
  refused(fileState, moveAtom({ p: filed.id, o: 2 }, { p: other.id, o: 0 }))
  refused(initialState(makeDoc([locked])), resizeImage({ p: locked.id, o: 0 }, 40, 20))

  const preview = paragraph('preview', [picture])
  const previewDoc = makeDoc([opaqueBlock('box', [preview], false)])
  assert.equal(
    imageChrome(previewDoc, { kind: 'text', anchor: { p: preview.id, o: 0 }, head: { p: preview.id, o: 1 } }),
    'outline',
    'a preview picture has an outline and no handles',
  )
  assert.equal(
    imageChrome(makeDoc([locked]), { kind: 'text', anchor: { p: locked.id, o: 0 }, head: { p: locked.id, o: 1 } }),
    'outline',
    'a locked picture has an outline and no handles',
  )
})

test('move, select, selected text, and word count', () => {
  const preview = paragraph('v', [text('你')], {}, { lock: { reason: 'preview' } })
  const first = paragraph('a', [text('a😀b hello')])
  const second = paragraph('b', [text('there')])
  const doc = makeDoc([
    opaqueBlock('o', [preview], false),
    first,
    second,
    table('t', [row([cell([paragraph('c', [text('你a')])]), cell([paragraph('d', [])]), cell([paragraph('e', [text('c')])])])]),
  ], { stories: new Map([[storyKey('s0/header/default'), [paragraph('h', [text('world world')])]]]) })
  const state = initialState(doc)
  assert.equal(state.sel.kind === 'text' && state.sel.head.p, first.id, 'the caret starts in the first editable paragraph')
  const all = run(state, selectAll)
  assert.equal(selectedText(all), 'a😀b hello\nthere\n你a\t\tc', 'select all skips the preview and joins cells with tabs')
  assert.equal(wordCount(doc), 7, 'Chinese characters count once, Latin words by whitespace, and the header is omitted')
  assert.equal(wordCount(makeDoc([paragraph('p', [text('hello,world')])])), 1, 'punctuation does not split a Latin word')
  assert.equal(wordCount(makeDoc([paragraph('p', [text('a😀b')])])), 1, 'an emoji inside a word is not whitespace')
  assert.equal(wordCount(makeDoc([paragraph('p', [text('\u{20000}')])])), 1, 'one extension-B ideograph counts once')
  assert.equal(wordCount(makeDoc([paragraph('p', [fieldAtom('12'), imageAtom('k')])])), 0, 'atoms contribute no words')

  let moved = run(caretAt(state, first.id, 1), move('forward', 'char', false))
  assert.equal(moved.sel.kind === 'text' && moved.sel.head.o, 3, 'a forward character step skips the whole emoji')
  moved = run(caretAt(state, first.id, 3), move('backward', 'char', false))
  assert.equal(moved.sel.kind === 'text' && moved.sel.head.o, 1, 'a backward character step returns to the emoji start')
  const hello = paragraph('h', [text('hello world')])
  moved = run(initialState(makeDoc([hello, second])), move('forward', 'word', false))
  assert.equal(moved.sel.kind === 'text' && moved.sel.head.o, 6, 'a forward word step lands on the next word')
  const both = makeDoc([hello, second])
  moved = run(caretAt(initialState(both), hello.id, 2), move('forward', 'paragraph', false))
  assert.equal(moved.sel.kind === 'text' && moved.sel.head.o, 11, 'a forward paragraph step goes to the end')
  moved = run(moved, move('forward', 'paragraph', false))
  assert.equal(moved.sel.kind === 'text' && moved.sel.head.p, second.id, 'the next paragraph step leaves the paragraph')
  moved = run(initialState(both), move('forward', 'document', false))
  assert.equal(moved.sel.kind === 'text' && moved.sel.head.p, second.id, 'document forward goes to the last paragraph')
  assert.equal(moved.sel.kind === 'text' && moved.sel.head.o, 5, 'document forward goes to the end')
  const selected = rangeAt(initialState(makeDoc([hello])), { p: hello.id, o: 0 }, { p: hello.id, o: 2 })
  const collapsed = run(selected, move('forward', 'char', false))
  assert.equal(collapsed.sel.kind === 'text' && collapsed.sel.anchor.o, 2, 'a move without shift collapses to the forward end')
  assert.equal(collapsed.sel.kind === 'text' && collapsed.sel.head.o, 2, 'the collapsed caret does not take an extra step')
  const around = run(caretAt(initialState(makeDoc([hello])), hello.id, 1), selectAround({ p: hello.id, o: 1 }, 'word'))
  assert.equal(selectedText(around), 'hello', 'select around a word selects that word')
  const marked = paragraph('m', [text('hello'), fieldAtom('12'), imageAtom('k'), text('!')])
  const follow = paragraph('f', [text('world')])
  const slice = rangeAt(initialState(makeDoc([marked, follow])), { p: marked.id, o: 1 }, { p: follow.id, o: 0 })
  assert.equal(selectedText(slice), 'ello12!\n', 'a selection that ends on a paragraph mark gains a trailing newline')
})

test('toolbar state uses the look callback and reports mixed', () => {
  const heading = paragraph('h', [text('Title')], { style: HEADING })
  const head = toolbarState(initialState(makeDoc([heading])), look)
  assert.equal(head.vertAlign, 'none', 'a heading without a vertical alignment mark is none')
  assert.equal(head.bold, true, 'a heading reads bold from the look, without a direct mark')
  assert.equal(head.sizePt, pt(16), 'the heading size comes from the look')
  assert.equal(head.font, '宋体', 'the font slot shown is East Asian')
  assert.equal(head.canUndo, false, 'a fresh state cannot undo')
  assert.equal(head.inTable, false, 'a body paragraph is not in a table')
  assert.equal(head.rowTools, false, 'row tools are off outside a table')
  assert.equal(head.columnTools, false, 'column tools are off outside a table')

  const para = paragraph('p', [text('a')])
  let state = initialState(makeDoc([para]))
  const stored = toggleMark('bold')(state)
  assert.ok(stored)
  state = apply(state, stored, 0)
  assert.equal(toolbarState(state, look).bold, true, 'stored bold is what the caret shows')
  const raised = toggleVertAlign('super')(state)
  assert.ok(raised)
  assert.equal(toolbarState(apply(state, raised, 0), look).vertAlign, 'super', 'stored superscript is what the caret shows')
  const splitAlign = paragraph('m', [text('a', 1, { vertAlign: 'super' }), text('b', 2, { vertAlign: 'sub' })])
  const mixedState = rangeAt(initialState(makeDoc([splitAlign])), { p: splitAlign.id, o: 0 }, { p: splitAlign.id, o: 2 })
  assert.equal(toolbarState(mixedState, look).vertAlign, 'mixed', 'different vertical alignments are mixed')
  const lowered = paragraph('d', [text('ab', 1, { vertAlign: 'sub' })])
  assert.equal(
    toolbarState(rangeAt(initialState(makeDoc([lowered])), { p: lowered.id, o: 0 }, { p: lowered.id, o: 2 }), look).vertAlign,
    'sub',
  )

  const mixed = paragraph('m', [text('a', 1, { bold: true, fontEastAsia: '宋体' }), text('b', 2, { fontEastAsia: '黑体' })], {}, { src: 3 })
  const other = paragraph('o', [text('c')], { align: 'center' })
  const span = rangeAt(initialState(makeDoc([mixed, other])), { p: mixed.id, o: 0 }, { p: other.id, o: 1 })
  const bar = toolbarState(span, look)
  assert.equal(bar.bold, 'mixed', 'bold is mixed when the selection disagrees')
  assert.equal(bar.font, 'mixed', 'font is mixed when the East Asian slot disagrees')
  assert.equal(bar.align, 'mixed', 'alignment is mixed across paragraphs')
  const inner = paragraph('c', [text('cell')])
  const inside = toolbarState(initialState(makeDoc([table('t', [row([cell([inner])])])])), look)
  assert.equal(inside.inTable, true, 'a caret in a cell reports inTable')
  assert.equal(inside.rowTools, true, 'row tools are on in a table without a horizontal merge')
  assert.equal(inside.columnTools, true, 'column tools are on when every span is 1')
  const edited = run(state, insertText('x'))
  assert.equal(toolbarState(edited, look).canUndo, true, 'a recorded command can undo')
  const back = undo(edited)
  assert.ok(back)
  assert.equal(toolbarState(back, look).canRedo, true, 'an undone command can redo')
})

test('history coalesces typing and deleting, and undo returns the same document', () => {
  const para = paragraph('p', [])
  let state = run(initialState(makeDoc([para])), insertText('a'), 0)
  state = run(state, insertText('b'), 900)
  state = run(state, insertText('c'), 1800)
  assert.equal(state.history.past.length, 1, 'keystrokes within a second of the last one coalesce')
  const original = makeDoc([paragraph('p', [])])
  let burst = run(initialState(original), insertText('a'), 0)
  burst = run(burst, insertText('b'), 900)
  burst = run(burst, insertText('c'), 1800)
  assert.equal(undo(burst)?.doc, original, 'undo of a coalesced burst is the original object')

  let split = run(initialState(makeDoc([paragraph('p', [])])), insertText('a'), 0)
  split = run(split, insertText('b'), 1001)
  assert.equal(split.history.past.length, 2, 'a gap longer than a second starts a new entry')

  const left = paragraph('a', [])
  const right = paragraph('b', [])
  let pair = run(initialState(makeDoc([left, right])), insertText('a'), 0)
  pair = run(caretAt(pair, right.id, 0), insertText('b'), 100)
  assert.equal(pair.history.past.length, 2, 'typing in another paragraph does not coalesce')

  const host = paragraph('p', [text('xy')])
  let typed = run(caretAt(initialState(makeDoc([host])), host.id, 1), insertText('a'), 0)
  const typedDoc = typed.doc
  const undone = undo(typed)
  assert.ok(undone)
  const moved = caretAt(undone, host.id, 0)
  assert.equal(redo(moved)?.doc, typedDoc, 'a selection change does not clear the redo document')
  const cleared = undo(run(initialState(makeDoc([paragraph('q', [])])), insertText('a'), 0))
  assert.ok(cleared)
  const replaced = run(cleared, insertText('c'), 50)
  assert.equal(redo(replaced), null, 'a new edit clears redo')
  assert.equal(flatText(asParagraph(replaced.doc.body[0])), 'c', 'the new text replaces the undone text')

  const word = paragraph('p', [text('abcd')])
  const wordDoc = makeDoc([word])
  let deleting = run(caretAt(initialState(wordDoc), word.id, 4), deleteBackward('char'), 0)
  deleting = run(deleting, deleteBackward('char'), 100)
  assert.equal(deleting.history.past.length, 1, 'deletes in the same paragraph coalesce')
  assert.equal(undo(deleting)?.doc, wordDoc, 'one undo restores both deleted characters')
  let both = run(initialState(makeDoc([paragraph('p', [])])), insertText('a'), 0)
  both = run(both, deleteBackward('char'), 50)
  assert.equal(both.history.past.length, 2, 'a delete does not coalesce into typing')

  let limited = initialState(makeDoc([paragraph('p', [text('a')])]))
  for (let i = 0; i < HISTORY_LIMIT + 1; i++) limited = run(limited, setAlign(i % 2 === 0 ? 'center' : 'left'), i)
  assert.equal(limited.history.past.length, HISTORY_LIMIT, 'history keeps the newest 500 entries')
})

function linkKey(value: string): NonNullable<Marks['link']> {
  return value as NonNullable<Marks['link']>
}

test('replace keeps the first character marks, skips links and locks, and undoes as one entry', () => {
  const mixed = paragraph('p', [
    text('Aa', null, { bold: true }),
    text('Bb', null, { italic: true }),
  ])
  const mixedDoc = makeDoc([mixed])
  let state = run(initialState(mixedDoc), replaceMatch({ para: mixed.id, start: 1, end: 3 }, 'X', null))
  const rewritten = asParagraph(state.doc.body[0])
  assert.equal(flatText(rewritten), 'AXb')
  const covering = rewritten.inlines.find(inline => inline.kind === 'text' && inline.text.includes('X'))
  assert.equal(covering?.kind, 'text')
  if (covering?.kind === 'text') {
    assert.equal(covering.marks.bold, true)
    assert.equal(covering.marks.italic, undefined)
  }
  assert.equal(undo(state)?.doc, mixedDoc)

  const linked = paragraph('p', [
    text('Go ', null, { bold: true }),
    text('here', null, { link: linkKey('to') }),
  ])
  const linkDoc = makeDoc([linked])
  refused(initialState(linkDoc), replaceMatch({ para: linked.id, start: 3, end: 7 }, 'there', null))
  refused(initialState(linkDoc), replaceMatch({ para: linked.id, start: 2, end: 4 }, 'X', null))
  const kept = run(initialState(linkDoc), replaceMatch({ para: linked.id, start: 0, end: 2 }, 'Do', null))
  assert.equal(flatText(asParagraph(kept.doc.body[0])), 'Do here')
  const open = asParagraph(kept.doc.body[0]).inlines[0]
  assert.equal(open?.kind, 'text')
  if (open?.kind === 'text') assert.equal(open.marks.bold, true)

  const locked = paragraph('lock', [text('secret')], {}, { lock: { reason: 'revisions' } })
  const openPara = paragraph('open', [text('secret tale')])
  const both = makeDoc([locked, openPara])
  const counts = replaceCounts(both, [
    { para: locked.id, start: 0, end: 6 },
    { para: openPara.id, start: 0, end: 6 },
  ], 'hidden')
  assert.deepEqual(counts, { replaced: 1, skipped: 1 })
  const replaced = run(initialState(both), replaceAllText('secret', 'hidden', { matchCase: true, wholeWord: false }))
  assert.equal(replaced.history.past.length, 1)
  assert.equal(flatText(asParagraph(replaced.doc.body[0])), 'secret')
  assert.equal(flatText(asParagraph(replaced.doc.body[1])), 'hidden tale')
  assert.equal(asParagraph(replaced.doc.body[0]), locked)
  assert.equal(undo(replaced)?.doc, both)

  const loop = makeDoc([paragraph('p', [text('cat')])])
  const once = run(initialState(loop), replaceAllText('cat', 'catcat', { matchCase: false, wholeWord: false }))
  assert.equal(flatText(asParagraph(once.doc.body[0])), 'catcat')
  assert.equal(undo(once)?.doc, loop)

  const same = makeDoc([paragraph('p', [text('same')])])
  refused(initialState(same), replaceAllText('same', 'same', { matchCase: true, wholeWord: false }))
  const emptied = run(initialState(same), replaceMatch({ para: 'p' as Paragraph['id'], start: 0, end: 4 }, '', null))
  assert.equal(flatText(asParagraph(emptied.doc.body[0])), '')

  const two = makeDoc([paragraph('a', [text('one one')])])
  const moved = run(initialState(two), replaceMatch(
    { para: 'a' as Paragraph['id'], start: 0, end: 3 },
    'two',
    { para: 'a' as Paragraph['id'], start: 4, end: 7 },
  ))
  assert.equal(flatText(asParagraph(moved.doc.body[0])), 'two one')
  assert.deepEqual(moved.sel, {
    kind: 'text',
    anchor: { p: 'a', o: 4 },
    head: { p: 'a', o: 7 },
  })
})

test('random command sequences keep the document valid and round-trip through undo', { timeout: 120_000 }, () => {
  const rng = mulberry32(0xc0ffee)
  for (let seq = 0; seq < 300; seq++) {
    const doc = randomDoc(rng)
    const before = JSON.stringify(project(doc))
    checkDoc(doc)
    let state = initialState(doc)
    const steps = 1 + Math.floor(rng() * 60)
    let now = 0
    for (let step = 0; step < steps; step++) {
      now += rng() < 0.5 ? 10 : COALESCE_MS + 5
      const generated = randomCommand(rng, state)
      let tr: Transaction | null
      try {
        tr = generated.command(state)
      } catch (error) {
        throw new Error(`seq ${seq} step ${step} ${generated.label} threw ${String(error)}`)
      }
      if (!tr) continue
      try {
        state = apply(state, tr, now)
        checkDoc(state.doc, doc)
        assertNoLoneSurrogate(state.doc)
        assertFits(state)
      } catch (error) {
        throw new Error(`seq ${seq} step ${step} ${generated.label}: ${String(error)}`)
      }
    }
    const finalDoc = state.doc
    let undone = state
    while (true) {
      const prev = undo(undone)
      if (!prev) break
      undone = prev
      checkDoc(undone.doc, doc)
    }
    assert.equal(undone.doc, doc, `sequence ${seq} undoes to the original document`)
    assert.equal(JSON.stringify(project(undone.doc)), before, `sequence ${seq} leaves the original content unchanged`)
    let redone = undone
    while (true) {
      const next = redo(redone)
      if (!next) break
      redone = next
    }
    assert.equal(redone.doc, finalDoc, `sequence ${seq} redoes to the final document`)
  }
})

function runSrcOf(para: Paragraph): TextRun['src'] {
  const run = para.inlines[0]
  assert.ok(run && run.kind === 'text')
  return run.src
}

function assertNoLoneSurrogate(doc: Doc): void {
  const visit = (para: Paragraph): void => {
    for (const inline of para.inlines) {
      if (inline.kind !== 'text') continue
      for (let i = 0; i < inline.text.length; i++) {
        const code = inline.text.charCodeAt(i)
        if (code >= 0xd800 && code <= 0xdbff) {
          const next = inline.text.charCodeAt(i + 1)
          assert.ok(next >= 0xdc00 && next <= 0xdfff, 'a high surrogate is paired')
          i += 1
        } else {
          assert.ok(code < 0xdc00 || code > 0xdfff, 'a low surrogate is paired')
        }
      }
    }
  }
  const walk = (blocks: readonly Block[]): void => {
    for (const block of blocks) {
      if (block.kind === 'p') visit(block)
      else if (block.kind === 'table') for (const entry of block.rows) for (const item of entry.cells) walk(item.blocks)
      else for (const para of block.preview) visit(para)
    }
  }
  walk(doc.body)
  for (const blocks of doc.stories.values()) walk(blocks)
}

function assertFits(state: EditorState): void {
  const sel = state.sel
  if (sel.kind === 'text') {
    for (const pos of [sel.anchor, sel.head]) {
      const loc = indexOf(state.doc).get(pos.p)
      assert.equal(loc.story, 'body', 'the selection stays in the body')
      assert.ok(pos.o >= 0 && pos.o <= flatText(loc.para).length, 'the selection offset is inside the paragraph')
    }
    return
  }
  const found = findTable(state.doc.body, sel.table)
  assert.ok(found, 'a cell selection names a table in the document')
  assert.ok(found.rows[sel.anchor.row]?.cells[sel.anchor.cell], 'the anchor cell exists')
  assert.ok(found.rows[sel.head.row]?.cells[sel.head.cell], 'the head cell exists')
}

function findTable(blocks: readonly Block[], id: Table['id']): Table | null {
  for (const block of blocks) {
    if (block.kind !== 'table') continue
    if (block.id === id) return block
    for (const entry of block.rows) for (const item of entry.cells) {
      const found = findTable(item.blocks, id)
      if (found) return found
    }
  }
  return null
}

function project(doc: Doc): unknown {
  const blocks: unknown[] = []
  const walk = (items: readonly Block[]): void => {
    for (const block of items) {
      if (block.kind === 'p') {
        blocks.push({ id: block.id, src: block.src, props: block.props, lock: block.lock, sectionEnd: block.sectionEnd, inlines: block.inlines })
      } else if (block.kind === 'table') {
        blocks.push({ id: block.id, grid: block.grid, rows: block.rows.length })
        for (const entry of block.rows) for (const item of entry.cells) walk(item.blocks)
      } else {
        blocks.push({ id: block.id, removable: block.removable })
        walk(block.preview)
      }
    }
  }
  walk(doc.body)
  return { blocks, added: doc.styles.added, lists: [...doc.lists.byNum.keys()], pages: doc.sections.map(item => item.page.widthPt), media: [...doc.media.byKey.keys()] }
}
