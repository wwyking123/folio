import assert from 'node:assert/strict'
import { test } from 'node:test'
import { flatText } from '../model.ts'
import type { Block, BlockId, Doc, Paragraph, Pos, Pt, Section, StoryKey, StyleId } from '../model.ts'
import { listMarkers } from './derive.ts'
import {
  apply,
  deleteBackward,
  deleteSelection,
  deleteTable,
  initialState,
  insertPageBreak,
  insertTable,
  insertText,
  selectAll,
  setPageSetup,
  setSelection,
  splitParagraph,
  toggleList,
  undo,
  type Command,
  type EditorState,
} from './edit.ts'
import { atomCaretBox, caretSliceY, layoutPlan, nextVerticalLine, pageAtSliceY, pageOfBlock, paginate, resolveImageHit, tableRowCuts, type ImageHitCandidate, type Layout, type LayoutPlan, type Measure } from './pages.ts'
import { PT_TO_PX } from './render.ts'
import { cell, makeDoc, mulberry32, numId, opaqueBlock, page, paragraph, row, section, table, text } from './testkit.ts'

function pt(n: number): Pt {
  return n as Pt
}

function para(id: string, sectionEnd: number | null = null): Paragraph {
  return {
    kind: 'p',
    id: id as Paragraph['id'],
    src: null,
    props: {},
    inlines: [{ kind: 'text', text: 'x', marks: {}, src: null }],
    lock: null,
    sectionEnd,
  }
}

function sectionBox(pxHeight: number, breakType: Section['breakType'] = 'nextPage'): Section {
  const edge = 0 as Pt
  return {
    src: null,
    page: {
      widthPt: pt((200 * 3) / 4),
      heightPt: pt((pxHeight * 3) / 4),
      margin: { top: edge, right: edge, bottom: edge, left: edge, header: edge, footer: edge, gutter: edge },
    },
    breakType,
    titlePage: false,
    headers: {},
    footers: {},
    grid: null,
  }
}

function docOf(blocks: readonly Block[], sections: readonly Section[], stories?: Map<StoryKey, readonly Block[]>): Doc {
  return {
    body: blocks,
    sections,
    stories: stories ?? new Map(),
    styles: { paragraph: new Map(), character: new Map(), defaultParagraph: 'Normal' as StyleId, added: [] },
    lists: { byNum: new Map() },
    media: { byKey: new Map() },
  }
}

function pm(height: number, lines: number[], extra: Partial<Measure> = {}): Measure {
  return {
    widthPx: 100,
    heightPx: height,
    lineBottoms: extra.lineBottoms ?? Float32Array.from(lines),
    ...(extra.rowBottoms ? { rowBottoms: extra.rowBottoms, cantSplitRows: extra.cantSplitRows } : {}),
    ...(extra.rowCuts ? { rowCuts: extra.rowCuts } : {}),
    ...(extra.rowCutSplit ? { rowCutSplit: extra.rowCutSplit } : {}),
    ...(extra.rowMinPx ? { rowMinPx: extra.rowMinPx } : {}),
    headerRows: extra.headerRows ?? 0,
    spaceBeforePx: extra.spaceBeforePx ?? 0,
    spaceAfterPx: extra.spaceAfterPx ?? 0,
    keepNext: extra.keepNext ?? false,
    keepLines: extra.keepLines ?? false,
    pageBreakBefore: extra.pageBreakBefore ?? false,
    widowControl: extra.widowControl ?? false,
    ...(extra.inlineBreaks !== undefined ? { inlineBreaks: extra.inlineBreaks } : {}),
  }
}

function sig(layout: Layout): string[] {
  return layout.pages.map((page) => {
    const body = page.slices
      .map((slice) => `${slice.block}:${slice.y0}-${slice.y1}${slice.repeatHeaderRows ? `/r${slice.repeatHeaderRows}` : ''}`)
      .join('|')
    return body.length > 0 ? body : '-'
  })
}

function pagesFor(layout: Layout, id: string): number[] {
  const pages: number[] = []
  layout.pages.forEach((page, index) => {
    if (page.slices.some((slice) => slice.block === id)) pages.push(index)
  })
  return pages
}

function assertSamePages(actual: Layout, full: Layout, body: readonly Paragraph[]): void {
  assert.equal(actual.pages.length, full.pages.length)
  for (let i = 0; i < full.pages.length; i++) assert.deepEqual(actual.pages[i], full.pages[i])
  for (const block of body) assert.deepEqual(pagesFor(actual, block.id), pagesFor(full, block.id), block.id)
}

function layoutOf(blocks: readonly Paragraph[], measures: Map<string, Measure>, height: number, previous: Layout | null = null, changed: readonly string[] = []): Layout {
  return paginate(
    docOf(blocks, [sectionBox(height)]),
    (block) => measures.get(block.id),
    previous,
    changed.map((id) => id as BlockId),
  )
}

test('blocks flow onto the next page when the next one does not fit', () => {
  const blocks = [para('a'), para('b')]
  const measures = new Map([
    ['a', pm(80, [40, 80])],
    ['b', pm(80, [40, 80])],
  ])
  const layout = layoutOf(blocks, measures, 100)
  assert.equal(layout.complete, true)
  assert.deepEqual(sig(layout), ['a:0-80', 'b:0-80'])
})

test('a paragraph splits on line bottoms', () => {
  const measures = new Map([['a', pm(80, [20, 40, 60, 80])]])
  assert.deepEqual(sig(layoutOf([para('a')], measures, 50)), ['a:0-40', 'a:40-80'])
})

test('space after that does not fit is clipped instead of opening an empty continuation', () => {
  const measures = new Map([['a', pm(100, [40, 80])]])
  assert.deepEqual(sig(layoutOf([para('a')], measures, 90)), ['a:0-80'])
})

test('widow control pulls the last line back so a page does not start with one line', () => {
  const measures = new Map([['a', pm(100, [20, 40, 60, 80, 100], { widowControl: true })]])
  assert.deepEqual(sig(layoutOf([para('a')], measures, 90)), ['a:0-60', 'a:60-100'])
})

test('orphan control moves a paragraph when only its first line would fit', () => {
  const measures = new Map([
    ['a', pm(70, [70])],
    ['b', pm(80, [20, 40, 60, 80], { widowControl: true })],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('b')], measures, 100)), ['a:0-70', 'b:0-80'])
})

test('without widow control a single line may stay at the bottom', () => {
  const measures = new Map([
    ['a', pm(70, [70])],
    ['b', pm(80, [20, 40, 60, 80])],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('b')], measures, 100)), ['a:0-70|b:0-20', 'b:20-80'])
})

test('space before is dropped at the top of a page', () => {
  const only = new Map([['b', pm(40, [25, 40], { spaceBeforePx: 10 })]])
  assert.equal(sig(layoutOf([para('b')], only, 100))[0], 'b:10-40')
  const measures = new Map([
    ['a', pm(80, [80])],
    ['b', pm(40, [25, 40], { spaceBeforePx: 10 })],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('b')], measures, 100)), ['a:0-80', 'b:10-40'])
})

test('pageBreakBefore starts a new page and drops the whole space before', () => {
  const measures = new Map([
    ['a', pm(40, [40], { spaceAfterPx: 4 })],
    ['b', pm(30, [20, 30], { pageBreakBefore: true, spaceBeforePx: 10 })],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('b')], measures, 100)), ['a:0-40', 'b:10-30'])
})

test('an inline page break starts the next page', () => {
  const only = new Map([
    ['a', pm(16, [16], { inlineBreaks: Float32Array.from([16]) })],
    ['b', pm(20, [20])],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('b')], only, 100)), ['a:0-16', 'b:0-20'], 'a paragraph that holds only the break takes its line, then the next block')
  const around = new Map([['a', pm(60, [20, 40, 60], { inlineBreaks: Float32Array.from([40]) })]])
  assert.deepEqual(sig(layoutOf([para('a')], around, 100)), ['a:0-40', 'a:40-60'], 'text after the break starts the next page')
  const end = new Map([
    ['a', pm(40, [40], { inlineBreaks: Float32Array.from([40]) })],
    ['b', pm(20, [20])],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('b')], end, 100)), ['a:0-40', 'b:0-20'], 'a break at the end of a paragraph starts the next block on a new page')
  const twice = new Map([['a', pm(60, [20, 40, 60], { inlineBreaks: Float32Array.from([20, 40]) })]])
  assert.deepEqual(sig(layoutOf([para('a')], twice, 100)), ['a:0-20', 'a:20-40', 'a:40-60'], 'two breaks in a row are two page starts')
})

test('a trailing inline break absorbs the next paragraph pageBreakBefore', () => {
  const absorbed = new Map([
    ['a', pm(20, [20], { inlineBreaks: Float32Array.from([20]) })],
    ['b', pm(20, [20], { pageBreakBefore: true })],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('b')], absorbed, 100)), ['a:0-20', 'b:0-20'])
  const mid = new Map([
    ['a', pm(40, [20, 40], { inlineBreaks: Float32Array.from([20]) })],
    ['b', pm(20, [20], { pageBreakBefore: true })],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('b')], mid, 100)), ['a:0-20', 'a:20-40', 'b:0-20'], 'a break inside the paragraph does not absorb the following pageBreakBefore')
})

test('keepLines does not cross an inline page break', () => {
  const measures = new Map([
    ['a', pm(50, [50])],
    ['b', pm(80, [20, 40, 60, 80], { inlineBreaks: Float32Array.from([40]), keepLines: true })],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('b')], measures, 100)), ['a:0-50|b:0-40', 'b:40-80'])
})

test('insertPageBreak inserts the atom and the page view starts there', () => {
  const source = paragraph('p', [text('ab')])
  let state = initialState(makeDoc([source]))
  const moved = setSelection({ kind: 'text', anchor: { p: source.id, o: 1 }, head: { p: source.id, o: 1 } })(state)
  assert.ok(moved, 'the caret moves between the letters')
  state = apply(state, moved, 0)
  const broken = insertPageBreak(state)
  assert.ok(broken, 'insertPageBreak applies')
  state = apply(state, broken, 0)
  const next = state.doc.body[0]
  assert.ok(next && next.kind === 'p', 'the paragraph stays one block')
  if (!next || next.kind !== 'p') return
  assert.deepEqual(next.inlines.map((inline) => inline.kind === 'text' ? inline.text : inline.atom.type), ['a', 'pageBreak', 'b'])
  assert.equal(state.sel.kind, 'text')
  if (state.sel.kind === 'text') {
    assert.equal(state.sel.head.o, 2, 'the caret sits after the atom')
    assert.equal(state.sel.head.p, next.id)
  }
  const measures = new Map<string, Measure>([[next.id, pm(40, [20, 40], { inlineBreaks: Float32Array.from([20]) })]])
  const layout = paginate(state.doc, (block) => measures.get(block.id), null, [])
  assert.deepEqual(sig(layout), [`${next.id}:0-20`, `${next.id}:20-40`])
})

test('adjacent body paragraphs keep the larger of space after and space before', () => {
  const measures = new Map([
    ['a', pm(40, [28], { spaceAfterPx: 12 })],
    ['b', pm(30, [22], { spaceBeforePx: 8 })],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('b')], measures, 100)), ['a:0-40|b:8-30'])
})

test('keepNext measures the collapsed gap, so an overlap that fits stays on the page', () => {
  const measures = new Map([
    ['a', pm(60, [60], { spaceAfterPx: 30 })],
    ['b', pm(50, [50], { spaceBeforePx: 30, keepNext: true })],
    ['c', pm(20, [20])],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('b'), para('c')], measures, 100)), ['a:0-60|b:30-50|c:0-20'])
})

test('keepNext moves a chain that fits on an empty page but not in the remaining space', () => {
  const measures = new Map([
    ['a', pm(50, [50])],
    ['b', pm(40, [40], { keepNext: true })],
    ['c', pm(40, [40])],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('b'), para('c')], measures, 100)), ['a:0-50', 'b:0-40|c:0-40'])
})

test('keepLines moves a paragraph that would otherwise split', () => {
  const measures = new Map([
    ['a', pm(50, [50])],
    ['b', pm(80, [40, 80], { keepLines: true })],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('b')], measures, 100)), ['a:0-50', 'b:0-80'])
})

test('keepLines still splits a paragraph taller than a page', () => {
  const measures = new Map([['a', pm(80, [40, 80], { keepLines: true })]])
  assert.deepEqual(sig(layoutOf([para('a')], measures, 50)), ['a:0-40', 'a:40-80'])
})

test('table rows do not split and header rows repeat on the continuation', () => {
  const measures = new Map([['t', pm(120, [], { rowBottoms: Float32Array.from([30, 60, 90, 120]), headerRows: 1 })]])
  assert.deepEqual(sig(layoutOf([para('t')], measures, 70)), ['t:0-60', 't:60-90/r1', 't:90-120/r1'])
})

test('an unmeasured block stops pagination and a later chunk continues', () => {
  const blocks = [para('a'), para('b'), para('c')]
  const measures = new Map<string, Measure>([
    ['a', pm(80, [80])],
  ])
  const first = layoutOf(blocks, measures, 100)
  assert.equal(first.complete, false)
  assert.deepEqual(sig(first), ['a:0-80'])
  measures.set('b', pm(80, [80]))
  const second = layoutOf(blocks, measures, 100, first, ['b'])
  assert.equal(second.complete, false)
  assert.deepEqual(sig(second), ['a:0-80', 'b:0-80'])
  measures.set('c', pm(80, [80]))
  const third = layoutOf(blocks, measures, 100, second, ['c'])
  assert.equal(third.complete, true)
  assert.deepEqual(sig(third), ['a:0-80', 'b:0-80', 'c:0-80'])
})

test('an unchanged tail page is reused after an edit that does not move it', () => {
  const blocks = [para('a'), para('b'), para('c'), para('d')]
  const measures = new Map([
    ['a', pm(100, [100])],
    ['b', pm(40, [40])],
    ['c', pm(40, [40])],
    ['d', pm(40, [40])],
  ])
  const full = layoutOf(blocks, measures, 100)
  assert.deepEqual(sig(full), ['a:0-100', 'b:0-40|c:0-40', 'd:0-40'])
  const again = layoutOf(blocks, measures, 100, full, ['b'])
  assert.deepEqual(sig(again), sig(full))
  assert.equal(again.pages[2], full.pages[2])
})

test('a taller block pushes the following blocks onto later pages', () => {
  const blocks = [para('a'), para('b'), para('c'), para('d')]
  const measures = new Map<string, Measure>([
    ['a', pm(100, [100])],
    ['b', pm(40, [40])],
    ['c', pm(40, [40])],
    ['d', pm(40, [40])],
  ])
  const full = layoutOf(blocks, measures, 100)
  measures.set('b', pm(80, [80]))
  const again = layoutOf(blocks, measures, 100, full, ['b'])
  assert.deepEqual(sig(again), ['a:0-100', 'b:0-80', 'c:0-40|d:0-40'])
})

test('deleting a block pulls the next block onto the previous page', () => {
  const a = para('a')
  const c = para('c')
  const measures = new Map([
    ['a', pm(60, [60])],
    ['b', pm(50, [50])],
    ['c', pm(30, [30])],
  ])
  const first = layoutOf([a, para('b'), c], measures, 100)
  assert.deepEqual(sig(first), ['a:0-60', 'b:0-50|c:0-30'])
  const second = layoutOf([a, c], measures, 100, first, ['c'])
  assert.deepEqual(sig(second), ['a:0-60|c:0-30'])
})

test('an odd-page section break inserts a blank page', () => {
  const first = sectionBox(100)
  const second = sectionBox(100, 'oddPage')
  const measures = new Map([
    ['a', pm(40, [40])],
    ['b', pm(40, [40])],
  ])
  const layout = paginate(
    docOf([para('a', 0), para('b')], [first, second]),
    (block) => measures.get(block.id),
    null,
    [],
  )
  assert.deepEqual(sig(layout), ['a:0-40', '-', 'b:0-40'])
  assert.deepEqual(
    layout.pages.map((page) => page.section),
    [0, 1, 1],
  )
})

test('a continuous section with the same paper stays on the page', () => {
  const measures = new Map([
    ['a', pm(40, [40])],
    ['b', pm(40, [40])],
  ])
  const layout = paginate(
    docOf([para('a', 0), para('b')], [sectionBox(100), sectionBox(100, 'continuous')]),
    (block) => measures.get(block.id),
    null,
    [],
  )
  assert.deepEqual(sig(layout), ['a:0-40|b:0-40'])
  assert.equal(layout.pages[0]?.section, 0)
})

test('a header taller than the top margin pushes the body down', () => {
  const edge = 0 as Pt
  const section: Section = {
    src: null,
    page: {
      widthPt: pt((200 * 3) / 4),
      heightPt: pt((200 * 3) / 4),
      margin: { top: pt(30), right: edge, bottom: pt(30), left: edge, header: pt(7.5), footer: edge, gutter: edge },
    },
    breakType: 'nextPage',
    titlePage: false,
    headers: { default: 's0/header/default' as StoryKey },
    footers: {},
    grid: null,
  }
  const header = para('h')
  const body = para('a')
  const measures = new Map([
    ['h', pm(80, [80])],
    ['a', pm(50, [50])],
  ])
  const layout = paginate(
    docOf([body], [section], new Map([['s0/header/default' as StoryKey, [header]]])),
    (block) => measures.get(block.id),
    null,
    [],
  )
  assert.equal(layout.pages[0]?.bodyTopPx, 90)
  assert.deepEqual(sig(layout), ['a:0-50'])
})

function margined(paperPx: number, marginPx: number): Section {
  const edge = (px: number): Pt => pt((px * 3) / 4)
  return {
    src: null,
    page: {
      widthPt: edge(600),
      heightPt: edge(paperPx),
      margin: {
        top: edge(marginPx),
        right: edge(marginPx),
        bottom: edge(marginPx),
        left: edge(marginPx),
        header: edge(0),
        footer: edge(0),
        gutter: edge(0),
      },
    },
    breakType: 'nextPage',
    titlePage: false,
    headers: {},
    footers: {},
    grid: null,
  }
}

test('setPageSetup with narrower margins changes the page count of a long document', () => {
  const blocks = Array.from({ length: 24 }, (_, index) => para(`p${index}`))
  const measures = new Map<string, Measure>(blocks.map((block) => [block.id, pm(100, [100])]))
  const measure = (block: Block): Measure | undefined => measures.get(block.id)
  const doc = docOf(blocks, [margined(900, 300)])
  const first = paginate(doc, measure, null, [])
  assert.equal(first.pages.length, 8, 'tall margins hold three blocks on a page')
  const state = initialState(doc)
  const narrow = margined(900, 50).page
  const tr = setPageSetup(0, narrow)(state)
  assert.ok(tr, 'setPageSetup applies')
  const next = apply(state, tr, 0).doc
  assert.equal(next.body, doc.body, 'page setup does not replace a block')
  const second = paginate(next, measure, first, [])
  assert.equal(second.pages.length, 3, 'narrower margins fit more blocks on a page')
  const plan = layoutPlan(doc, next, new Map(), new Map())
  assert.equal(plan.remeasure.length, blocks.length, 'the section page change invalidates every block measure')
})

test('changing nothing but titlePage measures no block again', () => {
  const blocks = [para('a'), para('b'), para('c')]
  const measures = new Map<string, Measure>(blocks.map((block) => [block.id, pm(40, [40])]))
  let calls = 0
  const measure = (block: Block): Measure | undefined => {
    calls += 1
    return measures.get(block.id)
  }
  const section = margined(400, 40)
  const doc = docOf(blocks, [section])
  const first = paginate(doc, measure, null, [])
  const seen = calls
  assert.ok(seen > 0, 'the first layout measures the blocks')
  const titled: Section = { ...section, titlePage: true }
  const next = docOf(blocks, [titled])
  const second = paginate(next, measure, first, [])
  assert.equal(calls, seen, 'titlePage does not measure a block again')
  assert.equal(second.pages.length, first.pages.length)
  assert.deepEqual(sig(second), sig(first))
  const plan = layoutPlan(doc, next, new Map(), new Map())
  assert.deepEqual(plan.remeasure, [], 'titlePage leaves measures in place')
})

test('deleting the first numbered item changes the marker painted for the third', () => {
  const a = paragraph('a', [text('one')])
  const b = paragraph('b', [text('two')])
  const c = paragraph('c', [text('three')])
  let state = initialState(makeDoc([a, b, c]))
  const all = selectAll(state)
  assert.ok(all, 'select all applies')
  state = apply(state, all, 0)
  const listed = toggleList('decimal')(state)
  assert.ok(listed, 'the three paragraphs become one list')
  state = apply(state, listed, 0)
  assert.equal(listMarkers(state.doc).get(c.id), '3.')
  const caret = setSelection({ kind: 'text', anchor: { p: a.id, o: 0 }, head: { p: a.id, o: 0 } })(state)
  assert.ok(caret, 'the caret moves to the first item')
  state = apply(state, caret, 0)
  const cleared = deleteBackward('char')(state)
  assert.ok(cleared, 'backspace at the start of a list item leaves the list')
  const next = apply(state, cleared, 0)
  const third = next.doc.body[2]
  assert.equal(third, state.doc.body[2], 'the third paragraph object is unchanged')
  assert.equal(listMarkers(next.doc).get(c.id), '2.', 'the third item is renumbered')
  const plan = layoutPlan(state.doc, next.doc, listMarkers(state.doc), listMarkers(next.doc))
  assert.ok(plan.paint.includes(c.id), 'the third item is painted again')
  assert.ok(!plan.remeasure.includes(c.id), 'a marker change does not remeasure the paragraph')
})

test('a paragraph taller than a page splits on a line bottom inside the body', () => {
  const linePx = 19.32
  const lines = Array.from({ length: 55 }, (_, index) => (index + 1) * linePx)
  const textPx = 1043
  const paddingPx = 10.67
  const body = 864
  const layout = layoutOf([para('a')], new Map([['a', pm(textPx + paddingPx, lines)]]), body)
  assert.equal(layout.complete, true)
  const first = layout.pages[0]?.slices[0]
  const next = layout.pages[1]?.slices[0]
  assert.ok(first && next, 'the paragraph occupies two pages')
  assert.ok(first.y1 <= body, 'page 0 ends at or under the body')
  const bottoms = Float32Array.from(lines)
  assert.ok(Array.from(bottoms).some((bottom) => bottom === first.y1), 'page 0 ends on a line bottom')
  assert.equal(next.y0, first.y1, 'page 1 starts at that line')
})

function fivePages(): { blocks: Paragraph[], measures: Map<string, Measure> } {
  const blocks = [para('a'), para('b'), para('c'), para('d'), para('e')]
  const measures = new Map(blocks.map((block) => [block.id, pm(100, [100])]))
  return { blocks, measures }
}

test('a block inserted after an unchanged block on page 1 is placed on that page or the next', () => {
  const { blocks, measures } = fivePages()
  const previous = layoutOf(blocks, measures, 100)
  assert.ok(previous.pages.length >= 3, 'the document covers at least three pages')
  assert.deepEqual(pagesFor(previous, 'b'), [1], 'b is the block on page 1')
  const inserted = para('n')
  measures.set(inserted.id, pm(100, [100]))
  const a = blocks[0]
  const b = blocks[1]
  const c = blocks[2]
  const d = blocks[3]
  const e = blocks[4]
  assert.ok(a && b && c && d && e)
  const next = [a, b, inserted, c, d, e]
  const again = layoutOf(next, measures, 100, previous, [inserted.id])
  const full = layoutOf(next, measures, 100)
  const placed = pagesFor(again, inserted.id)
  assert.ok(placed.some((page) => page === 1 || page === 2), `the new block is on page 1 or 2, got ${placed.join(',')}`)
  assertSamePages(again, full, next)
})

test('a block inserted after an unchanged block on a middle page is placed on that page or the next', () => {
  const { blocks, measures } = fivePages()
  const previous = layoutOf(blocks, measures, 100)
  assert.ok(previous.pages.length >= 5, 'a middle page is neither the first nor the last')
  assert.deepEqual(pagesFor(previous, 'c'), [2], 'c is the block on the middle page')
  const inserted = para('n')
  measures.set(inserted.id, pm(100, [100]))
  const a = blocks[0]
  const b = blocks[1]
  const c = blocks[2]
  const d = blocks[3]
  const e = blocks[4]
  assert.ok(a && b && c && d && e)
  const next = [a, b, c, inserted, d, e]
  const again = layoutOf(next, measures, 100, previous, [inserted.id])
  const full = layoutOf(next, measures, 100)
  const placed = pagesFor(again, inserted.id)
  assert.ok(placed.some((page) => page === 2 || page === 3), `the new block is on page 2 or 3, got ${placed.join(',')}`)
  assertSamePages(again, full, next)
})

test('a block inserted at the start of the body is placed on the first page or the next', () => {
  const { blocks, measures } = fivePages()
  const previous = layoutOf(blocks, measures, 100)
  assert.ok(previous.pages.length >= 3)
  const inserted = para('n')
  measures.set(inserted.id, pm(100, [100]))
  const next = [inserted, ...blocks]
  const again = layoutOf(next, measures, 100, previous, [inserted.id])
  const full = layoutOf(next, measures, 100)
  const placed = pagesFor(again, inserted.id)
  assert.ok(placed.some((page) => page === 0 || page === 1), `the new block is on page 0 or 1, got ${placed.join(',')}`)
  assertSamePages(again, full, next)
})

test('removing a block from page 1 whose neighbors are unchanged matches a layout from scratch', () => {
  const a = para('a')
  const b = para('b')
  const c = para('c')
  const d = para('d')
  const measures = new Map<string, Measure>([
    [a.id, pm(60, [60])],
    [b.id, pm(100, [100])],
    [c.id, pm(30, [30])],
    [d.id, pm(100, [100])],
  ])
  const previous = layoutOf([a, b, c, d], measures, 100)
  assert.ok(previous.pages.length >= 3, 'the document covers at least three pages')
  assert.deepEqual(pagesFor(previous, b.id), [1], 'b is the block removed from page 1')
  const next = [a, c, d]
  assert.equal(next[0], a, 'the previous neighbor keeps its object')
  assert.equal(next[1], c, 'the following neighbor keeps its object')
  const again = layoutOf(next, measures, 100, previous, [c.id, d.id])
  const full = layoutOf(next, measures, 100)
  assert.equal(pagesFor(again, b.id).length, 0, 'the removed block is on no page')
  assertSamePages(again, full, next)
})

test('removing a block that spans pages drops every slice of it', () => {
  const a = para('a')
  const b = para('b')
  const c = para('c')
  const measures = new Map<string, Measure>([
    [a.id, pm(100, [100])],
    [b.id, pm(250, [100, 200, 250])],
    [c.id, pm(100, [100])],
  ])
  const previous = layoutOf([a, b, c], measures, 100)
  assert.ok(previous.pages.length >= 3)
  const bPages = pagesFor(previous, b.id)
  assert.ok(bPages.includes(1) && bPages.length >= 2, 'b starts on page 1 and continues')
  const next = [a, c]
  assert.equal(next[0], a)
  assert.equal(next[1], c)
  const again = layoutOf(next, measures, 100, previous, [c.id])
  const full = layoutOf(next, measures, 100)
  assert.equal(pagesFor(again, b.id).length, 0, 'no page keeps a slice of the removed block')
  assertSamePages(again, full, next)
})

test('pageOfBlock is the page of the first slice, or null when the layout has not placed the id', () => {
  const body = para('p')
  const inner = para('in-cell')
  const grid = table('t', [row([cell([inner])])])
  const box = opaqueBlock('o', [para('preview')], true)
  const later = para('later')
  const measures = new Map<string, Measure>([
    [body.id, pm(80, [80])],
    [grid.id, pm(80, [], { rowBottoms: Float32Array.from([80]) })],
    [box.id, pm(80, [80])],
  ])
  const blocks = [body, grid, box, later]
  const doc = docOf(blocks, [sectionBox(100)])
  const incomplete = paginate(doc, (block) => measures.get(block.id), null, [])
  assert.equal(incomplete.complete, false)
  assert.equal(incomplete.pages.length, 3, sig(incomplete).join(' / '))
  assert.equal(pageOfBlock(incomplete, body.id), 0, 'the paragraph is on the first page')
  assert.equal(pageOfBlock(incomplete, grid.id), 1, 'the table is on the second page')
  assert.equal(pageOfBlock(incomplete, box.id), 2, 'the opaque block is on the third page')
  assert.equal(pageOfBlock(incomplete, later.id), null, 'a block past the measured part is unplaced')
  assert.equal(pageOfBlock(incomplete, inner.id), null, 'a cell paragraph is not its own slice')
  assert.equal(pageOfBlock(incomplete, 'missing'), null)

  measures.set(later.id, pm(80, [80]))
  const complete = paginate(doc, (block) => measures.get(block.id), incomplete, [later.id])
  assert.equal(complete.complete, true)
  assert.equal(pageOfBlock(complete, later.id), 3)
  assert.equal(pageOfBlock(complete, inner.id), null)

  const split = layoutOf([para('s')], new Map([['s', pm(80, [40, 80])]]), 50)
  assert.deepEqual(sig(split), ['s:0-40', 's:40-80'])
  assert.equal(pageOfBlock(split, 's'), 0, 'a block that continues answers the page of its first slice')
})

function blockSlices(layout: Layout, id: string): { y0: number; y1: number; repeatHeaderRows: number }[] {
  const slices: { y0: number; y1: number; repeatHeaderRows: number }[] = []
  for (const page of layout.pages) {
    for (const slice of page.slices) {
      if (slice.block === id) slices.push(slice)
    }
  }
  return slices
}

function assertPartition(layout: Layout, id: string, height: number): void {
  const slices = blockSlices(layout, id)
  assert.ok(slices.length > 0)
  assert.equal(slices[0]?.y0, 0)
  for (let i = 1; i < slices.length; i++) assert.equal(slices[i]?.y0, slices[i - 1]?.y1)
  assert.equal(slices[slices.length - 1]?.y1, height)
}

function assertNoSliceTaller(layout: Layout, id: string, bottoms: Float32Array, body: number): void {
  for (const slice of blockSlices(layout, id)) {
    const header = slice.repeatHeaderRows > 0 ? (bottoms[slice.repeatHeaderRows - 1] ?? 0) : 0
    const visual = header + (slice.y1 - slice.y0)
    assert.ok(visual <= body + 0.51, `${slice.y0}-${slice.y1} visual ${visual}`)
  }
}

function assertLegalEnd(layout: Layout, id: string, bottoms: Float32Array, cuts: Float32Array): void {
  const height = bottoms[bottoms.length - 1] ?? 0
  for (const slice of blockSlices(layout, id)) {
    if (Math.abs(slice.y1 - height) <= 0.51) continue
    if ([...bottoms, ...cuts].some((y) => Math.abs(y - slice.y1) <= 0.51)) continue
    let top = 0
    let rowTop = 0
    let found = false
    for (let i = 0; i < bottoms.length; i++) {
      const bottom = bottoms[i] ?? top
      if (slice.y1 > top + 0.51 && slice.y1 < bottom - 0.51) {
        rowTop = top
        found = true
        break
      }
      top = bottom
    }
    assert.equal(found, true, `end ${slice.y1} is not inside a row`)
    const from = Math.max(slice.y0, rowTop)
    for (const y of [...cuts, ...bottoms]) {
      assert.ok(!(y > from + 0.51 && y < slice.y1 - 0.01), `end ${slice.y1} passed ${y}`)
    }
  }
}

test('a row that fits the remainder is placed whole even when it has cuts', () => {
  const measures = new Map([
    ['a', pm(20, [20])],
    ['t', pm(50, [], { rowBottoms: Float32Array.from([50]), rowCuts: Float32Array.from([25]) })],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('t')], measures, 100)), ['a:0-20|t:0-50'])
})

test('a row taller than the remainder and shorter than the page is cut where a cut fits', () => {
  const bottoms = Float32Array.from([80])
  const cuts = Float32Array.from([40])
  const measures = new Map([
    ['a', pm(50, [50])],
    ['t', pm(80, [], { rowBottoms: bottoms, rowCuts: cuts })],
  ])
  const layout = layoutOf([para('a'), para('t')], measures, 100)
  assert.equal(layout.complete, true)
  assert.deepEqual(sig(layout), ['a:0-50|t:0-40', 't:40-80'])
  assertPartition(layout, 't', 80)
  assertLegalEnd(layout, 't', bottoms, cuts)
})

test('a row moves whole when no allowed cut fits the remainder', () => {
  const measures = new Map([
    ['a', pm(50, [50])],
    ['t', pm(80, [], { rowBottoms: Float32Array.from([80]), rowCuts: Float32Array.from([70]) })],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('t')], measures, 100)), ['a:0-50', 't:0-80'])
})

test('a row moves whole when the fitting cut is not one the kit would split at', () => {
  const measures = new Map([
    ['a', pm(50, [50])],
    ['t', pm(80, [], {
      rowBottoms: Float32Array.from([80]),
      rowCuts: Float32Array.from([40]),
      rowCutSplit: Uint8Array.from([0]),
    })],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('t')], measures, 100)), ['a:0-50', 't:0-80'])
})

test('a minimum keeps the largest fitting fragment and moves the row when that fragment is short', () => {
  const cut = new Map([
    ['a', pm(40, [40])],
    ['t', pm(80, [], {
      rowBottoms: Float32Array.from([80]),
      rowCuts: Float32Array.from([30, 60]),
      rowMinPx: Float32Array.from([50]),
    })],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('t')], cut, 100)), ['a:0-40|t:0-60', 't:60-80'])
  const move = new Map([
    ['a', pm(50, [50])],
    ['t', pm(80, [], {
      rowBottoms: Float32Array.from([80]),
      rowCuts: Float32Array.from([30, 60]),
      rowMinPx: Float32Array.from([50]),
    })],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('t')], move, 100)), ['a:0-50', 't:0-80'])
})

test('a minimum is ignored at the top of an empty page and applied once the page has content', () => {
  const bottoms = Float32Array.from([20, 110])
  const measure = pm(110, [], {
    rowBottoms: bottoms,
    headerRows: 1,
    rowCuts: Float32Array.from([50]),
    rowMinPx: Float32Array.from([0, 70]),
  })
  const atTop = layoutOf([para('t')], new Map([['t', measure]]), 100)
  assert.deepEqual(sig(atTop), ['t:0-50', 't:50-110/r1'])
  const after = layoutOf(
    [para('a'), para('t')],
    new Map([['a', pm(10, [10])], ['t', measure]]),
    100,
  )
  assert.deepEqual(sig(after), ['a:0-10', 't:0-50', 't:50-110/r1'])
})

test('a row three pages tall repeats the header and covers the table once', () => {
  const bottoms = Float32Array.from([20, 260])
  const cuts = Float32Array.from([100, 180])
  const measures = new Map([['t', pm(260, [], { rowBottoms: bottoms, headerRows: 1, rowCuts: cuts })]])
  const layout = layoutOf([para('t')], measures, 100)
  assert.equal(layout.complete, true)
  assert.deepEqual(sig(layout), ['t:0-100', 't:100-180/r1', 't:180-260/r1'])
  assertPartition(layout, 't', 260)
  assertNoSliceTaller(layout, 't', bottoms, 100)
  assertLegalEnd(layout, 't', bottoms, cuts)
  assert.equal(layout.pages[1]?.slices[0]?.repeatHeaderRows, 1)
  assert.equal(layout.pages[2]?.slices[0]?.repeatHeaderRows, 1)
})

test('a cantSplit row shorter than a page moves whole and one taller than a page is cut', () => {
  const short = new Map([
    ['a', pm(50, [50])],
    ['t', pm(80, [], {
      rowBottoms: Float32Array.from([80]),
      rowCuts: Float32Array.from([40]),
      cantSplitRows: Uint8Array.from([1]),
    })],
  ])
  assert.deepEqual(sig(layoutOf([para('a'), para('t')], short, 100)), ['a:0-50', 't:0-80'])
  const bottoms = Float32Array.from([250])
  const cuts = Float32Array.from([100, 200])
  const tall = new Map([['t', pm(250, [], { rowBottoms: bottoms, rowCuts: cuts, cantSplitRows: Uint8Array.from([1]) })]])
  const layout = layoutOf([para('t')], tall, 100)
  assert.equal(layout.complete, true)
  assert.deepEqual(sig(layout), ['t:0-100', 't:100-200', 't:200-250'])
  assertPartition(layout, 't', 250)
  assertNoSliceTaller(layout, 't', bottoms, 100)
})

test('a loose cut still splits a row taller than the page', () => {
  const bottoms = Float32Array.from([250])
  const cuts = Float32Array.from([100, 200])
  const measures = new Map([['t', pm(250, [], {
    rowBottoms: bottoms,
    rowCuts: cuts,
    rowCutSplit: Uint8Array.from([0, 0]),
  })]])
  const layout = layoutOf([para('t')], measures, 100)
  assert.deepEqual(sig(layout), ['t:0-100', 't:100-200', 't:200-250'])
  assertLegalEnd(layout, 't', bottoms, cuts)
})

test('a header row is not cut even when it is taller than the page', () => {
  const measures = new Map([['t', pm(150, [], {
    rowBottoms: Float32Array.from([150]),
    headerRows: 1,
    rowCuts: Float32Array.from([50, 100]),
  })]])
  assert.deepEqual(sig(layoutOf([para('t')], measures, 100)), ['t:0-150'])
})

test('a cut with no allowed position falls at the page end', () => {
  const bottoms = Float32Array.from([250])
  const none = new Map([['t', pm(250, [], { rowBottoms: bottoms })]])
  const open = layoutOf([para('t')], none, 100)
  assert.equal(open.complete, true)
  assert.deepEqual(sig(open), ['t:0-100', 't:100-200', 't:200-250'])
  assertPartition(open, 't', 250)
  assertNoSliceTaller(open, 't', bottoms, 100)
  assertLegalEnd(open, 't', bottoms, new Float32Array())
  const later = Float32Array.from([150, 250])
  const skipped = new Map([['t', pm(300, [], { rowBottoms: Float32Array.from([300]), rowCuts: later })]])
  const layout = layoutOf([para('t')], skipped, 100)
  assert.deepEqual(sig(layout), ['t:0-100', 't:100-150', 't:150-250', 't:250-300'])
  assertLegalEnd(layout, 't', Float32Array.from([300]), later)
  assertNoSliceTaller(layout, 't', Float32Array.from([300]), 100)
})

test('interior cuts that do not fit the remainder leave the old row packing', () => {
  const measures = new Map([['t', pm(120, [], {
    rowBottoms: Float32Array.from([30, 60, 90, 120]),
    headerRows: 1,
    rowCuts: Float32Array.from([15, 45, 75, 105]),
  })]])
  assert.deepEqual(sig(layoutOf([para('t')], measures, 70)), ['t:0-60', 't:60-90/r1', 't:90-120/r1'])
})

test('a cut that fits the remainder splits that row', () => {
  const bottoms = Float32Array.from([30, 60, 90, 120])
  const cuts = Float32Array.from([65])
  const measures = new Map([['t', pm(120, [], { rowBottoms: bottoms, headerRows: 1, rowCuts: cuts })]])
  const layout = layoutOf([para('t')], measures, 70)
  assert.deepEqual(sig(layout), ['t:0-65', 't:65-90/r1', 't:90-120/r1'])
  assertLegalEnd(layout, 't', bottoms, cuts)
})

test('an edit in a cut row reuses the pages before the table', () => {
  const measures = new Map([
    ['a', pm(100, [100])],
    ['b', pm(100, [100])],
    ['t', pm(250, [], { rowBottoms: Float32Array.from([250]), rowCuts: Float32Array.from([100, 200]) })],
  ])
  const blocks = [para('a'), para('b'), para('t')]
  const full = layoutOf(blocks, measures, 100)
  assert.deepEqual(sig(full), ['a:0-100', 'b:0-100', 't:0-100', 't:100-200', 't:200-250'])
  const again = layoutOf(blocks, measures, 100, full, ['t'])
  assert.equal(again.pages[0], full.pages[0])
  const first = pagesFor(full, 't')[0] ?? 0
  assert.ok(first >= 2, 'the table starts after a page the restart can keep')
  for (let i = 0; i < first; i++) assert.deepEqual(again.pages[i]?.slices, full.pages[i]?.slices)
  assert.deepEqual(sig(again), sig(full))
})

test('cuts fall only where every covering cell has a line bottom or has ended', () => {
  const both = tableRowCuts(Float32Array.from([40]), [
    { top: 0, bottom: 40, contentEnd: 30, breaks: [{ y: 10, tight: true }, { y: 20, tight: true }, { y: 30, tight: true }] },
    { top: 0, bottom: 40, contentEnd: 30, breaks: [{ y: 15, tight: true }, { y: 30, tight: true }] },
  ])
  assert.deepEqual([...both.cuts], [30])
  const ended = tableRowCuts(Float32Array.from([40]), [
    { top: 0, bottom: 40, contentEnd: 20, breaks: [{ y: 10, tight: true }, { y: 20, tight: true }] },
    { top: 0, bottom: 40, contentEnd: 30, breaks: [{ y: 10, tight: true }, { y: 20, tight: true }, { y: 30, tight: true }] },
  ])
  assert.deepEqual([...ended.cuts], [10, 20, 30])
})

test('a row-spanning cell blocks a cut through its text in every row it covers', () => {
  const found = tableRowCuts(Float32Array.from([40, 80]), [
    { top: 0, bottom: 80, contentEnd: 72, breaks: [{ y: 16, tight: true }, { y: 56, tight: true }, { y: 72, tight: true }] },
    { top: 0, bottom: 40, contentEnd: 32, breaks: [{ y: 16, tight: true }, { y: 32, tight: true }] },
    { top: 40, bottom: 80, contentEnd: 64, breaks: [{ y: 48, tight: true }, { y: 56, tight: true }, { y: 64, tight: true }] },
  ])
  assert.deepEqual([...found.cuts], [16, 56, 72])
})

test('a cut is tight only when every covering cell accepts it tightly', () => {
  const one = tableRowCuts(Float32Array.from([100]), [
    {
      top: 0,
      bottom: 100,
      contentEnd: 100,
      breaks: [
        { y: 20, tight: false },
        { y: 40, tight: true },
        { y: 60, tight: false },
        { y: 80, tight: true },
      ],
    },
  ])
  assert.deepEqual([...one.cuts], [20, 40, 60, 80])
  assert.deepEqual([...one.split], [0, 1, 0, 1])
  const disagree = tableRowCuts(Float32Array.from([80]), [
    { top: 0, bottom: 80, contentEnd: 80, breaks: [{ y: 40, tight: true }] },
    { top: 0, bottom: 80, contentEnd: 80, breaks: [{ y: 40, tight: false }] },
  ])
  assert.deepEqual([...disagree.cuts], [40])
  assert.deepEqual([...disagree.split], [0])
})

test('a caret on the cut stays on the page that ends there', () => {
  const spans = [
    { page: 0, y0: 0, y1: 400 },
    { page: 1, y0: 400, y1: 800 },
    { page: 2, y0: 800, y1: 1100 },
  ]
  assert.equal(pageAtSliceY(spans, 400), 0, 'the last line of a cell ends on the cut')
  assert.equal(pageAtSliceY(spans, 399), 0)
  assert.equal(pageAtSliceY(spans, 401), 1, 'the next line starts after the cut')
  assert.equal(pageAtSliceY(spans, 800), 1)
  assert.equal(pageAtSliceY(spans, 800.5), 2)
  assert.equal(pageAtSliceY(spans, 200), 0)
})

test('the caret after the last character stays on that character\'s line', () => {
  const yAt = (offset: number): number | null => (offset >= 4 ? 420 : 380)
  assert.equal(caretSliceY(4, 4, yAt), 380, 'an end caret past the cut keeps the last character')
  assert.equal(caretSliceY(2, 4, yAt), 380, 'a caret inside the line keeps its own y')
  assert.equal(caretSliceY(0, 4, () => 10), 10)
  assert.equal(caretSliceY(3, 3, () => 50), 50, 'an end caret on the same y stays')
})

test('up and down across a cut stay in the body and keep the goal column\'s line', () => {
  const body = (globalMid: number, id: string) => ({ globalMid, left: 100, right: 140, story: 'body' as const, id })
  const from = body(900, 'from')
  const header = { globalMid: 1040, left: 80, right: 400, story: 'header' as const, id: 'header' }
  const to = body(1120, 'to')
  const after = body(1160, 'after')
  const lines = [from, header, to, after]
  const goal = 138
  const down = nextVerticalLine(lines, from, 'down', goal)
  assert.equal(down?.id, 'to', 'down crosses the cut into the next cell paragraph, not the header')
  const back = down ? nextVerticalLine(lines, down, 'up', goal) : null
  assert.equal(back?.id, 'from', 'up from the first line on the next page returns to the line before the cut')
  assert.equal(back?.left, from.left)
  assert.equal(back?.right, from.right)
})

test('a rejected candidate does not swallow the next cut', () => {
  const found = tableRowCuts(Float32Array.from([40]), [
    { top: 0, bottom: 40, contentEnd: 40, breaks: [{ y: 10, tight: true }, { y: 10.25, tight: true }] },
    { top: 0, bottom: 9.5, contentEnd: 20, breaks: [{ y: 4, tight: true }] },
  ])
  assert.deepEqual([...found.cuts], [10.25])
})

// PageView.update needs a document, layout boxes, and IntersectionObserver. The kit has none of those,
// so these tests drive paginate the way update does: measure the plan's blocks, then read the cache.
const LONG_COUNT = 1000

function contentWidthPx(part: Section): number {
  const margin = part.page.margin
  return Math.max(1, (part.page.widthPt - margin.left - margin.gutter - margin.right) * PT_TO_PX)
}

function synthetic(block: Block, part: Section): Measure {
  const width = contentWidthPx(part)
  const chars = block.kind === 'p' ? Math.max(1, flatText(block).length) : 12
  const lines = Math.max(1, Math.ceil((chars * 7) / width))
  const bottoms = Array.from({ length: lines }, (_, index) => (index + 1) * 18)
  return pm(lines * 18, bottoms)
}

function sectionOf(doc: Doc, block: Block): Section {
  let index = 0
  for (const item of doc.body) {
    if (item === block) {
      const found = doc.sections[index] ?? doc.sections[0]
      assert.ok(found, 'section')
      return found
    }
    if (item.kind === 'p' && item.sectionEnd !== null) index = item.sectionEnd + 1
  }
  const first = doc.sections[0]
  assert.ok(first, 'section')
  return first
}

function many(count: number, value: string): Paragraph[] {
  const blocks: Paragraph[] = []
  for (let i = 0; i < count; i++) blocks.push(paragraph(`p${i}`, [text(value)]))
  return blocks
}

function thousand(): { doc: Doc, state: EditorState } {
  const doc = makeDoc(many(LONG_COUNT, 'abcdef'))
  return { doc, state: initialState(doc) }
}

function place(state: EditorState, anchor: Pos, head: Pos = anchor): EditorState {
  const tr = setSelection({ kind: 'text', anchor, head })(state)
  if (!tr) {
    assert.deepEqual(anchor, head, 'selection did not apply')
    return state
  }
  return apply(state, tr, 0)
}

function run(state: EditorState, command: Command): EditorState {
  const tr = command(state)
  assert.ok(tr, 'command did not apply')
  return apply(state, tr, 0)
}

function newIds(prev: readonly Block[], next: readonly Block[]): string[] {
  const old = new Set<Block>(prev)
  const ids: string[] = []
  for (const block of next) if (!old.has(block)) ids.push(block.id)
  return ids
}

function findBlock(doc: Doc, id: string): Block | null {
  for (const block of doc.body) if (block.id === id) return block
  return null
}

interface CachedView {
  cache: WeakMap<Block, Measure>
  layout: Layout
}

function showLayout(doc: Doc): CachedView {
  const cache = new WeakMap<Block, Measure>()
  for (const block of doc.body) cache.set(block, synthetic(block, sectionOf(doc, block)))
  const layout = paginate(doc, (block) => cache.get(block), null, [])
  assert.equal(layout.complete, true, 'the fresh layout finished')
  return { cache, layout }
}

function assertLayouts(actual: Layout, full: Layout, label: string): void {
  assert.equal(actual.complete, full.complete, `${label} complete`)
  assert.equal(actual.pages.length, full.pages.length, `${label} pages ${actual.pages.length} vs ${full.pages.length}`)
  for (let i = 0; i < full.pages.length; i++) assert.deepEqual(actual.pages[i], full.pages[i], `${label} page ${i}`)
  assert.deepEqual(actual.sectionKeys, full.sectionKeys, `${label} section keys`)
}

function checkUpdate(before: Doc, after: Doc, view: CachedView | null, label: string): { measured: number, plan: LayoutPlan } {
  const base = view ?? showLayout(before)
  const plan = layoutPlan(before, after, new Map(), new Map())
  const overlay = new WeakMap<Block, Measure>()
  let measured = 0
  for (const id of plan.remeasure) {
    const block = findBlock(after, id)
    assert.ok(block, `${label} remeasure ${id} is not in the body`)
    overlay.set(block, synthetic(block, sectionOf(after, block)))
    measured += 1
  }
  const missing = new Set<string>()
  const layout = paginate(after, (block) => {
    const found = overlay.get(block) ?? base.cache.get(block)
    if (!found) missing.add(block.id)
    return found
  }, base.layout, plan.changed)
  assert.equal(missing.size, 0, `${label} uncached ${[...missing].join(',')}`)
  const fresh = paginate(after, (block, part) => synthetic(block, part), null, [])
  assertLayouts(layout, fresh, label)
  return { measured, plan }
}

interface Sample {
  name: string
  before: Doc
  after: Doc
}

function requireParagraph(block: Block | undefined, label: string): Paragraph {
  assert.ok(block?.kind === 'p', label)
  return block
}

function samples(state: EditorState): Sample[] {
  const body = state.doc.body
  const fifth = requireParagraph(body[4], 'paragraph 5')
  const hundredth = requireParagraph(body[99], 'paragraph 100')
  const victim = requireParagraph(body[40], 'deleted paragraph')
  const victimNext = requireParagraph(body[41], 'paragraph after the deletion')
  const pasted = requireParagraph(body[10], 'pasted paragraph')
  const host = requireParagraph(body[19], 'table host')

  const entered = run(place(state, { p: fifth.id, o: 3 }), splitParagraph)
  const undone = undo(entered)
  assert.ok(undone, 'undo of enter')
  const joined = run(place(state, { p: hundredth.id, o: 0 }), deleteBackward('char'))
  const deleted = run(
    place(state, { p: victim.id, o: 0 }, { p: victimNext.id, o: 0 }),
    deleteSelection,
  )
  const pastedState = run(
    place(state, { p: pasted.id, o: 0 }, { p: pasted.id, o: flatText(pasted).length }),
    insertText('one\ntwo\nthree'),
  )
  const withTable = run(place(state, { p: host.id, o: flatText(host).length }), insertTable(2, 2))
  const grid = withTable.doc.body.find((block) => block.kind === 'table')
  assert.ok(grid?.kind === 'table', 'inserted table')
  const withoutTable = run(withTable, deleteTable(grid.id))

  assert.equal(entered.doc.body.length, state.doc.body.length + 1, 'enter adds a paragraph')
  assert.equal(joined.doc.body.length, state.doc.body.length - 1, 'backspace removes a paragraph')
  assert.equal(deleted.doc.body.length, state.doc.body.length - 1, 'deleting a paragraph removes it')
  assert.equal(pastedState.doc.body.length, state.doc.body.length + 2, 'the paste makes three paragraphs of one')
  assert.equal(newIds(state.doc.body, pastedState.doc.body).length, 3, 'the paste adds three new paragraph objects')
  assert.equal(withTable.doc.body.length, state.doc.body.length + 1, 'the table is one new block')
  assert.equal(withoutTable.doc.body.length, withTable.doc.body.length - 1, 'deleting the table removes it')
  assert.equal(undone.doc.body.length, state.doc.body.length, 'undo restores the paragraph count')

  return [
    { name: 'enter', before: state.doc, after: entered.doc },
    { name: 'backspace join', before: state.doc, after: joined.doc },
    { name: 'delete paragraph', before: state.doc, after: deleted.doc },
    { name: 'paste', before: state.doc, after: pastedState.doc },
    { name: 'insert table', before: state.doc, after: withTable.doc },
    { name: 'delete table', before: withTable.doc, after: withoutTable.doc },
    { name: 'undo enter', before: entered.doc, after: undone.doc },
  ]
}

test('layoutPlan on a long document remeasures only the new blocks', () => {
  const { state } = thousand()
  for (const sample of samples(state)) {
    const plan = layoutPlan(sample.before, sample.after, new Map(), new Map())
    assert.deepEqual([...plan.remeasure], newIds(sample.before.body, sample.after.body), sample.name)
    if (sample.name === 'enter') assert.ok(plan.remeasure.length <= 2, `enter remeasures ${plan.remeasure.length}`)
    if (sample.name === 'delete table') assert.deepEqual([...plan.remeasure], [], 'deleting a table measures nothing')
  }
})

test('a section break remeasures later blocks only when the page width changes', () => {
  const paras = many(50, `${'abcdef '.repeat(40)}`)
  const wide = section(page(612, 792))
  const prev = makeDoc(paras, { sections: [wide] })
  const breaker = paragraph('break', [text('end')], {}, { sectionEnd: 0 })
  const moved = { ...prev, body: [breaker, ...paras], sections: [wide, section(page(280, 792))] }
  const stayed = { ...prev, body: [breaker, ...paras], sections: [wide, section(page(612, 792))] }
  const different = layoutPlan(prev, moved, new Map(), new Map())
  const equal = layoutPlan(prev, stayed, new Map(), new Map())
  assert.deepEqual([...different.remeasure], [breaker.id, ...paras.map((para) => para.id)], 'a narrower page measures the blocks that moved')
  assert.deepEqual([...equal.remeasure], [breaker.id], 'equal pages leave the following blocks cached')
  checkUpdate(prev, moved, null, 'narrower section')
  checkUpdate(prev, stayed, null, 'equal section')
})

test('incremental layout matches a fresh layout after each long-document edit', () => {
  const { doc, state } = thousand()
  const view = showLayout(doc)
  assert.ok(view.layout.pages.length > 5, 'the document covers more than one page')
  for (const sample of samples(state)) {
    const checked = checkUpdate(sample.before, sample.after, sample.before === doc ? view : null, sample.name)
    if (sample.name === 'enter') assert.ok(checked.measured <= 2, `measured ${checked.measured}`)
  }
})

test('typing one character remeasures only that paragraph', () => {
  const { doc, state } = thousand()
  const para = requireParagraph(doc.body[400], 'typed paragraph')
  const next = run(place(state, { p: para.id, o: 1 }), insertText('x'))
  const plan = layoutPlan(doc, next.doc, new Map(), new Map())
  assert.deepEqual([...plan.remeasure], [para.id])
  const checked = checkUpdate(doc, next.doc, null, 'type')
  assert.equal(checked.measured, 1)
})

test('a list indent change remeasures only the paragraph that uses the level', () => {
  const listed = paragraph('a', [text('abcdef')], { numbering: { num: numId(1), level: 0 } })
  const other = paragraph('b', [text('bbbbbb')])
  const lists = (hanging: number) => ({
    byNum: new Map([[numId(1), {
      fresh: true,
      levels: [{
        format: 'decimal' as const,
        text: '%1.',
        start: 1,
        indentLeft: pt(36),
        hanging: pt(hanging),
      }],
    }]]),
  })
  const prev = makeDoc([listed, other], { lists: lists(18) })
  const changed = layoutPlan(prev, { ...prev, lists: lists(36) }, new Map(), new Map())
  const same = layoutPlan(prev, { ...prev, lists: lists(18) }, new Map(), new Map())
  assert.deepEqual([...changed.remeasure], [listed.id], 'the level indent changes the paragraph measure')
  assert.deepEqual([...same.remeasure], [], 'the same indent leaves the measure cached')
})

test('a marker that appears on a hanging indent remeasures that paragraph', () => {
  const para = paragraph('a', [text('abcdef')], { indent: { left: pt(36), hanging: pt(18) } })
  const lists = {
    byNum: new Map([[numId(1), {
      fresh: true,
      levels: [{
        format: 'decimal' as const,
        text: '%1.',
        start: 1,
        indentLeft: pt(36),
        hanging: pt(18),
      }],
    }]]),
  }
  const prev = makeDoc([para], { lists })
  const styleId = prev.styles.defaultParagraph
  const entry = prev.styles.paragraph.get(styleId)
  assert.ok(entry, 'the default paragraph style')
  const paragraphStyles = new Map(prev.styles.paragraph)
  paragraphStyles.set(styleId, { ...entry, look: { ...entry.look, numbering: { num: numId(1), level: 0 } } })
  const next = { ...prev, styles: { ...prev.styles, paragraph: paragraphStyles } }
  assert.equal(listMarkers(prev).get(para.id), undefined, 'the paragraph has no marker yet')
  assert.equal(listMarkers(next).get(para.id), '1.', 'the style numbering supplies a marker')
  const plan = layoutPlan(prev, next, listMarkers(prev), listMarkers(next))
  assert.deepEqual([...plan.remeasure], [para.id], 'the marker changes the hanging line box')
})

test('a document grid change remeasures every block', () => {
  const blocks = [paragraph('g0', [text('a')]), paragraph('g1', [text('b')])]
  const plain = section()
  const prev = makeDoc(blocks, { sections: [plain] })
  const next = { ...prev, sections: [{ ...plain, grid: { type: 'lines' as const, linePitch: pt(18) } }] }
  const plan = layoutPlan(prev, next, new Map(), new Map())
  const first = blocks[0]
  const second = blocks[1]
  assert.ok(first && second)
  assert.deepEqual([...plan.remeasure], [first.id, second.id])
})

test('incremental layout matches a fresh layout for 200 random edits', { timeout: 60_000 }, () => {
  const { doc, state } = thousand()
  const view = showLayout(doc)
  const rng = mulberry32(1)
  const kinds = ['enter', 'join', 'delete', 'paste', 'insertTable', 'deleteTable', 'undo'] as const
  for (let n = 0; n < 200; n++) {
    const kind = kinds[Math.floor(rng() * kinds.length)] ?? 'enter'
    const body = state.doc.body
    const last = body.length - 1
    let before = state.doc
    let after = state.doc
    if (kind === 'enter') {
      const index = Math.floor(rng() * body.length)
      const para = requireParagraph(body[index], `enter ${n}`)
      const offset = Math.floor(rng() * (flatText(para).length + 1))
      after = run(place(state, { p: para.id, o: offset }), splitParagraph).doc
    } else if (kind === 'join') {
      const index = 1 + Math.floor(rng() * last)
      const para = requireParagraph(body[index], `join ${n}`)
      after = run(place(state, { p: para.id, o: 0 }), deleteBackward('char')).doc
    } else if (kind === 'delete') {
      const index = Math.floor(rng() * last)
      const para = requireParagraph(body[index], `delete ${n}`)
      const next = requireParagraph(body[index + 1], `delete next ${n}`)
      after = run(place(state, { p: para.id, o: 0 }, { p: next.id, o: 0 }), deleteSelection).doc
    } else if (kind === 'paste') {
      const index = Math.floor(rng() * body.length)
      const para = requireParagraph(body[index], `paste ${n}`)
      after = run(
        place(state, { p: para.id, o: 0 }, { p: para.id, o: flatText(para).length }),
        insertText('one\ntwo\nthree'),
      ).doc
    } else if (kind === 'insertTable') {
      const index = Math.floor(rng() * last)
      const para = requireParagraph(body[index], `table ${n}`)
      after = run(place(state, { p: para.id, o: flatText(para).length }), insertTable(2, 2)).doc
    } else if (kind === 'deleteTable') {
      const index = Math.floor(rng() * last)
      const para = requireParagraph(body[index], `drop table ${n}`)
      const inserted = run(place(state, { p: para.id, o: flatText(para).length }), insertTable(2, 2))
      const grid = inserted.doc.body.find((block) => block.kind === 'table')
      assert.ok(grid?.kind === 'table', `table ${n}`)
      before = inserted.doc
      after = run(inserted, deleteTable(grid.id)).doc
    } else {
      const index = Math.floor(rng() * body.length)
      const para = requireParagraph(body[index], `undo ${n}`)
      const offset = Math.floor(rng() * (flatText(para).length + 1))
      const entered = run(place(state, { p: para.id, o: offset }), splitParagraph)
      const undone = undo(entered)
      assert.ok(undone, `undo ${n}`)
      before = entered.doc
      after = undone.doc
    }
    const checked = checkUpdate(before, after, before === doc ? view : null, `${kind} ${n}`)
    assert.deepEqual([...checked.plan.remeasure], newIds(before.body, after.body), `${kind} ${n} remeasure`)
  }
})

test('atomCaretBox puts the caret on the requested side of an image at its bottom baseline', () => {
  const rect = { left: 96, right: 420, top: 100, bottom: 260 }
  assert.deepEqual(atomCaretBox(rect, false, true, 18), { left: 96, top: 242, height: 18 })
  assert.deepEqual(atomCaretBox(rect, true, true, 18), { left: 420, top: 242, height: 18 })
  assert.deepEqual(atomCaretBox({ left: 96, right: 98, top: 100, bottom: 102 }, true, true, 18), { left: 98, top: 100, height: 18 })
  assert.deepEqual(atomCaretBox({ left: 96, right: 120, top: 100, bottom: 116 }, true, false, 18), { left: 120, top: 100, height: 16 })
})

test('resolveImageHit places clicks beside and below an image after it without crossing table cells', () => {
  const id = 'p1' as Paragraph['id']
  const bodyImg: ImageHitCandidate = {
    p: id,
    base: 2,
    left: 96,
    right: 720,
    top: 100,
    bottom: 340,
    cell: null,
    inkBefore: true,
    inkAfter: false,
    contentAbove: false,
    contentBelow: false,
  }
  assert.deepEqual(resolveImageHit([bodyImg], 719, 220, 760, 220), { p: id, o: 3 }, 'right margin lands after the image')
  assert.deepEqual(resolveImageHit([bodyImg], 140, 349, 140, 520), { p: id, o: 3 }, 'blank area below the image lands after it, even on the left side')
  assert.equal(resolveImageHit([bodyImg], 97, 220, 80, 220), null, 'text before the image keeps its own hit')

  const cellImg: ImageHitCandidate = {
    ...bodyImg,
    base: 0,
    left: 110,
    right: 190,
    top: 100,
    bottom: 160,
    cell: { left: 100, right: 200, top: 96, bottom: 240 },
    inkBefore: false,
  }
  assert.equal(resolveImageHit([cellImg], 260, 130, 260, 130), null, 'a click in a neighbor cell does not hit this cell\'s image')
  assert.deepEqual(resolveImageHit([cellImg], 195, 130, 195, 130), { p: id, o: 1 }, 'space to the right inside the same cell lands after the image')
  assert.deepEqual(resolveImageHit([cellImg], 120, 200, 120, 200), { p: id, o: 1 }, 'empty space below the image inside a tall cell lands after it')
})
