import assert from 'node:assert/strict'
import { test } from 'node:test'
import type { BorderEdge, BorderLine, CellBorders, LineSpacing, Pt, Row, RunLook, Section, Table, TableBorders } from '../model.ts'
import { cell, paragraph, pt, row, table, text, hex } from './testkit.ts'
import { cellBoxShadow, cellEdges, cellPadPt, collapsedSpaceBefore, markerBoxPx, pageBoxStyle, paintFamily, paintedCellEdges, preeditRunStyle, rowFrame, scriptSegments, snapLinePx } from './render.ts'
import type { RenderContext } from './render.ts'

const look = {
  fonts: { ascii: 'Times', hAnsi: 'TimesHigh', eastAsia: 'SimSun', cs: 'Naskh' },
  sizePt: 10.5 as Pt,
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  color: 'auto',
  caps: false,
  charSpacingPt: 0 as Pt,
} satisfies RunLook

test('scriptSegments gives CJK the east Asia font and Latin the ASCII font', () => {
  const segments = scriptSegments('Hello世界', look)
  assert.deepEqual(segments, [
    { text: 'Hello', family: 'Times', slot: 'ascii' },
    { text: '世界', family: 'SimSun', slot: 'eastAsia' },
  ])
})

test('Han in a face without CJK glyphs takes 宋体, and Latin in that run keeps its face', () => {
  const mixed = { ...look, fonts: { ascii: 'Calibri', hAnsi: 'Calibri', eastAsia: 'Calibri', cs: 'Calibri' } }
  const segments = scriptSegments('A中', mixed)
  const latin = segments[0]
  const han = segments[1]
  if (latin === undefined || han === undefined) throw new Error('segments')
  assert.equal(paintFamily(latin.family, latin.slot, false), 'Calibri')
  assert.equal(paintFamily(han.family, han.slot, false), '宋体')
  assert.equal(paintFamily(han.family, han.slot, true), 'Calibri', 'an east Asia face that has the glyphs keeps them')
  assert.equal(paintFamily(han.family, han.slot, undefined), 'Calibri', 'unloaded metrics stay on the run face')
})

test('scriptSegments keeps fullwidth forms on the east Asia font and high ANSI on hAnsi', () => {
  const segments = scriptSegments('A１é', look)
  assert.equal(segments.map((segment) => segment.text).join(''), 'A１é')
  assert.equal(segments[0]?.family, 'Times')
  assert.equal(segments[0]?.slot, 'ascii')
  assert.equal(segments[1]?.family, 'SimSun')
  assert.equal(segments[1]?.slot, 'eastAsia')
  assert.equal(segments[2]?.family, 'TimesHigh')
  assert.equal(segments[2]?.slot, 'hAnsi')
})

test('exact line spacing is not snapped to the document grid', () => {
  const exact: LineSpacing = { rule: 'exact', pt: pt(14) }
  assert.equal(snapLinePx(14 * (4 / 3), 20, exact), 14 * (4 / 3))
  assert.equal(snapLinePx(15, 20, { rule: 'auto', multiple: 1 }), 20)
  assert.equal(snapLinePx(15, 20, { rule: 'atLeast', pt: pt(10) }), 20)
  assert.equal(snapLinePx(20, 20, { rule: 'auto', multiple: 1 }), 20)
  assert.equal(snapLinePx(15, null, { rule: 'auto', multiple: 1 }), 15)
})

test('cell paragraphs collapse facing spacing, and the first paragraph keeps its space before', () => {
  assert.equal(collapsedSpaceBefore(10, null), 10)
  assert.equal(collapsedSpaceBefore(8, 12), 0)
  assert.equal(collapsedSpaceBefore(10, 6), 4)
  assert.equal(collapsedSpaceBefore(10, 0), 10)
})

test('pageBoxStyle converts points to px at 4/3', () => {
  const section = {
    src: null,
    page: {
      widthPt: 75 as Pt,
      heightPt: 150 as Pt,
      margin: {
        top: 15 as Pt,
        right: 7.5 as Pt,
        bottom: 15 as Pt,
        left: 7.5 as Pt,
        header: 0 as Pt,
        footer: 0 as Pt,
        gutter: 0 as Pt,
      },
    },
    breakType: 'nextPage',
    titlePage: false,
    headers: {},
    footers: {},
    grid: null,
  } satisfies Section
  const box = pageBoxStyle(section)
  assert.equal(box.page.width, '100px')
  assert.equal(box.page.height, '200px')
  assert.equal(box.content.left, '10px')
  assert.equal(box.content.top, '20px')
})

function edge(width: number, color: string, style: BorderLine['style'] = 'single'): BorderLine {
  return { style, widthPt: pt(width), color: hex(color) }
}

function all(line: BorderEdge): TableBorders {
  return { top: line, left: line, bottom: line, right: line, insideH: line, insideV: line }
}

function square(borders: TableBorders | undefined, cellBorders?: CellBorders): Table {
  const cellAt = (id: string, borders?: CellBorders) => {
    const made = cell([paragraph(id, [text(id)])])
    return borders === undefined ? made : { ...made, borders }
  }
  const built = table('t', [
    row([cellAt('a', cellBorders), cellAt('b')]),
    row([cellAt('c'), cellAt('d')]),
  ], [100, 100])
  return borders === undefined ? built : { ...built, props: { borders } }
}

function everyEdge(made: Table, line: BorderEdge): void {
  for (let r = 0; r < made.rows.length; r++) {
    for (let c = 0; c < (made.rows[r]?.cells.length ?? 0); c++) {
      assert.deepEqual(cellEdges(made, r, c), { top: line, left: line, bottom: line, right: line }, `${r},${c}`)
    }
  }
}

function sides(shadow: string): string[] {
  if (shadow === 'none') return []
  const found: string[] = []
  for (const part of shadow.split(', ')) {
    const nums = [...part.matchAll(/-?[\d.]+px/g)].map((match) => Number(match[0].replace('px', '')))
    const x = nums[0] ?? 0
    const y = nums[1] ?? 0
    if (Math.abs(x) < 0.001 && y > 0) found.push('top')
    else if (Math.abs(y) < 0.001 && x > 0) found.push('left')
    else if (Math.abs(x) < 0.001 && y < 0) found.push('bottom')
    else if (Math.abs(y) < 0.001 && x < 0) found.push('right')
  }
  return found
}

test('the four table shapes resolve the edges the file gives them', () => {
  const plain = square(undefined)
  everyEdge(plain, null)
  assert.equal(cellBoxShadow(paintedCellEdges(plain, 0, 0)), 'none', 'a table with no borders draws none')

  const direct = edge(1, 'FF0000')
  everyEdge(square(all(direct)), direct)

  const grid = edge(0.5, '000000')
  everyEdge(square(all(grid)), grid)

  const blue = edge(1.5, '0000FF')
  const over = square(all(blue), { right: null })
  assert.deepEqual(cellEdges(over, 0, 0), { top: blue, left: blue, bottom: blue, right: null })
  assert.deepEqual(cellEdges(over, 0, 1), { top: blue, left: blue, bottom: blue, right: blue })
  assert.deepEqual(cellEdges(over, 1, 0).right, blue)
  assert.deepEqual(cellEdges(over, 1, 1).bottom, blue)
})

test('a line shared by two cells is drawn once', () => {
  const line = edge(0.5, '000000')
  const made = square(all(line))
  assert.deepEqual(paintedCellEdges(made, 0, 0).right, line, 'the left cell paints the shared vertical line')
  assert.equal(paintedCellEdges(made, 0, 1).left, null, 'the right cell does not paint it again')
  assert.deepEqual(cellEdges(made, 0, 1).left, line, 'the right cell still resolves the line')
  assert.deepEqual(paintedCellEdges(made, 0, 0).bottom, line, 'the upper cell paints the shared horizontal line')
  assert.equal(paintedCellEdges(made, 1, 0).top, null, 'the lower cell does not paint it again')
  const left = sides(cellBoxShadow(paintedCellEdges(made, 0, 0)))
  const right = sides(cellBoxShadow(paintedCellEdges(made, 0, 1)))
  assert.deepEqual(left, ['top', 'left', 'bottom', 'right'])
  assert.deepEqual(right, ['top', 'bottom', 'right'])

  const over = square(all(line), { right: null })
  assert.equal(cellEdges(over, 0, 0).right, null, 'the cell still says its own right edge is none')
  assert.deepEqual(paintedCellEdges(over, 0, 0).right, line, 'the neighbour keeps the line between the cells')
  assert.equal(paintedCellEdges(over, 0, 1).left, null, 'the right cell does not paint it again')
  assert.ok(sides(cellBoxShadow(paintedCellEdges(over, 0, 0))).includes('right'))
  assert.ok(!sides(cellBoxShadow(paintedCellEdges(over, 0, 1))).includes('left'))
})

function boxed(id: string, borders?: CellBorders) {
  const made = cell([paragraph(id, [text(id)])])
  return borders === undefined ? made : { ...made, borders }
}

test('a shared line uses both facing edges', () => {
  const line = edge(1, '00AA00')
  const named = table('t', [
    row([boxed('a'), boxed('b', { left: line })]),
  ], [100, 100])
  assert.deepEqual(paintedCellEdges(named, 0, 0).right, line, 'a left edge named in column 1 is drawn')
  assert.equal(paintedCellEdges(named, 0, 1).left, null, 'the cell that names it does not paint it again')
  assert.ok(sides(cellBoxShadow(paintedCellEdges(named, 0, 0))).includes('right'))
  assert.equal(cellBoxShadow(paintedCellEdges(named, 0, 1)), 'none')

  const namedTop = table('t', [
    row([boxed('c')]),
    row([boxed('d', { top: line })]),
  ], [100])
  assert.deepEqual(paintedCellEdges(namedTop, 0, 0).bottom, line, 'a top edge named below row 0 is drawn')
  assert.equal(paintedCellEdges(namedTop, 1, 0).top, null)

  const kept = table('t', [
    row([boxed('e', { right: line }), boxed('f', { left: null })]),
  ], [100, 100])
  assert.deepEqual(paintedCellEdges(kept, 0, 0).right, line, 'one line and one none draws the line')

  const neither = table('t', [
    row([boxed('g', { right: null, bottom: null }), boxed('h', { left: null })]),
    row([boxed('i', { top: null }), boxed('j')]),
  ], [100, 100])
  const cleared = { ...neither, props: { borders: all(line) } }
  assert.equal(paintedCellEdges(cleared, 0, 0).right, null, 'two facing edges that are both none draw no line')
  assert.equal(paintedCellEdges(cleared, 0, 1).left, null)
  assert.equal(paintedCellEdges(cleared, 0, 0).bottom, null, 'a bottom of none and a top of none draw no line')

  const red = edge(2, 'FF0000')
  const blue = edge(0.5, '0000FF')
  const differ = table('t', [
    row([boxed('k', { right: red }), boxed('l', { left: blue })]),
  ], [100, 100])
  assert.deepEqual(paintedCellEdges(differ, 0, 0).right, blue, 'the later stroke is the right neighbour')
  const differRow = table('t', [
    row([boxed('m', { bottom: red })]),
    row([boxed('n', { top: blue })]),
  ], [100])
  assert.deepEqual(paintedCellEdges(differRow, 0, 0).bottom, blue, 'the later stroke is the lower neighbour')
  assert.equal(paintedCellEdges(differRow, 1, 0).top, null, 'the lower cell does not paint it again')

  const across = table('t', [
    row([cell([paragraph('o', [text('o')])], 2)]),
    row([boxed('p', { top: red }), boxed('q', { top: blue })]),
  ], [100, 100])
  assert.deepEqual(paintedCellEdges(across, 0, 0).bottom, red, 'several cells along one edge use the first')
  const down = table('t', [
    row([
      { ...cell([paragraph('r', [text('r')])]), vMerge: 'restart' },
      boxed('s', { left: red }),
    ]),
    row([
      { ...cell([paragraph('u', [text('u')])]), vMerge: 'continue' },
      boxed('v', { left: blue }),
    ]),
  ], [100, 100])
  assert.deepEqual(paintedCellEdges(down, 0, 0).right, red, 'several cells down one edge use the first')
  assert.equal(paintedCellEdges(down, 0, 1).left, null, 'the spanning cell paints that line once')
})

test('row height is an exact clip, a minimum, or nothing', () => {
  const auto = row([cell([paragraph('a', [text('a')])])])
  assert.deepEqual(rowFrame(auto), { height: null, minHeight: null, clip: false })
  const atLeast: Row = { ...auto, props: { heightPt: pt(18), heightRule: 'atLeast' } }
  assert.deepEqual(rowFrame(atLeast), { height: '24px', minHeight: '24px', clip: false })
  const exact: Row = { ...auto, props: { heightPt: pt(24), heightRule: 'exact' } }
  assert.deepEqual(rowFrame(exact), { height: '32px', minHeight: null, clip: true })
  const implicit: Row = { ...auto, props: { heightPt: pt(12) } }
  assert.deepEqual(rowFrame(implicit), { height: '16px', minHeight: '16px', clip: false }, 'a height with no rule is a minimum')
})

test('cell padding is the Word default, then the table, then the cell', () => {
  const made = table('t', [row([cell([paragraph('a', [text('a')])])])])
  const plain = made.rows[0]?.cells[0]
  assert.ok(plain)
  if (!plain) return
  assert.deepEqual(cellPadPt(made, plain), { top: 0, right: 5.4, bottom: 0, left: 5.4 })
  const withTable: Table = { ...made, props: { margins: { top: pt(1), right: pt(2), bottom: pt(3), left: pt(4) } } }
  assert.deepEqual(cellPadPt(withTable, plain), { top: 1, right: 2, bottom: 3, left: 4 })
  const over = { ...plain, margins: { top: pt(0), left: pt(8) } }
  assert.deepEqual(cellPadPt(withTable, over), { top: 0, right: 2, bottom: 3, left: 8 }, 'an explicit 0 overrides the table')
})

test('preeditRunStyle uses the ASCII face for pinyin and the East Asian face for Han at the caret marks', () => {
  const para = paragraph('p', [text('你是谁')])
  const ctx: RenderContext = {
    look: (_p, marks) => ({
      ...look,
      fonts: { ascii: 'Calibri', hAnsi: 'Calibri', eastAsia: 'Calibri', cs: 'Times New Roman' },
      sizePt: marks.sizePt ?? pt(11),
      bold: marks.bold === true,
    }),
    paraLook: () => {
      throw new Error('unused')
    },
    fontStack: (family) => `stack:${family}`,
    lineHeightPx: () => 20,
    cjkGlyphs: () => false,
    marker: () => null,
    imageUrl: () => '',
    gridPitchPx: null,
  }
  const pinyin = preeditRunStyle(ctx, para, { sizePt: pt(12), bold: true }, "ni'zai'na")
  assert.equal(pinyin.slot, 'ascii')
  assert.equal(pinyin.fontFamily, 'stack:Calibri')
  assert.equal(pinyin.fontSizePx, 16)
  assert.equal(pinyin.fontWeight, '700')
  const han = preeditRunStyle(ctx, para, {}, '你在哪')
  assert.equal(han.slot, 'eastAsia')
  assert.equal(han.fontFamily, 'stack:宋体')
  assert.equal(han.fontSizePx, 11 * (4 / 3))
})

test('markerBoxPx keeps paragraph padding past the marker even when left or hanging is zero', () => {
  assert.deepEqual(markerBoxPx(48, 24), { markerLeft: 24, markerWidth: 24, paddingLeft: 48 })
  assert.deepEqual(markerBoxPx(0, 24), { markerLeft: 0, markerWidth: 24, paddingLeft: 24 })
  assert.deepEqual(markerBoxPx(16, 0), { markerLeft: 16, markerWidth: 24, paddingLeft: 40 })
})
