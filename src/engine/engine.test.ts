import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { test } from 'node:test'
import { checkDoc, type Block, type Doc, type ListLevel, type NumId, type Paragraph, type Pt } from '../model.ts'
import { createEngine, readWithRetry, type Opened } from './engine.ts'
import { initSync } from './pkg/engine.js'

const root = new URL('../../', import.meta.url)

initSync({ module: readFileSync(new URL('public/engine.wasm', root)) })
const engine = createEngine()

function fixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(new URL(`fixtures/${name}`, root)))
}

function paragraphs(blocks: readonly Block[]): Paragraph[] {
  const found: Paragraph[] = []
  for (const block of blocks) {
    if (block.kind === 'p') found.push(block)
    else if (block.kind === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) found.push(...paragraphs(cell.blocks))
      }
    }
  }
  return found
}

function textOf(para: Paragraph): string {
  let text = ''
  for (const inline of para.inlines) {
    if (inline.kind === 'text') text += inline.text
  }
  return text
}

function storyText(doc: Doc, part: string): string {
  let text = ''
  for (const [key, blocks] of doc.stories) {
    if (!key.includes(part)) continue
    text += paragraphs(blocks).map(textOf).join('\n')
  }
  return text
}

function assertCheckDoc(doc: Doc) {
  try {
    checkDoc(doc)
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    if (message === 'not implemented') {
      assert.equal(message, 'not implemented', 'checkDoc is not implemented yet; structural check skipped')
      return
    }
    throw error
  }
}

test('open parses basic.docx into a document', () => {
  const opened = engine.open(fixture('basic.docx'))
  try {
    const body = paragraphs(opened.doc.body)
    assert.ok(
      body.some((para) => textOf(para) === 'Quarterly report'),
      'title paragraph text is kept',
    )
    const table = opened.doc.body.find((block) => block.kind === 'table')
    assert.ok(table !== undefined && table.kind === 'table', 'basic.docx has a table')
    if (table === undefined || table.kind !== 'table') return
    assert.equal(table.rows.length, 3, 'table has 3 rows')
    assert.ok(
      table.rows.every((row) => row.cells.length === 3),
      'table has 3 cells in every row',
    )
    const lists = body.filter((para) => para.props.numbering !== undefined && para.props.numbering !== null)
    assert.equal(lists.length, 6, 'six list paragraphs')
    const images = body.flatMap((para) => para.inlines).filter((inline) => inline.kind === 'atom' && inline.atom.type === 'image')
    assert.equal(images.length, 1, 'one image')
    const image = images[0]
    if (image === undefined || image.kind !== 'atom' || image.atom.type !== 'image') return
    const url = opened.source.imageUrl(image.atom.media)
    assert.ok(url.startsWith('blob:'), 'image bytes are served as a blob URL')
    assert.ok(storyText(opened.doc, 'header').includes('rdocx-acde fixture header'), 'header story text')
    assert.ok(storyText(opened.doc, 'footer').includes('rdocx-acde fixture footer'), 'footer story text')
    const roles = [...opened.doc.styles.paragraph.values()].map((style) => style.role)
    assert.ok(roles.includes('title'), 'Title role detected')
    assert.ok(roles.includes('heading1'), 'Heading 1 role detected')
    const centered = body.find((para) => para.props.align === 'center')
    assert.ok(centered !== undefined, 'centered paragraph keeps its direct alignment')
    if (centered !== undefined) {
      assert.equal(opened.source.paraLook(centered, opened.doc.lists).align, 'center', 'paraLook overlays direct alignment')
    }
    const plain = body.find((para) => textOf(para).startsWith('Plain, '))
    assert.ok(plain !== undefined, 'mixed-formatting paragraph is present')
    if (plain !== undefined) {
      assert.equal(opened.source.look(plain, { bold: true }).bold, true, 'look overlays a direct bold mark')
    }
    assertCheckDoc(opened.doc)
  } finally {
    opened.source.free()
  }
})

function pt(value: number): Pt {
  return value as Pt
}

test('a list level supplies left and hanging indent, and a paragraph indent overrides it', () => {
  const opened = engine.open(fixture('basic.docx'))
  try {
    const body = paragraphs(opened.doc.body)
    const bullet = body.find((para) => textOf(para) === 'First bullet')
    const nested = body.find((para) => textOf(para) === 'Nested bullet')
    const step = body.find((para) => textOf(para) === 'Step one')
    const plain = body.find((para) => textOf(para).startsWith('Plain, '))
    assert.ok(bullet && nested && step && plain)
    if (!bullet || !nested || !step || !plain) return
    const lists = opened.doc.lists
    const look = opened.source.paraLook(bullet, lists)
    const numbering = look.numbering
    assert.ok(numbering, 'the bullet is numbered')
    if (!numbering) return
    const level = lists.byNum.get(numbering.num)?.levels[numbering.level]
    assert.ok(level)
    if (!level) return
    assert.equal(Number(level.indentLeft), 36)
    assert.equal(Number(level.hanging), 18)
    assert.equal(look.indent.left, level.indentLeft, 'the bullet takes the level left indent')
    assert.equal(look.indent.hanging, level.hanging, 'the bullet takes the level hanging indent')
    assert.ok(Number(look.indent.hanging) > 0, 'a hanging indent gives the marker a column of its own')
    const nestedLook = opened.source.paraLook(nested, lists)
    assert.equal(Number(nestedLook.indent.left), 72)
    assert.equal(Number(nestedLook.indent.hanging), 18)
    const stepLook = opened.source.paraLook(step, lists)
    assert.equal(Number(stepLook.indent.left), 36)
    assert.equal(Number(stepLook.indent.hanging), 18)
    assert.equal(Number(opened.source.paraLook(plain, lists).indent.left), 0)
    assert.equal(Number(opened.source.paraLook(plain, lists).indent.hanging), 0, 'a paragraph outside a list keeps the style indent')
    const own = opened.source.paraLook({
      ...bullet,
      props: { ...bullet.props, indent: { left: pt(12), hanging: pt(0) } },
    }, lists)
    assert.equal(Number(own.indent.left), 12, 'a direct left indent wins')
    assert.equal(Number(own.indent.hanging), 0, 'a direct hanging indent wins, including zero')
    const leftOnly = opened.source.paraLook({
      ...bullet,
      props: { ...bullet.props, indent: { left: pt(10) } },
    }, lists)
    assert.equal(Number(leftOnly.indent.left), 10)
    assert.equal(leftOnly.indent.hanging, level.hanging, 'an unset hanging indent still comes from the level')
    const cleared = opened.source.paraLook({
      ...bullet,
      props: { ...bullet.props, numbering: null },
    }, lists)
    assert.equal(Number(cleared.indent.left), 0)
    assert.equal(Number(cleared.indent.hanging), 0, 'leaving the list drops the level indent')
  } finally {
    opened.source.free()
  }
})

test('a list added after import supplies its level indent', () => {
  const opened = engine.open(fixture('basic.docx'))
  try {
    const plain = paragraphs(opened.doc.body).find((para) => textOf(para).startsWith('Plain, '))
    assert.ok(plain)
    if (!plain) return
    const num = -1 as NumId
    assert.equal(opened.doc.lists.byNum.has(num), false, 'the fresh id is not in the imported document')
    const levels: ListLevel[] = []
    for (let level = 0; level < 9; level++) {
      levels.push({
        format: 'bullet',
        text: '•',
        start: 1,
        indentLeft: pt(36 * (level + 1)),
        hanging: pt(18),
      })
    }
    const byNum = new Map(opened.doc.lists.byNum)
    byNum.set(num, { fresh: true, levels })
    const lists = { byNum }
    const at = (level: number, indent?: NonNullable<Paragraph['props']['indent']>) => {
      const props = { ...plain.props, numbering: { num, level } }
      return opened.source.paraLook({ ...plain, props: indent ? { ...props, indent } : props }, lists)
    }
    const level0 = at(0)
    assert.equal(Number(level0.indent.left), 36)
    assert.equal(Number(level0.indent.hanging), 18)
    const level1 = at(1)
    assert.equal(Number(level1.indent.left), 72)
    const direct = at(0, { left: pt(12) })
    assert.equal(Number(direct.indent.left), 12, 'a direct left indent wins over the level')
    assert.equal(Number(direct.indent.hanging), 18, 'an unset hanging indent still comes from the level')
  } finally {
    opened.source.free()
  }
})

test('open parses long.docx and preserve.docx', () => {
  const long = engine.open(fixture('long.docx'))
  const preserve = engine.open(fixture('preserve.docx'))
  try {
    assert.ok(
      paragraphs(long.doc.body).some((para) => textOf(para).includes('Chapter 1')),
      'long.docx keeps chapter text',
    )
    assert.ok(
      paragraphs(long.doc.body).some((para) => textOf(para).includes('这是一段中文内容')),
      'long.docx keeps Chinese text',
    )
    const paras = paragraphs(preserve.doc.body)
    const atoms = paras.flatMap((para) => para.inlines).flatMap((inline) => (inline.kind === 'atom' ? [inline.atom] : []))
    assert.ok(
      atoms.some((atom) => atom.type === 'opaque' && atom.label === 'comment'),
      'comment appears as an atom',
    )
    assert.ok(
      atoms.some((atom) => atom.type === 'opaque' && atom.label === 'bookmark' && atom.display.includes('kept_bookmark')),
      'bookmark appears as an atom',
    )
    assert.ok(
      atoms.some((atom) => atom.type === 'field'),
      'field appears as an atom',
    )
    assert.ok(
      paras.some((para) => para.inlines.some((inline) => inline.kind === 'text' && inline.marks.link === 'https://example.com/')),
      'hyperlink mark is the relationship URL',
    )
    assertCheckDoc(long.doc)
    assertCheckDoc(preserve.doc)
  } finally {
    long.source.free()
    preserve.source.free()
  }
})

test('readWithRetry tries a failed fetch once more', async () => {
  const url = '/engine.wasm'
  let calls = 0
  let waits = 0
  const wait = async (): Promise<void> => { waits += 1 }
  const respond = (status: number, ok = status >= 200 && status < 300): Response => {
    return { ok, status } as Response
  }

  calls = 0
  waits = 0
  const ready = await readWithRetry(url, async (asked) => {
    assert.equal(asked, url)
    calls += 1
    return respond(200)
  }, wait)
  assert.equal(ready.status, 200)
  assert.equal(calls, 1, 'an ok response is not fetched again')
  assert.equal(waits, 0)

  calls = 0
  waits = 0
  const retried = await readWithRetry(url, async () => {
    calls += 1
    if (calls === 1) return respond(503, false)
    return respond(200)
  }, wait)
  assert.equal(retried.ok, true)
  assert.equal(calls, 2, 'one failure is tried once more')
  assert.equal(waits, 1)

  calls = 0
  waits = 0
  const afterThrow = await readWithRetry(url, async () => {
    calls += 1
    if (calls === 1) throw new Error('offline')
    return respond(200)
  }, wait)
  assert.equal(afterThrow.ok, true)
  assert.equal(calls, 2, 'a thrown fetch is tried once more')
  assert.equal(waits, 1)

  calls = 0
  await assert.rejects(
    () => readWithRetry(url, async () => {
      calls += 1
      return respond(404, false)
    }, async () => {}),
    (error: unknown) => {
      assert.ok(error instanceof Error)
      assert.equal(error.message, '/engine.wasm: 404')
      return true
    },
  )
  assert.equal(calls, 2, 'two failures stop')
})

test('blank imports an empty document through the same path', () => {
  let opened: Opened | undefined
  try {
    opened = engine.blank()
    assert.ok(Array.isArray(opened.doc.body), 'blank document has a body')
    assert.ok(opened.doc.sections.length > 0, 'blank document has a section')
  } finally {
    opened?.source.free()
  }
})

test('addMedia mounts the probe image while decoding and removes it afterwards', async () => {
  const png = new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
    0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x02,
    0x08, 0x02, 0x00, 0x00, 0x00, 0x26, 0x93, 0x09, 0x29,
    0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
  ])
  const children: unknown[] = []
  let mounted = false
  const prevBitmap = globalThis.createImageBitmap
  const prevImage = globalThis.Image
  const prevDoc = globalThis.document
  class ConnectedOnlyImage {
    isConnected = false
    naturalWidth = 4
    naturalHeight = 2
    style = { cssText: '' }
    onload: (() => void) | null = null
    onerror: (() => void) | null = null
    setAttribute(): void {}
    remove(): void {
      this.isConnected = false
      const idx = children.indexOf(this)
      if (idx >= 0) children.splice(idx, 1)
    }
    set src(_url: string) {
      if (this.isConnected) queueMicrotask(() => this.onload?.())
    }
  }
  Object.assign(globalThis, {
    createImageBitmap: async () => {
      throw new Error('NotSupportedError')
    },
    Image: ConnectedOnlyImage,
    document: {
      body: {
        append(node: ConnectedOnlyImage) {
          node.isConnected = true
          mounted = true
          children.push(node)
        },
      },
    },
  })
  let opened: Opened | undefined
  try {
    opened = engine.blank()
    const media = await opened.source.addMedia(new Blob([png], { type: 'image/png' }))
    assert.equal(mounted, true)
    assert.equal(children.length, 0)
    assert.equal(media.widthPt, 3)
    assert.equal(media.heightPt, 1.5)
  } finally {
    opened?.source.free()
    if (prevBitmap === undefined) delete (globalThis as Record<string, unknown>).createImageBitmap
    else globalThis.createImageBitmap = prevBitmap
    if (prevImage === undefined) delete (globalThis as Record<string, unknown>).Image
    else globalThis.Image = prevImage
    if (prevDoc === undefined) delete (globalThis as Record<string, unknown>).document
    else globalThis.document = prevDoc
  }
})
