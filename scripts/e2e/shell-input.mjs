import { readFileSync } from 'node:fs'
import { act, assert, buttonNamed, check, consoleErrors, ensureDev, evaluate, fileMenuItem, freshPage, namedRef, summary, until } from './acde.mjs'

// acde eval puts the expression on the command line. A long fixture is staged on the page instead.
const ARG_LIMIT = 100 * 1024
const CHUNK = 48 * 1024

const basic = readFileSync(new URL('../../fixtures/basic.docx', import.meta.url)).toString('base64')
const long = readFileSync(new URL('../../fixtures/long.docx', import.meta.url)).toString('base64')

await ensureDev()
await freshPage(`!!(window.__word && window.__word.editor)`, { label: 'editor' })

function stage(base64) {
  if (base64.length <= ARG_LIMIT) return JSON.stringify(base64)
  evaluate(`(() => { window.__openB64 = ''; return 1 })()`)
  for (let i = 0; i < base64.length; i += CHUNK) {
    const piece = JSON.stringify(base64.slice(i, i + CHUNK))
    evaluate(`(() => { window.__openB64 += ${piece}; return 1 })()`)
  }
  return 'window.__openB64'
}

function openDocument(base64) {
  const expr = stage(base64)
  evaluate(`(() => {
    window.__wordReady = false
    window.__wordError = ''
    try {
      Promise.resolve(window.__word.openBytes(${expr})).then(
        () => { window.__wordReady = true },
        (error) => {
          window.__wordError = String(error && error.message || error)
          window.__wordReady = true
        },
      )
    } catch (error) {
      window.__wordError = String(error && error.message || error)
      window.__wordReady = true
    }
    return 1
  })()`)
  return until(`window.__wordReady === true`, { timeout: 60000, label: 'layout complete' })
}

function readModel() {
  return evaluate(`(() => {
    const editor = window.__word.editor
    const sel = editor.state.sel
    const paras = []
    for (const block of editor.state.doc.body) {
      if (block.kind !== 'p') continue
      let text = ''
      const runs = []
      for (const inline of block.inlines) {
        if (inline.kind !== 'text') {
          text += '\\uFFFC'
          continue
        }
        text += inline.text
        runs.push({ text: inline.text, bold: inline.marks.bold === true })
      }
      paras.push({ id: block.id, text, runs })
    }
    const caret = sel.kind === 'text' ? { p: sel.head.p, o: sel.head.o } : null
    return {
      error: window.__wordError || '',
      paras,
      caret,
      pagesFinal: editor.status().pagesFinal,
    }
  })()`)
}

function paraById(model, id) {
  return model.paras.find((para) => para.id === id)
}

function dialogInfo() {
  return evaluate(`(() => {
    const dialog = document.querySelector('[role="dialog"]')
    if (!dialog) return null
    const id = dialog.getAttribute('aria-labelledby')
    const label = id ? document.getElementById(id) : null
    const backdrop = dialog.closest('.unsaved')
    return {
      name: label ? label.textContent : '',
      hidden: !backdrop || backdrop.hidden === true,
    }
  })()`)
}

// New opened a tab beside the dirty document. Close the dirty one without saving so one editor is left.
async function discardUnsaved() {
  const lang = evaluate('document.documentElement.lang')
  const dirty = evaluate(`(() => {
    const tab = document.querySelector('#tabs .tab.dirty:not(.active)')
    return tab ? { label: tab.querySelector('.tab-label').textContent, name: tab.title } : null
  })()`)
  assert(dirty, 'the dirty document kept its tab')
  const close = lang === 'zh' ? `关闭 ${dirty.label}` : `Close ${dirty.label}`
  const ask = lang === 'zh' ? `是否保存对 ${dirty.name} 的更改？` : `Save changes to ${dirty.name}?`
  const discard = lang === 'zh' ? '不保存' : "Don't save"
  act(buttonNamed(close), 'click')
  const info = dialogInfo()
  assert(info && info.hidden === false, 'closing a dirty tab asks')
  assert(info.name === ask, `dialog name ${info.name}`)
  act(buttonNamed(discard), 'click')
  await until(`document.querySelectorAll('#tabs .tab').length === 1`, { label: 'one tab left' })
}

function inputRef() {
  const errors = []
  for (const name of ['Document text', '文档正文']) {
    try {
      return namedRef(name)
    } catch (error) {
      errors.push(String(error.message || error))
    }
  }
  throw new Error(errors.join('\n'))
}

function caretOf(model) {
  assert(model.caret, `no caret: ${JSON.stringify(model)}`)
  return model.caret
}

await openDocument(basic)
evaluate(`(() => { const view = document.querySelector('.page-view'); if (view) view.scrollTop = 0; return 1 })()`)

const platform = evaluate(`window.__word.platform`)
assert(platform === 'mac' || platform === 'other', `platform ${platform}`)
const mod = platform === 'mac' ? 'meta' : 'ctrl'

let sink = inputRef()
act(sink, 'focus')
act(namedRef('Quarterly report'), 'click')
// placeSink moves the textarea onto the caret's page. Take the ref again in case the move minted a new one,
// and focus it so the following act type/key hits the sink if the click moved focus.
sink = inputRef()
act(sink, 'focus')

let before = readModel()
assert(!before.error, before.error)
const origin = caretOf(before)
const originPara = paraById(before, origin.p)
assert(originPara, `missing paragraph ${origin.p}`)
assert(originPara.text.includes('Quarterly report'), `clicked ${JSON.stringify(originPara.text)}`)

await check('typing English inserts at the caret', async () => {
  act(sink, 'type', 'Hello')
  const model = readModel()
  const para = paraById(model, origin.p)
  assert(para, `missing paragraph ${origin.p}`)
  const expected = originPara.text.slice(0, origin.o) + 'Hello' + originPara.text.slice(origin.o)
  assert(para.text === expected, `text ${JSON.stringify(para.text)} expected ${JSON.stringify(expected)}`)
  before = model
})

await check('typing Chinese inserts at the caret', async () => {
  const caret = caretOf(before)
  const paraBefore = paraById(before, caret.p)
  assert(paraBefore, `missing paragraph ${caret.p}`)
  act(sink, 'type', '中文输入')
  const model = readModel()
  const para = paraById(model, caret.p)
  assert(para, `missing paragraph ${caret.p}`)
  const expected = paraBefore.text.slice(0, caret.o) + '中文输入' + paraBefore.text.slice(caret.o)
  assert(para.text === expected, `text ${JSON.stringify(para.text)} expected ${JSON.stringify(expected)}`)
  before = model
})

let splitFrom
await check('enter splits the paragraph at the caret', async () => {
  splitFrom = before
  const caret = caretOf(splitFrom)
  const paraBefore = paraById(splitFrom, caret.p)
  assert(paraBefore, `missing paragraph ${caret.p}`)
  const count = splitFrom.paras.length
  act(sink, 'key', 'enter')
  const model = readModel()
  assert(model.paras.length === count + 1, `paragraphs ${count} -> ${model.paras.length}`)
  const head = paraById(model, caret.p)
  const headIndex = model.paras.findIndex((para) => para.id === caret.p)
  const next = model.paras[headIndex + 1]
  assert(head, `missing head ${caret.p}`)
  assert(next, 'enter did not add a following paragraph')
  assert(head.text === paraBefore.text.slice(0, caret.o), `head ${JSON.stringify(head.text)}`)
  assert(next.text === paraBefore.text.slice(caret.o), `tail ${JSON.stringify(next.text)}`)
  const after = caretOf(model)
  assert(after.p === next.id && after.o === 0, `caret ${after.p}@${after.o}`)
  before = model
})

await check('backspace at the new paragraph joins it with the previous one', async () => {
  const caret = caretOf(splitFrom)
  const paraBefore = paraById(splitFrom, caret.p)
  assert(paraBefore, `missing paragraph ${caret.p}`)
  act(sink, 'key', 'backspace')
  const model = readModel()
  assert(model.paras.length === splitFrom.paras.length, `paragraphs ${model.paras.length} vs ${splitFrom.paras.length}`)
  const para = paraById(model, caret.p)
  assert(para, `missing joined paragraph ${caret.p}`)
  assert(para.text === paraBefore.text, `joined ${JSON.stringify(para.text)}`)
  const after = caretOf(model)
  assert(after.p === caret.p && after.o === caret.o, `caret ${after.p}@${after.o} expected ${caret.p}@${caret.o}`)
  before = model
})

await check('left then right leaves the caret, and the next character lands there', async () => {
  const caret = caretOf(before)
  const paraBefore = paraById(before, caret.p)
  assert(paraBefore, `missing paragraph ${caret.p}`)
  act(sink, 'key', 'left')
  act(sink, 'key', 'right')
  const moved = readModel()
  const back = caretOf(moved)
  assert(back.p === caret.p && back.o === caret.o, `after arrows ${back.p}@${back.o} expected ${caret.p}@${caret.o}`)
  act(sink, 'type', 'Q')
  const model = readModel()
  const para = paraById(model, caret.p)
  assert(para, `missing paragraph ${caret.p}`)
  const expected = paraBefore.text.slice(0, caret.o) + 'Q' + paraBefore.text.slice(caret.o)
  assert(para.text === expected, `text ${JSON.stringify(para.text)}`)
  const after = caretOf(model)
  assert(after.o === caret.o + 1, `offset ${after.o} expected ${caret.o + 1}`)
  before = model
})

await check('arrow down twice then up twice returns to the same position', async () => {
  const caret = caretOf(before)
  act(sink, 'key', 'down')
  act(sink, 'key', 'down')
  act(sink, 'key', 'up')
  act(sink, 'key', 'up')
  const model = readModel()
  const after = caretOf(model)
  assert(after.p === caret.p && after.o === caret.o, `position ${after.p}@${after.o} expected ${caret.p}@${caret.o}`)
  before = model
})

await check('the platform bold chord marks only the text typed after it', async () => {
  const caret = caretOf(before)
  const paraBefore = paraById(before, caret.p)
  assert(paraBefore, `missing paragraph ${caret.p}`)
  act(sink, 'key', `${mod}+b`)
  act(sink, 'type', 'Bold')
  const model = readModel()
  const para = paraById(model, caret.p)
  assert(para, `missing paragraph ${caret.p}`)
  const expected = paraBefore.text.slice(0, caret.o) + 'Bold' + paraBefore.text.slice(caret.o)
  assert(para.text === expected, `text ${JSON.stringify(para.text)}`)
  let cursor = 0
  let boldRun = null
  let previous = null
  for (const run of para.runs) {
    if (cursor === caret.o) {
      boldRun = run
      break
    }
    previous = run
    cursor += run.text.length
  }
  assert(boldRun && boldRun.bold === true && boldRun.text.startsWith('Bold'), `run ${JSON.stringify(boldRun)}`)
  if (previous) assert(previous.bold !== true, `text before is bold: ${JSON.stringify(previous)}`)
  before = model
})

await check('undo removes the bold text and redo puts it back', async () => {
  act(sink, 'key', `${mod}+z`)
  const undone = readModel()
  assert(undone.paras.every((para) => !para.text.includes('Bold')), `still present: ${JSON.stringify(undone.paras.map((para) => para.text))}`)
  const redo = platform === 'mac' ? 'meta+shift+z' : 'ctrl+y'
  act(sink, 'key', redo)
  const redone = readModel()
  const back = redone.paras.some((para) => para.text.includes('Bold'))
  assert(back, `redo lost Bold: ${JSON.stringify(redone.paras.map((para) => para.text))}`)
})

await check('console has no errors', async () => {
  const errors = consoleErrors()
  assert(errors.length === 0, errors.join('\n'))
})

await check('a keystroke on long.docx stays under 8ms at p95', async () => {
  await openDocument(long)
  const placed = evaluate(`(() => {
    const editor = window.__word.editor
    const paras = editor.state.doc.body.filter((block) => block.kind === 'p')
    const para = paras[Math.floor(paras.length / 2)]
    if (!para) return { error: 'no paragraph' }
    let length = 0
    for (const inline of para.inlines) length += inline.kind === 'text' ? inline.text.length : 1
    const pos = { p: para.id, o: Math.min(40, length) }
    editor.run(window.__word.setSelection({ kind: 'text', anchor: pos, head: pos }))
    window.__benchPara = para.id
    return { id: para.id }
  })()`)
  assert(!placed.error, String(placed.error))
  // Typing happens in a paragraph the user can see. The jump above scrolls there, and the pages mount on the next frame.
  await until(`!!document.querySelector('.page[data-mounted="1"] .blk[data-id="' + window.__benchPara + '"]')`, {
    timeout: 8000,
    label: 'caret paragraph mounted',
  })
  const bench = evaluate(`(() => {
    const editor = window.__word.editor
    const para = { id: window.__benchPara }
    const times = []
    for (let i = 0; i < 100; i++) {
      const started = performance.now()
      editor.run(window.__word.insertText('x'))
      times.push(performance.now() - started)
    }
    times.sort((a, b) => a - b)
    const index = Math.ceil(0.95 * times.length) - 1
    return {
      p95: times[index] ?? 0,
      max: times[times.length - 1] ?? 0,
      samples: times.length,
      id: para.id,
      pages: editor.status().pages,
    }
  })()`)
  console.log(`  ${JSON.stringify(bench)}`)
  assert(!bench.error, String(bench.error))
  assert(bench.samples === 100, `samples ${bench.samples}`)
  assert(bench.p95 < 8, `p95 ${bench.p95} ms (max ${bench.max})`)
})

await check('enter and undo on long.docx stay under 16ms at p95', async () => {
  const placed = evaluate(`(() => {
    const editor = window.__word.editor
    const paras = editor.state.doc.body.filter((block) => block.kind === 'p')
    const para = paras[6]
    if (!para) return { error: 'no seventh paragraph' }
    let length = 0
    for (const inline of para.inlines) length += inline.kind === 'text' ? inline.text.length : 1
    if (length < 3) return { error: 'paragraph is shorter than 3', length }
    const pos = { p: para.id, o: 3 }
    editor.run(window.__word.setSelection({ kind: 'text', anchor: pos, head: pos }))
    window.__benchPara = para.id
    return { id: para.id, pages: editor.status().pages }
  })()`)
  assert(!placed.error, String(placed.error || JSON.stringify(placed)))
  await until(`!!document.querySelector('.page[data-mounted="1"] .blk[data-id="' + window.__benchPara + '"]')`, {
    timeout: 8000,
    label: 'seventh paragraph mounted',
  })
  evaluate(`(() => {
    window.__enterBench = null
    const editor = window.__word.editor
    const pagesBefore = editor.status().pages
    const enter = []
    const undo = []
    const frames = () => new Promise((resolve) => {
      requestAnimationFrame(() => requestAnimationFrame(resolve))
    })
    ;(async () => {
      try {
        for (let i = 0; i < 20; i++) {
          const t0 = performance.now()
          const entered = await window.word.call({ op: 'enter' })
          enter.push(performance.now() - t0)
          if (!entered.ok) {
            window.__enterBench = { error: JSON.stringify({ entered }) }
            return
          }
          await frames()
          const t1 = performance.now()
          const undone = await window.word.call({ op: 'undo' })
          undo.push(performance.now() - t1)
          if (!undone.ok) {
            window.__enterBench = { error: JSON.stringify({ undone }) }
            return
          }
          await frames()
        }
        const rank = (times) => {
          const sorted = times.slice().sort((a, b) => a - b)
          const index = Math.ceil(0.95 * sorted.length) - 1
          return { p95: sorted[index] ?? 0, max: sorted[sorted.length - 1] ?? 0 }
        }
        const status = editor.status()
        window.__enterBench = {
          enter: rank(enter),
          undo: rank(undo),
          pagesBefore,
          pagesAfter: status.pages,
          pagesFinal: status.pagesFinal,
          samples: enter.length,
        }
      } catch (error) {
        window.__enterBench = { error: String(error && error.message || error) }
      }
    })()
    return 1
  })()`)
  const bench = await until(`window.__enterBench`, { timeout: 20000, label: 'enter undo bench' })
  console.log(`  ${JSON.stringify(bench)}`)
  assert(!bench.error, String(bench.error))
  assert(bench.samples === 20, `samples ${bench.samples}`)
  assert(bench.enter.p95 < 16, `enter p95 ${bench.enter.p95} ms (max ${bench.enter.max})`)
  assert(bench.undo.p95 < 16, `undo p95 ${bench.undo.p95} ms (max ${bench.undo.max})`)
  assert(bench.pagesAfter === bench.pagesBefore, `pages ${bench.pagesBefore} -> ${bench.pagesAfter}`)
  assert(bench.pagesFinal === true, 'pages are not final')
})

await check('console has no errors after the long document', async () => {
  const errors = consoleErrors()
  assert(errors.length === 0, errors.join('\n'))
})

await check('a font change on long.docx keeps the selection where it was', async () => {
  await openDocument(long)
  const opened = evaluate(`window.__wordError || ''`)
  assert(!opened, opened)
  const placed = evaluate(`(() => {
    const editor = window.__word.editor
    const paras = editor.state.doc.body.filter((block) => block.kind === 'p')
    const para = paras[Math.floor(paras.length / 2)]
    if (!para) return { error: 'no paragraph' }
    let length = 0
    for (const inline of para.inlines) length += inline.kind === 'text' ? inline.text.length : 1
    if (length < 30) return { error: 'paragraph is shorter than 30', length }
    editor.run(window.__word.setSelection({ kind: 'text', anchor: { p: para.id, o: 10 }, head: { p: para.id, o: 30 } }))
    window.__fontPara = para.id
    return { id: para.id, length }
  })()`)
  assert(!placed.error, String(placed.error || JSON.stringify(placed)))
  await until(`!!document.querySelector('.page[data-mounted="1"] .blk[data-id="' + window.__fontPara + '"]')`, {
    timeout: 8000,
    label: 'selected paragraph mounted',
  })
  const before = evaluate(`(() => {
    const pages = document.querySelector('#pages')
    const host = pages && pages.getBoundingClientRect()
    const blocks = [...document.querySelectorAll('.page[data-mounted="1"] .blk[data-id="' + window.__fontPara + '"]')]
    if (!host) return { error: 'missing pages' }
    let block = null
    for (const el of blocks) {
      const clip = el.closest('.clip')
      const box = clip ? clip.getBoundingClientRect() : el.getBoundingClientRect()
      if (box.bottom > host.top + 0.5 && box.top < host.bottom - 0.5) {
        block = el
        break
      }
    }
    if (!block) return { error: 'block is not in view', count: blocks.length }
    return { top: block.getBoundingClientRect().top - host.top }
  })()`)
  assert(typeof before.top === 'number', String(before.error || JSON.stringify(before)))
  await until(`(() => {
    const want = ['Georgia', 'Verdana', 'Helvetica', 'Menlo']
    const select = document.querySelector('select.font')
    if (!select) return false
    return [...select.options].some((option) => want.includes(option.value))
  })()`, { timeout: 20000, label: 'probe font' })
  evaluate(`(() => { window.__fontArmed = 0; requestAnimationFrame(() => requestAnimationFrame(() => { window.__fontArmed = 1 })); return 1 })()`)
  await until(`window.__fontArmed === 1`, { timeout: 2000, label: 'font select ready' })
  const picked = evaluate(`(() => {
    const want = ['Georgia', 'Verdana', 'Helvetica', 'Menlo']
    const select = document.querySelector('select.font')
    if (!select) return { error: 'no font select' }
    const family = want.find((name) => [...select.options].some((option) => option.value === name))
    if (!family) return { error: 'none of the probe fonts are installed' }
    select.value = family
    select.dispatchEvent(new Event('change', { bubbles: true }))
    return { family }
  })()`)
  assert(!picked.error, String(picked.error))
  await until(`(() => {
    const editor = window.__word.editor
    return editor.status().pagesFinal === true && editor.toolbar().font === ${JSON.stringify(picked.family)}
  })()`, { timeout: 20000, label: 'font settled' })
  evaluate(`(() => { window.__fontFrames = 0; requestAnimationFrame(() => requestAnimationFrame(() => { window.__fontFrames = 1 })); return 1 })()`)
  await until(`window.__fontFrames === 1`, { timeout: 2000, label: 'two frames' })
  const after = evaluate(`(() => {
    const pages = document.querySelector('#pages')
    const host = pages && pages.getBoundingClientRect()
    const blocks = [...document.querySelectorAll('.page[data-mounted="1"] .blk[data-id="' + window.__fontPara + '"]')]
    if (!host) return { error: 'missing pages' }
    let block = null
    for (const el of blocks) {
      if (!el.isConnected) continue
      const clip = el.closest('.clip')
      const box = clip ? clip.getBoundingClientRect() : el.getBoundingClientRect()
      if (box.bottom > host.top + 0.5 && box.top < host.bottom - 0.5) {
        block = el
        break
      }
    }
    if (!block) return { error: 'block is not mounted', count: blocks.length }
    return { top: block.getBoundingClientRect().top - host.top, mounted: true }
  })()`)
  console.log(`  ${picked.family} top ${before.top} -> ${after.top}`)
  assert(after.mounted === true, String(after.error || JSON.stringify(after)))
  assert(Math.abs(after.top - before.top) <= 2, `block top ${before.top} -> ${after.top}`)
  const errors = consoleErrors()
  assert(errors.length === 0, errors.join('\n'))
})

await check('a point on the toolbar hits the button when the document is scrolled', async () => {
  const hit = evaluate(`(() => {
    const view = document.querySelector('.page-view')
    const toolbar = document.querySelector('.toolbar')
    const button = toolbar && toolbar.querySelector('button')
    if (!view || !toolbar || !button) return { error: 'missing toolbar or page view' }
    view.scrollTop = 20000
    const box = button.getBoundingClientRect()
    const x = (box.left + box.right) / 2
    const y = (box.top + box.bottom) / 2
    const target = document.elementFromPoint(x, y)
    return {
      scrollTop: view.scrollTop,
      inside: !!(target && toolbar.contains(target)),
      name: target && target.tagName,
    }
  })()`)
  assert(!hit.error, String(hit.error))
  assert(hit.scrollTop >= 20000, `scrollTop ${hit.scrollTop}`)
  assert(hit.inside, `hit ${hit.name} at scrollTop ${hit.scrollTop}`)
})

await check('a point on a tall paragraph hits the line under the pointer', async () => {
  const errors = []
  let opened = false
  for (const name of ['New', '新建']) {
    try {
      act(fileMenuItem(name), 'click')
      opened = true
      break
    } catch (error) {
      errors.push(String(error.message || error))
    }
  }
  assert(opened, `new document: ${errors.join('\n')}`)
  await discardUnsaved()
  sink = inputRef()
  act(sink, 'focus')
  const text = '改了一个技能，用户的使用方式就变了；'.repeat(120)
  assert(text.length === 2160, `fixture length ${text.length}`)
  act(sink, 'type', text)
  const typed = evaluate(`(() => {
    const body = window.__word.editor.state.doc.body
    const block = body.find((item) => item.kind === 'p')
    let chars = 0
    if (block) {
      for (const inline of block.inlines) chars += inline.kind === 'text' ? inline.text.length : 1
    }
    return { paras: body.length, chars }
  })()`)
  assert(typed.paras === 1 && typed.chars === 2160, `typed ${JSON.stringify(typed)}`)
  // Font metrics repaginate after the first paint. Wait until that settles so the two pages stay put.
  await until(
    `window.__word.editor.view.layout.pages.length >= 2 && window.__word.editor.status().pagesFinal`,
    { timeout: 20000, label: 'two pages' },
  )
  evaluate(`(() => {
    const view = document.querySelector('.page-view')
    if (!view) throw new Error('missing page view')
    view.scrollTop = 300
    return view.scrollTop
  })()`)
  await until(`document.querySelectorAll('.page[data-mounted="1"]').length >= 2`, {
    timeout: 8000,
    label: 'both pages mounted',
  })
  const hits = evaluate(`(() => {
    const editor = window.__word.editor
    const page = document.querySelector('.page')
    const body = document.querySelector('.page-body')
    if (!page || !body) return { error: 'missing page' }
    const box = body.getBoundingClientRect()
    const x = (box.left + box.right) / 2
    const stack = document.querySelector('.page-stack')
    const transform = stack ? getComputedStyle(stack).transform : 'none'
    const zoom = transform && transform !== 'none' ? new DOMMatrix(transform).a || 1 : 1
    const rows = []
    for (const y of [250, 400, 550]) {
      const pos = editor.positionAtPoint(x, y)
      if (!pos) {
        rows.push({ y, pos: null })
        continue
      }
      const caret = editor.view.caretRect(pos)
      // caretRect.top and height are page-local px at zoom 1. The page border box is already in client px.
      const pageTop = page.getBoundingClientRect().top
      const top = pageTop + caret.top * zoom
      const height = caret.height * zoom
      rows.push({ y, pos, page: caret.page, top, height })
    }
    return {
      x,
      zoom,
      pages: editor.view.layout.pages.length,
      bodyTop: box.top,
      bodyBottom: box.bottom,
      rows,
    }
  })()`)
  assert(!hits.error, String(hits.error))
  const failures = []
  for (const row of hits.rows) {
    if (!row.pos) {
      failures.push(`y=${row.y} missed`)
      continue
    }
    const low = row.top - 2
    const high = row.top + row.height + 2
    if (row.page !== 0 || row.y < low || row.y > high) {
      failures.push(
        `y=${row.y} page=${row.page} offset=${row.pos.o} top=${row.top} height=${row.height} band=[${low}, ${high}]`,
      )
    }
  }
  assert(
    failures.length === 0,
    `x=${hits.x} zoom=${hits.zoom} pages=${hits.pages} body=${hits.bodyTop}..${hits.bodyBottom} ${failures.join('; ')}`,
  )
})

await summary()
