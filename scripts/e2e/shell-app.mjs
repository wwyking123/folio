import { execFileSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { act, assert, buttonNamed, check, closeFileMenu, consoleErrors, ensureDev, evaluate, evidencePath, fileMenuItem, namedRef, ref, root, shot, summary, until } from './acde.mjs'

// Strings mirror src/locales. The page's lang, set from ?lang= or the system language, picks the table.
const STRINGS = {
  en: {
    new: 'New',
    open: 'Open',
    undo: 'Undo',
    redo: 'Redo',
    bold: 'Bold',
    italic: 'Italic',
    underline: 'Underline',
    strike: 'Strikethrough',
    font: 'Font',
    size: 'Font size',
    mixed: 'Mixed',
    input: 'Document text',
    style: 'Paragraph style',
    heading1: 'Heading 1',
    alignLeft: 'Align left',
    alignCenter: 'Align center',
    alignRight: 'Align right',
    alignJustify: 'Justify',
    bullet: 'Bulleted list',
    number: 'Numbered list',
    zoom: 'Zoom',
    zoomOut: 'Zoom out',
    zoomIn: 'Zoom in',
    untitled: 'Untitled.docx',
    discard: "Don't save",
    ask: 'Save changes to {name}?',
    closeTab: 'Close {name}',
    statusPage: 'Page {page} of {pages}',
    statusWords: '{words} words',
  },
  zh: {
    new: '新建',
    open: '打开',
    undo: '撤销',
    redo: '重做',
    bold: '加粗',
    italic: '倾斜',
    underline: '下划线',
    strike: '删除线',
    font: '字体',
    size: '字号',
    mixed: '混合',
    input: '文档正文',
    style: '段落样式',
    heading1: '标题 1',
    alignLeft: '左对齐',
    alignCenter: '居中',
    alignRight: '右对齐',
    alignJustify: '两端对齐',
    bullet: '项目符号',
    number: '编号',
    zoom: '缩放',
    zoomOut: '缩小',
    zoomIn: '放大',
    untitled: '未命名.docx',
    discard: '不保存',
    ask: '是否保存对 {name} 的更改？',
    closeTab: '关闭 {name}',
    statusPage: '第 {page} 页，共 {pages} 页',
    statusWords: '{words} 个字',
  },
}

const ALIGN = [
  ['left', 'alignLeft'],
  ['center', 'alignCenter'],
  ['right', 'alignRight'],
  ['justify', 'alignJustify'],
]

const MARKS = ['bold', 'italic', 'underline', 'strike']

const ARG_LIMIT = 100 * 1024
const CHUNK = 48 * 1024

const basic = readFileSync(new URL('../../fixtures/basic.docx', import.meta.url)).toString('base64')

function go(search) {
  try {
    evaluate(`location.assign(location.pathname + ${JSON.stringify(search)})`)
  } catch {
    // Reloading the page aborts the eval.
  }
}

// The token has to be visible to until(), so go() cannot hide it. Keep the last search's boot value here.
let bootToken = ''
function goTracked(search) {
  const match = search.match(/(?:^|[?&])boot=([^&]+)/)
  bootToken = match ? decodeURIComponent(match[1]) : ''
  go(search)
}

function names() {
  const lang = evaluate('document.documentElement.lang')
  const table = STRINGS[lang]
  if (!table) throw new Error(`no strings for lang ${JSON.stringify(lang)}`)
  return table
}

function stage(base64) {
  if (base64.length <= ARG_LIMIT) return JSON.stringify(base64)
  evaluate(`(() => { window.__openB64 = ''; return 1 })()`)
  for (let i = 0; i < base64.length; i += CHUNK) {
    const piece = JSON.stringify(base64.slice(i, i + CHUNK))
    evaluate(`(() => { window.__openB64 += ${piece}; return 1 })()`)
  }
  return 'window.__openB64'
}

function openDocument(base64, name) {
  const expr = stage(base64)
  const nameArg = name ? `, ${JSON.stringify(name)}` : ''
  evaluate(`(() => {
    window.__wordReady = false
    window.__wordError = ''
    try {
      Promise.resolve(window.__word.openBytes(${expr}${nameArg})).then(
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

function labeled(name) {
  return evaluate(`(() => {
    const want = ${JSON.stringify(name)}
    const el = [...document.querySelectorAll('button, select')].find((node) => node.getAttribute('aria-label') === want)
    if (!el) return null
    const option = el.selectedOptions && el.selectedOptions.length ? el.selectedOptions[0] : null
    return {
      tag: el.tagName,
      pressed: el.getAttribute('aria-pressed'),
      disabled: el.disabled === true,
      value: 'value' in el ? el.value : null,
      title: el.getAttribute('title'),
      selected: option ? option.textContent : null,
      hidden: !!(option && option.hidden),
    }
  })()`)
}

function selectionFact() {
  return evaluate(`(() => {
    const editor = window.__word.editor
    const sel = editor.state.sel
    if (!sel || sel.kind !== 'text') return { kind: sel ? sel.kind : 'none' }
    const doc = editor.state.doc
    const para = doc.body.find((block) => block.kind === 'p' && block.id === sel.head.p)
    if (!para) return { kind: 'text', missing: true }
    const from = Math.min(sel.anchor.o, sel.head.o)
    const to = Math.max(sel.anchor.o, sel.head.o)
    const same = sel.anchor.p === sel.head.p
    let offset = 0
    const runs = []
    if (same) {
      for (const inline of para.inlines) {
        const chunk = inline.kind === 'text' ? inline.text : '\\uFFFC'
        const start = offset
        const end = offset + chunk.length
        offset = end
        if (inline.kind !== 'text' || end <= from || start >= to) continue
        const marks = inline.marks || {}
        runs.push({
          bold: marks.bold === true,
          italic: marks.italic === true,
          underline: marks.underline === true,
          strike: marks.strike === true,
        })
      }
    }
    const numbering = para.props.numbering || null
    let format = null
    if (numbering) {
      const def = doc.lists.byNum.get(numbering.num)
      const level = def && def.levels[numbering.level]
      format = level ? level.format : null
    }
    return {
      kind: 'text',
      id: para.id,
      same,
      from,
      to,
      runs,
      style: para.props.style || null,
      align: para.props.align || null,
      format,
    }
  })()`)
}

function fill(template, values) {
  return template.replace(/\{(\w+)\}/g, (all, name) => (values[name] === undefined ? all : String(values[name])))
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

function bodyText() {
  return evaluate(`(() => {
    const parts = []
    for (const block of window.__word.editor.state.doc.body) {
      if (block.kind !== 'p') continue
      for (const inline of block.inlines) if (inline.kind === 'text') parts.push(inline.text)
    }
    return parts.join('')
  })()`)
}

function chooseOpen(path) {
  try {
    execFileSync('acde', ['choose', 'open', path], { cwd: root, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] })
  } catch (error) {
    const stderr = error && error.stderr ? String(error.stderr) : ''
    throw new Error(`acde choose open failed: ${stderr || error.message}`)
  }
}

function requireControl(name, tag) {
  const ref = namedRef(name)
  const info = labeled(name)
  assert(info, `no control labeled ${JSON.stringify(name)}`)
  assert(info.tag === tag, `${name} is ${info.tag}`)
  return { ref, info }
}

await ensureDev()
// A reload is async: the old page still has an editor until the new document commits. The boot token is only on the new URL.
goTracked(`?boot=shell-${Date.now()}`)
await until(
  `new URLSearchParams(location.search).get('boot') === ${JSON.stringify(bootToken)} && !!(window.__word && window.__word.editor)`,
  { timeout: 60000, label: 'shell' },
)

const ui = names()

await check('basic.docx opens with a clean title', async () => {
  await openDocument(basic, 'basic.docx')
  const opened = evaluate(`({ error: window.__wordError || '', title: document.title })`)
  assert(!opened.error, opened.error)
  assert(!opened.title.includes('•'), `title ${opened.title}`)
  assert(opened.title.includes('basic.docx'), `title ${opened.title}`)
  evaluate(`(() => { const view = document.querySelector('.page-view'); if (view) view.scrollTop = 0; return 1 })()`)
})

await check('the toolbar is one row at the default width', async () => {
  const tops = evaluate(`(() => {
    const groups = [...document.querySelectorAll('#toolbar > .group')]
    return groups.map((el) => el.getBoundingClientRect().top)
  })()`)
  assert(Array.isArray(tops) && tops.length >= 8, `groups ${JSON.stringify(tops)}`)
  const first = tops[0]
  for (const top of tops) assert(Math.abs(top - first) <= 1, `group tops ${JSON.stringify(tops)}`)
})

await check('every toolbar control is found by its accessible name', async () => {
  fileMenuItem(ui.new)
  requireControl(ui.new, 'BUTTON')
  requireControl(ui.open, 'BUTTON')
  assert(labeled(ui.new).title === null, 'the New row shows its chord as a hint, not a tooltip')
  closeFileMenu()
  const buttons = [
    'undo', 'redo', 'bold', 'italic', 'underline', 'strike',
    'alignLeft', 'alignCenter', 'alignRight', 'alignJustify', 'bullet', 'number',
  ]
  for (const key of buttons) requireControl(ui[key], 'BUTTON')
  requireControl(ui.font, 'SELECT')
  requireControl(ui.size, 'SELECT')
  requireControl(ui.style, 'SELECT')
  requireControl(ui.zoom, 'SELECT')
  const bold = labeled(ui.bold)
  const italic = labeled(ui.italic)
  const underline = labeled(ui.underline)
  const undo = labeled(ui.undo)
  const redo = labeled(ui.redo)
  assert(bold.title && (bold.title.includes('Ctrl+B') || bold.title.includes('⌘B')), `bold title ${bold.title}`)
  assert(italic.title && (italic.title.includes('Ctrl+I') || italic.title.includes('⌘I')), `italic title ${italic.title}`)
  assert(underline.title && (underline.title.includes('Ctrl+U') || underline.title.includes('⌘U')), `underline title ${underline.title}`)
  assert(undo.title && (undo.title.includes('Ctrl+Z') || undo.title.includes('⌘Z')), `undo title ${undo.title}`)
  assert(redo.title && (redo.title.includes('Ctrl+Y') || redo.title.includes('⌘Y')), `redo title ${redo.title}`)
  assert(labeled(ui.strike).title === ui.strike, 'strike has no keymap chord')
})

await check('bold, italic, underline and strike toggle on a selected word', async () => {
  act(namedRef('Quarterly report'), 'double')
  const selected = selectionFact()
  console.log(`  selection ${JSON.stringify(selected)}`)
  assert(selected.kind === 'text' && selected.same && selected.from < selected.to, `selection ${JSON.stringify(selected)}`)
  for (const mark of MARKS) {
    act(namedRef(ui[mark]), 'click')
    let fact = selectionFact()
    let button = labeled(ui[mark])
    assert(button.pressed === 'true', `${mark} pressed ${button.pressed}`)
    assert(fact.id === selected.id && fact.from < fact.to, `${mark} lost the selection ${JSON.stringify(fact)}`)
    assert(fact.runs.length > 0 && fact.runs.every((run) => run[mark] === true), `${mark} on ${JSON.stringify(fact.runs)}`)
    act(namedRef(ui[mark]), 'click')
    fact = selectionFact()
    button = labeled(ui[mark])
    assert(button.pressed === 'false', `${mark} pressed ${button.pressed}`)
    assert(fact.runs.length > 0 && fact.runs.every((run) => run[mark] !== true), `${mark} off ${JSON.stringify(fact.runs)}`)
  }
})

await check('paragraph style and alignment follow the toolbar', async () => {
  act(namedRef(ui.style), 'select', 'Heading1')
  const styled = selectionFact()
  const style = labeled(ui.style)
  console.log(`  style ${JSON.stringify(styled)} shown ${JSON.stringify(style)}`)
  assert(styled.style === 'Heading1', `model style ${styled.style}`)
  assert(style.value === 'Heading1', `select value ${style.value}`)
  assert(style.selected === ui.heading1, `shown ${style.selected}`)
  for (const [align, key] of ALIGN) {
    act(namedRef(ui[key]), 'click')
    const fact = selectionFact()
    assert(fact.align === align, `${align} model ${fact.align}`)
    assert(labeled(ui[key]).pressed === 'true', `${align} pressed`)
    for (const [, other] of ALIGN) {
      if (other === key) continue
      assert(labeled(ui[other]).pressed === 'false', `${other} still pressed`)
    }
  }
})

await check('bullet list and numbered list set the paragraph numbering', async () => {
  act(namedRef(ui.bullet), 'click')
  const bulleted = selectionFact()
  console.log(`  bullet ${JSON.stringify(bulleted)}`)
  assert(bulleted.format === 'bullet', `format ${bulleted.format}`)
  assert(labeled(ui.bullet).pressed === 'true', 'bullet not pressed')
  assert(labeled(ui.number).pressed === 'false', 'number pressed during a bullet list')
  act(namedRef(ui.number), 'click')
  const numbered = selectionFact()
  console.log(`  number ${JSON.stringify(numbered)}`)
  assert(numbered.format === 'decimal', `format ${numbered.format}`)
  assert(labeled(ui.number).pressed === 'true', 'number not pressed')
  assert(labeled(ui.bullet).pressed === 'false', 'bullet stayed pressed')
})

await check('undo reverts the last change and enables redo', async () => {
  assert(labeled(ui.redo).disabled === true, 'redo is enabled before undo')
  act(namedRef(ui.undo), 'click')
  const undone = selectionFact()
  console.log(`  undone ${JSON.stringify(undone)}`)
  assert(undone.format === 'bullet', `format after undo ${undone.format}`)
  assert(labeled(ui.redo).disabled === false, 'redo stayed disabled')
  assert(labeled(ui.undo).disabled === false, 'undo disabled while earlier edits remain')
})

await check('the status bar shows the page count and the word count', async () => {
  const reported = evaluate(`window.__word.editor.status()`)
  const text = evaluate(`document.querySelector('[role="status"]').textContent`)
  const pages = reported.pagesFinal ? String(reported.pages) : `≥${reported.pages}`
  const line = `${fill(ui.statusPage, { page: reported.page, pages })} · ${fill(ui.statusWords, { words: reported.words })}`
  console.log(`  status ${JSON.stringify(text)} reported ${JSON.stringify(reported)}`)
  assert(text === line, `status ${JSON.stringify(text)} expected ${JSON.stringify(line)}`)
})

await check('the window title carries the dirty mark after an edit', async () => {
  const title = evaluate('document.title')
  console.log(`  title ${JSON.stringify(title)}`)
  assert(title.startsWith('• '), `title ${title}`)
  assert(title.includes('basic.docx'), `title ${title}`)
})

function tabsInfo() {
  return evaluate(`[...document.querySelectorAll('#tabs .tab')].map((tab) => ({
    label: tab.querySelector('.tab-label').textContent,
    active: tab.getAttribute('aria-selected') === 'true',
    dirty: tab.classList.contains('dirty'),
  }))`)
}

await check('new document opens a clean tab beside the dirty one', async () => {
  act(fileMenuItem(ui.new), 'click')
  assert(dialogInfo() === null, 'New does not ask: the dirty document stays in its tab')
  const title = evaluate('document.title')
  const text = bodyText()
  const tabs = tabsInfo()
  console.log(`  new title ${JSON.stringify(title)} text ${JSON.stringify(text)} tabs ${JSON.stringify(tabs)}`)
  assert(title.includes(ui.untitled), `title ${title}`)
  assert(!title.includes('•'), `title ${title}`)
  assert(text === '', `new document text ${JSON.stringify(text)}`)
  assert(tabs.length === 2, `two tabs ${JSON.stringify(tabs)}`)
  assert(tabs[0].label === 'basic' && tabs[0].dirty && !tabs[0].active, `first tab ${JSON.stringify(tabs[0])}`)
  assert(tabs[1].active && !tabs[1].dirty, `second tab ${JSON.stringify(tabs[1])}`)
})

await check('closing the dirty tab asks, and Don\'t save drops it', async () => {
  act(buttonNamed(fill(ui.closeTab, { name: 'basic' })), 'click')
  const asked = dialogInfo()
  assert(asked && asked.hidden === false, 'closing a dirty tab asks')
  assert(asked.name === fill(ui.ask, { name: 'basic.docx' }), `dialog name ${asked.name}`)
  act(buttonNamed(ui.discard), 'click')
  await until(`document.querySelectorAll('#tabs .tab').length === 1`, { label: 'one tab left' })
  const title = evaluate('document.title')
  assert(title.includes(ui.untitled) && !title.includes('•'), `title ${title}`)
  assert(bodyText() === '', 'the untitled tab is the one left')
  assert(evaluate(`document.querySelectorAll('#pages > .doc-host').length`) === 1, 'the closed tab took its host with it')
})

await check('open by picker loads the chosen document', async () => {
  chooseOpen(`${root}fixtures/basic.docx`)
  act(fileMenuItem(ui.open), 'click')
  await until(`(() => {
    const body = window.__word && window.__word.editor && window.__word.editor.state.doc.body
    if (!body) return false
    return body.some((block) => block.kind === 'p' && block.inlines.some((inline) => inline.kind === 'text' && inline.text.includes('Quarterly report')))
  })()`, { timeout: 20000, label: 'picked file' })
  const title = evaluate('document.title')
  assert(title.includes('basic.docx'), `title ${title}`)
  assert(!title.includes('•'), `title ${title}`)
})

function fontProbe() {
  return evaluate(`(() => {
    const editor = window.__word.editor
    const bar = editor.toolbar()
    const sel = editor.state.sel
    if (!sel || sel.kind !== 'text') return { kind: sel ? sel.kind : 'none' }
    const paras = []
    for (const block of editor.state.doc.body) if (block.kind === 'p') paras.push(block)
    const ai = paras.findIndex((para) => para.id === sel.anchor.p)
    const hi = paras.findIndex((para) => para.id === sel.head.p)
    if (ai < 0 || hi < 0) return { kind: 'text', missing: true }
    const forward = ai < hi || (ai === hi && sel.anchor.o <= sel.head.o)
    const start = forward ? sel.anchor : sel.head
    const end = forward ? sel.head : sel.anchor
    const lo = Math.min(ai, hi)
    const hiIndex = Math.max(ai, hi)
    const runs = []
    let covered = ''
    const home = paras[lo]
    let homeText = ''
    for (const inline of home.inlines) {
      if (inline.kind === 'text') homeText += inline.text
    }
    for (let i = lo; i <= hiIndex; i++) {
      const para = paras[i]
      let paraLength = 0
      for (const inline of para.inlines) paraLength += inline.kind === 'text' ? inline.text.length : 1
      const from = i === lo ? start.o : 0
      const to = i === hiIndex ? end.o : paraLength
      let offset = 0
      for (const inline of para.inlines) {
        const chunk = inline.kind === 'text' ? inline.text : '\\uFFFC'
        const runStart = offset
        const runEnd = offset + chunk.length
        offset = runEnd
        if (inline.kind !== 'text' || runEnd <= from || runStart >= to) continue
        const marks = inline.marks || {}
        const text = chunk.slice(Math.max(from, runStart) - runStart, Math.min(to, runEnd) - runStart)
        covered += text
        runs.push({
          text,
          font: marks.font || null,
          sizePt: typeof marks.sizePt === 'number' ? marks.sizePt : null,
        })
      }
    }
    return {
      kind: 'text',
      font: bar.font,
      sizePt: bar.sizePt,
      covered,
      runs,
      id: home.id,
      text: homeText,
    }
  })()`)
}

function paragraphMarks(id) {
  return evaluate(`(() => {
    const id = ${JSON.stringify(id)}
    const para = window.__word.editor.state.doc.body.find((block) => block.kind === 'p' && block.id === id)
    if (!para) return null
    const runs = []
    for (const inline of para.inlines) {
      if (inline.kind !== 'text') continue
      const marks = inline.marks || {}
      runs.push({
        text: inline.text,
        font: marks.font || null,
        sizePt: typeof marks.sizePt === 'number' ? marks.sizePt : null,
      })
    }
    return runs
  })()`)
}

function spanOf(text) {
  return evaluate(`(() => {
    const want = ${JSON.stringify(text)}
    for (const para of document.querySelectorAll('.blk.p')) {
      if (para.getAttribute('aria-hidden') === 'true') continue
      for (const span of para.children) {
        if (span.textContent !== want) continue
        const style = getComputedStyle(span)
        return { family: style.fontFamily, size: style.fontSize }
      }
    }
    return null
  })()`)
}

function px(size) {
  const value = Number.parseFloat(size)
  assert(Number.isFinite(value), `font-size ${size}`)
  return value
}

let shapedWord = ''
let shapedSize = 0
let shapedPara = ''
let shapedMarks = ''
let shapedFont = ''
let shapedPoint = 0

await check('the font and size selects show the title at the caret', async () => {
  await openDocument(basic, 'basic.docx')
  evaluate(`(() => { const view = document.querySelector('.page-view'); if (view) view.scrollTop = 0; return 1 })()`)
  requireControl(ui.font, 'SELECT')
  requireControl(ui.size, 'SELECT')
  act(namedRef('Quarterly report'), 'click')
  await until(`(() => {
    const editor = window.__word.editor
    const sel = editor.state.sel
    if (!sel || sel.kind !== 'text') return false
    const para = editor.state.doc.body.find((block) => block.kind === 'p' && block.id === sel.head.p)
    if (!para) return false
    let text = ''
    for (const inline of para.inlines) if (inline.kind === 'text') text += inline.text
    if (text !== 'Quarterly report') return false
    const bar = editor.toolbar()
    const read = (name) => [...document.querySelectorAll('select')].find((node) => node.getAttribute('aria-label') === name)
    const font = read(${JSON.stringify(ui.font)})
    const size = read(${JSON.stringify(ui.size)})
    if (!font || !size) return false
    const fontValue = bar.font === 'mixed' ? '' : String(bar.font)
    const sizeValue = bar.sizePt === 'mixed' ? '' : String(bar.sizePt)
    return font.value === fontValue && size.value === sizeValue && fontValue !== '' && sizeValue !== ''
  })()`, { timeout: 5000, label: 'title font and size' })
  const font = labeled(ui.font)
  const size = labeled(ui.size)
  const bar = evaluate(`window.__word.editor.toolbar()`)
  console.log(`  title font ${JSON.stringify(font)} size ${JSON.stringify(size)} bar ${JSON.stringify(bar)}`)
  assert(font.value === bar.font, `font select ${font.value} toolbar ${bar.font}`)
  assert(size.value === String(bar.sizePt), `size select ${size.value} toolbar ${bar.sizePt}`)
})

await check('choosing Arial sets the selected word and the font select', async () => {
  evaluate(`(() => {
    const el = [...document.querySelectorAll('.blk.p')].find((node) => node.getAttribute('aria-label') === 'Centered paragraph.' && node.getAttribute('aria-hidden') !== 'true')
    if (el && el.scrollIntoView) el.scrollIntoView({ block: 'center' })
    return !!el
  })()`)
  act(namedRef('Centered paragraph.'), 'double')
  const selected = fontProbe()
  console.log(`  word ${JSON.stringify(selected)}`)
  assert(selected.kind === 'text' && selected.text === 'Centered paragraph.', `paragraph ${JSON.stringify(selected)}`)
  assert(selected.covered.length > 0, `selection ${JSON.stringify(selected.covered)}`)
  shapedWord = selected.covered
  shapedPara = selected.id
  shapedMarks = JSON.stringify(paragraphMarks(shapedPara))
  shapedFont = selected.font
  shapedPoint = selected.sizePt
  act(namedRef(ui.font), 'select', 'Arial')
  const styled = fontProbe()
  const shown = labeled(ui.font)
  const box = spanOf(shapedWord)
  console.log(`  arial ${JSON.stringify(styled.runs)} shown ${JSON.stringify(shown)} box ${JSON.stringify(box)}`)
  assert(styled.runs.length > 0 && styled.runs.every((run) => run.font === 'Arial'), `marks ${JSON.stringify(styled.runs)}`)
  assert(shown.value === 'Arial' && shown.selected === 'Arial', `select ${JSON.stringify(shown)}`)
  assert(box && box.family.replace(/^["']+/, '').startsWith('Arial'), `font-family ${box && box.family}`)
  shapedSize = px(box.size)
})

await check('choosing 20 grows the selected word to that size', async () => {
  act(namedRef(ui.size), 'select', '20')
  const sized = fontProbe()
  const shown = labeled(ui.size)
  const box = spanOf(shapedWord)
  console.log(`  size ${JSON.stringify(sized.runs)} shown ${JSON.stringify(shown)} box ${JSON.stringify(box)}`)
  assert(sized.runs.length > 0 && sized.runs.every((run) => run.sizePt === 20), `marks ${JSON.stringify(sized.runs)}`)
  assert(shown.value === '20' && shown.selected === '20', `select ${JSON.stringify(shown)}`)
  const grown = px(box && box.size)
  assert(grown > shapedSize + 0.5, `font-size ${shapedSize} -> ${box && box.size}`)
  assert(Math.abs(grown - 20 * 4 / 3) < 0.35, `font-size ${box && box.size}`)
})

await check('a selection across the Arial word and its neighbor shows mixed', async () => {
  const sink = namedRef(ui.input)
  act(sink, 'focus')
  act(sink, 'key', 'shift+right')
  const spanned = fontProbe()
  console.log(`  mixed runs ${JSON.stringify(spanned.runs)} toolbar ${spanned.font}`)
  assert(spanned.runs.some((run) => run.font === 'Arial'), `no Arial run ${JSON.stringify(spanned.runs)}`)
  assert(spanned.runs.some((run) => run.font !== 'Arial'), `no neighbor ${JSON.stringify(spanned.runs)}`)
  assert(spanned.font === 'mixed', `toolbar ${spanned.font}`)
  await until(`(() => {
    const el = [...document.querySelectorAll('select')].find((node) => node.getAttribute('aria-label') === ${JSON.stringify(ui.font)})
    const option = el && el.selectedOptions && el.selectedOptions[0]
    return !!(el && el.value === '' && option && option.hidden)
  })()`, { timeout: 5000, label: 'mixed font' })
  const shown = labeled(ui.font)
  assert(shown.hidden === true && shown.value === '' && shown.selected === ui.mixed, `mixed select ${JSON.stringify(shown)}`)
})

await check('undo twice restores the original font and size', async () => {
  act(namedRef(ui.undo), 'click')
  act(namedRef(ui.undo), 'click')
  const marks = JSON.stringify(paragraphMarks(shapedPara))
  const font = labeled(ui.font)
  const size = labeled(ui.size)
  const bar = evaluate(`window.__word.editor.toolbar()`)
  console.log(`  restored marks ${marks} bar ${JSON.stringify(bar)} font ${JSON.stringify(font)} size ${JSON.stringify(size)}`)
  assert(marks === shapedMarks, `marks ${marks} expected ${shapedMarks}`)
  assert(bar.font === shapedFont, `toolbar font ${bar.font} expected ${shapedFont}`)
  assert(bar.sizePt === shapedPoint, `toolbar size ${bar.sizePt} expected ${shapedPoint}`)
  assert(font.value === shapedFont, `font select ${font.value}`)
  assert(size.value === String(shapedPoint), `size select ${size.value}`)
})

await check('zoom steps live in the status bar and stop at the ends', async () => {
  const place = evaluate(`(() => {
    const select = [...document.querySelectorAll('select')].find((node) => node.getAttribute('aria-label') === ${JSON.stringify(ui.zoom)})
    if (!select) return null
    return { inStatus: !!select.closest('#status'), inToolbar: !!select.closest('#toolbar') }
  })()`)
  assert(place && place.inStatus === true && place.inToolbar === false, `zoom place ${JSON.stringify(place)}`)
  assert(labeled(ui.zoomOut).disabled === false, 'zoom out disabled at 100')
  assert(labeled(ui.zoomIn).disabled === false, 'zoom in disabled at 100')
  act(ref('#status select.zoom'), 'select', '50')
  assert(labeled(ui.zoom).value === '50', `zoom ${labeled(ui.zoom).value}`)
  assert(labeled(ui.zoomOut).disabled === true, 'zoom out stays enabled at 50')
  assert(labeled(ui.zoomIn).disabled === false, 'zoom in disabled at 50')
  act(namedRef(ui.zoomIn), 'click')
  assert(labeled(ui.zoom).value === '75', `zoom after step in ${labeled(ui.zoom).value}`)
  assert(labeled(ui.zoomOut).disabled === false, 'zoom out disabled at 75')
  act(ref('#status select.zoom'), 'select', '200')
  assert(labeled(ui.zoomIn).disabled === true, 'zoom in stays enabled at 200')
  assert(labeled(ui.zoomOut).disabled === false, 'zoom out disabled at 200')
  act(namedRef(ui.zoomOut), 'click')
  assert(labeled(ui.zoom).value === '150', `zoom after step out ${labeled(ui.zoom).value}`)
  act(ref('#status select.zoom'), 'select', '100')
  assert(labeled(ui.zoom).value === '100', `zoom restored ${labeled(ui.zoom).value}`)
})

await check('console has no errors', async () => {
  const errors = consoleErrors()
  assert(errors.length === 0, errors.join('\n'))
})

async function showLocale(lang) {
  const token = `${lang}-${Date.now()}`
  goTracked(`?lang=${lang}&boot=${token}`)
  await until(
    `new URLSearchParams(location.search).get('boot') === ${JSON.stringify(token)} && document.documentElement.lang === ${JSON.stringify(lang)} && !!(window.__word && window.__word.editor)`,
    { timeout: 60000, label: `locale ${lang}` },
  )
  await openDocument(basic, 'basic.docx')
  evaluate(`(() => { const view = document.querySelector('.page-view'); if (view) view.scrollTop = 0; return 1 })()`)
}

await check('screenshot of the window in zh', async () => {
  await showLocale('zh')
  shot(evidencePath('shell/window-zh.png'))
})

await check('screenshot of the window in en', async () => {
  await showLocale('en')
  shot(evidencePath('shell/window-en.png'))
})

await summary()
