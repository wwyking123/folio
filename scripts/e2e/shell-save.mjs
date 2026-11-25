// Unit 4: Save As, Save, the dirty bit across undo, and the unsaved dialog.
// `acde act w1 close` waits until the window is gone, so a beforeclose that returns false cannot be
// driven from the CLI and still continue. The dev hook window.__word.requestClose() runs that handler.
// There is no separate production build. Run by the reviewer; this dispatch only syntax-checks the file.
import { execFileSync } from 'node:child_process'
import { copyFileSync, mkdtempSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { act, actClosing, assert, check, closeFileMenu, consoleErrors, ensureDev, evaluate, fileMenuItem, namedRef, root, summary, until } from './acde.mjs'

const STRINGS = {
  en: {
    save: 'Save',
    saveAs: 'Save As',
    discard: "Don't save",
    cancel: 'Cancel',
    ask: 'Save changes to {name}?',
    closeTab: 'Close {name}',
    error: 'Could not save this document.',
    undo: 'Undo',
    redo: 'Redo',
    bold: 'Bold',
    font: 'Font',
    size: 'Font size',
    style: 'Paragraph style',
    bullet: 'Bulleted list',
    input: 'Document text',
    new: 'New',
    untitled: 'Untitled.docx',
  },
  zh: {
    save: '保存',
    saveAs: '另存为',
    discard: '不保存',
    cancel: '取消',
    ask: '是否保存对 {name} 的更改？',
    closeTab: '关闭 {name}',
    error: '无法保存此文档。',
    undo: '撤销',
    redo: '重做',
    bold: '加粗',
    font: '字体',
    size: '字号',
    style: '段落样式',
    bullet: '项目符号',
    input: '文档正文',
    new: '新建',
    untitled: '未命名.docx',
  },
}

const ARG_LIMIT = 100 * 1024
const CHUNK = 48 * 1024
const basic = readFileSync(new URL('../../fixtures/basic.docx', import.meta.url)).toString('base64')
const dir = mkdtempSync(`${tmpdir()}/word-save-`)
const pathA = `${dir}/basic.docx`
const savedA = `${dir}/saved-a.docx`
const savedB = `${dir}/saved-b.docx`
const pathDialog = `${dir}/dialog.docx`
const pathClose = `${dir}/close.docx`
const pathTitle = `${dir}/title.docx`

let bootToken = ''

function go(search) {
  try {
    evaluate(`location.assign(location.pathname + ${JSON.stringify(search)})`)
  } catch {
    // Reloading the page aborts the eval.
  }
}

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
    return { tag: el.tagName, disabled: el.disabled === true, title: el.getAttribute('title') }
  })()`)
}

function inspect(name) {
  try {
    return execFileSync('acde', ['inspect', '--find', name], { cwd: root, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] })
  } catch (error) {
    return `${error && error.stdout ? error.stdout : ''}\n${error && error.stderr ? error.stderr : ''}`
  }
}

function buttonRef(name, dialog) {
  const line = inspect(name).split('\n').find((item) => {
    const quoted = item.match(/button "([^"]*)"/)
    if (!quoted || quoted[1] !== name) return false
    const inDialog = item.includes(' in dialog ')
    return dialog ? inDialog : !inDialog
  })
  if (!line) throw new Error(`no ${dialog ? 'dialog ' : ''}button ${JSON.stringify(name)}`)
  const match = line.match(/\[(w\d+\.[\w.]+)\]/)
  if (!match) throw new Error(line)
  return match[1]
}

function chooseSave(path) {
  execFileSync('acde', ['choose', 'save', path], { cwd: root, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] })
}

function fill(template, values) {
  return template.replace(/\{(\w+)\}/g, (all, key) => (values[key] === undefined ? all : String(values[key])))
}

function installProject() {
  evaluate(`(() => {
    window.__project = () => {
      const doc = window.__word.editor.state.doc
      function textOf(para) {
        let text = ''
        for (const inline of para.inlines) if (inline.kind === 'text') text += inline.text
        return text
      }
      function runs(para) {
        const out = []
        for (const inline of para.inlines) {
          if (inline.kind === 'atom') {
            out.push({ atom: inline.atom.type })
            continue
          }
          const marks = inline.marks || {}
          const key = JSON.stringify(marks, Object.keys(marks).sort())
          const last = out[out.length - 1]
          if (last && last.text !== undefined && last.key === key) last.text += inline.text
          else out.push({ text: inline.text, marks, key })
        }
        return out.map((run) => {
          const rest = { ...run }
          delete rest.key
          return rest
        })
      }
      function listOf(para) {
        const numbering = para.props.numbering
        if (!numbering) return null
        const def = doc.lists.byNum.get(numbering.num)
        const level = def && def.levels[numbering.level]
        return level ? level.format + ':' + level.text : 'missing'
      }
      function walk(blocks) {
        return blocks.map((block) => {
          if (block.kind === 'p') {
            return {
              text: textOf(block),
              runs: runs(block),
              style: block.props.style || null,
              align: block.props.align || null,
              list: listOf(block),
            }
          }
          if (block.kind === 'table') {
            return { kind: 'table', rows: block.rows.map((row) => row.cells.map((cell) => walk(cell.blocks))) }
          }
          return { kind: block.kind, label: block.label }
        })
      }
      return { body: walk(doc.body), stories: [...doc.stories].map(([key, blocks]) => [key, walk(blocks)]) }
    }
    return 1
  })()`)
}

function project() {
  return evaluate(`window.__project()`)
}

function norm(value) {
  if (Array.isArray(value)) return value.map(norm)
  if (value && typeof value === 'object') {
    const out = {}
    for (const key of Object.keys(value).sort()) out[key] = norm(value[key])
    return out
  }
  return value
}

function same(actual, expected, label) {
  assert(JSON.stringify(norm(actual)) === JSON.stringify(norm(expected)), label)
}

function pageAlive() {
  try {
    evaluate('1')
    return true
  } catch {
    return false
  }
}

async function boot() {
  await ensureDev()
  const token = `save-${Date.now()}`
  goTracked(`?lang=en&boot=${token}`)
  await until(
    `new URLSearchParams(location.search).get('boot') === ${JSON.stringify(bootToken)} && !!(window.__word && window.__word.requestClose)`,
    { timeout: 60000, label: 'shell' },
  )
  installProject()
}

function sink() {
  return namedRef(names().input)
}

function dialogInfo() {
  return evaluate(`(() => {
    const dialog = document.querySelector('[role="dialog"]')
    if (!dialog) return null
    const id = dialog.getAttribute('aria-labelledby')
    const label = id ? document.getElementById(id) : null
    const backdrop = dialog.closest('.unsaved')
    return {
      modal: dialog.getAttribute('aria-modal'),
      name: label ? label.textContent : '',
      hidden: !backdrop || backdrop.hidden === true,
    }
  })()`)
}

function errorBand() {
  return evaluate(`(() => {
    const el = document.querySelector('#open-error')
    return { hidden: !el || el.hidden === true, text: el ? el.textContent : '' }
  })()`)
}

function dirty() {
  return evaluate(`document.title.startsWith('• ')`)
}

await boot()
const ui = names()

await check('Save and Save As sit after Open and show their chords', async () => {
  await openDocument(basic, 'basic.docx')
  const opened = evaluate(`window.__wordError || ''`)
  assert(!opened, opened)
  fileMenuItem(ui.saveAs)
  const save = labeled(ui.save)
  const saveAs = labeled(ui.saveAs)
  assert(save && save.tag === 'BUTTON', 'Save button')
  assert(saveAs && saveAs.tag === 'BUTTON', 'Save As button')
  assert(save.disabled === false, 'Save stays enabled until the document has a handle')
  assert(save.title && (save.title.includes('Ctrl+S') || save.title.includes('⌘S')), `save title ${save.title}`)
  assert(save.title && !save.title.includes('Shift') && !save.title.includes('⇧'), `save title ${save.title}`)
  assert(saveAs.title && (saveAs.title.includes('Ctrl+Shift+S') || saveAs.title.includes('⌘⇧S')), `save as title ${saveAs.title}`)
  closeFileMenu()
})

await check('edits save through Save As, and undo past that point is dirty again', async () => {
  act(sink(), 'focus')
  act(sink(), 'type', 'Hello 你好')
  act(sink(), 'key', 'enter')
  const split = evaluate(`(() => {
    const editor = window.__word.editor
    const para = editor.state.doc.body.find((block) => block.kind === 'p' && block.inlines.some((inline) => inline.kind === 'text' && inline.text === 'Overview'))
    if (!para) return false
    editor.run(window.__word.setSelection({ kind: 'text', anchor: { p: para.id, o: 0 }, head: { p: para.id, o: 0 } }))
    return true
  })()`)
  assert(split === true, 'Overview is still its own paragraph after the split')
  act(sink(), 'focus')
  act(sink(), 'key', 'backspace')
  act(namedRef('Plain, '), 'double')
  act(namedRef(ui.bold), 'click')
  act(namedRef(ui.font), 'select', 'Arial')
  act(namedRef(ui.size), 'select', '20')
  act(namedRef(ui.style), 'select', 'Heading1')
  act(namedRef(ui.bullet), 'click')
  const before = project()
  const joined = before.body.some((block) => block.text && block.text.includes('Quarterly reportOverview'))
  const typed = before.body.some((block) => block.text && block.text.includes('Hello 你好'))
  assert(typed, 'typed English and Chinese')
  assert(joined, 'join kept both paragraphs text')
  assert(before.body.some((block) => block.list && block.list.startsWith('bullet:')), 'list format')
  assert(before.body.some((block) => block.style === 'Heading1'), 'style')
  chooseSave(pathA)
  act(fileMenuItem(ui.saveAs), 'click')
  await until(`!document.title.startsWith('• ') && document.title.includes('basic.docx')`, { label: 'clean after Save As' })
  assert(labeled(ui.save).disabled === true, 'a clean document with a handle disables Save')
  copyFileSync(pathA, savedA)
  act(namedRef(ui.undo), 'click')
  await until(`document.title.startsWith('• ')`, { label: 'undo is dirty' })
  act(namedRef(ui.redo), 'click')
  await until(`!document.title.startsWith('• ')`, { label: 'redo is clean' })
  assert(JSON.stringify(norm(project())) === JSON.stringify(norm(before)), 'redo restored the saved document')
  act(sink(), 'focus')
  act(sink(), 'type', 'Again ')
  const during = project()
  assert(dirty() === true, 'the new edit is dirty')
  act(buttonRef(ui.save, false), 'click')
  await until(`!document.title.startsWith('• ')`, { timeout: 20000, label: 'clean after Save' })
  copyFileSync(pathA, savedB)
  assert(!readFileSync(savedA).equals(readFileSync(savedB)), 'Save wrote a new file without a picker')
  await openDocument(readFileSync(savedA).toString('base64'), 'saved-a.docx')
  same(project(), before, 'Save As bytes match the model from before that save')
  await openDocument(readFileSync(savedB).toString('base64'), 'saved-b.docx')
  same(project(), during, 'Save bytes match the model from before that save')
})

await check('a verify failure shows the error, writes nothing and stays dirty', async () => {
  const before = readFileSync(savedB)
  act(sink(), 'focus')
  act(sink(), 'type', 'Z')
  assert(dirty() === true, 'dirty before the tampered save')
  const detail = evaluate(`window.__word.saveTampered()`)
  const band = errorBand()
  assert(typeof detail === 'string' && detail.length > 0, 'no verify detail')
  assert(band.hidden === false && band.text.includes(ui.error) && band.text.includes(detail), `band ${band.text}`)
  assert(dirty() === true, 'still dirty')
  assert(readFileSync(savedB).equals(before), 'the file on disk was not replaced')
})

await check('beforeclose on a dirty document keeps the window and the three buttons answer', async () => {
  const kept = evaluate(`window.__word.requestClose()`)
  assert(kept === false, 'beforeclose returned false')
  assert(pageAlive(), 'the window is still open')
  const info = dialogInfo()
  assert(info && info.hidden === false, 'dialog is showing')
  assert(info.modal === 'true', 'aria-modal')
  assert(info.name === fill(ui.ask, { name: 'saved-b.docx' }), `dialog name ${info.name}`)
  const save = buttonRef(ui.save, true)
  const discard = buttonRef(ui.discard, true)
  const cancel = buttonRef(ui.cancel, true)
  assert(evaluate(`document.activeElement && document.activeElement.textContent`) === ui.save, 'Save is focused')
  act(save, 'key', 'tab')
  assert(evaluate(`document.activeElement && document.activeElement.textContent`) === ui.discard, 'Tab moves to Don\'t save')
  act(discard, 'key', 'shift+tab')
  assert(evaluate(`document.activeElement && document.activeElement.textContent`) === ui.save, 'Shift+Tab stays in the dialog')
  act(save, 'key', 'escape')
  await until(`document.querySelector('[role="dialog"]') === null || document.querySelector('.unsaved').hidden === true`, { label: 'escape closes the dialog' })
  assert(pageAlive() && dirty() === true, 'Escape left the dirty document open')
  assert(evaluate(`window.__word.requestClose()`) === false)
  act(buttonRef(ui.cancel, true), 'click')
  await until(`document.querySelector('.unsaved').hidden === true`, { label: 'cancel' })
  assert(pageAlive() && dirty() === true, 'Cancel left the dirty document open')
  const text = evaluate(`window.__project().body.map((block) => block.text || '').join('\\n')`)
  // Closing the document's tab asks the same question. The last tab gives way to a new untitled one.
  const closeTab = (name) => act(buttonRef(fill(ui.closeTab, { name }), false), 'click')
  closeTab('saved-b')
  assert(dialogInfo() && dialogInfo().hidden === false, 'closing a dirty tab asks')
  act(buttonRef(ui.cancel, true), 'click')
  await until(`document.querySelector('.unsaved').hidden === true`, { label: 'close cancel' })
  const still = evaluate(`window.__project().body.map((block) => block.text || '').join('\\n')`)
  assert(still === text, 'Cancel kept the document')
  assert(evaluate(`document.querySelectorAll('#tabs .tab').length`) === 1, 'Cancel kept the tab')
  const snap = project()
  chooseSave(pathDialog)
  closeTab('saved-b')
  act(buttonRef(ui.save, true), 'click')
  await until(`!document.title.startsWith('• ') && document.title.includes(${JSON.stringify(ui.untitled)})`, { timeout: 20000, label: 'saved then closed' })
  await openDocument(readFileSync(pathDialog).toString('base64'), 'dialog.docx')
  same(project(), snap, 'the dialog Save button wrote the document')
  act(sink(), 'focus')
  act(sink(), 'type', 'Q')
  closeTab('dialog')
  act(buttonRef(ui.discard, true), 'click')
  await until(`!document.title.startsWith('• ') && document.title.includes(${JSON.stringify(ui.untitled)})`, { label: 'discard then closed' })
  const fresh = evaluate(`window.__project().body.map((block) => block.text || '').join('')`)
  assert(!fresh.includes('Q'), 'Don\'t save dropped the edit')
})

await check('Save on close writes and closes, Don\'t save closes without writing', async () => {
  act(sink(), 'focus')
  act(sink(), 'type', 'Close me')
  assert(evaluate(`window.__word.requestClose()`) === false)
  chooseSave(pathClose)
  await actClosing(buttonRef(ui.save, true), 'click')
  assert(readFileSync(pathClose).subarray(0, 2).toString() === 'PK', 'close Save wrote a docx')
  await boot()
  await openDocument(readFileSync(pathClose).toString('base64'), 'close.docx')
  const text = evaluate(`window.__project().body.map((block) => block.text || '').join('')`)
  assert(text.includes('Close me'), 'the closed Save kept the edit')
  act(sink(), 'focus')
  act(sink(), 'type', 'Drop me')
  assert(evaluate(`window.__word.requestClose()`) === false)
  const disk = readFileSync(pathClose)
  await actClosing(buttonRef(ui.discard, true), 'click')
  assert(readFileSync(pathClose).equals(disk), 'Don\'t save left the file as it was')
  await boot()
  await openDocument(disk.toString('base64'), 'close.docx')
  const kept = evaluate(`window.__project().body.map((block) => block.text || '').join('')`)
  assert(kept.includes('Close me') && !kept.includes('Drop me'), 'Don\'t save did not write the discarded edit')
})

await check('Save As keeps a title split into Hello 你好 and Quarterly report', async () => {
  await openDocument(basic, 'basic.docx')
  const opened = evaluate(`window.__wordError || ''`)
  assert(!opened, opened)
  const title = project().body[0]
  assert(title && title.text === 'Quarterly report', `title ${title && title.text}`)
  assert(title.style, 'the title has a style')
  act(sink(), 'focus')
  act(sink(), 'type', 'Hello 你好')
  act(sink(), 'key', 'enter')
  chooseSave(pathTitle)
  act(fileMenuItem(ui.saveAs), 'click')
  await until(`!document.title.startsWith('• ')`, { timeout: 20000, label: 'clean after the title split' })
  await openDocument(readFileSync(pathTitle).toString('base64'), 'title.docx')
  const again = evaluate(`window.__wordError || ''`)
  assert(!again, again)
  const body = project().body
  const first = body[0]
  const second = body[1]
  assert(first && first.text === 'Hello 你好', `first ${first && first.text}`)
  assert(second && second.text === 'Quarterly report', `second ${second && second.text}`)
  assert(first.style === title.style && second.style === title.style, `styles ${first && first.style} ${second && second.style}`)
})

await check('console has no errors', async () => {
  const errors = consoleErrors()
  assert(errors.length === 0, errors.join('\n'))
})

await summary()
