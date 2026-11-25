// The document tab strip above the toolbar. One tab per open document; the app owns the documents.
import { t } from '../i18n.ts'
import { icon } from './icons.ts'
import { tabLabel } from './text.ts'

export interface TabItem {
  readonly id: number
  /** The file name, with its extension. The tab shows it without `.docx`. */
  readonly name: string
  readonly dirty: boolean
}

export interface TabsHandlers {
  onSelect(id: number): void
  onClose(id: number): void
  onNew(): void
}

export interface TabsView {
  render(items: readonly TabItem[], active: number): void
  /** The print preview covers the editing chrome, so the tabs stop taking clicks while it is up. */
  setCovered(covered: boolean): void
}

export function createTabs(root: HTMLElement, handlers: TabsHandlers): TabsView {
  root.setAttribute('role', 'tablist')
  root.setAttribute('aria-label', t('tabs'))
  const list = document.createElement('div')
  list.className = 'tabs-list'
  const add = document.createElement('button')
  add.type = 'button'
  add.className = 'tab-new'
  add.setAttribute('aria-label', t('tab.new'))
  add.title = t('tab.new')
  add.append(icon('zoom-in'))
  add.addEventListener('click', () => handlers.onNew())
  root.append(list, add)

  const rows = new Map<number, { el: HTMLElement; label: HTMLElement; close: HTMLButtonElement; key: string }>()
  let order = ''

  function row(id: number): { el: HTMLElement; label: HTMLElement; close: HTMLButtonElement; key: string } {
    const found = rows.get(id)
    if (found) return found
    const el = document.createElement('div')
    el.className = 'tab'
    el.setAttribute('role', 'tab')
    el.dataset.id = String(id)
    const dot = document.createElement('span')
    dot.className = 'tab-dirty'
    dot.setAttribute('aria-hidden', 'true')
    const label = document.createElement('span')
    label.className = 'tab-label'
    const close = document.createElement('button')
    close.type = 'button'
    close.className = 'tab-close'
    close.tabIndex = -1
    close.append(icon('close'))
    el.append(dot, label, close)
    // mousedown keeps the editor's focus where it was; the app moves focus to the chosen tab's editor.
    el.addEventListener('mousedown', (event) => {
      if (event.button !== 0 || close.contains(event.target as Node)) return
      event.preventDefault()
      handlers.onSelect(id)
    })
    // A middle click closes, as in a browser.
    el.addEventListener('auxclick', (event) => {
      if (event.button !== 1) return
      event.preventDefault()
      handlers.onClose(id)
    })
    close.addEventListener('mousedown', event => event.preventDefault())
    close.addEventListener('click', (event) => {
      event.stopPropagation()
      handlers.onClose(id)
    })
    const made = { el, label, close, key: '' }
    rows.set(id, made)
    return made
  }

  return {
    render(items, active) {
      const live = new Set(items.map(item => item.id))
      for (const [id, made] of rows) {
        if (live.has(id)) continue
        made.el.remove()
        rows.delete(id)
      }
      for (const item of items) {
        const made = row(item.id)
        const selected = item.id === active
        const key = `${item.name}\u0000${item.dirty}\u0000${selected}`
        if (made.key === key) continue
        made.key = key
        const label = tabLabel(item.name)
        made.label.textContent = label
        made.el.title = item.name
        made.el.classList.toggle('dirty', item.dirty)
        made.el.classList.toggle('active', selected)
        made.el.setAttribute('aria-selected', String(selected))
        made.close.setAttribute('aria-label', t('tab.close', { name: label }))
        made.close.title = t('tab.close', { name: label })
      }
      const nextOrder = items.map(item => item.id).join(',')
      if (nextOrder !== order) {
        order = nextOrder
        list.replaceChildren(...items.map(item => row(item.id).el))
      }
      root.classList.toggle('single', items.length === 1)
    },
    setCovered(covered) {
      root.classList.toggle('covered', covered)
      add.disabled = covered
    },
  }
}
