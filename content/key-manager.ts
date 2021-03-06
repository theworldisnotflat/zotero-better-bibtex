declare const Zotero: any
declare const window: any

import ETA = require('node-eta')
import { kuroshiro } from './key-manager/kuroshiro'

import { log } from './logger'
import { sleep } from './sleep'
import { flash } from './flash'
import { Events, itemsChanged as notifiyItemsChanged } from './events'
import { arXiv } from './arXiv'
import * as Extra from './extra'

import * as ZoteroDB from './db/zotero'

import { getItemsAsync } from './get-items-async'

import { Preferences as Prefs } from './prefs'
import { Formatter } from './key-manager/formatter'
import { DB } from './db/main'
import { DB as Cache } from './db/cache'

import { patch as $patch$ } from './monkey-patch'

import { sprintf } from 'sprintf-js'
import { intToExcelCol } from 'excel-column-name'

// export singleton: https://k94n.com/es6-modules-single-instance-pattern
export let KeyManager = new class { // tslint:disable-line:variable-name
  public keys: any
  public query: {
    field: { extra?: number }
    type: {
      note?: number,
      attachment?: number
    }
  }

  private itemObserverDelay: number = Prefs.get('itemObserverDelay')
  private scanning: any[]
  private started = false

  private async inspireHEP(url) {
    try {
      const results = await (await fetch(url, { method: 'GET', cache: 'no-cache', redirect: 'follow' })).json()
      if (results.metadata.texkeys.length !== 1) throw new Error(`expected 1 key, got ${results.metadata.texkeys.length}`)
      return results.metadata.texkeys[0]
    } catch (err) {
      log.error('inspireHEP', url, err)
      return null
    }
  }

  private getField(item, field): string {
    try {
      return item.getField(field) || ''
    } catch (err) {
      return ''
    }
  }
  public async pin(ids, inspireHEP = false) {
    ids = this.expandSelection(ids)

    for (const item of await getItemsAsync(ids)) {
      if (item.isNote() || item.isAttachment()) continue

      const extra = this.getField(item, 'extra')
      const parsed = Extra.get(extra, 'zotero')
      let citationKey: string = null

      if (inspireHEP) {
        const doi = (this.getField(item, 'DOI') || parsed.extraFields.kv.DOI || '').replace(/^https?:\/\/doi.org\//i, '')
        const arxiv = ((['arxiv.org', 'arxiv'].includes((this.getField(item, 'libraryCatalog') || '').toLowerCase())) && arXiv.parse(this.getField(item, 'publicationTitle')).id) || arXiv.parse(parsed.extraFields.tex.arxiv).id

        if (!doi && !arxiv) continue

        if (doi) citationKey = await this.inspireHEP(`https://inspirehep.net/api/doi/${doi}`)
        if (!citationKey && arxiv) citationKey = await this.inspireHEP(`https://inspirehep.net/api/arxiv/${arxiv}`)

        if (!citationKey) continue

        if (parsed.extraFields.citationKey === citationKey) continue

      } else {
        if (parsed.extraFields.citationKey) continue

        citationKey = this.get(item.id).citekey || this.update(item)
      }

      item.setField('extra', Extra.set(extra, { citationKey }))
      await item.saveTx() // this should cause an update and key registration
    }
  }

  public async unpin(ids) {
    ids = this.expandSelection(ids)

    for (const item of await getItemsAsync(ids)) {
      if (item.isNote() || item.isAttachment()) continue

      const parsed = Extra.get(item.getField('extra'), 'zotero', { citationKey: true })
      if (!parsed.extraFields.citationKey) continue

      item.setField('extra', parsed.extra) // citekey is stripped here but will be regenerated by the notifier
      item.saveTx()
    }

  }

  public async refresh(ids, manual = false) {
    ids = this.expandSelection(ids)

    Cache.remove(ids, `refreshing keys for ${ids}`)

    const warnAt = manual ? Prefs.get('warnBulkModify') : 0
    if (warnAt > 0 && ids.length > warnAt) {
      const affected = this.keys.find({ itemID: { $in: ids }, pinned: false }).length
      if (affected > warnAt) {
        const params = { treshold: warnAt, response: null }
        window.openDialog('chrome://zotero-better-bibtex/content/bulk-keys-confirm.xul', '', 'chrome,dialog,centerscreen,modal', params)
        switch (params.response) {
          case 'ok':
            break
          case 'whatever':
            Prefs.set('warnBulkModify', 0)
            break
          default:
            return
        }
      }
    }

    const updates = []
    for (const item of await getItemsAsync(ids)) {
      if (item.isNote() || item.isAttachment()) continue

      const extra = item.getField('extra')

      let citekey = Extra.get(extra, 'zotero', { citationKey: true }).extraFields.citationKey
      if (citekey) continue // pinned, leave it alone

      this.update(item)

      // remove the new citekey from the aliases if present
      citekey = this.get(item.id).citekey
      const aliases = Extra.get(extra, 'zotero', { aliases: true })
      if (aliases.extraFields.aliases.includes(citekey)) {
        aliases.extraFields.aliases = aliases.extraFields.aliases.filter(alias => alias !== citekey)

        if (aliases.extraFields.aliases.length) {
          item.setField('extra', Extra.set(aliases.extra, { aliases: aliases.extraFields.aliases }))
        } else {
          item.setField('extra', aliases.extra)
        }
      }

      if (manual) updates.push(item)
    }

    if (manual) notifiyItemsChanged(updates)
  }

  public async init() {
    await kuroshiro.init()

    this.keys = DB.getCollection('citekey')

    this.query = {
      field: {},
      type: {},
    }

    for (const type of await ZoteroDB.queryAsync('select itemTypeID, typeName from itemTypes')) { // 1 = attachment, 14 = note
      this.query.type[type.typeName] = type.itemTypeID
    }

    for (const field of await ZoteroDB.queryAsync('select fieldID, fieldName from fields')) {
      this.query.field[field.fieldName] = field.fieldID
    }

    Formatter.update('init')
  }

  public async start() {
    await this.rescan()

    await ZoteroDB.queryAsync('ATTACH DATABASE ":memory:" AS betterbibtexcitekeys')
    await ZoteroDB.queryAsync('CREATE TABLE betterbibtexcitekeys.citekeys (itemID PRIMARY KEY, itemKey, citekey)')
    await Zotero.DB.executeTransaction(async () => {
      for (const key of this.keys.data) {
        await ZoteroDB.queryAsync('INSERT INTO betterbibtexcitekeys.citekeys (itemID, itemKey, citekey) VALUES (?, ?, ?)', [ key.itemID, key.itemKey, key.citekey ])
      }
    })

    const citekeySearchCondition = {
      name: 'citationKey',
      operators: {
        is: true,
        isNot: true,
        contains: true,
        doesNotContain: true,
      },
      table: 'betterbibtexcitekeys.citekeys',
      field: 'citekey',
      localized: 'Citation Key',
    }
    $patch$(Zotero.Search.prototype, 'addCondition', original => function addCondition(condition, operator, value, required) {
      // detect a quick search being set up
      if (condition.match(/^quicksearch/)) this.__add_bbt_citekey = true
      // creator is always added in a quick search so use it as a trigger
      if (condition === 'creator' && this.__add_bbt_citekey) {
        original.call(this, citekeySearchCondition.name, operator, value, false)
        delete this.__add_bbt_citekey
      }
      return original.apply(this, arguments)
    })
    $patch$(Zotero.SearchConditions, 'hasOperator', original => function hasOperator(condition, operator) {
      if (condition === citekeySearchCondition.name) return citekeySearchCondition.operators[operator]
      return original.apply(this, arguments)
    })
    $patch$(Zotero.SearchConditions, 'get', original => function get(condition) {
      if (condition === citekeySearchCondition.name) return citekeySearchCondition
      return original.apply(this, arguments)
    })
    $patch$(Zotero.SearchConditions, 'getStandardConditions', original => function getStandardConditions() {
      return original.apply(this, arguments).concat({
        name: citekeySearchCondition.name,
        localized: citekeySearchCondition.localized,
        operators: citekeySearchCondition.operators,
      }).sort((a, b) => a.localized.localeCompare(b.localized))
    })
    $patch$(Zotero.SearchConditions, 'getLocalizedName', original => function getLocalizedName(str) {
      if (str === citekeySearchCondition.name) return citekeySearchCondition.localized
      return original.apply(this, arguments)
    })

    Events.on('preference-changed', pref => {
      if (['autoAbbrevStyle', 'citekeyFormat', 'citekeyFold', 'skipWords'].includes(pref)) {
        Formatter.update('pref-change')
      }
    })

    this.keys.on(['insert', 'update'], async citekey => {
      await ZoteroDB.queryAsync('INSERT OR REPLACE INTO betterbibtexcitekeys.citekeys (itemID, itemKey, citekey) VALUES (?, ?, ?)', [ citekey.itemID, citekey.itemKey, citekey.citekey ])

      // async is just a heap of fun. Who doesn't enjoy a good race condition?
      // https://github.com/retorquere/zotero-better-bibtex/issues/774
      // https://groups.google.com/forum/#!topic/zotero-dev/yGP4uJQCrMc
      await sleep(this.itemObserverDelay)

      try {
        await Zotero.Items.getAsync(citekey.itemID)
      } catch (err) {
        // assume item has been deleted before we could get to it -- did I mention I hate async? I hate async
        log.error('could not load', citekey.itemID, err)
        return
      }

      if (Prefs.get('autoPin') && !citekey.pinned) {
        this.pin([citekey.itemID])
      } else {
        // update display panes by issuing a fake item-update notification
        Zotero.Notifier.trigger('modify', 'item', [citekey.itemID], { [citekey.itemID]: { bbtCitekeyUpdate: true } })
      }
    })
    this.keys.on('delete', async citekey => {
      await ZoteroDB.queryAsync('DELETE FROM betterbibtexcitekeys.citekeys WHERE itemID = ?', [ citekey.itemID ])
    })

    this.started = true
  }

  public async rescan(clean?: boolean) {
    if (Prefs.get('scrubDatabase')) {
      for (const item of this.keys.where(i => i.hasOwnProperty('extra'))) { // 799
        delete item.extra
        this.keys.update(item)
      }
    }

    if (Array.isArray(this.scanning)) {
      let left
      if (this.scanning.length) {
        left = `, ${this.scanning.length} items left`
      } else {
        left = ''
      }
      flash('Scanning still in progress', `Scan is still running${left}`)
      return
    }

    this.scanning = []

    if (clean) this.keys.removeDataOnly()

    const marker = '\uFFFD'

    let bench = this.bench('cleanup')
    const ids = []
    const items = await ZoteroDB.queryAsync(`
      SELECT item.itemID, item.libraryID, item.key, extra.value as extra, item.itemTypeID
      FROM items item
      LEFT JOIN itemData field ON field.itemID = item.itemID AND field.fieldID = ${this.query.field.extra}
      LEFT JOIN itemDataValues extra ON extra.valueID = field.valueID
      WHERE item.itemID NOT IN (select itemID from deletedItems)
      AND item.itemTypeID NOT IN (${this.query.type.attachment}, ${this.query.type.note})
    `)
    for (const item of items) {
      ids.push(item.itemID)
      // if no citekey is found, it will be '', which will allow it to be found right after this loop
      const extra = Extra.get(item.extra, 'zotero', { citationKey: true })

      // don't fetch when clean is active because the removeDataOnly will have done it already
      const existing = clean ? null : this.keys.findOne({ itemID: item.itemID })
      if (!existing) {
        // if the extra doesn't have a citekey, insert marker, next phase will find & fix it
        this.keys.insert({ citekey: extra.extraFields.citationKey || marker, pinned: !!extra.extraFields.citationKey, itemID: item.itemID, libraryID: item.libraryID, itemKey: item.key })

      } else if (extra.extraFields.citationKey && ((extra.extraFields.citationKey !== existing.citekey) || !existing.pinned)) {
        // we have an existing key in the DB, extra says it should be pinned to the extra value, but it's not.
        // update the DB to have the itemkey if necessaru
        this.keys.update({ ...existing, citekey: extra.extraFields.citationKey, pinned: true, itemKey: item.key })

      } else if (!existing.itemKey) {
        this.keys.update({ ...existing, itemKey: item.key })
      }
    }

    this.keys.findAndRemove({ itemID: { $nin: ids } })
    this.bench(bench)

    bench = this.bench('regenerate')
    // find all references without citekey
    this.scanning = this.keys.find({ citekey: marker })

    if (this.scanning.length !== 0) {
      const progressWin = new Zotero.ProgressWindow({ closeOnClick: false })
      progressWin.changeHeadline('Better BibTeX: Assigning citation keys')
      progressWin.addDescription(`Found ${this.scanning.length} references without a citation key`)
      const icon = `chrome://zotero/skin/treesource-unfiled${Zotero.hiDPI ? '@2x' : ''}.png`
      const progress = new progressWin.ItemProgress(icon, 'Assigning citation keys')
      progressWin.show()

      const eta = new ETA(this.scanning.length, { autoStart: true })
      for (let done = 0; done < this.scanning.length; done++) {
        let key = this.scanning[done]
        const item = await getItemsAsync(key.itemID)

        if (key.citekey === marker) {
          if (key.pinned) {
            const parsed = Extra.get(item.getField('extra'), 'zotero', { citationKey: true })
            item.setField('extra', parsed.extra)
            await item.saveTx({ [key.itemID]: { bbtCitekeyUpdate: true } })
          }
          key = null
        }

        try {
          this.update(item, key)
        } catch (err) {
          log.error('KeyManager.rescan: update', done, 'failed:', err)
        }

        eta.iterate()

        // tslint:disable-next-line:no-magic-numbers
        if ((done % 10) === 1) {
          // tslint:disable-next-line:no-magic-numbers
          progress.setProgress((eta.done * 100) / eta.count)
          progress.setText(eta.format(`${eta.done} / ${eta.count}, {{etah}} remaining`))
        }
      }

      // tslint:disable-next-line:no-magic-numbers
      progress.setProgress(100)
      progress.setText('Ready')
      // tslint:disable-next-line:no-magic-numbers
      progressWin.startCloseTimer(500)
    }
    this.bench(bench)

    this.scanning = null
  }

  public update(item, current?) {
    if (item.isNote() || item.isAttachment()) return null

    current = current || this.keys.findOne({ itemID: item.id })

    const proposed = this.propose(item)

    if (current && (current.pinned || !Prefs.get('autoPin')) && (current.pinned === proposed.pinned) && (current.citekey === proposed.citekey)) return current.citekey

    if (current) {
      current.pinned = proposed.pinned
      current.citekey = proposed.citekey
      this.keys.update(current)
    } else {
      this.keys.insert({ itemID: item.id, libraryID: item.libraryID, itemKey: item.key, pinned: proposed.pinned, citekey: proposed.citekey })
    }

    return proposed.citekey
  }

  public remove(ids) {
     if (!Array.isArray(ids)) ids = [ids]

     this.keys.findAndRemove({ itemID : { $in : ids } })
   }

  public get(itemID) {
    // I cannot prevent being called before the init is done because Zotero unlocks the UI *way* before I'm getting the
    // go-ahead to *start* my init.
    if (!this.keys || !this.started) return { citekey: '', pinned: false, retry: true }

    const key = this.keys.findOne({ itemID })
    if (key) return key
    return { citekey: '', pinned: false, retry: true }
  }

  public propose(item) {
    const citekey: string = Extra.get(item.getField('extra'), 'zotero', { citationKey: true }).extraFields.citationKey

    if (citekey) return { citekey, pinned: true }

    const proposed = Formatter.format(item)

    const conflictQuery = { libraryID: item.libraryID, itemID: { $ne: item.id } }
    if (Prefs.get('keyScope') === 'global') delete conflictQuery.libraryID

    let postfix
    const seen = {}
    for (let n = proposed.postfix.start; true; n += 1) {
      if (n) {
        const alpha = intToExcelCol(n)
        postfix = sprintf(proposed.postfix.format, { a: alpha.toLowerCase(), A: alpha, n })
      } else {
        postfix = ''
      }

      // this should never happen, it'd mean the postfix pattern doesn't have placeholders, which should have been caught by parsePattern
      if (seen[postfix]) throw new Error(`${JSON.stringify(proposed.postfix)} does not generate unique postfixes`)
      seen[postfix] = true

      const postfixed = proposed.citekey + postfix

      const conflict = this.keys.findOne({ ...conflictQuery, citekey: postfixed })
      if (conflict) continue

      return { citekey: postfixed, pinned: false }
    }
  }

  public async tagDuplicates(libraryID) {
    const tag = '#duplicate-citation-key'
    const scope = Prefs.get('keyScope')

    const tagged = (await ZoteroDB.queryAsync(`
      SELECT items.itemID
      FROM items
      JOIN itemTags ON itemTags.itemID = items.itemID
      JOIN tags ON tags.tagID = itemTags.tagID
      WHERE (items.libraryID = ? OR 'global' = ?) AND tags.name = ? AND items.itemID NOT IN (select itemID from deletedItems)
    `, [ libraryID, scope, tag ])).map(item => item.itemID)

    const citekeys: {[key: string]: any[]} = {}
    for (const item of this.keys.find(scope === 'global' ? undefined : { libraryID })) {
      if (!citekeys[item.citekey]) citekeys[item.citekey] = []
      citekeys[item.citekey].push({ itemID: item.itemID, tagged: tagged.includes(item.itemID), duplicate: false })
      if (citekeys[item.citekey].length > 1) citekeys[item.citekey].forEach(i => i.duplicate = true)
    }

    const mistagged = Object.values(citekeys).reduce((acc, val) => acc.concat(val), []).filter(i => i.tagged !== i.duplicate).map(i => i.itemID)
    for (const item of await getItemsAsync(mistagged)) {
      if (tagged.includes(item.id)) {
        item.removeTag(tag)
      } else {
        item.addTag(tag)
      }

      await item.saveTx()
    }
  }

  private expandSelection(ids) {
    if (Array.isArray(ids)) return ids

    if (ids === 'selected') {
      try {
        return Zotero.getActiveZoteroPane().getSelectedItems(true)
      } catch (err) { // zoteroPane.getSelectedItems() doesn't test whether there's a selection and errors out if not
        log.error('Could not get selected items:', err)
        return []
      }
    }

    return [ids]
  }

  private bench(id) {
    if (typeof id === 'string') return { id, start: Date.now() }
  }
}
