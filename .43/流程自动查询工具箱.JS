// ==UserScript==
// @name         流程自动查询工具箱
// @namespace    http://tampermonkey.net
// @version      2.0
// @description  重构
// @author       Adaptive Collaborator
// @include      *://192.168.0.43*
// @grant        GM_registerMenuCommand
// @run-at       document-start
// ==/UserScript==

;(function () {
  'use strict'
  if (window.top !== window.self) return
  let running = false

  const CONFIG = {
    unplanned: {
      title: '未策划工作包统计表',
      nav: '未策划工作项统计表',
      dateFields: ['$FR$PLANFINISHDATE', '$TO$PLANFINISHDATE'],
      majorName: '$IN$MAJORCODE',
      fallbackMajor: true,
      majorVal: '230n,241n,243n',
      majorText: '电力(能环),自动化(能环),通讯(能环)'
    },
    detail: {
      title: '专业进度检查统计表',
      nav: '工作项进度检查明细表',
      dateFields: ['$FR$PLANENDDATE', '$TO$PLANENDDATE'],
      majorName: '$IN$WBSVALUE',
      fallbackMajor: false,
      majorVal: 'a72000f3-e4e5-4106-a1d9-90c17078b893:230n,a72000f3-e4e5-4106-a1d9-90c17078b893:241n,a72000f3-e4e5-4106-a1d9-90c17078b893:243n',
      majorText: '电力(能环),自动化(能环),通讯(能环)'
    }
  }

  GM_registerMenuCommand('⚡ 1. 未策划工作项查询', () => startAutomation('unplanned'))
  GM_registerMenuCommand('⚡ 2. 工作项进度明细查询', () => startAutomation('detail'))

  const sleep = t => new Promise(r => setTimeout(r, t)),
        fmt = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`

  async function startAutomation(type) {
    if (running) return
    const input = prompt('范围：1=本周，2=本月', '1')
    if (!input) return
    running = true

    try {
      const cfg = CONFIG[type], mainDoc = window.top.document
      let tw = findWindowByTitle(cfg.title)

      // 1. 导航与激活
      const tabBtn = mainDoc.getElementById(cfg.nav) ||
                     [...mainDoc.querySelectorAll('.mini-tab-text, .mini-button-text')].find(e => e.innerText.trim() === cfg.nav)

      if (tw || tabBtn) {
        tabBtn?.click(); await sleep(500)
      } else {
        const click = t => [...mainDoc.querySelectorAll('li, a, .mini-button-text')].find(e => e.innerText.trim() === t)?.click()
        click('查询统计'); await sleep(800)
        click('设计管理统计查询'); await sleep(500)
        click('进度检查'); await sleep(500)
        click(cfg.nav)
      }

      for (let i = 0; i < 15; i++) {
        tw = findWindowByTitle(cfg.title)
        if (tw?.mini?.getbyName(cfg.dateFields[0])) break
        await sleep(1000)
      }

      if (!tw?.mini) return alert(`❌ 无法定位页面: ${cfg.title}`)

      // 2. 填值
      const mini = tw.mini, now = new Date(), first = fmt(new Date(now.getFullYear(), now.getMonth(), 1)), end = new Date(now)
      input === '2' ? end.setMonth(now.getMonth() + 1, 0) : end.setDate(now.getDate() + (now.getDay() === 0 ? 0 : 7 - now.getDay()))

      mini.getbyName(cfg.dateFields[0])?.setValue(first)
      mini.getbyName(cfg.dateFields[1])?.setValue(fmt(end))

      const major = mini.getbyName(cfg.majorName) || (cfg.fallbackMajor ? mini.findControls(c => c.type === 'buttonedit')[0] : null)
      if (major) {
        major.setValue(cfg.majorVal); major.setText(cfg.majorText); major.doValueChanged?.()
      }

      await sleep(300)

      // 3. 监控加载并自动分页
      const grid = mini.findControls(c => c.type === 'datagrid')[0]
      if (grid && !grid._hasMonitored) {
        grid._hasMonitored = true
        grid.on('load', () => {
          const total = grid.getTotalCount(), size = grid.getPageSize(),
                target = [50, 100, 200, 300, 500].find(s => s >= total) || 500
          console.log(`%c✅ 同步成功 | 总数: ${total} | 当前页: ${size}`, 'color: #4CAF50; font-weight: bold;')
          if (total > size && size < target) {
            console.log(`🔄 扩容至 -> ${target}`); grid.setPageSize(target); grid.reload()
          }
        })
        grid.on('beforeload', () => console.log('%c📡 数据传输中...', 'color: #2196F3;'))
      }

      // 4. 发起查询
      console.log('🚀 发出查询指令...')
      if (typeof tw.search === 'function') tw.search()
      else (tw.document.querySelector('.icon-search') || tw.document.querySelector('.mini-button'))?.click()

    } finally { running = false }
  }

  function findWindowByTitle(title) {
    const scan = win => {
      try {
        if (win.document?.title === title) return win
        for (let i = 0; i < (win.frames?.length || 0); i++) {
          const f = scan(win.frames[i])
          if (f) return f
        }
      } catch (e) {}
      return null
    }
    return scan(window.top)
  }
})()
