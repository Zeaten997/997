// ==UserScript==
// @name         流程自动查询工具箱
// @namespace    http://tampermonkey.net
// @version      1.2
// @description  增加时间范围选择，修复重复弹窗，并确保查询页签切换到前台
// @author       YourName
// @include      *://192.168.0.43*
// @include      http://192.168.0.43*
// @grant        GM_registerMenuCommand
// @run-at       document-start
// ==/UserScript==

;(function () {
  'use strict'

  if (window.top !== window.self) return

  let running = false

  GM_registerMenuCommand('⚡ 1. 未策划工作项查询', () =>
    startAutomation('unplanned'),
  )
  GM_registerMenuCommand('⚡ 2. 工作项进度明细查询', () =>
    startAutomation('detail'),
  )

  const sleep = (t) => new Promise((r) => setTimeout(r, t))
  const fmt = (d) =>
    `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`

  function getTabIdByType(type) {
    return type === 'unplanned' ? '未策划工作项统计表' : '工作项进度检查明细表'
  }

  // 核心修复：根据类型返回唯一的特征控件名称
  function getFeatureName(type) {
    return type === 'unplanned' ? '$FR$PLANFINISHDATE' : '$IN$WBSVALUE'
  }

  async function activateTabFront(type) {
    const mainDoc = window.top.document
    const tabId = getTabIdByType(type)
    const tab = mainDoc.getElementById(tabId)
    if (tab) {
      console.log(`📌 激活现有页签: ${tabId}`)
      tab.click()
      await sleep(1000) // 给 iframe 留出显示和响应时间
      return true
    }
    return false
  }

  function getDateRange(mode) {
    const now = new Date()
    const firstDayDate = new Date(now.getFullYear(), now.getMonth(), 1)
    let endDate
    if (mode === 'month') {
      endDate = new Date(now.getFullYear(), now.getMonth() + 1, 0)
    } else {
      endDate = new Date(
        now.getTime() + (now.getDay() === 0 ? 0 : 7 - now.getDay()) * 86400000,
      )
    }
    return { firstDay: fmt(firstDayDate), endDay: fmt(endDate) }
  }

  async function startAutomation(type) {
    if (running) {
      alert('⏳ 当前已有查询流程正在执行')
      return
    }

    const dateMode = (() => {
      const input = prompt(
        '请选择查询时间范围：\n\n1 = 查询至本周\n2 = 查询本月',
        '1',
      )
      if (input === '1') return 'week'
      if (input === '2') return 'month'
      return null
    })()

    if (!dateMode) return
    running = true

    try {
      console.log(`🚀 启动 [${getTabIdByType(type)}] 流程...`)
      const mainDoc = window.top.document

      // 1. 尝试直接激活已存在的页签
      await activateTabFront(type)

      // 2. 轮询查找目标窗口 (增加重试)
      let tw = null
      for (let i = 0; i < 3; i++) {
        tw = findWindowByFeature(type)
        if (tw) break
        await sleep(500)
      }

      // 3. 如果找不到，才执行导航
      if (!tw) {
        console.log('📍 未发现目标页面，开始执行左侧菜单导航...')
        mainDoc.getElementById('a6d300dd-de02-497a-bd1e-8e810aead9ad')?.click()
        await sleep(800)

        const clickByText = (t) =>
          Array.from(mainDoc.querySelectorAll('a'))
            .find((a) => a.innerText.trim() === t)
            ?.click()

        clickByText('设计管理统计查询')
        await sleep(500)
        clickByText('进度检查')
        await sleep(500)

        const tabId = getTabIdByType(type)
        mainDoc.getElementById(tabId)?.click()

        // 等待新页面加载
        for (let i = 0; i < 15; i++) {
          await sleep(1000)
          tw = findWindowByFeature(type)
          if (tw) break
        }
      }

      if (!tw) {
        alert('❌ 无法定位查询页面，请手动打开页签后重试')
        return
      }

      // --- 执行具体业务逻辑 ---
      const mini = tw.mini
      const { firstDay, endDay } = getDateRange(dateMode)

      if (type === 'unplanned') {
        mini.getbyName('$FR$PLANFINISHDATE')?.setValue(firstDay)
        mini.getbyName('$TO$PLANFINISHDATE')?.setValue(endDay)

        // 弹出专业选择树
        const majorComp = mini
          .getComponents()
          .find((c) => c.type === 'buttonedit')
        if (majorComp) {
          majorComp.showPopup
            ? majorComp.showPopup()
            : majorComp.fire?.('buttonclick')
        }

        await sleep(2000)
        // 查找弹出的 Tree 窗口
        let treeWin = null
        const findTree = (win) => {
          try {
            if (win.document?.querySelector?.('#dataTree') && win.mini) {
              treeWin = win
              return
            }
            for (let i = 0; i < (win.frames?.length || 0); i++) {
              findTree(win.frames[i])
              if (treeWin) return
            }
          } catch (e) {}
        }
        findTree(window.top)

        if (treeWin) {
          const tree = treeWin.mini.get('dataTree')
          const targetNames = ['电力(能环)', '自动化(能环)', '通讯(能环)']
          const list = tree.getList?.() || tree.getData?.() || []
          list
            .filter((n) => targetNames.includes(n.NAME || n.text))
            .forEach((node) => tree.checkNode(node))

          const selBtn = Array.from(
            treeWin.document.querySelectorAll('.mini-button-text'),
          ).find((e) => e.innerText.trim() === '选择')
          selBtn ? selBtn.click() : treeWin.returnValue?.()
          await sleep(1000)
        }
      } else {
        // 明细表逻辑
        mini.getbyName('$FR$PLANENDDATE')?.setValue(firstDay)
        mini.getbyName('$TO$PLANENDDATE')?.setValue(endDay)
        let major = mini.getbyName('$IN$WBSVALUE')
        if (major) {
          major.setValue(
            'a72000f3-e4e5-4106-a1d9-90c17078b893:230n,a72000f3-e4e5-4106-a1d9-90c17078b893:241n,a72000f3-e4e5-4106-a1d9-90c17078b893:243n',
          )
          major.setText('电力(能环),自动化(能环),通讯(能环)')
          major.fire?.('valuechanged')
        }
      }

      await sleep(500)
      if (typeof tw.search === 'function') {
        tw.search()
      } else {
        tw.document.querySelector('.icon-search, .mini-button')?.click()
      }
      console.log('🏁 查询指令已发出')
    } finally {
      running = false
    }
  }

  function findWindowByFeature(type) {
    const feature = getFeatureName(type)
    function scan(win) {
      try {
        // 校验该 window 是否包含当前业务特有的 mini 控件
        if (win.mini && typeof win.mini.getbyName === 'function') {
          if (win.mini.getbyName(feature)) return win
        }
        for (let i = 0; i < win.frames.length; i++) {
          const f = scan(win.frames[i])
          if (f) return f
        }
      } catch (e) {}
      return null
    }
    return scan(window.top)
  }
})()
