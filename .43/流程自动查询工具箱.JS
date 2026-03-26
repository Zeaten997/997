// ==UserScript==
// @name         流程自动查询工具箱1.1
// @namespace    http://tampermonkey.net
// @version      1.1
// @description  增加时间范围选择，修复重复弹窗，并确保查询页签切换到前台
// @author       YourName
// @include      *://192.168.0.43*
// @include      http://192.168.0.43*
// @grant        GM_registerMenuCommand
// @run-at       document-start
// ==/UserScript==

(function() {
    'use strict';

    // 只在顶层窗口执行，避免 iframe 中重复注册菜单、重复弹窗、重复执行
    if (window.top !== window.self) return;

    // 防止重复并发执行
    let running = false;

    GM_registerMenuCommand("⚡ 1. 未策划工作项查询", () => startAutomation('unplanned'));
    GM_registerMenuCommand("⚡ 2. 工作项进度明细查询", () => startAutomation('detail'));

    const sleep = (t) => new Promise(r => setTimeout(r, t));
    const fmt = (d) => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;

    function getTabIdByType(type) {
        return type === 'unplanned' ? "未策划工作项统计表" : "工作项进度检查明细表";
    }

    async function activateTabFront(type) {
        const mainDoc = window.top.document;
        const tabId = getTabIdByType(type);
        const tab = mainDoc.getElementById(tabId);
        if (tab) {
            console.log(`📌 切换前台页签: ${tabId}`);
            tab.click();
            await sleep(800);
        }
    }

    function chooseDateMode() {
        const input = prompt(
            "请选择查询时间范围：\n\n1 = 查询至本周\n2 = 查询本月",
            "1"
        );

        if (input === null) return null;

        const val = input.trim();
        if (val === '1') return 'week';
        if (val === '2') return 'month';

        alert("❌ 输入无效，请输入 1 或 2");
        return null;
    }

    function getDateRange(mode) {
        const now = new Date();
        const firstDayDate = new Date(now.getFullYear(), now.getMonth(), 1);
        let endDate;

        if (mode === 'month') {
            // 本月最后一天
            endDate = new Date(now.getFullYear(), now.getMonth() + 1, 0);
        } else {
            // 保持原来的算法：本周日
            endDate = new Date(
                now.getTime() + (now.getDay() === 0 ? 0 : 7 - now.getDay()) * 86400000
            );
        }

        return {
            firstDay: fmt(firstDayDate),
            endDay: fmt(endDate)
        };
    }

    async function startAutomation(type) {
        if (running) {
            alert("⏳ 当前已有查询流程正在执行，请稍候再试");
            return;
        }

        running = true;

        try {
            const dateMode = chooseDateMode();
            if (!dateMode) return;

            console.log(`🚀 启动 [${type === 'unplanned' ? '未策划' : '进度明细'}] 流程...`);
            console.log(`🗓️ 查询模式：${dateMode === 'week' ? '本月1号至本周日' : '本月1号至本月最后一天'}`);

            const mainDoc = window.top.document;

            // 先深度扫描 iframe，查找目标页面
            let tw = findWindowByFeature(type);

            // 如果已经找到了目标页面，则直接把对应页签切到前台，不重新导航
            if (tw) {
                await activateTabFront(type);
                tw = findWindowByFeature(type); // 前台切换后重新确认一次
            }

            // 如果 iframe 没搜到，再检查是否有现成页签可激活
            if (!tw) {
                const tabId = getTabIdByType(type);
                const existingTab = mainDoc.getElementById(tabId);
                if (existingTab) {
                    console.log("📍 检测到页签已打开，尝试激活...");
                    existingTab.click();
                    await sleep(1000);
                    tw = findWindowByFeature(type);
                }
            }

            // 只有完全找不到才执行左侧菜单导航
            if (!tw) {
                console.log("📍 全局未匹配到页面，执行完整导航...");
                mainDoc.getElementById("a6d300dd-de02-497a-bd1e-8e810aead9ad")?.click();
                await sleep(800);

                const clickByText = (t) =>
                    Array.from(mainDoc.querySelectorAll('a'))
                        .find(a => a.innerText.trim() === t)?.click();

                clickByText("设计管理统计查询");
                await sleep(500);
                clickByText("进度检查");
                await sleep(500);

                const tabId = getTabIdByType(type);
                mainDoc.getElementById(tabId)?.click();

                // 轮询等待加载
                for (let i = 0; i < 10; i++) {
                    await sleep(1000);
                    tw = findWindowByFeature(type);
                    if (tw) break;
                }
            }

            if (!tw) {
                alert("❌ 无法正确定位页面，请确保页面已完全加载");
                return;
            }

            // 再次确保当前查询对应的页签显示在前台
            await activateTabFront(type);
            tw = findWindowByFeature(type) || tw;

            const mini = tw.mini;
            const { firstDay, endDay } = getDateRange(dateMode);

            // --- 逻辑执行部分 ---
            if (type === 'unplanned') {
                mini.getbyName("$FR$PLANFINISHDATE")?.setValue(firstDay);
                mini.getbyName("$TO$PLANFINISHDATE")?.setValue(endDay);

                const majorComp = mini.getComponents().find(c => c.type === 'buttonedit');
                if (majorComp) {
                    if (typeof majorComp.showPopup === 'function') {
                        majorComp.showPopup();
                    } else if (typeof majorComp.fire === 'function') {
                        majorComp.fire('buttonclick');
                    }
                }

                await sleep(2000);

                let treeWin = null;
                const findTree = (win) => {
                    try {
                        if (win.document?.querySelector?.('#dataTree') && win.mini) {
                            treeWin = win;
                            return;
                        }
                        for (let i = 0; i < (win.frames?.length || 0); i++) {
                            findTree(win.frames[i]);
                            if (treeWin) return;
                        }
                    } catch (e) {}
                };
                findTree(window.top);

                if (treeWin) {
                    const targetNames = ['电力(能环)', '自动化(能环)', '通讯(能环)'];
                    const tree = treeWin.mini.get('dataTree');
                    const list = tree.getList?.() || tree.getData?.() || [];

                    list
                        .filter(n => targetNames.includes(n.NAME || n.text || n.name))
                        .forEach(node => tree.checkNode(node));

                    if (typeof treeWin.returnValue === 'function') {
                        treeWin.returnValue();
                    } else {
                        const selBtn = Array.from(treeWin.document.querySelectorAll('.mini-button-text'))
                            .find(e => e.innerText.trim() === '选择');
                        selBtn?.click();
                    }

                    await sleep(1000);
                }
            } else {
                mini.getbyName("$FR$PLANENDDATE")?.setValue(firstDay);
                mini.getbyName("$TO$PLANENDDATE")?.setValue(endDay);

                let major = mini.getbyName("$IN$WBSVALUE");
                if (major) {
                    major.setValue("a72000f3-e4e5-4106-a1d9-90c17078b893:230n,a72000f3-e4e5-4106-a1d9-90c17078b893:241n,a72000f3-e4e5-4106-a1d9-90c17078b893:243n");
                    major.setText("电力(能环),自动化(能环),通讯(能环)");
                    if (major.fire) major.fire("valuechanged");
                }
            }

            await sleep(500);

            if (typeof tw.search === 'function') {
                tw.search();
            } else {
                (tw.document.querySelector(".icon-search, .mini-button"))?.click();
            }

            console.log("🏁 自动化查询闭环完成");
        } finally {
            running = false;
        }
    }

    /**
     * 深度扫描 iframe，识别 MiniUI 表单窗口
     */
    function findWindowByFeature(type) {
        const feature = type === 'unplanned' ? "$TO$PLANFINISHDATE" : "$IN$WBSVALUE";

        function scan(win) {
            try {
                if (win.mini && typeof win.mini.getbyName === 'function' && win.mini.getbyName(feature)) {
                    return win;
                }
                for (let i = 0; i < win.frames.length; i++) {
                    const f = scan(win.frames[i]);
                    if (f) return f;
                }
            } catch (e) {}
            return null;
        }

        return scan(window.top);
    }
})();
