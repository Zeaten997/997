// ==UserScript==
// @name         流程自动查询工具箱
// @namespace    http://tampermonkey.net
// @version      1.0
// @description  解决第二查询仍触发导航的问题；优化iframe识别逻辑
// @author       YourName
// @include      *://192.168.0.43*
// @include      http://192.168.0.43*
// @grant        GM_registerMenuCommand
// @run-at       document-start
// ==/UserScript==

(function() {
    'use strict';

    GM_registerMenuCommand("⚡ 1. 未策划工作项查询", () => startAutomation('unplanned'));
    GM_registerMenuCommand("⚡ 2. 工作项进度明细查询", () => startAutomation('detail'));

    const sleep = (t) => new Promise(r => setTimeout(r, t));
    const fmt = (d) => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;

    async function startAutomation(type) {
        console.log(`🚀 启动 [${type === 'unplanned' ? '未策划' : '进度明细'}] 流程...`);
        const mainDoc = window.top.document;

        // --- 优化1: 双重检查 (检查iframe内容 OR 检查页签是否已存在) ---
        let tw = findWindowByFeature(type);

        // 如果iframe没搜到，检查是否有现成的页签可以点
        if (!tw) {
            const tabName = type === 'unplanned' ? "未策划工作项统计表" : "工作项进度检查明细表";
            const existingTab = mainDoc.getElementById(tabName);
            if (existingTab) {
                console.log("📍 检测到页签已打开，尝试激活...");
                existingTab.click();
                await sleep(1000);
                tw = findWindowByFeature(type); // 激活后再搜一次
            }
        }

        // --- 优化2: 只有完全找不到才执行左侧菜单导航 ---
        if (!tw) {
            console.log("📍 全局未匹配到页面，执行完整导航...");
            mainDoc.getElementById("a6d300dd-de02-497a-bd1e-8e810aead9ad")?.click();
            await sleep(800);
            const clickByText = (t) => Array.from(mainDoc.querySelectorAll('a')).find(a => a.innerText.trim() === t)?.click();
            clickByText("设计管理统计查询"); await sleep(500);
            clickByText("进度检查"); await sleep(500);
            const tabId = type === 'unplanned' ? "未策划工作项统计表" : "工作项进度检查明细表";
            mainDoc.getElementById(tabId)?.click();

            // 轮询等待加载
            for(let i=0; i<10; i++) {
                await sleep(1000);
                tw = findWindowByFeature(type);
                if (tw) break;
            }
        }

        if (!tw) return alert("❌ 无法正确定位页面，请确保页面已完全加载");

        // --- 逻辑执行部分 (保持原代码功能逻辑) ---
        const mini = tw.mini;
        const now = new Date();
        const firstDay = fmt(new Date(now.getFullYear(), now.getMonth(), 1));
        const sunday = new Date(now.getTime() + (now.getDay() === 0 ? 0 : 7 - now.getDay()) * 86400000);

        if (type === 'unplanned') {
            mini.getbyName("$FR$PLANFINISHDATE")?.setValue(firstDay);
            mini.getbyName("$TO$PLANFINISHDATE")?.setValue(fmt(sunday));

            const majorComp = mini.getComponents().find(c => c.type === 'buttonedit');
            if (majorComp) (majorComp.showPopup || majorComp.fire?.('buttonclick'))?.();
            await sleep(2000);

            let treeWin = null;
            const findTree = (win) => {
                try {
                    if (win.document?.querySelector?.('#dataTree') && win.mini) treeWin = win;
                    for (let i = 0; i < win.frames?.length; i++) findTree(win.frames[i]);
                } catch (e) {}
            };
            findTree(window.top);

            if (treeWin) {
                const targetNames = ['电力(能环)', '自动化(能环)', '通讯(能环)'];
                const tree = treeWin.mini.get('dataTree');
                const list = tree.getList?.() || tree.getData?.() || [];
                list.filter(n => targetNames.includes(n.NAME || n.text || n.name)).forEach(node => tree.checkNode(node));
                if (typeof treeWin.returnValue === 'function') treeWin.returnValue();
                else {
                    const selBtn = Array.from(treeWin.document.querySelectorAll('.mini-button-text')).find(e => e.innerText.trim() === '选择');
                    selBtn?.click();
                }
                await sleep(1000);
            }
        } else {
            mini.getbyName("$FR$PLANENDDATE")?.setValue(firstDay);
            mini.getbyName("$TO$PLANENDDATE")?.setValue(fmt(sunday));
            let major = mini.getbyName("$IN$WBSVALUE");
            if (major) {
                major.setValue("a72000f3-e4e5-4106-a1d9-90c17078b893:230n,a72000f3-e4e5-4106-a1d9-90c17078b893:241n,a72000f3-e4e5-4106-a1d9-90c17078b893:243n");
                major.setText("电力(能环),自动化(能环),通讯(能环)");
                if (major.fire) major.fire("valuechanged");
            }
        }

        await sleep(500);
        if (typeof tw.search === 'function') tw.search();
        else (tw.document.querySelector(".icon-search, .mini-button"))?.click();
        console.log("🏁 自动化查询闭环完成");
    }

    /**
     * 深度扫描 iframe，识别 MiniUI 表单窗口
     */
    function findWindowByFeature(type) {
        // 关键：未策划识别结束日期，明细识别专业ID框
        const feature = type === 'unplanned' ? "$TO$PLANFINISHDATE" : "$IN$WBSVALUE";
        function scan(win) {
            try {
                // 不仅检查 mini 对象，还要确保 mini.getbyName 能拿到东西
                if (win.mini && typeof win.mini.getbyName === 'function' && win.mini.getbyName(feature)) {
                    return win;
                }
                for (let i = 0; i < win.frames.length; i++) {
                    const f = scan(win.frames[i]);
                    if (f) return f;
                }
            } catch(e) {}
            return null;
        }
        return scan(window.top);
    }
})();
