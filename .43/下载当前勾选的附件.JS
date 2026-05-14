// ==UserScript==
// @name         下载当前勾选的附件
// @namespace    http://tampermonkey.net
// @version      1.1
// @description  自动扫描页面所有 MiniUI 表格行，批量静默穿透下载业务附件
// @author       Assistant
// @match        http://192.168.0.43:8090/PortalLTE/Main/Index
// @grant        GM_registerMenuCommand
// ==/UserScript==

(function() {
    'use strict';

    const ALLOWED_TITLES = ["文印委托出版申请", "线上项目文印委托查询"];
    const isTargetPage = (title) => title && ALLOWED_TITLES.some(t => title.includes(t));

    // 1. 递归查找符合白名单且含有 dataGrid 的窗口
    function findTargetWindows(win, list = []) {
        try {
            if (win.mini?.get("dataGrid") && isTargetPage(win.document.title)) list.push(win);
            for (let i = 0; i < win.frames.length; i++) findTargetWindows(win.frames[i], list);
        } catch {}
        return list;
    }

    // 2. 动态静默解锁多选
    setInterval(() => {
        findTargetWindows(window).forEach(win => {
            const grid = win.mini.get("dataGrid");
            if (!grid || grid.customUnlocked) return;

            grid.setMultiSelect(true);
            const columns = grid.getColumns() || grid.columns || [];

            if (!columns.some(c => c.type === "checkcolumn")) {
                columns.unshift({ type: "checkcolumn", width: 40, headerAlign: "center", align: "center" });
                win.setTimeout(() => grid.setColumns?.(columns) || (grid.columns = columns), 0);
            }

            grid.customUnlocked = true;
            console.log(`🔓 [油猴脚本] 成功匹配并解锁页面 "${win.document.title}"`);

            win.setTimeout(() => (grid.reload || grid.load || grid.refresh)?.call(grid), 50);
        });
    }, 1500);

    // 3. 勾选行批量穿透下载
    GM_registerMenuCommand("📥 下载附件", () => {
        const targetWin = findTargetWindows(window)[0];
        if (!targetWin) return alert("❌ 未能在白名单页面内找到表格组件，请确保页面已加载完毕。");

        const grid = targetWin.mini.get("dataGrid");
        const rows = grid.getSelecteds?.() || grid.getCheckedRows?.() || [];
        if (!rows.length) return alert(`⚠️ 请先在当前页面（${targetWin.document.title}）勾选你想要下载的行！`);
        if (!confirm(`确认要依次下载当前已勾选的 ${rows.length} 条数据的附件吗？`)) return;

        // 异步串行下载封装
        const request = (url, params) => new Promise(resolve => targetWin.execute(url, { ...params, onComplete: resolve }));
        const delay = (ms) => new Promise(r => setTimeout(r, ms));

        (async function downloadLoop() {
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                if (!row.ID) continue;

                console.log(`📡 [${i + 1}/${rows.length}] 正在处理: ${row.Name || row.Title || row.PrjName || `第${i + 1}行`}`);
                try {
                    targetWin.addExecuteParam("Flag", 1);
                    targetWin.addExecuteParam("ID", row.ID);

                    const res = await request("/Design/AutoUI/PrjMakeBlueprint/GetProductAttachment", { async: true, validateForm: false, showLoading: false });
                    if (res?.Attachment) {
                        targetWin.addExecuteParam("fileIds", res.Attachment);
                        const batchRes = await request("/Design/AutoUI/PrjMakeBlueprint/AddTempBatch", { async: true, validateForm: false, showLoading: false });
                        if (batchRes?.TempBatchID) window.open(`/FileStore/Download.aspx?TempBatchID=${batchRes.TempBatchID}`);
                        await delay(1500);
                    } else {
                        await delay(1000);
                    }
                } catch (err) {
                    console.error("❌ 请求网络出错:", err);
                    await delay(1500);
                }
            }
            alert(`✅ 批量下载任务派发完毕！共成功处理 ${rows.length} 条勾选数据。`);
        })();
    });
})();
