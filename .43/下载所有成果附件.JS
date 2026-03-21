// ==UserScript==
// @name         下载所有成果附件
// @namespace    http://tampermonkey.net
// @version      1.0
// @description  静默穿透下载业务附件
// @author       Assistant
// @match        http://192.168.0.43:8090/PortalLTE/Main/Index
// @grant        GM_registerMenuCommand
// ==/UserScript==

(function() {
    'use strict';

    // 注册油猴菜单
    GM_registerMenuCommand("📥 下载", function() {
        executeSilentDownload();
    });

    function executeSilentDownload() {
        console.log('🤫 启动静默穿透下载...');

        // 1. 跨框架寻找选中的行
        function getSelectedData(win) {
            try {
                // 尝试从 MiniUI 的 grid 对象直接获取数据
                const grids = win.mini ? win.mini.findControls(c => c.type === 'datagrid') : [];
                for (let grid of grids) {
                    const row = grid.getSelected();
                    // 这里的 row.ID 对应你之前提到的 UUID
                    if (row && row.ID) return { id: row.ID, win: win };
                }

                // 兜底：从 DOM 选中样式解析
                let selectedRow = win.document.querySelector('.mini-grid-row-selected');
                if (selectedRow) {
                    const link = selectedRow.querySelector('a[onclick*="onGridButtonClick"]');
                    // 正则提取 onclick 中的第三个参数 ID
                    const match = link?.getAttribute('onclick')?.match(/onGridButtonClick\(this,\d+,'[^']+',\s*'([^']+)'\)/);
                    if (match && match[1]) return { id: match[1], win: win };
                }

                // 递归进入子框架
                for (let i = 0; i < win.frames.length; i++) {
                    let res = getSelectedData(win.frames[i]);
                    if (res) return res;
                }
            } catch (e) {}
            return null;
        }

        const data = getSelectedData(window);
        if (!data) {
            alert('❌ 未找到选中行！请先在页面上手动点选一个小圆点。');
            return;
        }

        const targetID = data.id;
        const targetWin = data.win;
        console.log('🆔 提取到业务 ID:', targetID);

        try {
            console.log('📡 正在后台请求下载令牌...');

            // 步骤 A: 获取附件列表
            targetWin.addExecuteParam("Flag", 1);
            targetWin.addExecuteParam("ID", targetID);

            targetWin.execute("/Design/AutoUI/PrjMakeBlueprint/GetProductAttachment", {
                async: false,
                validateForm: false,
                showLoading: true,
                onComplete: function (res) {
                    if (res && res.Attachment) {
                        console.log('📎 找到附件，换取 TempBatchID...');

                        // 步骤 B: 换取 TempBatchID
                        targetWin.addExecuteParam("fileIds", res.Attachment);
                        targetWin.execute("/Design/AutoUI/PrjMakeBlueprint/AddTempBatch", {
                            async: false,
                            validateForm: false,
                            showLoading: true,
                            onComplete: function (batchRes) {
                                if (batchRes && batchRes.TempBatchID) {
                                    const downloadUrl = "/FileStore/Download.aspx?TempBatchID=" + batchRes.TempBatchID;
                                    console.log('🎯 穿透成功！下载地址:', downloadUrl);
                                    window.open(downloadUrl);
                                } else {
                                    alert('❌ 获取 TempBatchID 失败');
                                }
                            }
                        });
                    } else {
                        alert("⚠️ 该行数据没有附件。");
                    }
                }
            });
        } catch (err) {
            console.error('❌ 执行失败:', err);
            alert('运行出错，请查看控制台日志。');
        }
    }
})();
 //批量化潜力：基于现在的逻辑，如果你需要一次下载多行，只需要把 getSelectedData 改为循环获取所有选中的行，并依次执行后台请求即可。
