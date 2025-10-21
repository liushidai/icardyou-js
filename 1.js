// ==UserScript==
// @name         icardyou-batch
// @namespace    https://github.com/liushidai/icardyou-js
// @version      2025-10-21-1.0
// @description  批量导出为excel、pdf;批量确认
// @author       liushidai
// @match        https://icardyou.icu/sendpostcard/myPostCard/1*
// @match        https://www.icardyou.icu/sendpostcard/myPostCard/1*
// @match        https://www.icardyou.com/sendpostcard/myPostCard/1*
// @match        https://icardyou.com/sendpostcard/myPostCard/1*
// @icon         data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==
// @grant        GM_addElement
// @grant        GM_log
// @require      https://code.jquery.com/jquery-3.6.0.min.js
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @run-at       document-body
// ==/UserScript==

(function () {
    'use strict';
    GM_log("Hello World");

    function addCheckbox() {
        // 每行添加多选框
        $('tbody tr:not(:first)').each(function () {
            var checkbox = $('<input>', {
                type: 'checkbox', class: 'row-checkbox'
            });
            $(this).prepend('<td></td>').find('td:first').append(checkbox);
        });

        // 为首行添加全选框并设置点击事件
        var allCheckbox = $('<input>', {
            type: 'checkbox', id: 'all-checkbox'
        });
        $('tbody tr:first').prepend('<th></th>').find('th:first').append(allCheckbox);

        // 全选框点击处理
        $('#all-checkbox').click(function () {
            var isChecked = $(this).is(':checked');
            $('.row-checkbox').prop('checked', isChecked);
        });

        // 行内多选框点击处理
        $('.row-checkbox').click(function () {
            var allChecked = $('.row-checkbox').length === $('.row-checkbox:checked').length;
            $('#all-checkbox').prop('checked', allChecked);
        });

    }

    /**
     * 添加导出选项框
     */
    function addExportModal() {
        if (document.getElementById('exportModal')) return;

        // 插入样式
        const style = document.createElement('style');
        style.textContent = `
        #exportModal {
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(0,0,0,0.5);
            display: none; /* 👈 默认隐藏！非常重要 */
            align-items: center;
            justify-content: center;
            z-index: 9999;
        }
        /* 其他样式保持不变 */
        #exportModal .modal-content {
            background: white;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            max-width: 300px;
            width: 90%;
        }
        #exportModal button {
            margin: 8px;
            padding: 8px 16px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
        }
        #exportModal #exportExcelBtn {
            background: #5cb85c;
            color: white;
        }
        #exportModal #exportPdfBtn {
            background: #d9534f;
            color: white;
        }
        #exportModal .close {
            position: absolute;
            top: 10px;
            right: 15px;
            font-size: 24px;
            cursor: pointer;
            color: #999;
        }
    `;
        document.head.appendChild(style);

        const modal = document.createElement('div');
        modal.id = 'exportModal';
        modal.innerHTML = `
        <div class="modal-content">
            <div class="close">&times;</div>
            <h4>选择导出格式</h4>
            <button id="exportExcelBtn">Excel</button>
            <button id="exportPdfBtn">PDF</button>
        </div>
    `;
        document.body.appendChild(modal);

        // 关闭逻辑
        modal.querySelector('.close').onclick = () => modal.style.display = 'none';
        modal.onclick = (e) => {
            if (e.target === modal) modal.style.display = 'none';
        };

        // 按钮逻辑
        modal.querySelector('#exportExcelBtn').onclick = () => {
            modal.style.display = 'none';
            if (window.__selectedRowDatas) {
                getAddressAndExportExcel(window.__selectedRowDatas);
                window.__selectedRowDatas = null;
            }
        };

        modal.querySelector('#exportPdfBtn').onclick = () => {
            modal.style.display = 'none';
            if (window.__selectedRowDatas) {
                const rowDatas = window.__selectedRowDatas;
                const resultArray = [];
                let completed = 0;

                if (rowDatas.length === 0) {
                    alert('无数据可导出');
                    return;
                }

                rowDatas.forEach(param => {
                    getAddress(param.id, function (response) {
                        resultArray.push({
                            cardType: param.cardType,
                            no: param.no,
                            status: param.status,
                            zip: response.zip || '',
                            realName: response.realName || '',
                            address: response.address || ''
                        });
                        completed++;
                        if (completed === rowDatas.length) {
                            resultArray.sort((a, b) => a.id - b.id);
                            previewForPrint(resultArray); // 👈 改为调用打印预览
                        }
                    });
                });
                window.__selectedRowDatas = null;
            }
        };
    }

    /**
     * 添加导出按钮 与 批量确认按钮
     */
    function addButton() {
        // 选择tbody中的第一行tr元素
        var firstRow = $('tbody tr:first');
        // 创建按钮元素
        var exportExcel = $('<button>', {
            text: '导出', class: 'exportExcel'
        });
        var batchVerify = $('<button>', {
            text: '批量确认', class: 'batchVerify'
        });

        // 为按钮1添加点击事件
        exportExcel.click(function () {
            // 获取表格数据
            var rowDatas = getTableData();
            // 临时存储
            window.__selectedRowDatas = rowDatas;
            document.getElementById('exportModal').style.display = 'flex';
        });
        // 为按钮2添加点击事件
        batchVerify.click(function () {
            // 获取表格数据
            var rowDatas = getTableData();
            // 调用接口确认
            batchConfirmCard(rowDatas);
        });

        // 将按钮添加到首行的最后一个单元格中
        var lastTd = firstRow.find('th:last');
        lastTd.append(exportExcel).append(batchVerify);
    }

    // 获取表格数据
    function getTableData() {
        let rowDatas = [];
        $("tbody tr:not(:first)").each(function () {
            var checkbox = $(this).find('input[type="checkbox"]');
            if (checkbox.is(':checked')) {
                // 1.获取卡片编号
                var cardNo = $(this).find('td').eq(2).find('a');
                // 2.获取id
                var cardId = cardNo.attr('href').split('/').pop();
                // 3.获取状态
                var cardStatus = $(this).find('td').eq(3).find('span');
                // 4.获取类型
                var cardTypeCell = $(this).find('td').eq(1);
                var cardType = '';
                var firstCellLink = cardTypeCell.find('a');
                if (firstCellLink.length > 0 && firstCellLink.text().trim() === '活动') {
                    // 是活动链接，提取 ID
                    var href = firstCellLink.attr('href'); // 如 "/games/detail/141823"
                    var match = href.match(/\/games\/detail\/(\d+)/);
                    cardType = match ? match[1] : '未知';
                } else {
                    // 非活动，直接取文本（如“配对”、“赠送”）
                    cardType = cardTypeCell.text().trim();
                }
                // 保存数据
                let rowData = {
                    'cardType': cardType, 'no': cardNo.text(), 'id': cardId, 'status': cardStatus.text()
                };
                rowDatas.push(rowData);
            }
        });
        if (!Array.isArray(rowDatas) || rowDatas.length === 0) {
            alert('至少勾选一行数据！');
            return;
        }
        return rowDatas;
    }

    // 获取地址数据并导出为excel表
    function getAddressAndExportExcel(rowDatas) {
        const resultArray = [];
        if (rowDatas.length === 0) {
            alert('至少勾选一行数据！');
            return;
        }
        // 3.请求接口
        rowDatas.forEach(function (param) {
            getAddress(param.id, function (response) {
                resultArray.push({
                    'cardType': param.cardType,
                    'no': param.no,
                    'id': param.id,
                    'status': param.status,
                    'zip': response.zip,
                    'realName': response.realName,
                    'address': response.address
                });
                if (resultArray.length === rowDatas.length) {
                    // 排序
                    resultArray.sort((a, b) => a.id - b.id);
                    console.log("所有数据获取完成，结果:", resultArray);
                    // 导出数据
                    download(resultArray)
                }
            });
        });
    }

    // 批量调用确认接口
    function batchConfirmCard(rowDatas) {
        const deferredArray = [];
        rowDatas.forEach((param) => {
            const dfd = $.Deferred();
            confirmCard(param.id).done(() => {
                dfd.resolve();
            }).fail(() => {
                dfd.reject();
            });
            deferredArray.push(dfd.promise());
        });

        $.when.apply($, deferredArray).done(() => {
            console.log('所有接口调用成功');
            alert('确认寄出成功-点击确认，刷新页面');
            location.reload();
        }).fail(() => {
            console.error('有接口调用出现错误');
            alert('存在错误-点击确认，刷新页面');
            location.reload();
        });
    }

    // 获取地址
    function getAddress(id, callback) {
        $.ajax({
            type: "POST", url: "/sendpostcard/findLostAddress", data: {id: id}, success: function (data) {
                callback(data);
            }
        });
    }

    // 确认寄出接口
    function confirmCard(id) {
        return $.ajax({
            type: "POST", url: "/sendpostcard/confirmSendCard", data: {id: id}, success: function (data) {
                console.log("id:" + id + ";data:" + data);
                // 假设根据返回数据中某个字段判断是否真正成功，这里只是示例
                if (data === 1) {
                    // 进行后续操作
                } else {
                    console.error('操作未成功');
                }
            }, error: function (xhr, status, error) {
                console.error('接口调用出现错误：', error);
            }
        });
    }

    /**
     * 导出excel逻辑
     * @param data
     */
    function download(data) {
        console.log(data);
        // 将数据转换为SheetJS需要的格式（二维数组）
        var worksheetData = XLSX.utils.json_to_sheet(data, {
            header: ["cardType", "id", "no", "status", "zip", "realName", "address"],
        });
        console.log(worksheetData);

        var workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheetData, "Sheet1");
        XLSX.writeFile(workbook, "icardyou.xlsx");

    }

    /**
     * 打印逻辑
     * @param data
     */
    function previewForPrint(data) {
        const printWin = window.open('', '_blank', 'width=900,height=700');
        const doc = printWin.document;
        let fontSize = 14;
        let cardWidth = 300; // 默认宽度 300px

        const updateContent = () => {
            doc.open();
            doc.write(`
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>明信片地址打印预览</title>
    <style>
        body {
            font-family: "Microsoft YaHei", sans-serif;
            font-size: ${fontSize}px;
            line-height: 1.4;
            padding: 10px;
            background: #f5f5f5;
        }
        .address-card {
            border: 1px dashed #333;
            padding: 12px;
            margin: 10px auto; /* 居中 */
            background: white;
            width: ${cardWidth}px; /* 👈 动态宽度 */
            page-break-inside: avoid;
            box-sizing: border-box;
        }
        .info-top {
            color: #888;
            font-size: 0.9em;
            margin-bottom: 8px;
            padding-bottom: 6px;
            border-bottom: 1px dashed #ccc;
        }
        .info-bottom {
            font-weight: bold;
        }
        .control-bar {
            position: fixed;
            top: 10px;
            right: 10px;
            z-index: 1000;
            background: white;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 6px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.1);
        }
        .control-bar button, .control-bar input {
            margin: 0 4px;
            padding: 4px 8px;
            font-size: 14px;
        }
        @media print {
            .control-bar { display: none !important; }
            body { background: white; padding: 0; }
            .address-card { margin: 5mm auto !important; }
        }
        pre {
            margin: 4px 0 !important;
            font-family: inherit !important;
            white-space: pre-wrap;
        }
    </style>
</head>
<body>
    <div class="control-bar">
        <button onclick="window.parent.setFontSize(${fontSize + 1})">A+</button>
        <button onclick="window.parent.setFontSize(${fontSize - 1})">A-</button>
        <button onclick="window.parent.setCardWidth(${cardWidth - 10})">←</button>
        <span id="width-display">${cardWidth}px</span>
        <button onclick="window.parent.setCardWidth(${cardWidth + 10})">→</button>
        <button onclick="window.print()">打印</button>
    </div>
`);

            data.forEach(item => {
                const top = `类型: ${item.cardType} | 状态: ${item.status}`;
                const bottomLines = [
                    `卡片编号: ${item.no || ''}`,
                    `邮编: ${item.zip || ''}`,
                    `姓名: ${item.realName || ''}`,
                    `地址: ${item.address || ''}`
                ].join('\n');

                doc.write(`
    <div class="address-card">
        <div class="info-top">${top}</div>
        <div class="info-bottom"><pre>${bottomLines}</pre></div>
    </div>
`);
            });

            doc.write(`
</body>
</html>
        `);
            doc.close();
        };

        updateContent();

        // 提供外部调用接口
        printWin.setFontSize = (size) => {
            if (size >= 8 && size <= 32) {
                fontSize = size;
                updateContent();
            }
        };

        printWin.setCardWidth = (width) => {
            // 限制合理范围，比如 150px ~ 600px
            cardWidth = Math.max(150, Math.min(600, width));
            updateContent();
            // 更新显示
            const display = printWin.document.getElementById('width-display');
            if (display) display.textContent = cardWidth + 'px';
        };
    }

    // 添加多选框
    addCheckbox();
    // 添加导出选项框
    addExportModal();
    // 添加操作按钮
    addButton();
    // Your code here...
})();
