// ==UserScript==
// @name         icardyou-batch
// @namespace    http://tampermonkey.net/
// @version      2024-12-14
// @description  批量导出
// @author       You
// @match        https://www.icardyou.icu/sendpostcard/myPostCard/1**
// @icon         data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==
// @grant        GM_addElement
// @grant        GM_log
// @require      https://code.jquery.com/jquery-3.6.0.min.js
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @run-at       document-body
// ==/UserScript==
(function() {
    'use strict';
    GM_log("Hello World");
    function addCheckbox(){
        // 每行添加多选框
        $('tbody tr:not(:first)').each(function () {
            var checkbox = $('<input>', {
                type: 'checkbox',
                class: 'row-checkbox'
            });
            $(this).prepend('<td></td>').find('td:first').append(checkbox);
        });

        // 为首行添加全选框并设置点击事件
        var allCheckbox = $('<input>', {
            type: 'checkbox',
            id: 'all-checkbox'
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
    function addButton(){
        // 选择tbody中的第一行tr元素
        var firstRow = $('tbody tr:first');
        // 创建按钮元素
        var exportExcel = $('<button>', {
            text: '导出',
            class: 'exportExcel'
        });
        var batchVerify = $('<button>', {
            text: '批量确认',
            class: 'batchVerify'
        });

        // 为按钮1添加点击事件
        exportExcel.click(function () {
            getTableData();
        });
        // 为按钮2添加点击事件
        batchVerify.click(function () {
            alert('未完工！');
        });

        // 将按钮添加到首行的最后一个单元格中
        var lastTd = firstRow.find('th:last');
        lastTd.append(exportExcel).append(batchVerify);
    }

    // 获取表格数据
    function getTableData() {
        let rowDatas = [];
        const resultArray = [];
        $("tbody tr:not(:first)").each(function() {
            var checkbox = $(this).find('input[type="checkbox"]');
            if (checkbox.is(':checked')) {
                // 1.获取卡片编号
                // 2.获取id
                var column3A = $(this).find('td').eq(2).find('a');
                var id = column3A.attr('href').split('/').pop();
                var column4S = $(this).find('td').eq(3).find('span');
                let rowData = {'no':column3A.text(),'id':id,'status':column4S.text()};
                rowDatas.push(rowData);
            }
        });
        if (rowDatas.length === 0) {
            alert('至少勾选一行数据！');
            return;
        }
        // 3.请求接口
        rowDatas.forEach(function (param) {
            getAddress(param.id, function (response) {
                resultArray.push({'no':param.no,'id':param.id,'status':param.status,'zip':response.zip,'realName': response.realName,'address':response.address});
                if (resultArray.length === rowDatas.length) {
                    console.log("所有数据获取完成，结果:", resultArray);
                    // 导出数据
                    download(resultArray)
                }
            });
        });
    }

    // 获取地址
    function getAddress(id,callback){
        $.ajax({
            type:"POST",
            url:"/sendpostcard/findLostAddress",
            data:{id:id},
            success:function(data){
                callback(data);
            }
        });
    }

    function download(data) {
        console.log(data);
        // 将数据转换为SheetJS需要的格式（二维数组）
        var worksheetData = XLSX.utils.json_to_sheet(data,{
            header: ["id","no","status","zip","realName","address"],
        });
        console.log(worksheetData);

        var workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheetData, "Sheet1");
        XLSX.writeFile(workbook, "icardyou.xlsx");

    }
    // 添加多选框
    addCheckbox();
    // 添加操作按钮
    addButton();
    // Your code here...
})();
