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
            // 获取表格数据
            var rowDatas= getTableData();
            // 调用接口导出
            getAddressAndExportExcel(rowDatas);
        });
        // 为按钮2添加点击事件
        batchVerify.click(function () {
            // 获取表格数据
            var rowDatas= getTableData();
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
        return rowDatas;
    }
    // 获取表格数据
    function getAddressAndExportExcel(rowDatas) {
        const resultArray = [];
        if (rowDatas.length === 0) {
            alert('至少勾选一行数据！');
            return;
        }
        // 3.请求接口
        rowDatas.forEach(function (param) {
            getAddress(param.id, function (response) {
                resultArray.push({'no':param.no,'id':param.id,'status':param.status,'zip':response.zip,'realName': response.realName,'address':response.address});
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
    // 确认寄出接口
    function confirmCard(id){
        return $.ajax({
            type:"POST",
            url:"/sendpostcard/confirmSendCard",
            data:{id:id},
            success:function(data){
                console.log("id:"+id+";data:" +data);
                // 假设根据返回数据中某个字段判断是否真正成功，这里只是示例
                if (data == 1) {
                    // 进行后续操作
                } else {
                    console.error('操作未成功');
                }
            },
            error:function(xhr, status, error) {
                console.error('接口调用出现错误：', error);
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
