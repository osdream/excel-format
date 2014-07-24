/***************************************************************************
 * 
 * Copyright (c) 2014 Baidu.com, Inc. All Rights Reserved
 * $Id$
 * 
 **************************************************************************/
 
 
/*
 * path:    parse.js
 * desc:    
 * author:  songao(songao@baidu.com)
 * version: $Revision$
 * date:    $Date: 2014/07/23 11:47:37$
 */

var fs = require('fs');
var ejs = require('ejs');
var mustache = require('mustache');
var xlsx = require('node-xlsx');

/**
 * 将十进制的列号转换为excel的AB表示方法
 * @param {number} number
 * @return {string}
 */
function toABCode(number) {
    var iAlpha = parseInt(number / 27, 10);
    var iRemainder = number - iAlpha * 26;
    var str = '';
    if (iAlpha > 0) {
        return toABCode(iAlpha) + String.fromCharCode(iRemainder + 64);
    }
    return String.fromCharCode(iRemainder + 64);
}

exports.process = function(excelFile, opt_tplFile, option) {
    option = option || {};

    var sheetIndex = option.sheetIndex != null ? option.sheetIndex : 0;
    var escape = option.escape || '';
    var xlsxData = xlsx.parse(excelFile);

    if (opt_tplFile) {
        var tpl = fs.readFileSync(opt_tplFile, 'utf-8');
        // convert to mustache template first
        var tpl4Mustache = ejs.render(tpl, {});
        var sheet = xlsxData.worksheets[sheetIndex];
        var data = sheet.data;
        var dataMap = {};
        data.forEach(function(row, x) {
            x = x + 1;
            row.forEach(function(col, y) {
                y = y + 1;
                var v = col.value || '';
                if (escape === 'quote') {
                    v = v.replace(/(["'])/g, '\\$1');
                }
                else if (escape === 'html') {
                    v = v.replace(/&/g,'&amp;')
                        .replace(/</g,'&lt;')
                        .replace(/>/g,'&gt;')
                        .replace(/"/g, "&quot;")
                        .replace(/'/g, "&#39;");
                }
                dataMap[x + ',' + y] = v;
                dataMap[x + ',' + toABCode(y)] = v;
            });
        });
        // put in the sheet meta data too
        dataMap.sheet = sheet;
        return mustache.render(tpl4Mustache, dataMap)
    }
    else {
        return JSON.stringify(xlsxData, null, 4);
    }
};
















/* vim: set ts=4 sw=4 sts=4 tw=100 : */
