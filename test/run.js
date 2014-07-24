/***************************************************************************
 * 
 * Copyright (c) 2014 Baidu.com, Inc. All Rights Reserved
 * $Id$
 * 
 **************************************************************************/
 
 
/*
 * path:    run.js
 * desc:    
 * author:  songao(songao@baidu.com)
 * version: $Revision$
 * date:    $Date: 2014/07/24 21:40:10$
 */

var excelFormat = require('../lib/format');

var excelFile = __dirname + '/input.xlsx';
var tplFile = __dirname + '/output.tpl';

console.log(
    excelFormat.process(excelFile, tplFile, {
        sheetIndex: 0
    })
);





















/* vim: set ts=4 sw=4 sts=4 tw=100 : */
