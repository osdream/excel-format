Excel Format Tool
======

## install

```javascript
npm install excel-format -g
```

## usage

### as node.js module

```javascript
var excelFormat = require('excel-format');
var output = excel.process('./in.xslx', './out.tpl', {sheetIndex: 0});
console.log(output);
```

### as command line

```javascript
  Usage: excelformat [options] <excel-file>

  Options:

    -h, --help             output usage information
    -V, --version          output the version number
    -t, --template <file>  template file for formating in mustache and ejs
    -s, --sheet <n>        the sheet index, default 0
    -o, --output <file>    the file to put the formating result
```

### template

simple way: direct use mustache

```javascript
== the sheet meta ==
{{{sheet.name}}}

== use excel axis ==
{{{1,A}}} {{{2,B}}} !

== use number indexes ==
Yes, {{{3,3}}} {{{4,2}}} {{{5,1}}}

```

complex way: use ejs to generate mustache template, then use generated template to format

```javascript
== complex case ==
<% for (var i = 'A'.charCodeAt(0); i <= 'E'.charCodeAt(0); i++) { %>
{{{7,<%= String.fromCharCode(i) %>}}}
<% } %>

```
