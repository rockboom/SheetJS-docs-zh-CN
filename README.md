# [SheetJS js-xlsx](http://sheetjs.com)
SheetJS是用于多种电子表格格式的解析器和编写器。通过官方规范、相关文档以及测试文件实现简洁的JS方法。SheetJS强调解析和编写的稳健，其跨格式的特点和统一的JS规范兼容，并且ES3/ES5浏览器向后兼容IE6。<br>
目前这个是社区版，我们也提供了性能增强的专业版，专业版提供样式和专业支持的附加功能。

[**专业版**](http://sheetjs.com/pro)

[**商业支持**](http://sheetjs.com/support)

[**介绍文档**](http://docs.sheetjs.com/)

[**浏览器示例**](http://sheetjs.com/demos)

[**源码**](http://git.io/xlsx)

[**问题和错误报告**](https://github.com/sheetjs/js-xlsx/issues)

[**常见的支持问题**](https://discourse.sheetjs.com)

[**支持的电子数据表的文件格式：**](#file-formats)

<details>
  <summary><b>支持格式的图表</b> (点击查看)</summary>

![circo graph of format support](formats.png)

![graph legend](legend.png)

</details>

[**浏览器测试**](http://oss.sheetjs.com/js-xlsx/tests/)

[![Build Status](https://saucelabs.com/browser-matrix/sheetjs.svg)](https://saucelabs.com/u/sheetjs)

[![Build Status](https://travis-ci.org/SheetJS/js-xlsx.svg?branch=master)](https://travis-ci.org/SheetJS/js-xlsx)
[![Build Status](https://semaphoreci.com/api/v1/sheetjs/js-xlsx/branches/master/shields_badge.svg)](https://semaphoreci.com/sheetjs/js-xlsx)
[![Coverage Status](http://img.shields.io/coveralls/SheetJS/js-xlsx/master.svg)](https://coveralls.io/r/SheetJS/js-xlsx?branch=master)
[![Dependencies Status](https://david-dm.org/sheetjs/js-xlsx/status.svg)](https://david-dm.org/sheetjs/js-xlsx)
[![npm Downloads](https://img.shields.io/npm/dt/xlsx.svg)](https://npmjs.org/package/xlsx)
[![ghit.me](https://ghit.me/badge.svg?repo=sheetjs/js-xlsx)](https://ghit.me/repo/sheetjs/js-xlsx)
[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)

## 目录表
<details>
  <summary><b>点击展示目录表</b></summary>

<!-- toc -->

- [安装](#installation)
  * [JS 生态示例](#js-ecosystem-demos)
  * [可选模块](#optional-modules)
  * [ECMAScript 5 兼容性](#ecmascript-5-compatibility)
- [原理](#philosophy)
- [解析工作簿](#parsing-workbooks)
  * [解析示例](#parsing-examples)
  * [流式读取](#streaming-read)
- [操作工作簿](#working-with-the-workbook)
  * [解析和编写示例](#parsing-and-writing-examples)
- [编写工作簿](#writing-workbooks)
  * [编写示例](#writing-examples)
  * [流式写入](#streaming-write)
- [接口](#interface)
  * [解析函数](#parsing-functions)
  * [编写函数](#writing-functions)
  * [工具函数](#utilities)
- [常用电子表格的格式](#common-spreadsheet-format)
  * [一般结构](#general-structures)
  * [单元格对象](#cell-object)
    + [数据类型](#data-types)
    + [日期类型](#dates)
  * [数据表对象](#sheet-objects)
    + [工作表对象](#worksheet-object)
    + [图表对象](#chartsheet-object)
    + [宏对象](#macrosheet-object)
    + [对话表对象](#dialogsheet-object)
  * [工作簿对象](#workbook-object)
    + [工作簿的文件属性](#workbook-file-properties)
  * [工作簿级别的属性](#workbook-level-attributes)
    + [定义名称](#defined-names)
    + [查看工作表](#workbook-views)
    + [其他的工作簿属性](#miscellaneous-workbook-properties)
  * [文档特点](#document-features)
    + [公式](#formulae)
    + [列属性](#column-properties)
    + [行属性](#row-properties)
    + [数字格式化](#number-formats)
    + [超链接](#hyperlinks)
    + [单元格注释](#cell-comments)
    + [表的可见性](#sheet-visibility)
    + [VBA和宏命令](#vba-and-macros)
- [解析选项](#parsing-options)
  * [输入类型](#input-type)
  * [猜测文件类型](#guessing-file-type)
- [编写选项](#writing-options)
  * [支持的输出格式](#supported-output-formats)
  * [输出类型](#output-type)
- [工具函数](#utility-functions)
  * [数组输入](#array-of-arrays-input)
  * [对象输入](#array-of-objects-input)
  * [HTML Table 输入](#html-table-input)
  * [公式输出](#formulae-output)
  * [分隔符输出](#delimiter-separated-output)
    + [UTF-16 Unicode 文本](#utf-16-unicode-text)
  * [HTML 输出](#html-output)
  * [JSON](#json)
- [文件格式](#file-formats)
  * [Excel 2007+ XML (XLSX/XLSM)](#excel-2007-xml-xlsxxlsm)
  * [Excel 2.0-95 (BIFF2/BIFF3/BIFF4/BIFF5)](#excel-20-95-biff2biff3biff4biff5)
  * [Excel 97-2004 Binary (BIFF8)](#excel-97-2004-binary-biff8)
  * [Excel 2003-2004 (SpreadsheetML)](#excel-2003-2004-spreadsheetml)
  * [Excel 2007+ Binary (XLSB, BIFF12)](#excel-2007-binary-xlsb-biff12)
  * [Delimiter-Separated Values (CSV/TXT)](#delimiter-separated-values-csvtxt)
  * [其它的工作簿格式](#other-workbook-formats)
    + [Lotus 1-2-3 (WKS/WK1/WK2/WK3/WK4/123)](#lotus-1-2-3-wkswk1wk2wk3wk4123)
    + [Quattro Pro (WQ1/WQ2/WB1/WB2/WB3/QPW)](#quattro-pro-wq1wq2wb1wb2wb3qpw)
    + [OpenDocument Spreadsheet (ODS/FODS)](#opendocument-spreadsheet-odsfods)
    + [Uniform Office Spreadsheet (UOS1/2)](#uniform-office-spreadsheet-uos12)
  * [其它的单表格式](#other-single-worksheet-formats)
    + [dBASE and Visual FoxPro (DBF)](#dbase-and-visual-foxpro-dbf)
    + [Symbolic Link (SYLK)](#symbolic-link-sylk)
    + [Lotus Formatted Text (PRN)](#lotus-formatted-text-prn)
    + [Data Interchange Format (DIF)](#data-interchange-format-dif)
    + [HTML](#html)
    + [Rich Text Format (RTF)](#rich-text-format-rtf)
    + [Ethercalc Record Format (ETH)](#ethercalc-record-format-eth)
- [测试](#testing)
  * [Node](#node)
  * [浏览器](#browser)
  * [测试环境](#tested-environments)
  * [测试文件](#test-files)
- [合作](#contributing)
  * [OSX/Linux](#osxlinux)
  * [Windows](#windows)
  * [测试](#tests)
- [证书](#license)
- [引用](#references)

<!-- tocstop -->

</details>

## 安装
在浏览器里使用，增加一个script标签：
```html
<script lang="javascript" src="dist/xlsx.full.min.js"></script>
```

<details>
  <summary><b>使用CDN</b> (点击显示详情)</summary>

|    CDN     | URL                                        |
|-----------:|:-------------------------------------------|
|    `unpkg` | <https://unpkg.com/xlsx/>                  |
| `jsDelivr` | <https://jsdelivr.com/package/npm/xlsx>    |
|    `CDNjs` | <http://cdnjs.com/libraries/xlsx>          |
|    `packd` | <https://bundle.run/xlsx@latest?name=XLSX> |

`unpkg`提供最新的版本:

```html
<script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
```

</details>


使用 [npm](https://www.npmjs.org/package/xlsx):

```bash
$ npm install xlsx
```

使用 [bower](http://bower.io/search/?q=js-xlsx):

```bash
$ bower install js-xlsx
```

### JS生态示例

[`示例` 目录](demos/) 包括了一些简单的项目:

**框架和APIS**
- [`angularjs`](demos/angular/)
- [`angular 2 / 4 / 5 / 6 and ionic`](demos/angular2/)
- [`knockout`](demos/knockout/)
- [`meteor`](demos/meteor/)
- [`react and react-native`](demos/react/)
- [`vue 2.x and weex`](demos/vue/)
- [`XMLHttpRequest and fetch`](demos/xhr/)
- [`nodejs server`](demos/server/)
- [`databases and key/value stores`](demos/database/)
- [`typed arrays and math`](demos/array/)

**打包工具**
- [`browserify`](demos/browserify/)
- [`fusebox`](demos/fusebox/)
- [`parcel`](demos/parcel/)
- [`requirejs`](demos/requirejs/)
- [`rollup`](demos/rollup/)
- [`systemjs`](demos/systemjs/)
- [`typescript`](demos/typescript/)
- [`webpack 2.x`](demos/webpack/)

**集成平台**
- [`electron application`](demos/electron/)
- [`nw.js application`](demos/nwjs/)
- [`Chrome / Chromium extensions`](demos/chrome/)
- [`Adobe ExtendScript`](demos/extendscript/)
- [`Headless Browsers`](demos/headless/)
- [`canvas-datagrid`](demos/datagrid/)
- [`Swift JSC and other engines`](demos/altjs/)
- [`"serverless" functions`](demos/function/)
- [`internet explorer`](demos/oldie/)

### 可选模块

<details>
    <summary><b>可选特点</b> (点击显示详情)</summary>
node版本自动要求模块提供其他的特性。某些模块的文件相当大而且仅在一些特殊的场景下才会用到，因此不应该把他们当做核心部分一起加载。在浏览器中用到这些模块时，可以用下面的方式进行加载：

```html
<!-- international support from js-codepage -->
<script src="dist/cpexcel.js"></script>
```
每一个依赖合适的版本可以放在 dist/directory 目录下。
完整的单文件版本在 `dist/xlsx.full.min.js` 文件里面。
默认情况下，Webpack和Browserify构建包含可选的模块。可以通过配置Webpack移除对 `resolve.alias` 的支持：

```js
  /* uncomment the lines below to remove support */
  resolve: {
    alias: { "./dist/cpexcel.js": "" } // <-- omit international support
  }
```
</details>

### ECMAScript 5 的兼容性

自从库使用了像 `Array#forEach` 这样的函数，老版本的浏览器需要[shim 提供缺少的函数](http://oss.sheetjs.com/js-xlsx/shim.js)。

要在加载 `xlsx.js` 的script标签之前添加shim，才能使用它。

```html
<!-- add the shim first -->
<script type="text/javascript" src="shim.min.js"></script>
<!-- after the shim is referenced, add the library -->
<script type="text/javascript" src="xlsx.full.min.js"></script>
```
shim.min.js也包括了在IE6-9中用于加载和保存文件的 `IE_LoadFile` 和 `IE_SaveFile`。对于适用于Photoshop和其它的Adobe产品的格式，`xlsx.extendscript.js`脚本会绑定shim。


## 原理

<details>
    <summary><b>原理</b> (点击显示详情)</summary>
在SheetJS之前，处理电子表格文件的接口只能用于特定的格式。许多第三方库要么支持一种格式，要么为每一种支持的文件类型提供一个不同的类集。虽然在Excel 2007里面引入了XLSB，但只有Sheet和Excel支持这种格式。

为了提高不可知格式的显示，js-xlsx使用了被称作["Common Spreadsheet Format"]的纯JS的显示方法(#common-spreadsheet-format)。强调一种统一的显示方式，能够有一些特点，比如格式转换和嵌套`class tap`。通过提取出各种格式的复杂性，工具没有必要担心特定的文件类型。

一个简单的的对象显示和细心的代码练习相结合，能让示例运行在较老的浏览器以及像`ExtendScript`和`Web Workers`这样可选择的环境里执行。虽然很想使用最新的和最好的特性，不过这些特性需要最新的浏览器，用以限制兼容性。

工具函数捕获通用的使用例子，比如生成JS对象或HTML。大多数简单例子的操作只要几行代码。大多数复杂的普遍的复杂操作应该直截了当的生成。

在Excel 2007种，Excel添加XSLX格式作为默认的起始端。然而，有一些其他格式会更多的出现上述的属性。例如，XLSB格式XLSX格式相似，不过文件会使用一半的空间，而且也会更开的打开文件。虽然XLSX编写器可以使用，但是其他格式的编写器也可以使用，因此使用者能够充分利用每一种格式独特的特点。社区版本的主要关注点在正确的数据转换，即从任意一个兼容的数据表示中提取数据，导出适用于任意第三方接口的各种数据格式。
</details>

## 解析工作簿

对于解析，第一步是读取文件。这一步包括获取数据并且导入数据库。这里有一些常用的例子。
<details>
    <summary><b>nodejs读取文件</b> (点击显示详情)</summary>
`readFile` 只能在服务器环境中使用。浏览器没有用于读取任意指定路径文件的API，因此必须使用另外的策略。
```js
if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile('test.xlsx');
/* DO SOMETHING WITH workbook HERE */
```
</details>

<details>
  <summary><b>Photoshop ExtendScript读取文件</b> (点击显示详情)</summary>

`readFile` 用Photoshop和其他的ExtendScript目标把逻辑`File`包起来。需要指定文件的绝对路径

```js
#include "xlsx.extendscript.js"
/* Read test.xlsx from the Documents folder */
var workbook = XLSX.readFile(Folder.myDocuments + '/' + 'test.xlsx');
/* DO SOMETHING WITH workbook HERE */
```

[`extendscript`](demos/extendscript/) 包含了一个更复杂的例子。

</details>

<details>
  <summary><b>浏览器从页面读取TABLE元素</b> (点击显示详情)</summary>
`table_to_book` 和 `table_to_sheet`工具函数获取DOM的TABLE元素，并且通过子节点进行迭代。

```js
var workbook = XLSX.utils.table_to_book(document.getElementById('tableau'));
/* DO SOMETHING WITH workbook HERE */
```

一个网页里面的多张表可以被转换成单个的工作表。

```js
/* create new workbook */
var workbook = XLSX.utils.book_new();

/* convert table 'table1' to worksheet named "Sheet1" */
var ws1 = XLSX.utils.table_to_sheet(document.getElementById('table1'));
XLSX.utils.book_append_sheet(workbook, ws1, "Sheet1");

/* convert table 'table2' to worksheet named "Sheet2" */
var ws2 = XLSX.utils.table_to_sheet(document.getElementById('table2'));
XLSX.utils.book_append_sheet(workbook, ws2, "Sheet2");

/* workbook now has 2 worksheets */
```

另一种选择，HTML代码也可以被提取和解析。

```js
var htmlstr = document.getElementById('tableau').outerHTML;
var workbook = XLSX.read(htmlstr, {type:'string'});
```
</details>

<details>
  <summary><b>浏览器下载文件(ajax)</b> (点击显示详情)</summary>
注意：对于运行在老版浏览器里更完整的例子，请查看示例 <http://oss.sheetjs.com/js-xlsx/ajax.html>。[`xhr`示例](demos/xhr/)包含`XMLHttpRequest` 和 `fetch`更多的例子。

```js
var url = "http://oss.sheetjs.com/test_files/formula_stress_test.xlsx";

/* set up async GET request */
var req = new XMLHttpRequest();
req.open("GET", url, true);
req.responseType = "arraybuffer";

req.onload = function(e) {
  var data = new Uint8Array(req.response);
  var workbook = XLSX.read(data, {type:"array"});

  /* DO SOMETHING WITH workbook HERE */
}

req.send();
```

</details>

<details>
  <summary><b>浏览器拖拽</b> (点击显示详情)</summary>
拖拽使用了HTML5 的 `FileReader` API，加载数据时使用`readAsBinaryString` 或 `readAsArrayBuffer`。但并不是所有的浏览器都支持全部的 `FileReader` API，因此非常推荐动态的特性检测。

```js
var rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
function handleDrop(e) {
  e.stopPropagation(); e.preventDefault();
  var files = e.dataTransfer.files, f = files[0];
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = e.target.result;
    if(!rABS) data = new Uint8Array(data);
    var workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});

    /* DO SOMETHING WITH workbook HERE */
  };
  if(rABS) reader.readAsBinaryString(f); else reader.readAsArrayBuffer(f);
}
drop_dom_element.addEventListener('drop', handleDrop, false);
```
</details>

<details>
  <summary><b>浏览器通过form元素上传文件</b> (点击显示详情)</summary>

来自`file input`元素的数据能够被和拖拽例子中相同的`FileReader`API处理。

```js
var rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
function handleFile(e) {
  var files = e.target.files, f = files[0];
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = e.target.result;
    if(!rABS) data = new Uint8Array(data);
    var workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});

    /* DO SOMETHING WITH workbook HERE */
  };
  if(rABS) reader.readAsBinaryString(f); else reader.readAsArrayBuffer(f);
}
input_dom_element.addEventListener('change', handleFile, false);
```
[`oldie`示例](demos/oldie/)展示了一个IE兼容性的回退方案。

</details>

包括移动App文件处理等更多的使用例子可以在[included demos](demos/)中查看。

## 流式读取文件

<details>
  <summary><b>为什么没有流式读取API？</b> (点击显示详情)</summary>

最常用的和最令人感兴趣的格式(XLS, XLSX/M, XLSB, ODS)最终都是ZIP或CFB文件容器。两种格式都不会放目录结构在文件开头：ZIP文件把主要的目录记录放在逻辑文件的结尾，然而CFB文件可以把存储信息放在文件的任何地方。所以，为了正确地处理这些格式，流式函数必须在开始之前缓存整个文件。这样证明了流式的期待的错误，因此我们不提供任何流式阅读API。

</details>

当处理可读流时，最简单的方式是缓存流，并且最后再去处理整个文件。这可以通过临时文件或者是显式连接流来实现。

<details>
  <summary><b>显式连接流</b> (点击显示详情)</summary>

```js
var fs = require('fs');
var XLSX = require('xlsx');
function process_RS(stream/*:ReadStream*/, cb/*:(wb:Workbook)=>void*/)/*:void*/{
  var buffers = [];
  stream.on('data', function(data) { buffers.push(data); });
  stream.on('end', function() {
    var buffer = Buffer.concat(buffers);
    var workbook = XLSX.read(buffer, {type:"buffer"});

    /* DO SOMETHING WITH workbook IN THE CALLBACK */
    cb(workbook);
  });
}
```

使用像`concat-stream`这样的模块会有更多有效的解决办法可以使用。
</details>

<details>
  <summary><b>首先写入文件系统</b> (点击显示详情)</summary>

这个例子使用[`tempfile`](https://npm.im/tempfile)生成文件名。

```js
var fs = require('fs'), tempfile = require('tempfile');
var XLSX = require('xlsx');
function process_RS(stream/*:ReadStream*/, cb/*:(wb:Workbook)=>void*/)/*:void*/{
  var fname = tempfile('.sheetjs');
  console.log(fname);
  var ostream = fs.createWriteStream(fname);
  stream.pipe(ostream);
  ostream.on('finish', function() {
    var workbook = XLSX.readFile(fname);
    fs.unlinkSync(fname);

    /* DO SOMETHING WITH workbook IN THE CALLBACK */
    cb(workbook);
  });
}
```
</details>

## 使用工作簿

完整的对象格式会在本文件的后面部分进行介绍。

<details>
  <summary><b>读取指定的单元格</b> (点击显示详情)</summary>

这个例子提取first工作表中A1单元格的存储值：

```js
var first_sheet_name = workbook.SheetNames[0];
var address_of_cell = 'A1';

/* Get worksheet */
var worksheet = workbook.Sheets[first_sheet_name];

/* Find desired cell */
var desired_cell = worksheet[address_of_cell];

/* Get the value */
var desired_value = (desired_cell ? desired_cell.v : undefined);
```
</details>

<details>
  <summary><b>在工作簿中增加新的工作表</b> (点击显示详情)</summary>

例子中使用[`XLSX.utils.aoa_to_sheet`](#array-of-arrays-input)生成工作表，使用`XLSX.utils.book_append_sheet`把表添加到工作簿中。

```js
var new_ws_name = "SheetJS";

/* make worksheet */
var ws_data = [
  [ "S", "h", "e", "e", "t", "J", "S" ],
  [  1 ,  2 ,  3 ,  4 ,  5 ]
];
var ws = XLSX.utils.aoa_to_sheet(ws_data);

/* Add the worksheet to the workbook */
XLSX.utils.book_append_sheet(wb, ws, ws_name);
```
</details>

<details>
  <summary><b>从头开始创建工作簿</b> (点击显示详情)</summary>

工作簿对象包含一个`SheetNames`名称数组和一个`Sheets`对象，用来将表名映射到表对象。`XLSX.utils.book_new`工具函数创建一个新的工作簿对象：

```js
/* create a new blank workbook */
var wb = XLSX.utils.book_new();
```

新的工作簿是空白的而且不包含工作表。如果工作簿，那么写入函数将会出错。
</details>

### 解析和编写示例

- <http://sheetjs.com/demos/modify.html> read + modify + write files

- <https://github.com/SheetJS/js-xlsx/blob/master/bin/xlsx.njs> node

node安装一个能够读取电子数据表和输出各种格式的命令行工具 `xlsx`。源码可以在 `bin` 目录下的`xlsx.njs`里面找到。

`XLSX.utils`中的一些辅助函数会生成不同的工作表视图。

- `XLSX.utils.sheet_to_csv` 生成CSV文件
- `XLSX.utils.sheet_to_txt` 生成UTF16的格式化文本
- `XLSX.utils.sheet_to_html` 生成HTML
- `XLSX.utils.sheet_to_json` 生成一个对象数组
- `XLSX.utils.sheet_to_formulae` 生成一张公示列表

## 编写工作簿

对编写而言，第一步是生成导出数据。辅助函数`write` 和 `writeFile`将会生成各种适合分发的数据格式。第二步是和端点实际的共享数据。假设`workbook`是一个工作簿对象。

<details>
  <summary><b>nodejs写入文件</b> (点击显示详情)</summary>

`XLSX.writeFile` uses `fs.writeFileSync` in server environments:

```js
if(typeof require !== 'undefined') XLSX = require('xlsx');
/* output format determined by filename */
XLSX.writeFile(workbook, 'out.xlsb');
/* at this point, out.xlsb is a file that you can distribute */
```
</details>

<details>
  <summary><b>Photoshop ExtendScript 写入文件</b> (点击显示详情)</summary>

`writeFile` 把 `File`包裹在 Photoshop 和 other ExtendScript 目标里面。指定的路径应该是绝对路径。

```js
#include "xlsx.extendscript.js"
/* output format determined by filename */
XLSX.writeFile(workbook, 'out.xlsx');
/* at this point, out.xlsx is a file that you can distribute */
```

[`extendscript` 示例](demos/extendscript/)包含有更复杂的例子。

</details>

<details>
  <summary><b>浏览器将TABLE元素添加到页面</b> (点击显示详情)</summary>

`sheet_to_html`工具函数生成能被添加到任意DOM元素的HTML代码。

```js
var worksheet = workbook.Sheets[workbook.SheetNames[0]];
var container = document.getElementById('tableau');
container.innerHTML = XLSX.utils.sheet_to_html(worksheet);
```
</details>

<details>
  <summary><b>浏览器上传文件(ajax)</b> (点击显示详情)</summary>

用 `XHR` 的完整的复杂示例可以在 [`XHR`示例](demos/xhr/) 中查看，获取和包装器库的例子也可以包含在里面。例子中假设服务器能处理Base64编码的文件(查看基本的莫得服务器示例)。

```js
/* in this example, send a base64 string to the server */
var wopts = { bookType:'xlsx', bookSST:false, type:'base64' };

var wbout = XLSX.write(workbook,wopts);

var req = new XMLHttpRequest();
req.open("POST", "/upload", true);
var formdata = new FormData();
formdata.append('file', 'test.xlsx'); // <-- server expects `file` to hold name
formdata.append('data', wbout); // <-- `data` holds the base64-encoded data
req.send(formdata);
```
</details>

<details>
  <summary><b>浏览器保存文件</b> (点击显示详情)</summary>

`XLSX.writeFile` 包含了一些用于触发文件保存的方法。

- `URL`浏览器API为文件创建一个URL对象，通过创建a标签并给他添加click事件就可以使用URL对象。现代浏览器都支持这个方法。
- `msSaveBlob`是IE10及IE10以上用来触发文件保存的API。
- 对于Windows XP 和 Windows 7里面的IE6和IE6以上的以上浏览器，`IE_FileSave` 使用 VBScript 和 ActiveX 来写入文件。s补充程序（shim)必须包含在包含的HTML页面中。
  
并没有标准的方法判断是否真实的文件已经被下载了。

```js
/* output format determined by filename */
XLSX.writeFile(workbook, 'out.xlsb');
/* at this point, out.xlsb will have been downloaded */
```

</details>

<details>
  <summary><b>浏览器保存文件(兼容性)</b> (点击显示详情)</summary>

`XLSX.writeFile`方法在大多数的现代浏览器以及老版本的浏览器中都能使用。对于更老的浏览器，wrapper库里面有变通的方法可以应用。

[`FileSaver.js`](https://github.com/eligrey/FileSaver.js/) 执行 `saveAs`方法。

注意：如果`saveAs`方法可以使用，`XLSX.writeFile`会自动调用。

```js
/* bookType can be any supported output type */
var wopts = { bookType:'xlsx', bookSST:false, type:'array' };

var wbout = XLSX.write(workbook,wopts);

/* the saveAs call downloads a file on the local machine */
saveAs(new Blob([wbout],{type:"application/octet-stream"}), "test.xlsx");
```

[`Downloadify`](https://github.com/dcneiner/downloadify)使用Flash SWF按钮生成本地文件，即使是ActiveX不能使用的环境也适用。

```js
Downloadify.create(id,{
	/* other options are required! read the downloadify docs for more info */
	filename: "test.xlsx",
	data: function() { return XLSX.write(wb, {bookType:"xlsx", type:'base64'}); },
	append: false,
	dataType: 'base64'
});
```

[`oldie`示例](demos/oldie/)展示了IE向后兼容的场景。
</details>

[included demos](demos/)包含了移动app和其他专门的部署。