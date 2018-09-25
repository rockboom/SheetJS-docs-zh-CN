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

### 写入示例

- <http://sheetjs.com/demos/table.html> 导出一个 HTML table
- <http://sheetjs.com/demos/writexlsx.html> 生成一个简单文件

### 流式写入

`XLSX.stream`对象中可以使用流式写入函数。流式函数像普通的函数一样传入相同的参数，不过返回一个可读流。但是他们只暴露在Nodejs中。

- `XLSX.stream.to_csv` 是 `XLSX.utils.sheet_to_csv`的流式版本。
- `XLSX.stream.to_html` 是 `XLSX.utils.sheet_to_html`的流式版本。
- `XLSX.stream.to_json` 是 `XLSX.utils.sheet_to_json`的流式版本。

<details>
  <summary><b>nodejs转换成CSV并写入文件</b> (点击显示详情)</summary>

```js
var output_file_name = "out.csv";
var stream = XLSX.stream.to_csv(worksheet);
stream.pipe(fs.createWriteStream(output_file_name));
```

</details>

<details>
  <summary><b>nodejs将JSON流输出到屏幕</b> (点击显示详情)</summary>

```js
/* to_json returns an object-mode stream */
var stream = XLSX.stream.to_json(worksheet, {raw:true});

/* the following stream converts JS objects to text via JSON.stringify */
var conv = new Transform({writableObjectMode:true});
conv._transform = function(obj, e, cb){ cb(null, JSON.stringify(obj) + "\n"); };

stream.pipe(conv); conv.pipe(process.stdout);
```

</details>

<https://github.com/sheetjs/sheetaki> pips将可写流写入nodejs响应。

## 接口

`XLSX`是浏览器暴露出来可以使用的方法，导出到node中能够使用的`XLSX.version`是`XLSX`库的版本(通过构建脚本添加)。`XLSX.SSF`是[格式化库](http://版本git.io/ssf)的嵌入版本。

### 解析函数

`XLSX.read(data, read_opts)` 用来解析数据  `data`。
`XLSX.readFile(filename, read_opts)` 用来读取文件名 `filename` 并且解析。解析选项会在[解析选项](#parsing-options)部分阐述。

### 写入函数

`XLSX.write(wb, write_opts)` 用来写入工作簿 `wb`。
`XLSX.writeFile(wb, filename, write_opts)` 把 `wb` 写入到特定的文件 `filename` 中。如果是基于浏览器的环境，此函数会强制浏览器端下载。
`XLSX.writeFileAsync(filename, wb, o, cb)` 把 `wb` 写入到特定的文件 `filename` 中。如果 `o` 被省略，写入函数会使用第三个参数作为回调函数。

`XLSX.stream` 包含一组流式写入函数的集合。

写入选项会在[写入选项部分](#writing-options)部分进行阐述。

### 工具

`XLSX.utils`对象中的工具函数都可以使用，工具函数在[工具函数](#utility-functions)部分进行阐述。

**导入:**

- `aoa_to_sheet` 把转换JS数据数组的数组为工作表。
- `json_to_sheet` 把JS对象数组转换为工作表。
- `table_to_sheet` 把DOM TABLE元素转换为工作表。
- `sheet_add_aoa` 把JS数据数组的数组添加到已存在的工作表中。
- `sheet_add_json` 把JS对象数组添加到已存在的工作表中。

**导出:**

- `sheet_to_json` 把工作表转换为JS对象数组。
- `sheet_to_csv` 生成分隔符隔开值的输出。
- `sheet_to_txt` 生成UTF16格式化的文本。
- `sheet_to_html` 生成HTML输出。
- `sheet_to_formulae` 生成公式列表(带有值回退)。

**单元格和单元格地址的操作:**

- `format_cell` 生成文本类型的单元格值(使用数字格式)。
- `encode_row / decode_row` 在0索引行和1索引行之间转换。
- `encode_col / decode_col` 在0索引列和列名之间转换。
- `encode_cell / decode_cell` 转换单元格地址。
- `encode_range / decode_range` 转换单元格的范围。

## 常用的数据表格式(Common Spreadsheet Format)

js-xlsx符合常用的数据表格式(CSF)。

### 一般结构

单元格地址对象的存储格式为`{c:C, r:R}`，其中`C`和`R`分别代表的是0索引列和行号。例如单元格地址`B5`用对象`{c:1, r:4}`表示。

单元格范围对象存储格式为`{s:S, e:E}`，其中`S`是第一个单元格，`E`是最后一个单元格。范围是包含关系。例如范围 `A3:B7`用对象`{s:{c:0, r:2}, e:{c:1, r:6}}`表示。当遍历数据表范围时，工具函数执行行优先命令。

```js
for(var R = range.s.r; R <= range.e.r; ++R) {
  for(var C = range.s.c; C <= range.e.c; ++C) {
    var cell_address = {c:C, r:R};
    /* if an A1-style address is needed, encode the address */
    var cell_ref = XLSX.utils.encode_cell(cell_address);
  }
}
```

### 单元格对象

单元格对象是纯粹的JS对象，它的keys和values遵循下列的约定：

| Key | Description|
| --- | ---------- |
| `v` | 原始值(查看数据类型部分获取更多的信息) |
| `w` | 格式化文本(如果可以使用) |
| `t` | 内行: `b` Boolean, `e` Error, `n` Number, `d` Date, `s` Text, `z` Stub |
| `f` | 单元格公式编码为A1样式的字符串(如果可以使用) |
| `F` | 如果公式是数组公式，则包围数组的范围(如果可以使用) |
| `r` | 富文本编码 (如果可以使用) |
| `h` | 富文本渲染成HTML (如果可以使用) |
| `c` | 与单元格关联的注释 |
| `z` | 与单元格关联的数字格式字符串(如果有必要) |
| `l` | 单元格的超链接对象 (`.Target` 长联接, `.Tooltip` 是提示消息) |
| `s` | 单元格的样式/主题 (如果可以使用) |

如果`w`文本可以使用，内置的导出工具(比如CSV导出方法)就会使用它。要想改变单元格的值，在打算导出之前确保删除`cell.w`(或者设置 `cell.w`为`undefined`)。工具函数会根据数字格式(`cell.z`)和原始值(如果可用)重新生成`w`文本。

真实的数组公式存储在数组范围中第一个单元个的`f`字段内。此范围内的其他单元格会省略`f`字段。

### 数据类型

原始值被存储在`v`值属性中，用来解释基于`t`类型的属性。这样的区别允许用于数字和数字类型文本的展示。下面有6种有效的单元格类型。

| Type | Description |
| :--: | :---------- |
| `b`  | Boolean: 值可以理解为JS `boolean`                            |
| `e`  | Error: 值是数字类型的编码，而且`w`属性存储共同的名称 ** |
| `n`  | Number: 值是JS `number` ** |
| `d`  | Date: 值是 JS `Date` 对象或者是被解析为Date的字符串 **|
| `s`  | Text: 值可以理解为 JS `string` 并且被写成文本 **|
| `z`  | Stub: 被数据处理工具函数忽略的空白子单元格 ** |

<details>
  <summary><b>Error 值以及含义</b> (点击显示详情)</summary>

|  Value | Error 含义   |
| -----: | :-------------- |
| `0x00` | `#NULL!`        |
| `0x07` | `#DIV/0!`       |
| `0x0F` | `#VALUE!`       |
| `0x17` | `#REF!`         |
| `0x1D` | `#NAME?`        |
| `0x24` | `#NUM!`         |
| `0x2A` | `#N/A`          |
| `0x2B` | `#GETTING_DATA` |

</details>

`n`表示Number类型。`n`包括了所有被Excel存储为数字的数据表，比如dates/times和Boolean字段。Excel专门使用能够被IEEE754浮点数表示的数据，比如JS Number，所以字段 `v` 保存原始数字。`w`字段保持格式化文本。Dates 默认存储为数字，使用`XLSX.SSF.parse_date_code`进行转换。

类型 `d`表示日期类型，只有当选项为`cellDates`才会生成日期类型。因为JSON没有普通的日期类型，所以希望解析器存储的日期字符串像从`date.toISOString()`中获取的一样。另一方面，写入函数和导出函数也可以处理日期字符串和JS日期对象。需要注意Excel会忽略时区修饰符，并且处理所有本地时区的日期。代码库没有改正这个错误。

类型`z`表示空白的存根单元格。生成存根单元格是为了以防万一单元格没有被赋予指定值，但是保留了注释或者是其他的元数据。存根单元格会被核心库的数据处理工具函数忽略。默认情况下不会生成存根单元格，只有当解析器`sheetStubs`的选项被设为`true`时才会生成。


#### 日期

<details>
  <summary><b>Excel 日期编码的细节</b> (点击显示详情)</summary>

默认情况下，Excel把日期存储为数字，并用指定的日期处理格式编码进行处理。例如，日期`19-Feb-17`被存储为数字`42785`，数字格式为`d-mmm-yy`。`SSF`模块了解数字格式并进行适当的转换。

XSLX也支持特定的日期类型`d`，它的数据是ISO 8601日期字符串。格式化工具把日期还原为数字。

所有解析器的默认行为是生成数字单元格。设置`cellDates`为true会强制生成器存储日期。

</details>

<details>
  <summary><b>时区和日期</b> (点击显示详情)</summary>

Excel没有原生的通用时间的概念。所有的时间都会在本地时区指定。Excel限制指定真正的绝对日期。

对于下面的Excel，代码库将所有的日期视为相对于当地时区的日期。
</details>

<details>
  <summary><b>时期：1900年和1904年</b> (点击显示详情)</summary>

Excel支持两种时期(January 1 1900和January 1 1904)，查看["1900 vs. 1904 Date System" article](http://support2.microsoft.com/kb/180162)。工作簿的时期可以通过测试工作簿的`wb.Workbook.WBProps.date1904`属性来决定：

```js
!!(((wb.Workbook||{}).WBProps||{}).date1904)
```
</details>

### 数据表对象

每一个不以`!`开始的key都会映射到一个单元格(用`A-1`符合)。`sheet[address]` 返回指定地址的单元格对象。

**指定的数据表属性(通过 `sheet[key]`访问, 每一个都以 `!`开始):**

- `sheet['!ref']`：A-1的范围是基于表示数据表的范围。操作数据表的函数应该使用这个参数来决定操作范围。此范围之外指定的单元格不会被处理。尤其是手动编写数据表时，范围之外的单元格不会被包含在其中。

处理数据表的函数应该对`!ref`的存在进行检测。如果`!ref`被忽略或者不是有效的范围，函数就可以吧数据表看做是空表或者尝试猜范围。词库附带的工具函数会将工作表视为空(例如CSV的输出是空字符串)。

当用`sheetRows`属性集读取工作表时，ref参数会使用限制的范围。最初的范围通过`ws['!fullref']`设置。

- `sheet['!margins']`：`sheet['!margins']`对象表示页面的边距。默认值遵循Excel的常规设置。Excel也有"wide"和"narrow"的设置，不过他们都被存储为原生的尺寸。主要的属性已在下表列出：

<details>
  <summary><b>页面边距详情</b> (点击显示详情)</summary>

| key      | description            | "normal" | "wide" | "narrow" |
|----------|------------------------|:---------|:-------|:-------- |
| `left`   | left margin (inches)   | `0.7`    | `1.0`  | `0.25`   |
| `right`  | right margin (inches)  | `0.7`    | `1.0`  | `0.25`   |
| `top`    | top margin (inches)    | `0.75`   | `1.0`  | `0.75`   |
| `bottom` | bottom margin (inches) | `0.75`   | `1.0`  | `0.75`   |
| `header` | header margin (inches) | `0.3`    | `0.5`  | `0.3`    |
| `footer` | footer margin (inches) | `0.3`    | `0.5`  | `0.3`    |

```js
/* Set worksheet sheet to "normal" */
ws["!margins"]={left:0.7, right:0.7, top:0.75,bottom:0.75,header:0.3,footer:0.3}
/* Set worksheet sheet to "wide" */
ws["!margins"]={left:1.0, right:1.0, top:1.0, bottom:1.0, header:0.5,footer:0.5}
/* Set worksheet sheet to "narrow" */
ws["!margins"]={left:0.25,right:0.25,top:0.75,bottom:0.75,header:0.3,footer:0.3}
```
</details>

#### 工作表对象

除了基本的数据表关键字之外，工作表还增加了下面的内容：

- `ws['!cols']`：返回列属性对象的数组。实际上的列宽使用统一的方式存储在文件里，宽度的测量依据最大数字宽度(在像素中，最大的渲染宽度是数字0-9)。渲染时，列对象用`wpx`字段存储像素宽度，用`wch`存储字符宽度，用`MDW`字段存储最大数字宽度。

- `ws['!rows']`: 返回行属性对象的数组，后面的文档会进行阐述。每一个行对象编码属性包括行高和能见度。

- `ws['!merges']`: 返回与工作表中合并单元格相对应的范围对象的数组。纯文本格式不支持合并单元格。如果合并的单元格存在，CSV导出将会把所有的单元格写入合并范围，因此确保在合并的范围内只有第一个单元格(左上角的单元格)被设置。

- `ws['!protect']`: 写入数据表保护属性的对象。`password`键为支持密码保护的数据表(XLSX/XLSB/XLS)指定密码。写入函数会使用XOR模糊方式。下面`key`控制数据表保护--数据表被锁定时设置key为false可以使用feature，或者设为true禁用feature。

<details>
  <summary><b>工作表保护详情</b> (点击显示详情)</summary>

| key                   | feature (true=disabled / false=enabled) | default    |
|:----------------------|:----------------------------------------|:-----------|
| `selectLockedCells`   | Select locked cells                     | enabled    |
| `selectUnlockedCells` | Select unlocked cells                   | enabled    |
| `formatCells`         | Format cells                            | disabled   |
| `formatColumns`       | Format columns                          | disabled   |
| `formatRows`          | Format rows                             | disabled   |
| `insertColumns`       | Insert columns                          | disabled   |
| `insertRows`          | Insert rows                             | disabled   |
| `insertHyperlinks`    | Insert hyperlinks                       | disabled   |
| `deleteColumns`       | Delete columns                          | disabled   |
| `deleteRows`          | Delete rows                             | disabled   |
| `sort`                | Sort                                    | disabled   |
| `autoFilter`          | Filter                                  | disabled   |
| `pivotTables`         | Use PivotTable reports                  | disabled   |
| `objects`             | Edit objects                            | enabled    |
| `scenarios`           | Edit scenarios                          | enabled    |
</details>

- `ws['!autofilter']`: 自动筛选下面的模式：

```typescript
type AutoFilter = {
  ref:string; // A-1 based range representing the AutoFilter table range
}
```

#### 图表对象

图表会被显示为标准的数据表。要注意和`!type`被设置为`"chart"`的属性进行区分。

底层数据和`!ref`指的是图表中的缓存数据。 图表的第一行是底层标题。

#### 宏对象

宏对象会被显示为标准的数据表。注意与`!type`设置为`"macro"`的属性进行区分。

#### 对话框对象

对话框对象会被显示为标准的数据表。注意与`!type`设置为`"dialog"`的属性进行区分。

### 工作簿对象

`workbook.SheetNames` 是工作簿内工作表的有序列表。

`wb.Sheets[sheetname]` 返回一个表示工作表的对象。

`wb.Props` 是一个存储标准属性的对象。`wb.Custprops` 存储自定义的属性。因为XLS标准属性偏离了XLSX标准，所以XLS解析把核心的属性存储在两个属性中。

`wb.Workbook` 存储[工作簿级别的特性](#workbook-level-attributes).

#### 工作簿文件属性

各种各样的文件格式为不同的文件属性使用不同的内置名称。工作簿的`Props`对象用来规范这些名称。

<details>
  <summary><b>文件属性</b> (点击显示详情)</summary>

| JS Name       | Excel Description              |
|:--------------|:-------------------------------|
| `Title`       | Summary tab "Title"            |
| `Subject`     | Summary tab "Subject"          |
| `Author`      | Summary tab "Author"           |
| `Manager`     | Summary tab "Manager"          |
| `Company`     | Summary tab "Company"          |
| `Category`    | Summary tab "Category"         |
| `Keywords`    | Summary tab "Keywords"         |
| `Comments`    | Summary tab "Comments"         |
| `LastAuthor`  | Statistics tab "Last saved by" |
| `CreatedDate` | Statistics tab "Created"       |

</details>

例如设置工作簿的title属性：

```js
if(!wb.Props) wb.Props = {};
wb.Props.Title = "Insert Title Here";
```

自定义的属性会被添加到工作簿的`Custom`对象中：

```js
if(!wb.Custprops) wb.Custprops = {};
wb.Custprops["Custom Property"] = "Custom Value";
```

写入函数将会处理选项对象的`Props`键：

```js
/* force the Author to be "SheetJS" */
XLSX.write(wb, {Props:{Author:"SheetJS"}});
```

### 工作簿级别的特性

`wb.Workbook` 存储工作簿级别的特性。

#### 定义名称

`wb.Workbook.Names` 是一个定义名称对象的数组，这些名称对象都有键：

<details>
  <summary><b>定义名称属性</b> (点击显示)</summary>

| Key       | Description    |
|:----------|:--------------|
| `Sheet`   | 名称的范围。  数据表的索引 (0 = 第一章张数据表) 或`null` (工作簿)  |
| `Name`    | 区分大小写的名称。 标准的规则应用** |
| `Ref`     | A1单元格样式的引用 (`"Sheet1!$A$1:$D$20"`) |
| `Comment` | 注释 (只适用于XLS/XLSX/XLSB) |

</details>

Excel 允许两个表格范围定义的名称共享相同的名称。但是一个表格范围的名称不能和一个工作簿范围的名称相冲突。工作簿写入函数不强制这样的约束。

#### 工作簿视图

`wb.Workbook.Views` 是一个工作簿视图对象数组，这些视图对象有keys。

| Key | Description |
|:----|:------------|
| `RTL` | 如果值为true，从左到右的显示 |

#### 混合的工作簿属性

| Key  | Description   |
|:-----|:--------------|
| `CodeName`        | [VBA Project Workbook Code Name](#vba-and-macros)            |
| `date1904`        | 时间: 0/false 表示1900系统时间, 1/true 表示1904系统时间        |
| `filterPrivacy`   | 警告或去除存储的个人验证信息                                   |

### 文档特点

即使是想存储数据这样的基本特点，官方的Excel格式也会用不同的方式存储相同的内容。期望解析器从底层文件格式转换为通用的电子表格格式。期望编写器将CSF格式转换回基本的文件格式。

#### 公式

A1单元格样式字符串被存储在`f`字段中。虽然不同的文件格式用不同的方式存储文件格式，不过这些格式都需要被翻译。虽然一些格式存储的公式有一个前导等号，但是CSF公式不以`=`开始。

<details>
  <summary><b>A1=1, A2=2, A3=A1+A2的显示</b> (点击展示详情)</summary>

```js
{
  "!ref": "A1:A3",
  A1: { t:'n', v:1 },
  A2: { t:'n', v:2 },
  A3: { t:'n', v:3, f:'A1+A2' }
}
```
</details>

共享的公式会被解压缩，并且每一个单元格都有相应的公式。编写器通常不会尝试去生成共享公式。

有公式记录但是没有值的单元格会被序列化，序列化的方式能够被Excel和其他电子表格工具将会识别。这个代码库将不会自动计算公式结果！例如去计算工作表中的`BESSELJ`。

<details>
  <summary><b>没有已知值的公式</b> (点击显示详情)</summary>

```js
{
  "!ref": "A1:A3",
  A1: { t:'n', v:3.14159 },
  A2: { t:'n', v:2 },
  A3: { t:'n', f:'BESSELJ(A1,A2)' }
}
```
</details>

**数组公式**

数组公式被存储在数组块左上角的单元格内。一个数组公式的所有单元格都会有于该范围对应的`F`字段。一个单一的单元格公式要注意于`F`字段所存储的纯公式进行区分。

<details>
  <summary><b>数组公式示例</b> (点击显示详情)</summary>

例如设置单元格`C1`为数组公式`{=SUM(A1:A3*B1:B3)}`：

```js
worksheet['C1'] = { t:'n', f: "SUM(A1:A3*B1:B3)", F:"C1:C1" };
```

对于多个单元格的数组公式，每一个单元格都有相同的数组范围，不过只有第一个单元格指定公式。考虑`D1:D3=A1:A3*B1:B3`：

```js
worksheet['D1'] = { t:'n', F:"D1:D3", f:"A1:A3*B1:B3" };
worksheet['D2'] = { t:'n', F:"D1:D3" };
worksheet['D3'] = { t:'n', F:"D1:D3" };
```
</details>

工具函数和编写器被用来检查`F`字段的存在，并且忽略单元格内任何可能的公式元素`f`，这些单元格并不包含起始单元格。这些操作函数并不会被要求执行公式的校验。

<details>
  <summary><b>公式输出工具函数</b> (点击显示详情)</summary>

`sheet_to_formulae`方法生成为每个公式或者是数组公式生成一行。数组公式被渲染在`range=formula`的表格内，而纯单元格被渲染在`cell=formula or value`的表格内。注意字符串的迭代会有前缀符号`'`，与Excel公式栏显示的一致。

<details>
  <summary><b>公式文件格式细节</b> (点击显示详情)</summary>

| Storage Representation | Formats                  | Read  | Write |
|:-----------------------|:-------------------------|:-----:|:-----:|
| A1-style strings       | XLSX                     |  :o:  |  :o:  |
| RC-style strings       | XLML and plain text      |  :o:  |  :o:  |
| BIFF Parsed formulae   | XLSB and all XLS formats |  :o:  |       |
| OpenFormula formulae   | ODS/FODS/UOS             |  :o:  |  :o:  |

因为Excel禁止单元格的命名与A1的名称或者是RC样式单元格引用相冲突，可能会进行不是那么简单的正则变化。DIFF解析的公式必须被明确的解开。OpenFormula可以转换正则表达式。
</details>

#### 列属性

每张表都会有`!cols`数组，如果展开的话就是`ColInfo`的一个集合，有下列的属性：

```typescript
type ColInfo = {
  /* visibility */
  hidden?: boolean; // if true, the column is hidden

  /* column width is specified in one of the following ways: */
  wpx?:    number;  // width in screen pixels
  width?:  number;  // width in Excel's "Max Digit Width", width*256 is integral
  wch?:    number;  // width in characters

  /* other fields for preserving features from files */
  MDW?:    number;  // Excel's "Max Digit Width" unit, always integral
};
```
</details>

<details>
  <summary><b>为什么有三种宽度类型？</b> (点击显示详情)</summary>
有三种不同的宽度类型对应于电子数据表存储列宽的三种不同方式。

SYLK和其他的纯文本格式使用原生的字符计算。像Visicalc和Multiplan这样的同时期的工具是基于字符的。因为字符有相同的宽度，足以存储一个计数。这样的传统也延续到了BIFF格式。

SpreadsheetML (2003) 尝试通过标准化整个文件中的屏幕像素计数来与HTML对齐。列宽、行高以及其他的测量使用像素。当像素和字符数量不一致时，Excel四舍五入结果。

XLSX内部用一个模糊的"最大数位宽度"表存储列宽。最大数字宽度是渲染时最大数字的宽度，通常字符"0"是最宽的。内部的宽度必须是宽度除以256的整数倍。ECMA-376介绍了一个公式用于像素和内部宽度之间的转换。这代表一种混合的方式。

读取函数尝试去填充全部的三种属性。写函数努力尝试将指定值循环到所需类型。为了阻止潜在的冲突。首先操作应该要删除其他的属性。列入，当改变像素宽度时，删除`wch` 和 `width`属性。
</details>

<details>
  <summary><b>执行细节</b> (点击显示详情)</summary>

给出的这些约束可能决定了MDW没有检查字体！解析器通过从宽度转换为像素并返回来猜测像素宽度，重复所有可能的MDW并选择最小化村务的MDW。XLML实际上存储额像素宽度，所以猜想会在相反的方向运行。

即使所有的信息都是可用的，也会期望写入函数遵循下面的优先级顺序：

1) 如果 `width` 字段可用，优先使用`width`。
2) 如果 `wpx` 字段可用，请使用`wpx`。
3) 如果 `wch` 字段可用，请使用`wch`。

</details>

#### 行属性

如果`!rows` 数组在每张电子表中都存在，那就是一个`RowInfo`对象的集合，集合包含一下的属性：

```typescript
type RowInfo = {
  /* visibility */
  hidden?: boolean; // if true, the row is hidden

  /* row height is specified in one of the following ways: */
  hpx?:    number;  // height in screen pixels
  hpt?:    number;  // height in points

  level?:  number;  // 0-indexed outline / group level
};
```
注意：Excel UI显示基本大纲级别为`1`,最大级别为`8`。`level`字段存储基本大纲级别为`0`，最大级别为`7`。

<details>
  <summary><b>实现细节</b> (点击展示详情)</summary>

Excel内部以点为单位存储行高。默认的分辨率是72DPI或者是96DPI，所以像素和点的大小应该相同。不同的分辨率他们可能不同，因此库分开了这些概念：

即使所有的信息都可用，写入函数也应该遵循下面的优先级顺序：
1)如果`hpx`可用，就使用 `hpx`像素高度。 
2) 如果`hpx`可用，就使用 `hpx`像素高度。
</details>

#### 数字格式

对于每一个单元格而言，`cell.w`的文本来自于`cell.v` 和 `cell.z`格式。如果格式没有指定，Excel`General`格式就会被使用。格式要么是指定的的字符串要么是格式表内的一个索引。解析器应该用数字格式表来填充`workbook.SSF`。写入函数用来序列化这个表。

自定义的工具应该确保本地表的表内有各自的格式字符串。Excel约定规定自定义的格式以索引164开头。下面的例子从头创建了一个自定义的格式：

<details>
  <summary><b>自定义格式的新工作簿</b> (点击显示详情)</summary>

```js
var wb = {
  SheetNames: ["Sheet1"],
  Sheets: {
    Sheet1: {
      "!ref":"A1:C1",
      A1: { t:"n", v:10000 },                    // <-- General format
      B1: { t:"n", v:10000, z: "0%" },           // <-- Builtin format
      C1: { t:"n", v:10000, z: "\"T\"\ #0.00" }  // <-- Custom format
    }
  }
}
```
</details>

这些规则和Excel如何显示自定义的数字格式稍微有些区别。特别是文字字符必须必包含在双引号里面或者在反斜杠之前。更多信息，查看Excel文档`Create or delete a custom number format` 或者是ECMA-376 18.8.31(数字格式)。

<details>
  <summary><b>默认的数字格式</b> (点击展示详情)</summary>

ECMA-376 18.8.30里列出的默认格式：

| ID | Format                     |
|---:|:---------------------------|
|  0 | `General`                  |
|  1 | `0`                        |
|  2 | `0.00`                     |
|  3 | `#,##0`                    |
|  4 | `#,##0.00`                 |
|  9 | `0%`                       |
| 10 | `0.00%`                    |
| 11 | `0.00E+00`                 |
| 12 | `# ?/?`                    |
| 13 | `# ??/??`                  |
| 14 | `m/d/yy` (see below)       |
| 15 | `d-mmm-yy`                 |
| 16 | `d-mmm`                    |
| 17 | `mmm-yy`                   |
| 18 | `h:mm AM/PM`               |
| 19 | `h:mm:ss AM/PM`            |
| 20 | `h:mm`                     |
| 21 | `h:mm:ss`                  |
| 22 | `m/d/yy h:mm`              |
| 37 | `#,##0 ;(#,##0)`           |
| 38 | `#,##0 ;[Red](#,##0)`      |
| 39 | `#,##0.00;(#,##0.00)`      |
| 40 | `#,##0.00;[Red](#,##0.00)` |
| 45 | `mm:ss`                    |
| 46 | `[h]:mm:ss`                |
| 47 | `mmss.0`                   |
| 48 | `##0.0E+0`                 |
| 49 | `@`                        |

</details>

格式14(`m/d/yy`)被Excel本地化：即使文件指明了数字格式，也会根据系统设置用不同的方式绘制。当文件的的生产者和使用者都在相同的区域时这会很有用，不过对于网络上的例子就会不同。为了避免歧义，解析函数接受`dateNF`选项覆盖指定格式字符串的解释。

#### 超链接

超链接存储在单元格对象的`l`关键字内。超链接对象的`Target`字段是连接目标，包括了URI片段。工具提示被存储在`Tooltip`字段内，当移动鼠标到文字上方就会显示。

例如下方的片段在单元格`A3`内创建了一个指向<http://sheetjs.com>的链接，提示信息是`"Find us @ SheetJS.com!"`：

```js
ws['A3'].l = { Target:"http://sheetjs.com", Tooltip:"Find us @ SheetJS.com!" };
```

注意Excel并不会自动为超链接添加样式--他们通常会向普通文本一样显示。

如果链接的目标是一个单元格或者是范围又或者是在相同的工作簿内定义名字("Internal Links")，那么链接的开头会有一个哈希字符标识：

```js
ws['A2'].l = { Target:"#E2" }; /* link to cell E2 */
```

#### 单元格注释

单元格注释是对象，被存储在单元格对象的`c`数组内。实际上注释的内容根据注释的作者被分成了小块。每一个注释对象的`a`字段存储注释的作者，`t`字段是注释的纯文字展示。

例如下面的片段在单元格`A1`内添加了单元格注释：

```js
if(!ws.A1.c) ws.A1.c = [];
ws.A1.c.push({a:"SheetJS", t:"I'm a little comment, short and stout!"});
```
注意：XLSB对作者的名字施加54个字符的限制。名字的长度超过54个字符可能造成其他的格式问题。

把注释标记为普通的隐藏，只需设置`hidden`属性：

```js
if(!ws.A1.c) ws.A1.c = [];
ws.A1.c.push({a:"SheetJS", t:"This comment is visible"});

if(!ws.A2.c) ws.A2.c = [];
ws.A2.c.hidden = true;
ws.A2.c.push({a:"SheetJS", t:"This comment will be hidden"});
```

#### 数据表能见度

Excel支持将表格隐藏在更低的标签栏。表格数据存储文件内，但是UI不容易让它可以使用。标准的隐藏表格会被显示在"Unhide"菜单内。Excel也有"very hidden"表格，这些表格不能被显示在菜单内。只可以通过Vb编辑器访问。

能见度的设置被存储在表格属性数组的`Hidden`属性当中。

<details>
  <summary><b>更多细节</b> (点击显示详情)</summary>

| Value | Definition  |
|:-----:|:------------|
|   0   | Visible     |
|   1   | Hidden      |
|   2   | Very Hidden |

更多详情请查看<https://rawgit.com/SheetJS/test_files/master/sheet_visibility.xlsx>：

```js
> wb.Workbook.Sheets.map(function(x) { return [x.name, x.Hidden] })
[ [ 'Visible', 0 ], [ 'Hidden', 1 ], [ 'VeryHidden', 2 ] ]
```

非Excel格式不支持"Very Hidden"状态。测试一个数据比哦啊是否可见的最好方式是检查是否`Hidden`属性为逻辑truth：

```js
> wb.Workbook.Sheets.map(function(x) { return [x.name, !x.Hidden] })
[ [ 'Visible', true ], [ 'Hidden', false ], [ 'VeryHidden', false ] ]
```
</details>

#### VBA和宏命令

VBA宏命令存储在特殊的数据blob中，当`bookVBA`选项为true时，blob会暴露在工作簿对象的`vbaraw`属性中。VBA宏命令支持 `XLSM`, `XLSB`, 和 `BIFF8 XLS` 格式。如果blob存在于工作簿中，并且和工作簿的名字有关联，支持的格式写入函数会自动插入数据blob。

<details>
	<summary><b>自定义编码名称</b> (点击显示)</summary>

工作簿编码名称存储在`wb.Workbook.WBProps.CodeName`中。默认情况下Excel将会设置成`ThisWorkbook`或者是一个翻译的短语比如`DieseArbeitsmappe`。工作表和图表的编码名称在工作表属性对象的`wb.Workbook.Sheets[i].CodeName`中。宏数据表和对话数据表会被忽略。

读取函数和写入函数会保护编码名称，但是当在一个不同的工作簿内增加一个VBA blob时，编码名称必须被手动设置。

</details>

<details>
	<summary><b>宏数据表</b> (点击显示)</summary>

老版本的Excel也支持非VBA的宏数据表表格类型，宏数据表存储了一些自动命令。他们暴露在`!type`设置成`"macro"`的对象中。
</details>

<details>
	<summary><b>检测工作簿内宏指令</b> (点击显示)</summary>
如果宏指令存在，那么`vbaraw`字段就可以被设置，所以测试简单：

```js
function wb_has_macro(wb/*:workbook*/)/*:boolean*/ {
	if(!!wb.vbaraw) return true;
	const sheets = wb.SheetNames.map((n) => wb.Sheets[n]);
	return sheets.some((ws) => !!ws && ws['!type']=='macro');
}
```
</details>

## 解析选项

导出的`read` 和 `readFile`函数接受选项参数：
| Option Name | Default | Description                                          |
| :---------- | ------: | :--------------------------------------------------- |
|`type`       |         | 输入数据编码 (查看下方的输入类型)           |
|`raw`        | false   | 如果为true，纯文本解析不会解析值 ** |
|`codepage`   |         | 如果指定， 合适的时候使用编码页面**      |
|`cellFormula`| true    | 保存公式到`f`字段                       |
|`cellHTML`   | true    | 解析富文本并把HTML保存到`.h` 字段      |
|`cellNF`     | false   | 把数字格式的字符串保存到 `.z` 字段          |
|`cellStyles` | false   | 把样式/主题保存到 `.s` 字段              |
|`cellText`   | true    | 生成格式化文本`.w` 字段           |
|`cellDates`  | false   | 把日期存储为类型 `d` (默认是 `n`)             |
|`dateNF`     |         | 如果指定，使用代码日期14的字符串 **     |
|`sheetStubs` | false   | 为子单元格创建`z`类型的单元格对象       |
|`sheetRows`  | 0       | 如果`sheetRows`的值 >0, 读取第一个`sheetRows` 行 **  |
|`bookDeps`   | false   | 值为true，解析计算链                   |
|`bookFiles`  | false   | 如果值为true， 添加原始文件到工作簿对象 **  |
|`bookProps`  | false   | 如果值为true， 只有足够的解析才能得到工作簿的元数据**   |
|`bookSheets` | false   | 如果值为true，只有足够的解析才能得到表格名称   |
|`bookVBA`    | false   | 如果值为true，复制 VBA blob 到 `vbaraw` 字段 **          |
|`password`   | ""      | 如果定义了密码并且文件已经加密，就会使用密码 **    |
|`WTF`        | false   | 如果值为true， 对意外的文件特性抛出错误 ** |

- 虽然`cellNF`为false，但是格式化的文本也会被生成并且保存到`.w`字段。
- 在一些情况下，即使`bookSheets`为false，数据表也可能被解析。
- Excel积极尝试从CSV和其他纯文本中解释值。这会导致意外的行为！`raw`选项抑制值解析。
- `bookSheets` 和 `bookProps` 结合起来提供两套信息集合。
- `Deps`将会是一个空对象，如果`bookDeps`为false。
- `bookFiles`的行为依赖于文件类型：
    * `keys` 数组(ZIP里面的路径)用于基于ZIP基础的格式
    * `files`哈希(将路径映射到表示文件的的对象)用于ZIP
    * `cfb`对象用于使用CFB容器的格式
- `sheetRows-1`将会在查看JSON对象输出时生成(因为解析数据时，数据头行会被计算成一行)
- `bookVBA`仅仅在`xl/vbaProject.bin`展示原始的VBA CFB对象。不解析数据。XLSM 和 XLSB把VBA CFB对象存储在`xl/vbaProject.bin`。BIFF8 XLS将VBA条目与核心工作簿条目混合在一起，因此库从XLS CFB容器生成了一个新的XLSB兼容blob。
- `codepage`用于没有`CodePage`记录的BIFF2 - BIFF5文件以及在`type:"binary"`内没有BOM的CSV文件。
- 目前仅支持XOR加密。当文件使用其他的加密方法时会抛出不支持的错误。
- WTF主要用于发展。默认情况下，单一的工作表内，解析器将会抑制读取错误，允许你从解析正确的工作表内读取。设置`WTF:1`强制这些错误被抛出。

### 输入类型

字符串能用很多种方式解释。`read`的`type`参数告诉库如何解析数据参数：

| `type`     | expected input                                                  |
|------------|-----------------------------------------------------------------|
| `"base64"` | 字符串: 文件的Base64编码                             |
| `"binary"` | 字符串: 二进制字符串 (字节 `n` 是 `data.charCodeAt(n)`)        |
| `"string"` | 字符串: JS字符串 (字符被解释为UTF8)              |
| `"buffer"` | nodejs Buffer                                                   |
| `"array"`  | 数组: 8位的无符号型整数数组 (字节 `n` 是 `data[n]`)      |
| `"file"`   | 字符串: 将要被读取的文件的路径 (只在nodejs中可用)            |

### 猜测文件类型

<details>
  <summary><b>实现细节</b> (点击显示详情)</summary>

Excel和其他的电子数据表格工具读取前几个字节并且应用试探法确定稳定类型。这个支持文件类型的双关：用`.xls`扩展的重命名文件会告诉你的电脑使用Excel打开文件而且Excel知道如何去处理它。这个库应用了相似的逻辑：

| Byte 0 | Raw File Type | Spreadsheet Types                                   |
|:-------|:--------------|:----------------------------------------------------|
| `0xD0` | CFB Container | BIFF 5/8 or password-protected XLSX/XLSB or WQ3/QPW |
| `0x09` | BIFF Stream   | BIFF 2/3/4/5                                        |
| `0x3C` | XML/HTML      | SpreadsheetML / Flat ODS / UOS1 / HTML / plain text |
| `0x50` | ZIP Archive   | XLSB or XLSX/M or ODS or UOS2 or plain text         |
| `0x49` | Plain Text    | SYLK or plain text                                  |
| `0x54` | Plain Text    | DIF or plain text                                   |
| `0xEF` | UTF8 Encoded  | SpreadsheetML / Flat ODS / UOS1 / HTML / plain text |
| `0xFF` | UTF16 Encoded | SpreadsheetML / Flat ODS / UOS1 / HTML / plain text |
| `0x00` | Record Stream | Lotus WK\* or Quattro Pro or plain text             |
| `0x7B` | Plain text    | RTF or plain text                                   |
| `0x0A` | Plain text    | SpreadsheetML / Flat ODS / UOS1 / HTML / plain text |
| `0x0D` | Plain text    | SpreadsheetML / Flat ODS / UOS1 / HTML / plain text |
| `0x20` | Plain text    | SpreadsheetML / Flat ODS / UOS1 / HTML / plain text |

DBF文件会基于前几个字节以及第三个和第四个字节进行检测(对应于文件日期的月和天)。

纯文本格式的猜测遵循下面的优先级顺序：

| Format | Test                                                                |
|:-------|:--------------------------------------------------------------------|
| XML    | `<?xml` 出现在前1024个字符                        |
| HTML   | 以 `<`开头并且HTML标签出现在前1024个字符 * |
| XML    | 以 `<`开头                      |
| RTF    | 以 `{\rt`开头                           |
| DSV    | 以`/sep=.$/`开头，分隔符是指定的字符串      |
| DSV    | 前1024个字符中未引用的`";"` 字符比 `"\t"` 或者 `","` 多|
| TSV    | 前1024个字符中未引用的`"\t"`字符比 `","` 多      |
| CSV    | 前1024个字符中的一个是逗号`","`                   |
| ETH    | 以`socialcalc:version:`开头                                   |
| PRN    | (默认)                            |

- HTML 标签包括：`html`, `table`, `head`, `meta`, `script`, `style`, `div`

<details>
  <summary><b>为什么随机的文本文件合法？</b> (点击显示详情)</summary>
Excel在读取文件方面非常积极。添加一个XLS 扩展到任意的显示文件，让Excel认为该文件可能是一个CSV或者是TSV文件，即使它仅仅是一列！这个库尝试去复制那样的行为。

最好的方法是去校验想要得到的工作表并且确保它有期待的行数或列数。提取范围非常简单：

```js
var range = XLSX.utils.decode_range(worksheet['!ref']);
var ncols = range.e.c - range.s.c + 1, nrows = range.e.r - range.s.r + 1;
```

</details>

## 写入选项

导出的`write` 和 `writeFile`函数接受一个选型参数：

| Option Name |  Default | Description                                         |
| :---------- | -------: | :-------------------------------------------------- |
|`type`       |          | 输出数据编码(查看下面的输出类型)        |
|`cellDates`  |  `false` | 把字节存储为类型`d` (默认是 `n`)            |
|`bookSST`    |  `false` | 生成共享的字符串表格 **                     |
|`bookType`   | `"xlsx"` | 工作簿的类型 (查看下方支持的格式)  |
|`sheet`      |     `""` | 单页格式的工作表名称 **       |
|`compression`|  `false` | 基于ZIP的格式使用ZIP压缩 **        |
|`Props`      |          | 写入时覆盖工作簿的属性 **        |
|`themeXLSX`  |          | 写入XLSX/XLSB/XLSM时，覆盖主题XML **   |
|`ignoreEC`   |   `true` | 禁止"数字作为文本"错误 **                 |

- `bookSST` 较慢并且有更多的内存密集型，不过与iOS数字老版本有更好的兼容性。
- 原始数据时唯一保证存储的东西。在README文件中没有描述的功能可能无法序列化。
- `cellDates` 仅用于XLSX输出并且不能保证与第三方读取器一起工作。Excel自身不经常用类型`d`编写单元格，因此非Excel工具会忽视数据或者是有日期错误。
- `Props`是一个备份工作簿`Props`字段的对象。从[工作簿文件属性](#workbook-file-properties) 部分查看表格。
- 如果指定，来自`themeXLSX`的字符串会被存储为XLSX/XLSB/XLSM文件的基本主题(ZIP中的`xl / theme / theme1.xml`)。
- 由于在程序中有一个bug，一些功能比如"分列"会在忽略错误条件的工作表上使Excel崩溃。默认情况下写入函数将会标记文件忽略错误。设置`ignoreEC`为`false`来禁止。

### 支持的输出格式

与第三方工具的广泛兼容性，这个库支持很多种输出格式。明确的文件类型被`bookType`选项控制：

| `bookType` | file ext | container | sheets | Description                     |
| :--------- | -------: | :-------: | :----- |:------------------------------- |
| `xlsx`     | `.xlsx`  |    ZIP    | multi  | Excel 2007+ XML Format(XML 格式)   |
| `xlsm`     | `.xlsm`  |    ZIP    | multi  | Excel 2007+ Macro XML Format(宏 XML 格式)    |
| `xlsb`     | `.xlsb`  |    ZIP    | multi  | Excel 2007+ Binary Format(二进制格式)     |
| `biff8`    | `.xls`   |    CFB    | multi  | Excel 97-2004 Worksheet Format(工作簿格式)   |
| `biff5`    | `.xls`   |    CFB    | multi  | Excel 5.0/95 Worksheet Format(工作簿格式)    |
| `biff2`    | `.xls`   |   none    | single | Excel 2.0 Worksheet Format(工作簿格式)      |
| `xlml`     | `.xls`   |   none    | multi  | Excel 2003-2004 (SpreadsheetML) |
| `ods`      | `.ods`   |    ZIP    | multi  | OpenDocument Spreadsheet(开放文档格式的电子表格) |
| `fods`     | `.fods`  |   none    | multi  | Flat OpenDocument Spreadsheet(平滑的开放文档格式的电子表格)   |
| `csv`      | `.csv`   |   none    | single | Comma Separated Values(逗号分隔值)          |
| `txt`      | `.txt`   |   none    | single | UTF-16 Unicode Text (TXT)       |
| `sylk`     | `.sylk`  |   none    | single | Symbolic Link (SYLK)            |
| `html`     | `.html`  |   none    | single | HTML Document                   |
| `dif`      | `.dif`   |   none    | single | Data Interchange Format (DIF) (数据交换格式)  |
| `dbf`      | `.dbf`   |   none    | single | dBASE II + VFP Extensions (DBF)(dBASE II + VFP扩展) |
| `rtf`      | `.rtf`   |   none    | single | Rich Text Format (RTF)          |
| `prn`      | `.prn`   |   none    | single | Lotus Formatted Text(Lotus格式化文本。)         |
| `eth`      | `.eth`   |   none    | single | Ethercalc Record Format (ETH)(Ethercalc记录格式)   |

- `compression`仅用于带ZIP容器的格式。
- 格式只支持需要`sheet`选型指明工作表的单表。如果字符串为空，就会使用第一张工作表。
- 如果`bookType`未指定值，那么`writeFile`会自动根据文件扩展名来猜测输出文件格式。他就会在上表中选择匹配扩展名的第一个格式。

### 输出类型

`write`函数的`type`参数备份`read`函数的`type`参数：
| `type`     | output                                                          |
|------------|-----------------------------------------------------------------|
| `"base64"` | 字符串: 文件的Base64编码                             |
| `"binary"` | 字符串: 二进制字符串 (字节 `n` 是 `data.charCodeAt(n)`)        |
| `"string"` | 字符串: JS 字符串 (字符被解释成UTF8)              |
| `"buffer"` | nodejs Buffer                                                   |
| `"array"`  | ArrayBuffer, 8位无符号整数的回退数组              |
| `"file"`   | 字符串: 将要创建的文件的地址(仅用于nodejs)       |

## 工具函数

`sheet_to_*`函数接受一张工作表以及一个可选的选项对象。
`*_to_sheet`函数接受一个数据对象以及一个可选的选项对象。
示例都是基于下面的工作表：

```
XXX| A | B | C | D | E | F | G |
---+---+---+---+---+---+---+---+
 1 | S | h | e | e | t | J | S |
 2 | 1 | 2 | 3 | 4 | 5 | 6 | 7 |
 3 | 2 | 3 | 4 | 5 | 6 | 7 | 8 |
```

### 数组的数组输入

`XLSX.utils.aoa_to_sheet`获取JS值数组的数组，并且返回一个工作表寻找输入数据。Numbers，Booleans和Strings都被存储为相应的样式。Date被存储为date或者是numbers。跳过数组孔和显式“未定义”值。`null`值可能被剔除。所有其它值存储为字符串。函数获取选项参数：

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`dateNF`     |  FMT 14  | 字符串输出使用特定的日期格式          |
|`cellDates`  |  false   | 存储日期为类型 `d` (默认是 `n`)            |
|`sheetStubs` |  false   | 为`null`值创建类型为`z`的单元格对象  |

<details>
  <summary><b>例子</b> (点击显示)</summary>
生成实例表：

```js
var ws = XLSX.utils.aoa_to_sheet([
  "SheetJS".split(""),
  [1,2,3,4,5,6,7],
  [2,3,4,5,6,7,8]
]);
```
</details>

`XLSX.utils.sheet_add_aoa`获取JS值的数组的数组，并且更新一个已存在的工作表对象。它遵循和`aoa_to_sheet`一样的过程，并且接受一个选项参数：

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`dateNF`     |  FMT 14  | 字符串输出使用指定的日期格式          |
|`cellDates`  |  false   | 存储日期为类型 `d` (默认是 `n`)            |
|`sheetStubs` |  false   | 为`null`值创建类型为`z`的单元格对象   |
|`origin`     |          | 只用指定的单元格作为指定的起点 (查看下表)    |

`origin`应该是以下之一：

| `origin`         | Description                                               |
| :--------------- | :-------------------------------------------------------- |
| (cell object)    | 使用指定的单元格 (单元格对象)                          |
| (string)         | 使用指定的单元格 (A1样式的单元格)                        |
| (number >= 0)    | 从指定行的第一列开始 (0索引)  |
| -1               | 从第一列开始添加到工作表底部    |
| (default)        | 从单元格A1开始                                      |

<details>
  <summary><b>示例</b> (点击显示)</summary>

考虑工作表：

```
XXX| A | B | C | D | E | F | G |
---+---+---+---+---+---+---+---+
 1 | S | h | e | e | t | J | S |
 2 | 1 | 2 |   |   | 5 | 6 | 7 |
 3 | 2 | 3 |   |   | 6 | 7 | 8 |
 4 | 3 | 4 |   |   | 7 | 8 | 9 |
 5 | 4 | 5 | 6 | 7 | 8 | 9 | 0 |
```

此工作表可按照顺序`A1:G1, A2:B4, E2:G4, A5:G5`构建：

```js
/* Initial row */
var ws = XLSX.utils.aoa_to_sheet([ "SheetJS".split("") ]);

/* Write data starting at A2 */
XLSX.utils.sheet_add_aoa(ws, [[1,2], [2,3], [3,4]], {origin: "A2"});

/* Write data starting at E2 */
XLSX.utils.sheet_add_aoa(ws, [[5,6,7], [6,7,8], [7,8,9]], {origin:{r:1, c:4}});

/* Append row */
XLSX.utils.sheet_add_aoa(ws, [[4,5,6,7,8,9,0]], {origin: -1});
```
</details>

### 对象数组输入

`XLSX.utils.json_to_sheet`获取对象数组并且返回一张基于对象自动生成"headers"的工作表。默认的列顺序由第一次出现的字段决定，这些字段通过使用`Object.keys`得到，不过可以使用选项参数覆盖。

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`header`     |          | 使用指定的列顺序 (默认 `Object.keys`)  |
|`dateNF`     |  FMT 14  | 字符串输出使用指定的日期格式          |
|`cellDates`  |  false   | 存储日期为类型 `d` (默认是 `n`)            |
|`skipHeader` |  false   | 如果值为true, 输出不包含header行       |

<details>
  <summary><b>示例</b> (点击显示)</summary>

原始的表单不能以明显的方法复制，因为JS对象的keys必须是独一无二的。之后用`e_1` 和 `S_1`替换第二个`e` 和 `S`。

```js
var ws = XLSX.utils.json_to_sheet([
  { S:1, h:2, e:3, e_1:4, t:5, J:6, S_1:7 },
  { S:2, h:3, e:4, e_1:5, t:6, J:7, S_1:8 }
], {header:["S","h","e","e_1","t","J","S_1"]});
```

或者可以跳过header行：

```js
var ws = XLSX.utils.json_to_sheet([
  { A:"S", B:"h", C:"e", D:"e", E:"t", F:"J", G:"S" },
  { A: 1,  B: 2,  C: 3,  D: 4,  E: 5,  F: 6,  G: 7  },
  { A: 2,  B: 3,  C: 4,  D: 5,  E: 6,  F: 7,  G: 8  }
], {header:["A","B","C","D","E","F","G"], skipHeader:true});
```
</details>

`XLSX.utils.sheet_add_json`获取一个对象数组，并且更新一个已存在的工作表对象。与`json_to_sheet`一样有相同的过程，并且接受一个选项参数：

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`header`     |          | 使用指定的列排序 (默认 `Object.keys`)  |
|`dateNF`     |  FMT 14  | 字符串输出使用指定的日期格式      |
|`cellDates`  |  false   | 把存储日期为类型 `d` (默认是 `n`)            |
|`skipHeader` |  false   | 如果值为true, 输出不包含header行        |
|`origin`     |          | 使用指定的单元格作为起点 (查看下方表格)    |

`origin`应该是以下之一：

| `origin`         | Description                                               |
| :--------------- | :-------------------------------------------------------- |
| (cell object)    | 使用指定的单元格(单元格对象)                          |
| (string)         | 使用指定的单元格 (A1样式的单元格)                        |
| (number >= 0)    | 从指定行的第一列开始(0索引)  |
| -1               | 从第一列开始添加到工作表底部  |
| (default)        | 从单元格A1开始                                        |

<details>
  <summary><b>例子</b> (点击展示)</summary>
考虑工作表：

```
XXX| A | B | C | D | E | F | G |
---+---+---+---+---+---+---+---+
 1 | S | h | e | e | t | J | S |
 2 | 1 | 2 |   |   | 5 | 6 | 7 |
 3 | 2 | 3 |   |   | 6 | 7 | 8 |
 4 | 3 | 4 |   |   | 7 | 8 | 9 |
 5 | 4 | 5 | 6 | 7 | 8 | 9 | 0 |
```
工作表能够以`A1:G1, A2:B4, E2:G4, A5:G5`顺序构建：

```js
/* Initial row */
var ws = XLSX.utils.json_to_sheet([
  { A: "S", B: "h", C: "e", D: "e", E: "t", F: "J", G: "S" }
], {header: ["A", "B", "C", "D", "E", "F", "G"], skipHeader: true});

/* Write data starting at A2 */
XLSX.utils.sheet_add_json(ws, [
  { A: 1, B: 2 }, { A: 2, B: 3 }, { A: 3, B: 4 }
], {skipHeader: true, origin: "A2"});

/* Write data starting at E2 */
XLSX.utils.sheet_add_json(ws, [
  { A: 5, B: 6, C: 7 }, { A: 6, B: 7, C: 8 }, { A: 7, B: 8, C: 9 }
], {skipHeader: true, origin: { r: 1, c: 4 }, header: [ "A", "B", "C" ]});

/* Append row */
XLSX.utils.sheet_add_json(ws, [
  { A: 4, B: 5, C: 6, D: 7, E: 8, F: 9, G: 0 }
], {header: ["A", "B", "C", "D", "E", "F", "G"], skipHeader: true, origin: -1});
```

</details>

### HTML Table 输入

`XLSX.utils.table_to_sheet`获取一个table DOM元素，并且返回一个工作表寻找输入的table。Numbers会被解析。所有的数据将会被存储为字符串。

`XLSX.utils.table_to_book`基于工作表会产生一个最小的工作簿。

两个函数接受选项参数：

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`raw`        |          | 如果值为true,  每一个单元格将会保存原始的字符串          |
|`dateNF`     |  FMT 14  | 字符串输出使用指定的日期格式          |
|`cellDates`  |  false   | 把日期存储为类型 `d` (默认是 `n`)            |
|`sheetRows`  |    0     | 如果值 >0, 读取表格的第一个`sheetRows`行 |
|`display`    |  false   | 如果值为true, 隐藏的行和单元格将不会被解析   |

<details>
  <summary><b>例子</b> (点击显示)</summary>

生成示例表单，以HTML table开始：

```html
<table id="sheetjs">
<tr><td>S</td><td>h</td><td>e</td><td>e</td><td>t</td><td>J</td><td>S</td></tr>
<tr><td>1</td><td>2</td><td>3</td><td>4</td><td>5</td><td>6</td><td>7</td></tr>
<tr><td>2</td><td>3</td><td>4</td><td>5</td><td>6</td><td>7</td><td>8</td></tr>
</table>
```

处理表格:

```js
var tbl = document.getElementById('sheetjs');
var wb = XLSX.utils.table_to_book(tbl);
```
</details>

注意：`XLSX.read`能够处理表示为字符串的HTML。

### 公式输出

`XLSX.utils.sheet_to_formulae`生成一个命令数组，命令显示了一个人会怎样进入一个应用。每一个入口都是表格`A1-cell-address=formula-or-value`。字符串文字以"`"为前缀，符合Excel。

<details>
  <summary><b>例子</b> (点击显示)</summary>

示例表：

```js
> var o = XLSX.utils.sheet_to_formulae(ws);
> [o[0], o[5], o[10], o[15], o[20]];
[ 'A1=\'S', 'F1=\'J', 'D2=4', 'B3=3', 'G3=8' ]
```
</details>

### 定界分隔符输出

作为一个`writeFile` CSV 类型的替代，`XLSX.utils.sheet_to_csv`也会产生CSV输出。这个函数获取一个选项参数：

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`FS`         |  `","`   | "字段分隔符"表示字段之间的分隔符         |
|`RS`         |  `"\n"`  | "记录分隔符"表示行之间的分隔符           |
|`dateNF`     |  FMT 14  | 字符串输出使用指定的日期格式         |
|`strip`      |  false   | 删除每条记录中的尾随字段分隔符**  |
|`blankrows`  |  true    | 包含CSV输出的空白行              |
|`skipHidden` |  false   | 跳过CSV输出的隐藏行/列       |

- `strip`将删除默认`FS / RS`下每行的尾随逗号
`blankrows`必须设置为`false`才能跳过空白行。

<details>
  <summary><b>例子</b> (点击显示)</summary>

示例表：

```js
> console.log(XLSX.utils.sheet_to_csv(ws));
S,h,e,e,t,J,S
1,2,3,4,5,6,7
2,3,4,5,6,7,8
> console.log(XLSX.utils.sheet_to_csv(ws, {FS:"\t"}));
S	h	e	e	t	J	S
1	2	3	4	5	6	7
2	3	4	5	6	7	8
> console.log(XLSX.utils.sheet_to_csv(ws,{FS:":",RS:"|"}));
S:h:e:e:t:J:S|1:2:3:4:5:6:7|2:3:4:5:6:7:8|
```
</details>

#### #### UTF-16 Unicode 文本

`txt`输出类型使用tab字符作为字段分隔符。如果`codepage`可用(包含全部的分发但不是核心)，输出将会被编码为`CP1200`并且BOM会被预置。

`XLSX.utils.sheet_to_txt`获取和`sheet_to_csv`一样的参数。

### HTML 输出

作为' writeFile ' HTML类型的替代，`XLSX.utils.sheet_to_html`也会生成HTML输出。这个函数接受一个选项参数：

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`id`         |          | 为`TABLE`元素指定 `id` 特性 |
|`editable`   |  false   | 如果值为true, 为每一个TD设置 `contenteditable="true"`|
|`header`     |          | 覆盖 header (默认 `html body`)               |
|`footer`     |          | 覆盖 footer (默认 `/body /html`)             |

<details>
  <summary><b>例子</b> (点击显示)</summary>

示例表格：

```js
> console.log(XLSX.utils.sheet_to_html(ws));
// ...
```
</details>

### JSON

`XLSX.utils.sheet_to_json`生成不同类型的JS对象。这个函数接受一个选项参数：

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`raw`        | `true`   | 使用原生值 (true) 或者格式化字符串 (false)  |
|`range`      | from WS  | 覆盖 Range (查看下面的table)                    |
|`header`     |          | 控制输出格式 (查看下面的table)             |
|`dateNF`     |  FMT 14  | 字符串输出使用指定的日期格式          |
|`defval`     |          | 使用指定的值替代null或者undefined  |
|`blankrows`  |    **    | 包含输出的空白行 **                |

- `raw`只影响有格式编码(`.z`)字段或者格式化文本(`.w`)字段的单元格。
- 如果`header`被指定，第一行就会被当做数据行；如果`header`未指定，第一行是header并且不作为数据。
- 当header未指定时，转换将通过添加"_"和一个从"1"开始的计数自动消除标题条目的歧义。例如有三列的标题都是`foo`，那么输出字段是`foo1`，`foo_1`，`foo_2`。
- 当`raw`值为true时返回`null`，`raw`值为false会被跳过。
- 如果`defval`没有指定，通常`null`和`undefined`会被跳过。如果`defval`有指定值，所有的`null`和`undefined`嗲都将会用`defval`填充。
- 当`header`为`1`时，默认生成空白行。`blankrows`必须设置为`false`来跳过空白行。
- 当`header`不为`1`时，默认跳过空白行。`blankrows`必须设置为`true`来生成空白行。

`range`是以下之一：

| `range`          | Description       |
| :--------------- | :---------------- |
| (number)         | 使用工作表范围，但将起始行设置为值 |
| (string)         | 使用指定的范围 (A1类型的有界范围的字符串) |
| (default)        | 使用工作表范围 (`ws['!ref']`)     |

`header`是以下之一：

| `header`         | Description     |
| :--------------- | :-------------- |
| `1`              | 生成数组类型的数组 ("二维数组")   |
| `"A"`            | 行对象的键是文字的列标题           |
| array of strings | 在行对象内使用指定的字符串作为键   |
| (default)        | 读取并消除第一行的歧义作为键   |

如果`header`不为`1`，行对象将会包含不可枚举的属性``__rowNum__`，这个属性代表与条目相对应的工作表的行。

<details>
  <summary><b>示例</b> (点击显示)</summary>
示例表：

```js
> XLSX.utils.sheet_to_json(ws);
[ { S: 1, h: 2, e: 3, e_1: 4, t: 5, J: 6, S_1: 7 },
  { S: 2, h: 3, e: 4, e_1: 5, t: 6, J: 7, S_1: 8 } ]

> XLSX.utils.sheet_to_json(ws, {header:"A"});
[ { A: 'S', B: 'h', C: 'e', D: 'e', E: 't', F: 'J', G: 'S' },
  { A: '1', B: '2', C: '3', D: '4', E: '5', F: '6', G: '7' },
  { A: '2', B: '3', C: '4', D: '5', E: '6', F: '7', G: '8' } ]

> XLSX.utils.sheet_to_json(ws, {header:["A","E","I","O","U","6","9"]});
[ { '6': 'J', '9': 'S', A: 'S', E: 'h', I: 'e', O: 'e', U: 't' },
  { '6': '6', '9': '7', A: '1', E: '2', I: '3', O: '4', U: '5' },
  { '6': '7', '9': '8', A: '2', E: '3', I: '4', O: '5', U: '6' } ]

> XLSX.utils.sheet_to_json(ws, {header:1});
[ [ 'S', 'h', 'e', 'e', 't', 'J', 'S' ],
  [ '1', '2', '3', '4', '5', '6', '7' ],
  [ '2', '3', '4', '5', '6', '7', '8' ] ]
```

展示`row`效果的例子：

```js
> ws['A2'].w = "3";                          // set A2 formatted string value

> XLSX.utils.sheet_to_json(ws, {header:1, raw:false});
[ [ 'S', 'h', 'e', 'e', 't', 'J', 'S' ],
  [ '3', '2', '3', '4', '5', '6', '7' ],     // <-- A2 uses the formatted string
  [ '2', '3', '4', '5', '6', '7', '8' ] ]

> XLSX.utils.sheet_to_json(ws, {header:1});
[ [ 'S', 'h', 'e', 'e', 't', 'J', 'S' ],
  [ 1, 2, 3, 4, 5, 6, 7 ],                   // <-- A2 uses the raw value
  [ 2, 3, 4, 5, 6, 7, 8 ] ]
```
</details>

## 文件格式

虽然库的名称是`xlsx`，不过它支持多种电子表格文件格式：

| Format                                                       | Read  | Write |
|:-------------------------------------------------------------|:-----:|:-----:|
| **Excel Worksheet/Workbook Formats**                         |:-----:|:-----:|
| Excel 2007+ XML Formats (XLSX/XLSM)                          |  :o:  |  :o:  |
| Excel 2007+ Binary Format (XLSB BIFF12)                      |  :o:  |  :o:  |
| Excel 2003-2004 XML Format (XML "SpreadsheetML")             |  :o:  |  :o:  |
| Excel 97-2004 (XLS BIFF8)                                    |  :o:  |  :o:  |
| Excel 5.0/95 (XLS BIFF5)                                     |  :o:  |  :o:  |
| Excel 4.0 (XLS/XLW BIFF4)                                    |  :o:  |       |
| Excel 3.0 (XLS BIFF3)                                        |  :o:  |       |
| Excel 2.0/2.1 (XLS BIFF2)                                    |  :o:  |  :o:  |
| **Excel支持的文本格式**                             |:-----:|:-----:|
| Delimiter-Separated Values(定界分隔符的值 ) (CSV/TXT)                         |  :o:  |  :o:  |
| Data Interchange Format(数据交换格式 (DIF)                                |  :o:  |  :o:  |
| Symbolic Link(符号链接) (SYLK/SLK)                                     |  :o:  |  :o:  |
| Lotus Formatted Text(lotus格式的文本) (PRN)                                   |  :o:  |  :o:  |
| UTF-16 Unicode Text (TXT)                                    |  :o:  |  :o:  |
| **其他工作簿/工作表格式**                         |:-----:|:-----:|
| OpenDocument Spreadsheet(开放文档格式的电子表格) (ODS)                               |  :o:  |  :o:  |
| Flat XML ODF Spreadsheet (FODS)                              |  :o:  |  :o:  |
| Uniform Office Format Spreadsheet (标文通 UOS1/UOS2)         |  :o:  |       |
| dBASE II/III/IV / Visual FoxPro (DBF)                        |  :o:  |  :o:  |
| Lotus 1-2-3 (WKS/WK1/WK2/WK3/WK4/123)                        |  :o:  |       |
| Quattro Pro Spreadsheet (WQ1/WQ2/WB1/WB2/WB3/QPW)            |  :o:  |       |
| **其他常用的电子表格输出格式**                  |:-----:|:-----:|
| HTML Tables                                                  |  :o:  |  :o:  |
| Rich Text Format tables(富文本格式表) (RTF)                                |       |  :o:  |
| Ethercalc Record Format(Ethercalc记录格式) (ETH)                                |  :o:  |  :o:  |
不会写入给定文件格式不支持的功能。具有范围限制的格式将会被静默截断：

| Format                                    | Last Cell  | Max Cols | Max Rows |
|:------------------------------------------|:-----------|---------:|---------:|
| Excel 2007+ XML Formats (XLSX/XLSM)       | XFD1048576 |    16384 |  1048576 |
| Excel 2007+ Binary Format (XLSB BIFF12)   | XFD1048576 |    16384 |  1048576 |
| Excel 97-2004 (XLS BIFF8)                 | IV65536    |      256 |    65536 |
| Excel 5.0/95 (XLS BIFF5)                  | IV16384    |      256 |    16384 |
| Excel 2.0/2.1 (XLS BIFF2)                 | IV16384    |      256 |    16384 |

Excel 2003 电子表格的范围限制被Excel的版本控制，并且不会写入函数强制执行。

### Excel 2007+ XML (XLSX/XLSM)

<details>
  <summary>(点击显示)</summary>

XlSX和XLSM文件是ZIP容器包含的与开源打包约定(Open Packaging Conventions, OPC)一致的一系列文件。大多数XLSM格式与XLSX相同，被用作文件包含宏命令。

这个格式在ECMA-376以及随后的ISO/IEC 29500都进行了标准化。Excel没有加遵循这个规范，并且还有其他文件讨论Excel如何偏离规范。

</details>

### Excel 2.0-95 (BIFF2/BIFF3/BIFF4/BIFF5)

<details>
  <summary>(点击显示)</summary>

BIFF 2/3 XLS是二进制记录的单表流。Excel 4介绍了工作簿的原理，除了有单表的`XLS`格式。结构与Lotus 1-2-3文件格式非常相似。BIFF5/8/12用了多种方式扩展格式，不过很大程度上保持了相同的记录格式。

对于这些格式没有官方的规范。Excel 95在这些格式里面可以写入文件，因此记录的长度以及字段都是由所有支持的格式以及比较文件决定。Excel 2016可以生成BIFF5文件，从XLSX或BIFF2开始启用全套文件测试。

</details>

### Excel 97-2004 Binary (BIFF8)

<details>
  <summary>(点击显示)</summary>

BIFF8仅仅使用混合的文件二进制容器格式，将一些内容放在文件的流内。在它的核心，它将会使用来自BIFF的老版本的二进制记录格式的扩展版本。

`MS-XLS`规范覆盖了文件格式的基础，并且其他的规范扩展了属性(如特性)的规范化。

</details>

### Excel 2003-2004 (SpreadsheetML)

<details>
  <summary>(点击显示)</summary>

在XLSX之前，SpreadsheetML文件是简单的XML文件。没有官方的并且全面的的规范，虽然MS对于这种格式有发布的文档。因此Excel 2016 生成了电子表格文件，映射功能非常简单。

</details>

### Excel 2007+ Binary (XLSB, BIFF12)

<details>
  <summary>(点击显示)</summary>

XLSB格式与XLSX并行引入，将BIFF架构与内容分离和XLSX的ZIP容器相结合。XLSX子文件的大部分节点能用一个相应的子文件映射到XLSB记录中去。

`MS-XLSB`规范包含了文件格式的基础，并且其他的规范扩展了属性(如特性)的序列化。

</details>

### 定界分隔符的值 (CSV/TXT)

<details>
  <summary>(点击显示)</summary>

Excel CSV在许多重要的方法上背离了RFC4180。生成的CSV文件通常应该运行在Excel内，但是他们不能运行RFC4180兼容的读取器中。解析器通常理解Excel CSV。如果值不可用，写入器会为公式主动生成单元格。

Excel TXT 使用tab作为分隔符，编码页1200。

注意：

- 像在Excel中，以`0x49 0x44 ("ID")`开始的文件会被当做是符号链接(Symbolic Link)文件。不像Excel，如果文件没有一个有效的SYLK标题，他将会被主动解释为SYLK文件。为了广泛的兼容性，所有值为`ID`的单元格会自动用双引号包裹。

</details>

### 其他工作簿格式

<details>
  <summary>(点击显示)</summary>
对其他格式的支持通常远远超出XLS / XLSB / XLSX支持，这在很大程度上是由于缺乏公开可用的文档。文件通常是在各自的应用内产生，并且会与他们的导出文件相比较以确定结构。主要的关注点是数据提出。

</details>

#### Lotus 1-2-3 (WKS/WK1/WK2/WK3/WK4/123)

<details>
  <summary>(点击显示)</summary>

Lotus格式由与BIFF结构相似的二进制记录组成。Lotus几十年前发布了一份包含原始WK1格式的规范。通过生成文件和与Excel支持进行比较来推断其他功能。

</details>

#### Quattro Pro (WQ1/WQ2/WB1/WB2/WB3/QPW)

<details>
  <summary>(点击显示)</summary>

Quattro Pro格式使用与BIFF和Lotus一样的二进制记录。一些较新的格式(命名为WB3 和 QPW)使用像BIFF8 XLS一样的CFB附件。

</details>

#### OpenDocument Spreadsheet(开放文档格式的电子表格) (ODS/FODS)

<details>
  <summary>(点击显示)</summary>

ODS是一种类似于XLSX的XML-in-ZIP格式，而FODS是一种类似于SpreadsheetML的XML格式。两种格式都在OASIS标准中进行了详细的说明，不过像LO/OO工具被添加到了未公开的扩展中。解析器和编写器并没有实现全部的标准，反而重点实现了提取和存储行数据中重要的部分。

</details>

#### Uniform Office Spreadsheet(统一办公电子表格) (UOS1/2)

<details>
  <summary>(点击显示)</summary>

UOS是一种非常相似的格式，并且它有2个变种，分别对应ODS和FODS。大多数情况下，格式之间的区别是标签和属性的名称。

</details>

### Other Single-Worksheet Formats(其他单一工作表格式)

大多数较老的浏览器仅支持一种工作表：

#### dBASE and Visual FoxPro (DBF)

<details>
  <summary>(点击显示)</summary>

DBF实际上是一种类型化的表格格式：每一列只能保存一种数据类型，并且每条记录忽略类型信息。解析器生成标题行并且在工作表的第二行开始插入记录。编写器让文件和Visual FoxPro兼容。

多文件的扩展，比如内部示例和表格，目前不支持，会被在web浏览器中读取任意文件的普通能力所限制。读取器理解DBF level 7的扩展，比如DATETIME。

</details>

#### Symbolic Link(符号链接) (SYLK)

<details>
  <summary>(点击显示)</summary>
没有真正的文档。通过各种版本的Excel中保存文件来收集所有知识，以推断出字段的含义。注意：

- 简单的公式被存储在RC表单中。
- 列宽会被四舍五入成完整的字符。

</details>

#### Lotus Formatted Text (PRN)

<details>
  <summary>(点击显示)</summary>
没有真正的文档。事实上Excel把PRN视为一种只能输出的文件格式。然而我们能够猜测列宽并且反向还原原始布局。Excel 240个字符宽度的限制不会被强制执行。

</details>

#### Data Interchange Format(数据交换格式) (DIF)

<details>
  <summary>(点击显示)</summary>

没有统一标准的定义。 Visicalc DIF与Lotus DIF不同，并且两者都与Excel DIF不一样。在不明确的情况下，解析器/编写器遵循Excel中的预期行为。特别地，Excel以不兼容的方式扩展DIF：

- 由于Excel自动将数字字符串转换为数字，数字的字符串常量被转换成公式：`"0.3" -> "=""0.3""`
- DIF技术上期待数字的单元格保存原始的数字数据，不过Excel允许格式化数字(包括日期)。
- DIF技术上不支持公式，但是Excel将会转换简单公式。数组公式没有保存。
</details>

#### HTML
<details>
  <summary>(点击显示)</summary>
Excel HTML工作表包含以样式编码的特殊元数据。例如`mso-number-format`是一个包含数字格式的本地化字符串。尽管元数据的输出是有效的HTML，但是他不接受空的`&`符号。

编写器通过`t`标签添加类型元数据到TD元素中去。解析器检查这些标签，并且覆盖默认的解释。例如文本`<td>12345</td>`将会被解析成数字，不过`<td t="s">12345</td>`将会被解析成文本。
</details>

#### Rich Text Format(富文本格式) (RTF)

<details>
  <summary>(点击显示)</summary>

当复制工作表内的单元格或者范围时，Excel RTF工作表会被存储在剪贴板内。支持的编码是单词RTF支持的一个子集。

</details>

#### Ethercalc Record Format (ETH)

<details>
  <summary>(点击显示)</summary>

[Ethercalc](https://ethercalc.net/)是一种开源的web电子表格，由记录格式驱动，让人联想到包含在MIME多部分消息中的SYLK。

</details>

## 测试

### Node

<details>
  <summary>(点击显示)</summary>
`make test`将会运行node基础的测试。默认情况下，它以各种支持的格式对文件运行测试。要测试一种指定的文件类型，设置`FMTS`为你想要测试的类型。使用`make test_misc`可以获得指定功能的测试。

```bash
$ make test_misc   # run core tests
$ make test        # run full tests
$ make test_xls    # only use the XLS test files
$ make test_xlsx   # only use the XLSX test files
$ make test_xlsb   # only use the XLSB test files
$ make test_xml    # only use the XML test files
$ make test_ods    # only use the ODS test files
```

要想启用所有的错误，请设置环境变量`WTF=1`：

```bash
$ make test        # run full tests
$ WTF=1 make test  # enable all error messages
```

`flow` and `eslint` checks are available:

```bash
$ make lint        # eslint checks
$ make flow        # make lint + Flow checking
$ make tslint      # check TS definitions
```

</details>

### 浏览器

<details>
  <summary>(点击显示)</summary>
核心浏览器内测试可在此repo中的`tests/index.html`中找到。启动一个本地服务器并且导航到那个目录去运行测试。`make ctestserv`将会在8080端口启动一个服务。

`make ctest`将生成浏览器装置。要添加更多的文件，编辑`tests/fixtures.lst`并且添加路径。

要运行完整的浏览器内测试，从[`oss.sheetjs.com`](https://github.com/SheetJS/SheetJS.github.io)克隆这个repo，并且替换`xlsx.js`文件(然后打开一个浏览器窗口跳转到`stress.html`)。

```bash
$ cp xlsx.js ../SheetJS.github.io
$ cd ../SheetJS.github.io
$ simplehttpserver # or "python -mSimpleHTTPServer" or "serve"
$ open -a Chromium.app http://localhost:8000/stress.html
```
</details>

### 测试环境

<details>
  <summary>(点击显示)</summary>

- NodeJS `0.8`, `0.10`, `0.12`, `4.x`, `5.x`, `6.x`, `7.x`, `8.x`
- IE 6/7/8/9/10/11 (IE 6-9 require shims)
- Chrome 24+ (including Android 4.0+)
- Safari 6+ (iOS and Desktop)
- Edge 13+, FF 18+, and Opera 12+

测试使用mocha测试框架。Travis-CI 和 Sauce Labs 链接：

- <https://travis-ci.org/SheetJS/js-xlsx> 用于nodejs内的XLSX模块
- <https://semaphoreci.com/sheetjs/js-xlsx> 用于nodejs内的XLSX模块
- <https://travis-ci.org/SheetJS/SheetJS.github.io> 用于 XLS\* 模块
- <https://saucelabs.com/u/sheetjs> 用于使用Sauce Labs的 XLS\* 模块

Travis-CI测试组合也包括用于多种时区的测试。改变本地的时区，设置TZ环境可用：

```bash
$ env TZ="Asia/Kolkata" WTF=1 make test_misc
```
</details>

### 测试文件

测试文件被封装在[另一个仓库](https://github.com/SheetJS/test_files)。

运行`make init`将会刷新`test_files`子模块并获取子模块的文件。注意这个可能需要`svn`, `git`, `hg`以及其它可能不可用的命令。如果`make init`失败，请从[仓库](https://github.com/SheetJS/test_files/releases)下载测试文件快照的最新版本。

<details>
  <summary><b>最新快照</b> (点击显示)</summary>
最新的测试文件快照：
<http://github.com/SheetJS/test_files/releases/download/20170409/test_files.zip>

(下载并解压到`test_files`子目录)
</details>

## 贡献

由于开放规范承诺的不稳定性，确保代码是洁净室非常重要。[贡献记录](CONTRIBUTING.md)


<details>
  <summary><b>文件组织</b> (点击显示)</summary>
在最高级别，最终脚本是`bits`文件夹中各个文件的串联。运行`make`应该在所有平台上重现最终输出。同样，README被分成了`docbits`文件夹中的位。

文件夹：

| folder       | contents              |
|:-------------|:----------------------------|
| `bits`       | 组成最终脚本的原生源文件                |
| `docbits`    | 组成`README.md`的原生markdown文件                |
| `bin`        | 服务器端bin脚本 (`xlsx.njs`)                          |
| `dist`       | 用于Web浏览器和非标准JS环境的dist文件  |
| `demos`      | 针对ExtendScript和Webpack等平台的演示项目     |
| `tests`      | 浏览器测试 (运行 `make ctest` 进行构建)    |
| `types`      | typescript定义和测试      |
| `misc`       | 各种各样的支持脚本      |
| `test_files` | 测试文件 (从测试文件仓库拉取)            |

</details>

克隆仓库之后，运行`make help`将会显示一个命令列表。

### OSX/Linux

<details>
  <summary>(点击显示)</summary>

`xlsx.js`文件由来自于`bits`子目录的文件构建。构建脚本(运行`make`)将会连接各个位来产生脚本。提交一个贡献之前，确保运行将会准确地产生`xlsx.js`文件。测试的最简单方式就是添加下面的脚本：

```bash
$ git add xlsx.js
$ make clean
$ make
$ git diff xlsx.js
```

运行`make dist`产生dist文件。每一个版本中的dist文件都会被更新，并且*不应该在版本之间提交*。
</details>

### Windows

<details>
  <summary>(点击显示)</summary>

包含`make.cmd`的脚本将会从`bits`目录中构建`xlsx.js`。构建很简单：

```cmd
> make
```

准备开发环境：

```cmd
> make init
```

windows中可用命令的完整列表显示在`make help`中：

```
make init -- 安装依赖和全局模块
make lint -- 运行 eslint linter
make test -- 运行mocha测试组合
make misc -- 运行更小的测试组合
make book -- 重新构建README 和 summary
make help -- 显示命令信息
```

与[测试文件](#test-files)中解释的一样，在windows中发布ZIP文件必须要下载和提取。如果Bash在windows内可用，可能会运行 OSX/Linux工作流。下面额步骤准备环境：

```bash
# Install support programs for the build and test commands
sudo apt-get install make git subversion mercurial

# Install nodejs and NPM within the WSL
wget -qO- https://deb.nodesource.com/setup_8.x | sudo bash
sudo apt-get install nodejs

# Install dev dependencies
sudo npm install -g mocha voc blanket xlsjs
```

</details>

### 测试

<details>
  <summary>(点击显示)</summary>
`test_misc`(Linux/OSX用`make test_misc`/windows用`make misc`)目标运行定向的功能测试。执行功能测试需要5-10秒而无需对整个测试电池进行测试。新功能应附带相关文件格式和功能的测试。

对于涉及读取端的测试，一个合适的功能测试会包括读取一个存在的文件并且检查工作簿对象的结果。如果涉及参数，文件应该读取不同的值以确保功能如预期所料工作。

对于涉及已经可以解析的新写入功能的测试，恰当的功能测试包括用这个功能写入工作簿，在之后打开并确认功能已经被保存。

对于涉及没有现有读取能力的新写入功能的测试，请添加功能测试到kitchen sink`tests/write.js`。

## 证书

更多细节请查阅相关的证书。原始作者保留未由Apache 2.0许可证明确授予的所有权利。

## 引用

<details>
  <summary><b>OSP覆盖的规格(OSP-covered Specifications)</b> (点击显示)</summary>

- `MS-CFB`: 复合文件二进制文件格式(Compound File Binary File Format)
- `MS-CTXLS`: Excel自定义工具栏二进制文件格式(Excel Custom Toolbar Binary File Format)
- `MS-EXSPXML3`: Excel计算版本2 Web服务XML架构(Excel Calculation Version 2 Web Service XML Schema)
- `MS-ODATA`: 开源的数据协议(Open Data Protocol) (OData)
- `MS-ODRAW`: office绘图二进制文件格式(Office Drawing Binary File Format)
- `MS-ODRAWXML`: Office开源XML结构的Office绘图扩展(Office Drawing Extensions to Office Open XML Structure)
- `MS-OE376`: Office对ECMA-376标准支持的执行信息(Office Implementation Information for ECMA-376 Standards Support)
- `MS-OFFCRYPTO`: Office文档密码学结构(Office Document Cryptography Structure)
- `MS-OI29500`: Office对ISO/IEC 29500标准支持的执行信息(Office Implementation Information for ISO/IEC 29500 Standards Support)
- `MS-OLEDS`: 对象链接和嵌入数据结构(Object Linking and Embedding (OLE) Data Structures)
- `MS-OLEPS`: 对象链接和嵌入属性设置数据结构(Object Linking and Embedding (OLE) Property Set Data Structures)
- `MS-OODF3`: Office对ODF 1.2标准支持的执行信息(Office Implementation Information for ODF 1.2 Standards Support)
- `MS-OSHARED`: Office常用数据类型和对象结构(Office Common Data Types and Objects Structures)
- `MS-OVBA`: Office VBA文件结构(Office VBA File Format Structure)
- `MS-XLDM`: 电子表格数据模型文件格式(Spreadsheet Data Model File Format)
- `MS-XLS`: Excel二进制文件格式(.xls)结构规范(Excel Binary File Format (.xls) Structure Specification)
- `MS-XLSB`: Excel (.xlsb)二进制文件格式(Excel (.xlsb) Binary File Format)
- `MS-XLSX`: Excel (.xlsx) Extensions to the Office Open XML SpreadsheetML File Format
- `XLS`: Microsoft Office Excel 97-2007 Binary File Format Specification
- `RTF`: 富文本(Rich Text Format)

</details>

- ISO/IEC 29500:2012(E) "信息技术 - 文档描述和处理语言 - Office开源XML文件格式"
- Office应用版本 1.2(2011/9/29)开源文档格式
- 工作表文件格式(来自于Lotus) 1984年12月