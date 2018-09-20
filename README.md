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