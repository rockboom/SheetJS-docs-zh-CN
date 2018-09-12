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
