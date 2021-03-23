![version](https://img.shields.io/badge/version-16%2B-8331AE)
![platform](https://img.shields.io/static/v1?label=platform&message=mac-intel%20|%20mac-arm%20|%20win-64&color=blue)
[![license](https://img.shields.io/github/license/miyako/4d-plugin-xls)](LICENSE)
![downloads](https://img.shields.io/github/downloads/miyako/4d-plugin-xls/total)

**Note**: for v17 and earlier, move `manifest.json` to `Contents`

4d-plugin-xls
=============

4D plugin to write XLS documents using [xlslib](https://sourceforge.net/projects/xlslib/) 2.5.0.

For reading cell values, you might want to consider [this](https://github.com/miyako/4d-plugin-free-xl).

### Library Information

* Notable build flags on Mac

```
-stdlib=libc++
-isysroot MacOSX10.9.sdk
-mmacosx-version-min=10.9
```

### Build notes

suppress `configure` error "cannot run test program while cross compiling"

* Library solution for Windows

https://github.com/miyako/msvc-xlslib

https://github.com/miyako/msvc-iconv

**Note**: XLSLIB has been modified to accept unicode path names on Windows.

## Examples

* Adding a SUM() function cell

```
$text:=Unicode_sample 

  //create a workbook
$book:=XLS WORKBOOK Create 

$sheetName:=$text
$sheet:=XLS WORKBOOK Create sheet ($book;$sheetName)

  //problem with Mac version of Excel 2010; OK on Windows
XLS WORKSHEET SET COL WIDTH ($sheet;0;20*256)

$row:=0  //zero-based
$col:=0  //zero-based
$format:=0  //NULL=default format (0x0F)
$cell:=XLS WORKSHEET Set cell text ($sheet;$row;$col;$format;$text)
XLS CELL RELEASE ($cell)  //we don't need this reference any more, so release it.

  //create a range reference node
$cell1:=XLS WORKSHEET Set cell real ($sheet;0;1;$format;1)
$cell2:=XLS WORKSHEET Set cell real ($sheet;1;1;$format;2)
$area:=XLS WORKBOOK Create area node ($book;$cell1;$cell2;XLS_CELL_ABSOLUTE_As1;XLS_CELLOP_AS_REFERENCE)
XLS CELL RELEASE ($cell1)
XLS CELL RELEASE ($cell2)

  //create a function node
$fn:=XLS WORKBOOK Create fn1 node ($book;XLS_FUNC_SUM;$area)
$cell:=XLS WORKSHEET Set cell fn ($sheet;2;1;$format;$fn)
XLS NODE RELEASE ($fn)
XLS NODE RELEASE ($area)
XLS CELL RELEASE ($cell)

XLS WORKSHEET RELEASE ($sheet)

$success:=XLS WORKBOOK Save document ($book;System folder(Desktop)+$text+".xls")

XLS WORKBOOK CLEAR ($book)
```

## Discussions

https://forums.4d.com/Post//28120080/1/

## 日本語特有の制限

日本語のフォント名は指定できないようです。たとえば`ＭＳ Ｐゴシック`は`MS PGothic`と指定しなければなりません。

## その他

フォントサイズは[TWIP](https://ja.wikipedia.org/wiki/Twip)（ポイントの20分の1）単位で指定します。

