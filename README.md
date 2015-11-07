4d-plugin-xls
=============

4D plugin to write XLS documents using [xlslib](http://xlslib.sourceforge.net/) 2.3.4.

For reading cell values, you might want to consider [this](https://github.com/miyako/4d-plugin-free-xl).

About
-----
v14 is for v14 and above, 32/64 bits for both platforms. (Mac requires 10.8+).

v11 is for v11 and above, 32/64 bits for Windows and 32 bits for Mac. (Mac requires 10.6+).

**Note**: XLSLIB has been modified to accept unicode path names on Windows.

Example
---
Adding a SUM() function cell
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
