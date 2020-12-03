//%attributes = {}
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




$area:=XLS WORKBOOK Create area node ($book;$cell1;$cell2;XLS_CELL_RELATIVE_A1;XLS_CELLOP_AS_REFERENCE)
XLS CELL RELEASE ($cell1)
XLS CELL RELEASE ($cell2)

  //create a function node
$fn:=XLS WORKBOOK Create fn1 node ($book;XLS_FUNC_SUM;$area)
$cell:=XLS WORKSHEET Set cell fn ($sheet;2;1;$format;$fn)
XLS NODE RELEASE ($fn)
XLS NODE RELEASE ($area)
XLS CELL RELEASE ($cell)

XLS WORKSHEET RELEASE ($sheet)

$success:=XLS WORKBOOK Save document ($book;System folder:C487(Desktop:K41:16)+$text+".xls")

XLS WORKBOOK CLEAR ($book)