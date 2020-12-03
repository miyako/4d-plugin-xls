//%attributes = {}
C_LONGINT:C283($vl_area;$vl_book;$vl_cell;$vl_cell1;$vl_cell2;$vl_col;$vl_fn;$vl_fn2;$vl_format;$vl_row;$vl_sheet)
C_LONGINT:C283($vl_success)
C_TEXT:C284($vt_sheetName;$vt_UnicodeText)


$vt_UnicodeText:="Osfas"

  //create a workbook
$vl_book:=XLS WORKBOOK Create 

$vt_sheetName:=$vt_UnicodeText
$vl_sheet:=XLS WORKBOOK Create sheet ($vl_book;$vt_sheetName)

  //problem with Mac version of Excel 2010; OK on Windows
XLS WORKSHEET SET COL WIDTH ($vl_sheet;0;20*256)

$vl_row:=0  //zero-based
$vl_col:=0  //zero-based
$vl_format:=0  //NULL=default format (0x0F)
$vl_cell:=XLS WORKSHEET Set cell text ($vl_sheet;$vl_row;$vl_col;$vl_format;$vt_UnicodeText)
XLS CELL RELEASE ($vl_cell)  //we don't need this reference any more, so release it.
$vl_cell:=XLS WORKSHEET Set cell text ($vl_sheet;2;$vl_col;$vl_format;"Summe")
XLS CELL RELEASE ($vl_cell)  //we don't need this reference any more, so release it.

  //create a range reference node
$vl_cell1:=XLS WORKSHEET Set cell real ($vl_sheet;0;1;$vl_format;3)
$vl_cell2:=XLS WORKSHEET Set cell real ($vl_sheet;1;1;$vl_format;2)
$vl_area1:=XLS WORKBOOK Create area node ($vl_book;$vl_cell1;$vl_cell2;XLS_CELL_ABSOLUTE_As1;XLS_CELLOP_AS_REFERENCE)
$vl_area2:=XLS WORKBOOK Create area node ($vl_book;$vl_cell1;$vl_cell2;XLS_CELL_ABSOLUTE_As1;XLS_CELLOP_AS_REFERENCE)

XLS CELL RELEASE ($vl_cell1)
XLS CELL RELEASE ($vl_cell2)

  //create a function node
$vl_fn:=XLS WORKBOOK Create fn1 node ($vl_book;XLS_FUNC_SUM;$vl_area1)
$vl_cell:=XLS WORKSHEET Set cell fn ($vl_sheet;2;1;$vl_format;$vl_fn)
XLS NODE RELEASE ($vl_fn)
XLS CELL RELEASE ($vl_cell)

$vl_fn:=XLS WORKBOOK Create fn1 node ($vl_book;XLS_FUNC_PRODUCT;$vl_area2)
$vl_cell:=XLS WORKSHEET Set cell fn ($vl_sheet;3;1;$vl_format;$vl_fn)
XLS NODE RELEASE ($vl_fn)
XLS CELL RELEASE ($vl_cell)

XLS NODE RELEASE ($vl_area1)
XLS NODE RELEASE ($vl_area2)

XLS WORKSHEET RELEASE ($vl_sheet)

$vl_success:=XLS WORKBOOK Save document ($vl_book;System folder:C487(Desktop:K41:16)+$vt_UnicodeText+".xls")

XLS WORKBOOK CLEAR ($vl_book)