//%attributes = {}
/*
塗り潰し＆フォント設定
*/

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

/*
塗り潰し（スタイルも指定すること！）
*/

XLS CELL SET COLOR ($cell;XLS_CLR_ROSE;XLS_CLR_ROSE)
XLS CELL SET FILL STYLE ($cell;XLS_FILL_SOLID)

/*
日本語フォントはこの方法で
*/

Case of 
	: (True:C214)
		$font:=XLS WORKBOOK Create font ($book;"MS PGothic")  //英語名で指定する
	: (False:C215)
		$font:=XLS WORKBOOK Create font ($book;"ＭＳ Ｐゴシック")  //日本語名は無視される（クラッシュ回避のため）
End case 

XLS CELL SET FONT ($cell;$font)
XLS FONT RELEASE ($font)

Case of 
	: (True:C214)
		XLS CELL SET FONT NAME ($cell;"MS PGothic")  //英語名で指定する
	: (False:C215)
		XLS CELL SET FONT NAME ($cell;"ＭＳ Ｐゴシック")  //日本語名は無視される（クラッシュ回避のため）
End case 

XLS CELL RELEASE ($cell)  //we don't need this reference any more, so release it.

XLS WORKSHEET RELEASE ($sheet)

$success:=XLS WORKBOOK Save document ($book;System folder:C487(Desktop:K41:16)+$text+".xls")

XLS WORKBOOK CLEAR ($book)