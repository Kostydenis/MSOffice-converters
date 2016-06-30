var XLS = WScript.CreateObject("Excel.Application") ;
XLS.Workbooks.open(WScript.CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)+"\\in_excel.xlsx");

var word = WScript.CreateObject("Word.Application");
var doc = word.Documents.Add();
var tbl = doc.Tables.Add(doc.Range(0,0),5,5);

RowCount = 6;
ColCount = 6;

for (i=1;i<RowCount;i++) {
	for (j=1;j<ColCount;j++) {
		tbl.Cell(i,j).Range.Text = XLS.Cells(i,j).Value;
	}
}
XLS.Quit();
word.Visible = true;





