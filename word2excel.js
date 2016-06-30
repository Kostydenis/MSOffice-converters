var WRD = WScript.CreateObject("Word.Application") ;
WRD.Documents.open("C:\\Users\\Denis\\Dropbox\\edu_local\\system_software\\labs\\lab4\\in_word.docx");


var excel = WScript.CreateObject("Excel.Application");
excel.Workbooks.Add();


for (i=1;i<WRD.ActiveDocument.Tables(1).Rows.Count+1;i++) {
	for (j=1;j<WRD.ActiveDocument.Tables(1).Rows(i).Cells.Count+1;j++) {
		excel.Cells(i,j).Value = "" + WRD.ActiveDocument.Tables(1).Rows(i).Cells(j).Range.Text.slice(0, -1) ;
	}
}
WRD.Quit();
excel.Visible = true;
