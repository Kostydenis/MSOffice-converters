var XLS = WScript.CreateObject("Excel.Application") ;
XLS.Workbooks.open(WScript.CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)+"\\in_excel.xlsx");

var pp = WScript.CreateObject("PowerPoint.Application");
var p = pp.Presentations.Add();

RowCount = 6;
ColCount = 6;

for (i=1;i<RowCount;i++) {
	for (j=1;j<ColCount;j++) {
		currentSlide = p.Slides.Add(p.Slides.Count+1,2);
		currentSlide.Shapes.Title.TextFrame.TextRange.Text = "Заголовок слайда " + i;
		currentSlide.Shapes.Placeholders(2).TextFrame.TextRange.Text = "Текст слайда " + j;
	}
}

XLS.Quit();
pp.Visible = true;





