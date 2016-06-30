var WRD = WScript.CreateObject("Word.Application") ;
WRD.Documents.open("C:\\Users\\Denis\\Dropbox\\edu_local\\system_software\\labs\\lab4\\in_word.docx");

var pp = WScript.CreateObject("PowerPoint.Application");
var p = pp.Presentations.Add();

RowCount = 6;
ColCount = 6;

for (i=1;i<WRD.ActiveDocument.Tables(1).Rows.Count+1;i++) {
	currentSlide = p.Slides.Add(p.Slides.Count+1,2);
	currentSlide.Shapes.Title.TextFrame.TextRange.Text = WRD.ActiveDocument.Tables(1).Rows(i).Cells(i).Range.Text.slice(0, -1);
	for (j=1;j<WRD.ActiveDocument.Tables(1).Rows(i).Cells.Count+1;j++) {
		currentSlide.Shapes.Placeholders(2).TextFrame.TextRange.Text += WRD.ActiveDocument.Tables(1).Rows(i).Cells(j).Range.Text.slice(0, -1);
	}
}

WRD.Quit();
pp.Visible = true;





