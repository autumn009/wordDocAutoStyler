using Microsoft.Office.Interop.Word;

Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
var doc = word.Documents.Open(args[0]);
foreach (Microsoft.Office.Interop.Word.Paragraph para in doc.Paragraphs)
{
    var s = para.Range.Text;
    if (s.StartsWith("●●"))
    {
        para.set_Style("見出し 2");
    }
    else if (s.StartsWith("●"))
    {
        para.set_Style("見出し 3");
    }
}
//foreach (Style item in doc.Styles)
//{
//Console.WriteLine(item.NameLocal);
//}
doc.Save();
doc.Close();
word.Quit();
Console.WriteLine("Done");
