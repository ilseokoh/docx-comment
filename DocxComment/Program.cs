using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.IO;

var filepath = $"{Directory.GetCurrentDirectory()}\\취업규칙.docx";

Application app = new Application();

Console.WriteLine($"{filepath} is processing");

Document wordDoc = app.Documents.Open(filepath, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

for (var i = 1; i <= wordDoc.Comments.Count; i++)
{
    
    Console.WriteLine($" ----------- {wordDoc.Comments[i].Author} / {wordDoc.Comments[i].Date} ----------- ");

    // Scope 는 Comment의 대상이 되는 본문
    Console.WriteLine($"본문: {wordDoc.Comments[i].Scope.Text}");

    // Range 는 Comment 내용인데 여러줄로 되어 있으면 마지막 줄 내용만 나온다.
    // Range.Sentences 라는 컬렉션에 (1부터 시작하는 Index)에서 가져와야 하네 
    //Console.WriteLine(wordDoc.Comments[i].Range.Text);
    Console.WriteLine("Comment");
    for (var j = 1; j <= wordDoc.Comments[i].Range.Sentences.Count; j++)
    {
        Console.WriteLine(wordDoc.Comments[i].Range.Sentences[j].Text);
    }
}

wordDoc.Close();
