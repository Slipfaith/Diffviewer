using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxCompare;

public sealed class DocComparer
{
    private readonly string _author;
    private readonly ParagraphMatcher _matcher;
    private readonly TrackChangeBuilder _builder;

    public DocComparer(string author)
    {
        _author = author;
        _matcher = new ParagraphMatcher(0.8);
        _builder = new TrackChangeBuilder(author);
    }

    public void Compare(string fileA, string fileB, string output)
    {
        if (!File.Exists(fileA))
        {
            throw new FileNotFoundException($"File not found: {fileA}");
        }
        if (!File.Exists(fileB))
        {
            throw new FileNotFoundException($"File not found: {fileB}");
        }

        var paragraphsA = ReadParagraphs(fileA);
        var paragraphsB = ReadParagraphs(fileB);
        var textsA = paragraphsA.Select(p => p.InnerText).ToList();
        var textsB = paragraphsB.Select(p => p.InnerText).ToList();

        var matches = _matcher.Match(textsA, textsB);
        var matchedA = new HashSet<int>(matches.Where(m => m.IndexA.HasValue).Select(m => m.IndexA!.Value));
        var matchedB = new HashSet<int>(matches.Where(m => m.IndexB.HasValue).Select(m => m.IndexB!.Value));

        File.Copy(fileA, output, true);
        using var doc = WordprocessingDocument.Open(output, true);
        var body = doc.MainDocumentPart!.Document.Body!;
        body.RemoveAllChildren<Paragraph>();

        for (int i = 0; i < paragraphsA.Count; i++)
        {
            var match = matches.FirstOrDefault(m => m.IndexA == i);
            if (match?.IndexB is int indexB)
            {
                var paraA = paragraphsA[i];
                var paraB = paragraphsB[indexB];
                if (paraA.InnerText == paraB.InnerText)
                {
                    body.Append(_builder.BuildUnchangedParagraph(paraA));
                }
                else
                {
                    body.Append(_builder.BuildModifiedParagraph(paraA, paraB));
                }
            }
            else
            {
                body.Append(_builder.BuildDeletedParagraph(paragraphsA[i]));
            }
        }

        for (int j = 0; j < paragraphsB.Count; j++)
        {
            if (!matchedB.Contains(j))
            {
                var paraB = paragraphsB[j];
                body.Append(_builder.BuildInsertedParagraph(paraB, paraB.InnerText));
            }
        }

        doc.MainDocumentPart.Document.Save();
    }

    private static List<Paragraph> ReadParagraphs(string filePath)
    {
        using var doc = WordprocessingDocument.Open(filePath, false);
        var body = doc.MainDocumentPart!.Document.Body!;
        return body.Elements<Paragraph>().ToList();
    }
}
