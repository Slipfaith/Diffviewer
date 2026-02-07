using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxCompare;

public sealed class TrackChangeBuilder
{
    private int _changeId;
    private readonly string _author;
    private readonly DateTime _timestamp;

    public TrackChangeBuilder(string author)
    {
        _author = author;
        _timestamp = DateTime.UtcNow;
        _changeId = 1;
    }

    public Paragraph BuildUnchangedParagraph(Paragraph source)
    {
        return (Paragraph)source.CloneNode(true);
    }

    public Paragraph BuildDeletedParagraph(Paragraph source)
    {
        var paragraph = new Paragraph();
        paragraph.ParagraphProperties = (ParagraphProperties?)source.ParagraphProperties?.CloneNode(true);
        var runProps = GetRunProps(source);
        var text = new string[] { source.InnerText };
        paragraph.Append(CreateDeletedRun(runProps, string.Join(" ", text)));
        return paragraph;
    }

    public Paragraph BuildInsertedParagraph(Paragraph template, string text)
    {
        var paragraph = new Paragraph();
        paragraph.ParagraphProperties = (ParagraphProperties?)template.ParagraphProperties?.CloneNode(true);
        var runProps = GetRunProps(template);
        paragraph.Append(CreateInsertedRun(runProps, text));
        return paragraph;
    }

    public Paragraph BuildModifiedParagraph(Paragraph source, Paragraph target)
    {
        var paragraph = new Paragraph();
        paragraph.ParagraphProperties = (ParagraphProperties?)source.ParagraphProperties?.CloneNode(true);

        var sourceProps = GetRunProps(source);
        var targetProps = GetRunProps(target);
        var ops = WordDiff.DiffWords(source.InnerText, target.InnerText);

        for (int i = 0; i < ops.Count; i++)
        {
            var op = ops[i];
            string text = op.Text;
            if (!string.IsNullOrWhiteSpace(text) && i < ops.Count - 1)
            {
                text += " ";
            }

            if (op.Type == DiffType.Equal)
            {
                paragraph.Append(CreateRun(sourceProps, text));
            }
            else if (op.Type == DiffType.Delete)
            {
                paragraph.Append(CreateDeletedRun(sourceProps, text));
            }
            else if (op.Type == DiffType.Insert)
            {
                paragraph.Append(CreateInsertedRun(targetProps, text));
            }
        }

        return paragraph;
    }

    private Run CreateRun(RunProperties? runProps, string text)
    {
        var run = new Run();
        if (runProps != null)
        {
            run.RunProperties = (RunProperties)runProps.CloneNode(true);
        }
        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private InsertedRun CreateInsertedRun(RunProperties? runProps, string text)
    {
        var inserted = new InsertedRun
        {
            Author = _author,
            Date = _timestamp,
            Id = _changeId++.ToString()
        };
        var run = new Run();
        if (runProps != null)
        {
            run.RunProperties = (RunProperties)runProps.CloneNode(true);
        }
        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        inserted.Append(run);
        return inserted;
    }

    private DeletedRun CreateDeletedRun(RunProperties? runProps, string text)
    {
        var deleted = new DeletedRun
        {
            Author = _author,
            Date = _timestamp,
            Id = _changeId++.ToString()
        };
        var run = new Run();
        if (runProps != null)
        {
            run.RunProperties = (RunProperties)runProps.CloneNode(true);
        }
        run.Append(new DeletedText(text) { Space = SpaceProcessingModeValues.Preserve });
        deleted.Append(run);
        return deleted;
    }

    private static RunProperties? GetRunProps(Paragraph paragraph)
    {
        return paragraph.Descendants<Run>().FirstOrDefault()?.RunProperties;
    }
}

public enum DiffType
{
    Equal,
    Insert,
    Delete
}

public sealed class DiffOp
{
    public DiffType Type { get; }
    public string Text { get; }

    public DiffOp(DiffType type, string text)
    {
        Type = type;
        Text = text;
    }
}

public static class WordDiff
{
    public static List<DiffOp> DiffWords(string source, string target)
    {
        var a = SplitWords(source);
        var b = SplitWords(target);
        int[,] dp = new int[a.Count + 1, b.Count + 1];

        for (int i = 1; i <= a.Count; i++)
        {
            for (int j = 1; j <= b.Count; j++)
            {
                if (a[i - 1] == b[j - 1])
                {
                    dp[i, j] = dp[i - 1, j - 1] + 1;
                }
                else
                {
                    dp[i, j] = Math.Max(dp[i - 1, j], dp[i, j - 1]);
                }
            }
        }

        var ops = new List<DiffOp>();
        int x = a.Count;
        int y = b.Count;
        while (x > 0 || y > 0)
        {
            if (x > 0 && y > 0 && a[x - 1] == b[y - 1])
            {
                ops.Add(new DiffOp(DiffType.Equal, a[x - 1]));
                x--;
                y--;
            }
            else if (y > 0 && (x == 0 || dp[x, y - 1] >= dp[x - 1, y]))
            {
                ops.Add(new DiffOp(DiffType.Insert, b[y - 1]));
                y--;
            }
            else if (x > 0)
            {
                ops.Add(new DiffOp(DiffType.Delete, a[x - 1]));
                x--;
            }
        }

        ops.Reverse();
        return MergeOps(ops);
    }

    private static List<string> SplitWords(string text)
    {
        return text.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList();
    }

    private static List<DiffOp> MergeOps(List<DiffOp> ops)
    {
        var merged = new List<DiffOp>();
        foreach (var op in ops)
        {
            if (merged.Count == 0 || merged[^1].Type != op.Type)
            {
                merged.Add(new DiffOp(op.Type, op.Text));
            }
            else
            {
                var last = merged[^1];
                merged[^1] = new DiffOp(last.Type, $"{last.Text} {op.Text}");
            }
        }
        return merged;
    }
}
