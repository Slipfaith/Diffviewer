using System;
using System.Collections.Generic;

namespace DocxCompare;

public sealed class ParagraphMatch
{
    public int? IndexA { get; }
    public int? IndexB { get; }

    public ParagraphMatch(int? indexA, int? indexB)
    {
        IndexA = indexA;
        IndexB = indexB;
    }
}

public sealed class ParagraphMatcher
{
    private readonly double _fuzzyThreshold;

    public ParagraphMatcher(double fuzzyThreshold)
    {
        _fuzzyThreshold = fuzzyThreshold;
    }

    public List<ParagraphMatch> Match(IReadOnlyList<string> paragraphsA, IReadOnlyList<string> paragraphsB)
    {
        int min = Math.Min(paragraphsA.Count, paragraphsB.Count);
        var matches = new List<ParagraphMatch>();
        var unmatchedA = new List<int>();
        var unmatchedB = new HashSet<int>();

        for (int i = 0; i < paragraphsB.Count; i++)
        {
            unmatchedB.Add(i);
        }

        for (int i = 0; i < min; i++)
        {
            matches.Add(new ParagraphMatch(i, i));
            unmatchedB.Remove(i);
        }

        for (int i = min; i < paragraphsA.Count; i++)
        {
            unmatchedA.Add(i);
        }

        var fuzzyMatches = new List<ParagraphMatch>();
        var usedB = new HashSet<int>();

        foreach (int indexA in unmatchedA)
        {
            double bestScore = _fuzzyThreshold;
            int bestIndex = -1;
            foreach (int indexB in unmatchedB)
            {
                if (usedB.Contains(indexB))
                {
                    continue;
                }
                double score = Similarity(paragraphsA[indexA], paragraphsB[indexB]);
                if (score >= bestScore)
                {
                    bestScore = score;
                    bestIndex = indexB;
                }
            }
            if (bestIndex >= 0)
            {
                usedB.Add(bestIndex);
                fuzzyMatches.Add(new ParagraphMatch(indexA, bestIndex));
            }
        }

        matches.AddRange(fuzzyMatches);

        return matches;
    }

    private static double Similarity(string a, string b)
    {
        if (a.Length == 0 && b.Length == 0)
        {
            return 1.0;
        }
        int distance = Levenshtein(a, b);
        int max = Math.Max(a.Length, b.Length);
        return max == 0 ? 1.0 : 1.0 - (double)distance / max;
    }

    private static int Levenshtein(string a, string b)
    {
        int[,] dp = new int[a.Length + 1, b.Length + 1];
        for (int i = 0; i <= a.Length; i++)
        {
            dp[i, 0] = i;
        }
        for (int j = 0; j <= b.Length; j++)
        {
            dp[0, j] = j;
        }
        for (int i = 1; i <= a.Length; i++)
        {
            for (int j = 1; j <= b.Length; j++)
            {
                int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                dp[i, j] = Math.Min(
                    Math.Min(dp[i - 1, j] + 1, dp[i, j - 1] + 1),
                    dp[i - 1, j - 1] + cost
                );
            }
        }
        return dp[a.Length, b.Length];
    }
}
