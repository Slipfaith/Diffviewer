using System;

namespace DocxCompare;

public static class Program
{
    public static int Main(string[] args)
    {
        try
        {
            if (args.Length < 3)
            {
                Console.Error.WriteLine("Usage: docx_compare.exe <file_a> <file_b> <output> [--author \"Change Tracker\"]");
                return 1;
            }

            string fileA = args[0];
            string fileB = args[1];
            string output = args[2];
            string author = "Change Tracker";

            for (int i = 3; i < args.Length; i++)
            {
                if (args[i] == "--author" && i + 1 < args.Length)
                {
                    author = args[i + 1];
                    i++;
                }
            }

            var comparer = new DocComparer(author);
            comparer.Compare(fileA, fileB, output);
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 1;
        }
    }
}
