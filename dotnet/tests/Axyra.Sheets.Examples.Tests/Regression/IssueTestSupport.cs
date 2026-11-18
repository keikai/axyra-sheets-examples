using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace Axyra.Sheets.Examples.Tests.Regression;

/// <summary>
/// Helpers for the issue tests: reading parts of a saved package and PDF content streams.
/// </summary>
/// <remarks>
/// The tests run on the evaluation build, which adds an "Axyra Evaluation Copy" sheet to every
/// saved workbook. Sheets are therefore looked up by name, and worksheet XML is searched across
/// every worksheet part instead of a fixed <c>sheet1.xml</c>.
/// </remarks>
internal static class IssueTestSupport
{
    internal static CellValue N(double value) => new CellValue.Number(value);

    internal static CellValue T(string value) => new CellValue.Text(value);

    /// <summary>The text of one package part, or <c>null</c> when the part does not exist.</summary>
    internal static string? ZipEntry(string file, string name)
    {
        using ZipArchive zip = ZipFile.OpenRead(file);
        ZipArchiveEntry? entry = zip.GetEntry(name);
        if (entry is null) return null;
        using StreamReader reader = new(entry.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    internal static List<string> ZipNames(string file)
    {
        using ZipArchive zip = ZipFile.OpenRead(file);
        return zip.Entries.Select(e => e.FullName).ToList();
    }

    /// <summary>Every worksheet part of the package, concatenated.</summary>
    internal static string SheetsXml(string file)
    {
        StringBuilder xml = new();
        foreach (string name in ZipNames(file))
            if (name.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal))
                xml.Append(ZipEntry(file, name));
        return xml.ToString();
    }

    /// <summary>Every match of <paramref name="regex"/> in <paramref name="text"/>, in order (DOTALL).</summary>
    internal static List<string> FindAll(string text, string regex) =>
        Regex.Matches(text, regex, RegexOptions.Singleline).Select(m => m.Value).ToList();

    internal static int Count(string text, string part)
    {
        int count = 0;
        for (int i = text.IndexOf(part, StringComparison.Ordinal); i >= 0;
             i = text.IndexOf(part, i + part.Length, StringComparison.Ordinal))
            count++;
        return count;
    }

    /// <summary>The numeric value of a cell, or <c>NaN</c> when it does not hold a number.</summary>
    internal static double Number(CellValue value) => value is CellValue.Number n ? n.Value : double.NaN;

    /// <summary>Inflates every FlateDecode stream of a PDF and returns them in file order.</summary>
    internal static List<string> PdfStreams(byte[] pdf)
    {
        Encoding latin1 = Encoding.Latin1;
        string text = latin1.GetString(pdf);
        List<string> streams = [];
        int i = 0;
        while ((i = text.IndexOf("stream", i, StringComparison.Ordinal)) >= 0)
        {
            if (i >= 3 && string.CompareOrdinal(text, i - 3, "end", 0, 3) == 0)
            {
                i += 6;
                continue;
            }
            int start = i + 6;
            if (start < text.Length && text[start] == '\r') start++;
            if (start < text.Length && text[start] == '\n') start++;
            int end = text.IndexOf("endstream", start, StringComparison.Ordinal);
            if (end < 0) break;
            using MemoryStream inflated = new();
            try
            {
                using ZLibStream z = new(new MemoryStream(pdf, start, end - start), CompressionMode.Decompress);
                z.CopyTo(inflated);
            }
            catch (InvalidDataException)
            {
                // not a FlateDecode stream (or truncated); keep what was inflated
            }
            streams.Add(latin1.GetString(inflated.ToArray()));
            i = end + 9;
        }
        return streams;
    }
}
