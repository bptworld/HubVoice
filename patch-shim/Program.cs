using System.Text;
using System.Text.RegularExpressions;

record Hunk(int OriginalStart, List<string> Lines);
record FilePatch(string OriginalPath, string UpdatedPath, List<Hunk> Hunks);
record HunkLine(char Operation, string Content);

public static class Program
{
    public static int Main(string[] args)
    {
        try
        {
            var options = ParseArgs(args);
            if (string.IsNullOrWhiteSpace(options.PatchFile))
            {
                Console.Error.WriteLine("patch: missing -i <patchfile>");
                return 1;
            }

            var filePatches = ParsePatchFile(options.PatchFile!);
            foreach (var filePatch in filePatches)
            {
                ApplyFilePatch(filePatch, options.StripCount);
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"patch: {ex.Message}");
            return 1;
        }
    }

    private static PatchOptions ParseArgs(string[] args)
    {
        var options = new PatchOptions();

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg == "--binary")
            {
                continue;
            }

            if (arg == "-i" && i + 1 < args.Length)
            {
                options.PatchFile = args[++i];
                continue;
            }

            if (arg.StartsWith("-i", StringComparison.Ordinal) && arg.Length > 2)
            {
                options.PatchFile = arg[2..];
                continue;
            }

            if (arg == "-p" && i + 1 < args.Length)
            {
                options.StripCount = int.Parse(args[++i]);
                continue;
            }

            if (arg.StartsWith("-p", StringComparison.Ordinal) && arg.Length > 2)
            {
                options.StripCount = int.Parse(arg[2..]);
                continue;
            }
        }

        return options;
    }

    private static List<FilePatch> ParsePatchFile(string patchPath)
    {
        var lines = File.ReadAllLines(patchPath);
        var patches = new List<FilePatch>();
        var i = 0;

        while (i < lines.Length)
        {
            if (!lines[i].StartsWith("--- ", StringComparison.Ordinal))
            {
                i++;
                continue;
            }

            var originalPath = lines[i][4..].Split('\t')[0].Trim();
            i++;
            if (i >= lines.Length || !lines[i].StartsWith("+++ ", StringComparison.Ordinal))
            {
                throw new InvalidOperationException($"Malformed patch near {originalPath}");
            }

            var updatedPath = lines[i][4..].Split('\t')[0].Trim();
            i++;

            var hunks = new List<Hunk>();
            while (i < lines.Length && lines[i].StartsWith("@@ ", StringComparison.Ordinal))
            {
                var match = Regex.Match(lines[i], @"^@@ -(?<old>\d+)(,(?<oldCount>\d+))? \+(?<new>\d+)(,(?<newCount>\d+))? @@");
                if (!match.Success)
                {
                    throw new InvalidOperationException($"Invalid hunk header: {lines[i]}");
                }

                var originalStart = int.Parse(match.Groups["old"].Value);
                i++;

                var hunkLines = new List<string>();
                while (i < lines.Length &&
                       !lines[i].StartsWith("@@ ", StringComparison.Ordinal) &&
                       !lines[i].StartsWith("--- ", StringComparison.Ordinal))
                {
                    hunkLines.Add(lines[i]);
                    i++;
                }

                hunks.Add(new Hunk(originalStart, hunkLines));
            }

            patches.Add(new FilePatch(originalPath, updatedPath, hunks));
        }

        return patches;
    }

    private static void ApplyFilePatch(FilePatch patch, int stripCount)
    {
        var targetPath = ResolveTargetPath(patch, stripCount);
        var fullPath = Path.GetFullPath(targetPath);
        var exists = File.Exists(fullPath);

        List<string> sourceLines;
        string newline;
        bool hadTrailingNewline;

        if (exists)
        {
            var text = File.ReadAllText(fullPath);
            newline = text.Contains("\r\n", StringComparison.Ordinal) ? "\r\n" : "\n";
            hadTrailingNewline = text.EndsWith("\n", StringComparison.Ordinal);
            sourceLines = SplitLines(text);
        }
        else
        {
            newline = "\n";
            hadTrailingNewline = true;
            sourceLines = new List<string>();
        }

        var result = new List<string>();
        var sourceIndex = 0;

        foreach (var hunk in patch.Hunks)
        {
            var hunkStart = FindHunkStart(sourceLines, sourceIndex, hunk);
            while (sourceIndex < hunkStart && sourceIndex < sourceLines.Count)
            {
                result.Add(sourceLines[sourceIndex]);
                sourceIndex++;
            }

            foreach (var hunkLine in ParseHunkLines(hunk.Lines))
            {
                switch (hunkLine.Operation)
                {
                    case ' ':
                        EnsureSourceLine(sourceLines, sourceIndex, hunkLine.Content, fullPath);
                        result.Add(sourceLines[sourceIndex]);
                        sourceIndex++;
                        break;
                    case '-':
                        EnsureSourceLine(sourceLines, sourceIndex, hunkLine.Content, fullPath);
                        sourceIndex++;
                        break;
                    case '+':
                        result.Add(hunkLine.Content);
                        break;
                    case '\\':
                        break;
                }
            }
        }

        while (sourceIndex < sourceLines.Count)
        {
            result.Add(sourceLines[sourceIndex]);
            sourceIndex++;
        }

        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        var output = string.Join(newline, result);
        if (hadTrailingNewline)
        {
            output += newline;
        }
        File.WriteAllText(fullPath, output, new UTF8Encoding(false));
    }

    private static int FindHunkStart(List<string> sourceLines, int minimumIndex, Hunk hunk)
    {
        var guess = Math.Max(hunk.OriginalStart - 1, minimumIndex);
        var candidates = Enumerable.Range(minimumIndex, sourceLines.Count - minimumIndex + 1)
            .OrderBy(index => Math.Abs(index - guess));

        foreach (var candidate in candidates)
        {
            if (HunkMatchesAt(sourceLines, candidate, hunk))
            {
                return candidate;
            }
        }

        throw new InvalidOperationException(
            $"Unable to find matching context for hunk starting near line {hunk.OriginalStart}.");
    }

    private static bool HunkMatchesAt(List<string> sourceLines, int startIndex, Hunk hunk)
    {
        var sourceIndex = startIndex;
        foreach (var hunkLine in ParseHunkLines(hunk.Lines))
        {
            if (hunkLine.Operation is not (' ' or '-'))
            {
                continue;
            }

            if (sourceIndex >= sourceLines.Count)
            {
                return false;
            }

            if (!string.Equals(sourceLines[sourceIndex], hunkLine.Content, StringComparison.Ordinal))
            {
                return false;
            }

            sourceIndex++;
        }

        return true;
    }

    private static IEnumerable<HunkLine> ParseHunkLines(IEnumerable<string> rawLines)
    {
        foreach (var rawLine in rawLines)
        {
            if (rawLine.Length == 0)
            {
                yield return new HunkLine(' ', string.Empty);
                continue;
            }

            var op = rawLine[0];
            if (op is not (' ' or '-' or '+' or '\\'))
            {
                throw new InvalidOperationException($"Unsupported patch line '{rawLine}'");
            }

            var content = rawLine.Length > 1 ? rawLine[1..] : string.Empty;
            yield return new HunkLine(op, content);
        }
    }

    private static string ResolveTargetPath(FilePatch patch, int stripCount)
    {
        var candidate = patch.UpdatedPath != "/dev/null" ? patch.UpdatedPath : patch.OriginalPath;
        if (candidate == "/dev/null")
        {
            throw new InvalidOperationException("Deleting files is not supported by this local patch tool");
        }

        var normalized = candidate.Replace('\\', '/');
        var parts = normalized.Split('/', StringSplitOptions.RemoveEmptyEntries).ToList();
        if (parts.Count >= stripCount)
        {
            parts = parts.Skip(stripCount).ToList();
        }

        if (parts.Count == 0)
        {
            throw new InvalidOperationException($"Unable to resolve target path from '{candidate}'");
        }

        return Path.Combine(parts.ToArray());
    }

    private static List<string> SplitLines(string text)
    {
        var normalized = text.Replace("\r\n", "\n");
        var parts = normalized.Split('\n').ToList();
        if (parts.Count > 0 && parts[^1] == string.Empty)
        {
            parts.RemoveAt(parts.Count - 1);
        }
        return parts;
    }

    private static void EnsureSourceLine(List<string> sourceLines, int sourceIndex, string expected, string filePath)
    {
        if (sourceIndex >= sourceLines.Count)
        {
            throw new InvalidOperationException($"Patch ran past end of file: {filePath}");
        }

        if (!string.Equals(sourceLines[sourceIndex], expected, StringComparison.Ordinal))
        {
            throw new InvalidOperationException(
                $"Patch context mismatch in {filePath} at line {sourceIndex + 1}. Expected '{expected}' but found '{sourceLines[sourceIndex]}'.");
        }
    }

    private sealed class PatchOptions
    {
        public string? PatchFile { get; set; }
        public int StripCount { get; set; } = 0;
    }
}
