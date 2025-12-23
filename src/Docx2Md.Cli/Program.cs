using Docx2Md.Core;
using Docx2Md.Core.Models;

namespace Docx2Md.Cli;

class Program
{
    static int Main(string[] args)
    {
        Console.WriteLine("DOCX to Markdown Converter");
        Console.WriteLine("==========================");
        Console.WriteLine();

        if (args.Length < 1)
        {
            ShowUsage();
            return 1;
        }

        var inputPath = args[0];
        var outputPath = args.Length > 1 ? args[1] : GetDefaultOutputPath(inputPath);

        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"Error: Input file not found: {inputPath}");
            return 1;
        }

        try
        {
            Console.WriteLine($"Input:  {inputPath}");
            Console.WriteLine($"Output: {outputPath}");
            Console.WriteLine();

            var settings = ConversionSettings.Default;
            var converter = new Docx2MdConverter(settings);

            Console.WriteLine("Parsing DOCX file...");
            var document = converter.ConvertFile(inputPath, outputPath);

            Console.WriteLine($"✓ Converted {document.Segments.Count} segments");
            Console.WriteLine($"✓ Extracted {document.Images.Count} images");
            Console.WriteLine();

            // Show diagnostics summary
            var diagnostics = document.GetAllDiagnostics().ToList();
            if (diagnostics.Count > 0)
            {
                Console.WriteLine("Diagnostics:");
                var errors = diagnostics.Count(d => d.Level == DiagnosticLevel.Error);
                var warnings = diagnostics.Count(d => d.Level == DiagnosticLevel.Warning);
                var info = diagnostics.Count(d => d.Level == DiagnosticLevel.Info);

                if (errors > 0) Console.WriteLine($"  Errors:   {errors}");
                if (warnings > 0) Console.WriteLine($"  Warnings: {warnings}");
                if (info > 0) Console.WriteLine($"  Info:     {info}");

                if (settings.GenerateDiagnosticReport)
                {
                    var diagFile = Path.GetFileNameWithoutExtension(outputPath) + "_diagnostics.md";
                    Console.WriteLine($"  See {diagFile} for details");
                }
            }
            else
            {
                Console.WriteLine("✓ No diagnostics");
            }

            Console.WriteLine();
            Console.WriteLine("Conversion complete!");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            if (ex.InnerException != null)
            {
                Console.Error.WriteLine($"  {ex.InnerException.Message}");
            }
            return 1;
        }
    }

    static void ShowUsage()
    {
        Console.WriteLine("Usage: docx2md <input.docx> [output.md]");
        Console.WriteLine();
        Console.WriteLine("Arguments:");
        Console.WriteLine("  input.docx   Path to input DOCX file");
        Console.WriteLine("  output.md    Path to output Markdown file (optional)");
        Console.WriteLine("               If not specified, uses input filename with .md extension");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  docx2md document.docx");
        Console.WriteLine("  docx2md document.docx output.md");
    }

    static string GetDefaultOutputPath(string inputPath)
    {
        var directory = Path.GetDirectoryName(inputPath) ?? ".";
        var filename = Path.GetFileNameWithoutExtension(inputPath);
        return Path.Combine(directory, filename + ".md");
    }
}
