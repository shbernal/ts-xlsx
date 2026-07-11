using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace TsXlsx.OoxmlValidator;

internal static class Program
{
    private const int Success = 0;
    private const int ValidationFailure = 1;
    private const int ToolFailure = 2;
    private const int MaxErrorsPerFile = 1_000;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.Never,
        WriteIndented = true,
    };

    public static int Main(string[] args)
    {
        try
        {
            var options = CliOptions.Parse(args);
            var results = options.Files
                .Select(file => ValidateFile(file, options.Format))
                .OrderBy(result => result.File, StringComparer.Ordinal)
                .ToArray();

            var report = new ValidationReport(options.Format.ToString(), results);
            Console.WriteLine(JsonSerializer.Serialize(report, JsonOptions));
            return results.All(result => result.Valid) ? Success : ValidationFailure;
        }
        catch (CliException exception)
        {
            Console.Error.WriteLine(exception.Message);
            Console.Error.WriteLine("Usage: ooxml-validator [--format Microsoft365] <file.xlsx> [more.xlsx ...]");
            return ToolFailure;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine($"OOXML validator failed: {exception.Message}");
            return ToolFailure;
        }
    }

    private static FileValidationResult ValidateFile(string file, FileFormatVersions format)
    {
        try
        {
            using var document = SpreadsheetDocument.Open(file, false);
            var validator = new OpenXmlValidator(format) { MaxNumberOfErrors = MaxErrorsPerFile };
            var errors = validator.Validate(document)
                .Select(ToDiagnostic)
                .OrderBy(error => error.PartUri, StringComparer.Ordinal)
                .ThenBy(error => error.XPath, StringComparer.Ordinal)
                .ThenBy(error => error.Id, StringComparer.Ordinal)
                .ThenBy(error => error.Description, StringComparer.Ordinal)
                .ToArray();

            return new FileValidationResult(file, errors.Length == 0, errors);
        }
        catch (Exception exception)
        {
            var error = new ValidationDiagnostic(
                "PackageOpenError",
                "Package",
                exception.Message,
                null,
                null);
            return new FileValidationResult(file, false, [error]);
        }
    }

    private static ValidationDiagnostic ToDiagnostic(ValidationErrorInfo error)
    {
        return new ValidationDiagnostic(
            error.Id ?? "UnknownValidationError",
            error.ErrorType.ToString(),
            error.Description ?? "OpenXmlValidator returned no description.",
            error.Part?.Uri.ToString() ?? error.Path?.PartUri?.ToString(),
            error.Path?.XPath);
    }
}

internal sealed record ValidationReport(string Format, IReadOnlyList<FileValidationResult> Results);

internal sealed record FileValidationResult(
    string File,
    bool Valid,
    IReadOnlyList<ValidationDiagnostic> Errors);

internal sealed record ValidationDiagnostic(
    string Id,
    string Type,
    string Description,
    string? PartUri,
    [property: JsonPropertyName("xpath")] string? XPath);

internal sealed record CliOptions(FileFormatVersions Format, IReadOnlyList<string> Files)
{
    public static CliOptions Parse(IReadOnlyList<string> arguments)
    {
        var format = FileFormatVersions.Microsoft365;
        var files = new List<string>();

        for (var index = 0; index < arguments.Count; index += 1)
        {
            var argument = arguments[index];
            if (argument == "--format")
            {
                if (index + 1 >= arguments.Count)
                {
                    throw new CliException("--format requires a FileFormatVersions value.");
                }

                var value = arguments[index + 1];
                if (!Enum.TryParse(value, true, out format) || !Enum.IsDefined(format))
                {
                    throw new CliException($"Unsupported file format version: {value}");
                }

                index += 1;
                continue;
            }

            if (argument.StartsWith("-", StringComparison.Ordinal))
            {
                throw new CliException($"Unknown option: {argument}");
            }

            files.Add(Path.GetFullPath(argument));
        }

        if (files.Count == 0)
        {
            throw new CliException("At least one .xlsx file is required.");
        }

        foreach (var file in files)
        {
            if (!string.Equals(Path.GetExtension(file), ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                throw new CliException($"Only .xlsx files are supported: {file}");
            }

            if (!File.Exists(file))
            {
                throw new CliException($"File does not exist: {file}");
            }
        }

        return new CliOptions(format, files);
    }
}

internal sealed class CliException(string message) : Exception(message);
