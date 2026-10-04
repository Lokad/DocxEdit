using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace Lokad.DocxEdit;

internal sealed partial class DocxCommandExecution
{
    private sealed class ParsedOptions
    {
        public required string Command { get; init; }
        public required IReadOnlyList<string> Positionals { get; init; }
        public required HashSet<string> Flags { get; init; }
        public required bool Json { get; init; }
        public required bool Strict { get; init; }
        public required string? DiagnosticsPath { get; init; }
        public required string? ReportPath { get; init; }
        public required string? OperationReportPath { get; init; }
        public required string? OutputPath { get; init; }
        public required DocxPaperSize PaperSize { get; init; }
        public required DocxOrientation Orientation { get; init; }
        public required string? Id { get; init; }
        public required string? ExtractPath { get; init; }
        public required int? MaxText { get; init; }
        public required int? MaxCommentText { get; init; }
        public required int? MaxDiagnostics { get; init; }
        public required int? MaxPreviewChars { get; init; }
        public required int? Radius { get; init; }
        public required TrackChangesMode TrackChanges { get; init; }
        public required string? Author { get; init; }
        public required DateTimeOffset? TimestampUtc { get; init; }
        public required DocxTextView TextView { get; init; }
        public required DocxValidationProfile ValidationProfile { get; init; }
        public required string? Error { get; init; }

        public static ParsedOptions Parse(string[] args)
        {
            string command = args[0];
            var positionals = new List<string>();
            var flags = new HashSet<string>(StringComparer.Ordinal);
            bool json = false;
            bool strict = false;
            string? diagnosticsPath = null;
            string? reportPath = null;
            string? operationReportPath = null;
            string? outputPath = null;
            DocxPaperSize paperSize = DocxPaperSize.A4;
            DocxOrientation orientation = DocxOrientation.Portrait;
            string? id = null;
            string? extractPath = null;
            int? maxText = null;
            int? maxCommentText = null;
            int? maxDiagnostics = null;
            int? maxPreviewChars = null;
            int? radius = null;
            TrackChangesMode trackChanges = TrackChangesMode.Off;
            string? author = null;
            DateTimeOffset? timestampUtc = null;
            DocxTextView textView = DocxTextView.Final;
            DocxValidationProfile validationProfile = DocxValidationProfile.Structural;
            var seenFlags = new HashSet<string>(StringComparer.Ordinal);

            static bool TryParseTimestampUtc(string value, out DateTimeOffset timestamp)
            {
                string[] formats = ["O", "yyyy-MM-ddTHH:mm:ss.FFFFFFFK", "yyyy-MM-ddTHH:mm:ssK", "yyyy-MM-dd"];
                if (DateTimeOffset.TryParseExact(value, formats, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out timestamp))
                {
                    timestamp = timestamp.ToUniversalTime();
                    return true;
                }

                timestamp = default;
                return false;
            }

            for (int i = 1; i < args.Length; i++)
            {
                string arg = args[i];
                switch (arg)
                {
                    case "--paper":
                        seenFlags.Add(arg);
                        if (!TryReadValue(args, ref i, out string? paperValue))
                        {
                            return WithError(command, "Missing value for --paper.");
                        }
                        if (!DocxPaperSizeExtensions.TryParseWireValue(paperValue, out paperSize))
                        {
                            return WithError(command, "Invalid value for --paper. Expected a4 or letter.");
                        }
                        break;
                    case "--orientation":
                        seenFlags.Add(arg);
                        if (!TryReadValue(args, ref i, out string? orientationValue))
                        {
                            return WithError(command, "Missing value for --orientation.");
                        }
                        if (!DocxOrientationExtensions.TryParseWireValue(orientationValue, out orientation))
                        {
                            return WithError(command, "Invalid value for --orientation. Expected portrait or landscape.");
                        }
                        break;
                    case "--json":
                        seenFlags.Add("--json");
                        json = true;
                        break;
                    case "--strict":
                        seenFlags.Add("--strict");
                        strict = true;
                        break;
                    case "--compact":
                        seenFlags.Add(arg);
                        flags.Add(arg);
                        break;
                    case "--runs":
                    case "--headers-footers":
                    case "--summary":
                    case "--include-comment-text":
                        seenFlags.Add(arg);
                        flags.Add(arg);
                        break;
                    case "--diagnostics":
                        seenFlags.Add("--diagnostics");
                        if (!TryReadValue(args, ref i, out diagnosticsPath))
                        {
                            return WithError(command, "Missing value for --diagnostics.");
                        }

                        break;
                    case "--report":
                        seenFlags.Add("--report");
                        if (!TryReadValue(args, ref i, out reportPath))
                        {
                            return WithError(command, "Missing value for --report.");
                        }

                        break;
                    case "--operation-report":
                        seenFlags.Add("--operation-report");
                        if (!TryReadValue(args, ref i, out operationReportPath))
                        {
                            return WithError(command, "Missing value for --operation-report.");
                        }

                        break;
                    case "--extract":
                        seenFlags.Add("--extract");
                        if (!TryReadValue(args, ref i, out extractPath))
                        {
                            return WithError(command, "Missing value for --extract.");
                        }

                        break;
                    case "--max-text":
                        seenFlags.Add("--max-text");
                        if (!TryReadValue(args, ref i, out string? maxTextValue))
                        {
                            return WithError(command, "Missing value for --max-text.");
                        }

                        if (!int.TryParse(maxTextValue, out int parsedMaxText) || parsedMaxText < 0)
                        {
                            return WithError(command, "Invalid value for --max-text.");
                        }

                        maxText = parsedMaxText;
                        break;
                    case "--max-comment-text":
                        seenFlags.Add("--max-comment-text");
                        if (!TryReadValue(args, ref i, out string? maxCommentTextValue))
                        {
                            return WithError(command, "Missing value for --max-comment-text.");
                        }

                        if (!int.TryParse(maxCommentTextValue, out int parsedMaxCommentText) || parsedMaxCommentText < 0)
                        {
                            return WithError(command, "Invalid value for --max-comment-text.");
                        }

                        maxCommentText = parsedMaxCommentText;
                        break;
                    case "--max-diagnostics":
                        seenFlags.Add("--max-diagnostics");
                        if (!TryReadValue(args, ref i, out string? maxDiagnosticsValue))
                        {
                            return WithError(command, "Missing value for --max-diagnostics.");
                        }

                        if (!int.TryParse(maxDiagnosticsValue, out int parsedMaxDiagnostics) || parsedMaxDiagnostics < 1)
                        {
                            return WithError(command, "Invalid value for --max-diagnostics.");
                        }

                        maxDiagnostics = parsedMaxDiagnostics;
                        break;
                    case "--max-preview-chars":
                        seenFlags.Add("--max-preview-chars");
                        if (!TryReadValue(args, ref i, out string? maxPreviewCharsValue))
                        {
                            return WithError(command, "Missing value for --max-preview-chars.");
                        }

                        if (!int.TryParse(maxPreviewCharsValue, out int parsedMaxPreviewChars) || parsedMaxPreviewChars < 0)
                        {
                            return WithError(command, "Invalid value for --max-preview-chars.");
                        }

                        maxPreviewChars = parsedMaxPreviewChars;
                        break;
                    case "--radius":
                        seenFlags.Add("--radius");
                        if (!TryReadValue(args, ref i, out string? radiusValue))
                        {
                            return WithError(command, "Missing value for --radius.");
                        }

                        if (!int.TryParse(radiusValue, out int parsedRadius) || parsedRadius < 0)
                        {
                            return WithError(command, "Invalid value for --radius.");
                        }

                        radius = parsedRadius;
                        break;
                    case "--track-changes":
                        seenFlags.Add("--track-changes");
                        if (!TryReadValue(args, ref i, out string? trackChangesValue))
                        {
                            return WithError(command, "Missing value for --track-changes.");
                        }

                        if (!TryParseTrackChangesMode(trackChangesValue, out trackChanges))
                        {
                            return WithError(command, "Invalid value for --track-changes. Expected off, preserve, suggest, or require.");
                        }

                        break;
                    case "--view":
                        seenFlags.Add("--view");
                        if (!TryReadValue(args, ref i, out string? textViewValue))
                        {
                            return WithError(command, "Missing value for --view.");
                        }

                        if (!TryParseTextView(textViewValue, out textView))
                        {
                            return WithError(command, "Invalid value for --view. Expected final, original, or markup.");
                        }

                        break;
                    case "--profile":
                        seenFlags.Add("--profile");
                        if (!TryReadValue(args, ref i, out string? profileValue))
                        {
                            return WithError(command, "Missing value for --profile.");
                        }

                        if (!TryParseValidationProfile(profileValue, out validationProfile))
                        {
                            return WithError(command, "Invalid value for --profile. Expected structural or package.");
                        }

                        break;
                    case "--author":
                        seenFlags.Add("--author");
                        if (!TryReadValue(args, ref i, out author))
                        {
                            return WithError(command, "Missing value for --author.");
                        }

                        break;
                    case "--timestamp-utc":
                        seenFlags.Add("--timestamp-utc");
                        if (!TryReadValue(args, ref i, out string? timestampValue))
                        {
                            return WithError(command, "Missing value for --timestamp-utc.");
                        }

                        if (!TryParseTimestampUtc(timestampValue, out DateTimeOffset parsedTimestamp))
                        {
                            return WithError(command, $"Invalid value for --timestamp-utc '{timestampValue}'. Expected ISO-8601 UTC, e.g. 2026-01-01T00:00:00Z.");
                        }

                        timestampUtc = parsedTimestamp;
                        break;
                    case "-o":
                    case "--output":
                        seenFlags.Add("--output");
                        if (!TryReadValue(args, ref i, out outputPath))
                        {
                            return WithError(command, $"Missing value for {arg}.");
                        }

                        break;
                    case "--id":
                        seenFlags.Add("--id");
                        if (!TryReadValue(args, ref i, out id))
                        {
                            return WithError(command, "Missing value for --id.");
                        }

                        break;
                    default:
                        if (arg != "-" && arg.StartsWith('-'))
                        {
                            return WithError(command, $"Unknown option '{arg}'.");
                        }

                        positionals.Add(arg);
                        break;
                }
            }

            if (DocxHelp.TryGetCommand(command, out DocxCommandInfo commandInfo))
            {
                foreach (string seen in seenFlags.OrderBy(static flag => flag, StringComparer.Ordinal))
                {
                    if (!commandInfo.Options.Any(option => option.Flags.Contains(seen, StringComparer.Ordinal)))
                    {
                        return WithError(command, $"{seen} is not an option of the {command} command.");
                    }
                }

                if (positionals.Count < commandInfo.MinPositionals || positionals.Count > commandInfo.MaxPositionals)
                {
                    return WithError(command, ArityError(command, commandInfo, positionals.Count));
                }
            }

            static string ArityError(string commandName, DocxCommandInfo info, int actual)
            {
                if (info.MinPositionals == info.MaxPositionals && info.MinPositionals == 1)
                {
                    return $"The {commandName} command expects exactly 1 positional argument; received {actual}.";
                }

                string expected = info.MinPositionals == info.MaxPositionals
                    ? $"exactly {info.MinPositionals} positional arguments"
                    : $"between {info.MinPositionals} and {info.MaxPositionals} positional arguments";
                return $"The {commandName} command expects {expected}; received {actual}.";
            }

            if (diagnosticsPath == "-")
            {
                return WithError(command, "--diagnostics does not accept '-'; diagnostics output is file-only.");
            }

            if (reportPath == "-")
            {
                return WithError(command, "--report does not accept '-'; reports are file-only.");
            }

            if (operationReportPath == "-")
            {
                return WithError(command, "--operation-report does not accept '-'; operation reports are file-only.");
            }

            return new ParsedOptions
            {
                Command = command,
                Positionals = positionals,
                Flags = flags,
                Json = json,
                Strict = strict,
                DiagnosticsPath = diagnosticsPath,
                ReportPath = reportPath,
                OperationReportPath = operationReportPath,
                OutputPath = outputPath,
                PaperSize = paperSize,
                Orientation = orientation,
                Id = id,
                ExtractPath = extractPath,
                MaxText = maxText,
                MaxCommentText = maxCommentText,
                MaxDiagnostics = maxDiagnostics,
                MaxPreviewChars = maxPreviewChars,
                Radius = radius,
                TrackChanges = trackChanges,
                Author = author,
                TimestampUtc = timestampUtc,
                TextView = textView,
                ValidationProfile = validationProfile,
                Error = null
            };
        }

        private static bool TryReadValue(string[] args, ref int index, [NotNullWhen(true)] out string? value)
        {
            if (index + 1 >= args.Length)
            {
                value = null;
                return false;
            }

            value = args[++index];
            return true;
        }

        private static ParsedOptions WithError(string command, string message)
        {
            string detail = DocxHelp.TryGetCommand(command, out DocxCommandInfo _) ? CommandUsageError(command) : $"Valid commands: {ValidCommands()}\nSee 'docxedit --help'.";
            return new ParsedOptions
            {
                Command = command,
                Positionals = [],
                Flags = new HashSet<string>(StringComparer.Ordinal),
                Json = false,
                Strict = false,
                DiagnosticsPath = null,
                ReportPath = null,
                OperationReportPath = null,
                OutputPath = null,
                PaperSize = DocxPaperSize.A4,
                Orientation = DocxOrientation.Portrait,
                Id = null,
                ExtractPath = null,
                MaxText = null,
                MaxCommentText = null,
                MaxDiagnostics = null,
                MaxPreviewChars = null,
                Radius = null,
                TrackChanges = TrackChangesMode.Off,
                Author = null,
                TimestampUtc = null,
                TextView = DocxTextView.Final,
                ValidationProfile = DocxValidationProfile.Structural,
                Error = $"{message}\n{detail}"
            };
        }

        private static bool TryParseTrackChangesMode(string value, out TrackChangesMode mode)
        {
            string normalized = value.ToLowerInvariant();
            mode = normalized switch
            {
                "off" => TrackChangesMode.Off,
                "preserve" => TrackChangesMode.Preserve,
                "suggest" => TrackChangesMode.Suggest,
                "require" => TrackChangesMode.Require,
                _ => TrackChangesMode.Off
            };
            return normalized is "off" or "preserve" or "suggest" or "require";
        }

        private static bool TryParseValidationProfile(string value, out DocxValidationProfile profile)
        {
            switch (value.ToLowerInvariant())
            {
                case "structural":
                    profile = DocxValidationProfile.Structural;
                    return true;
                case "package":
                    profile = DocxValidationProfile.Package;
                    return true;
                default:
                    profile = DocxValidationProfile.Structural;
                    return false;
            }
        }

        private static bool TryParseTextView(string value, out DocxTextView textView)
        {
            switch (value.ToLowerInvariant())
            {
                case "final":
                    textView = DocxTextView.Final;
                    return true;
                case "original":
                    textView = DocxTextView.Original;
                    return true;
                case "markup":
                    textView = DocxTextView.Markup;
                    return true;
                default:
                    textView = DocxTextView.Final;
                    return false;
            }
        }
    }
}
