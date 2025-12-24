using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docx2Md.Core.Models;

namespace Docx2Md.Core.Parsing;

/// <summary>
/// Resolves Word numbering definitions to actual prefix text.
/// Handles multi-level numbering with LevelText patterns like "%1.", "Article %1", "%1.%2."
/// </summary>
public class NumberingResolver
{
    private readonly NumberingDefinitionsPart? _numberingPart;

    // Cache: abstractNumId -> AbstractNum
    private readonly Dictionary<int, AbstractNum> _abstractNumCache = new();

    // Cache: abstractNumId -> nsid (numbering style ID, used to link related numbering definitions)
    private readonly Dictionary<int, string> _abstractNumToNsid = new();

    // Cache: (numId, level) -> LevelInfo
    private readonly Dictionary<(int numId, int level), LevelInfo> _levelInfoCache = new();

    // Cache: numId -> abstractNumId
    private readonly Dictionary<int, int> _numIdToAbstractNumId = new();

    // Multi-level counter state: counterKey -> (level -> current count)
    // counterKey is either nsid (if available) or abstractNumId.ToString()
    // Using nsid ensures continuous numbering across sections that share the same numbering style
    private readonly Dictionary<string, Dictionary<int, int>> _levelCounters = new();

    public NumberingResolver(WordprocessingDocument wordDoc)
    {
        _numberingPart = wordDoc.MainDocumentPart?.NumberingDefinitionsPart;
        CacheAbstractNums();
    }

    private void CacheAbstractNums()
    {
        if (_numberingPart?.Numbering == null)
            return;

        foreach (var abstractNum in _numberingPart.Numbering.Elements<AbstractNum>())
        {
            var id = abstractNum.AbstractNumberId?.Value;
            if (id.HasValue)
            {
                _abstractNumCache[id.Value] = abstractNum;

                // Cache the nsid (numbering style ID) if present
                // nsid links related numbering definitions that should share state
                var nsid = abstractNum.Nsid?.Val?.Value;
                if (!string.IsNullOrEmpty(nsid))
                {
                    _abstractNumToNsid[id.Value] = nsid;
                }
            }
        }
    }

    /// <summary>
    /// Get the counter key for an abstractNumId (uses nsid if available for cross-section continuity)
    /// </summary>
    private string GetCounterKey(int abstractNumId)
    {
        if (_abstractNumToNsid.TryGetValue(abstractNumId, out var nsid))
        {
            return nsid;
        }
        return abstractNumId.ToString();
    }

    /// <summary>
    /// Get numbering info for a paragraph with the given numId and level
    /// </summary>
    public LevelInfo? GetLevelInfo(int numId, int level)
    {
        var key = (numId, level);
        if (_levelInfoCache.TryGetValue(key, out var cached))
            return cached;

        if (_numberingPart?.Numbering == null)
            return null;

        // Find the numbering instance
        var numberingInstance = _numberingPart.Numbering
            .Elements<NumberingInstance>()
            .FirstOrDefault(ni => ni.NumberID?.Value == numId);

        if (numberingInstance?.AbstractNumId?.Val == null)
            return null;

        var abstractNumId = numberingInstance.AbstractNumId.Val.Value;

        // Cache the numId -> abstractNumId mapping
        _numIdToAbstractNumId[numId] = abstractNumId;

        if (!_abstractNumCache.TryGetValue(abstractNumId, out var abstractNum))
            return null;

        // Find the level definition
        var levelDef = abstractNum.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == level);

        if (levelDef == null)
            return null;

        var info = new LevelInfo
        {
            LevelText = levelDef.LevelText?.Val?.Value,
            Format = levelDef.NumberingFormat?.Val?.Value ?? NumberFormatValues.Decimal,
            StartValue = levelDef.StartNumberingValue?.Val?.Value ?? 1,
            LevelIndex = level,
            AbstractNumId = abstractNumId
        };

        // Cache format info for ALL levels (needed for %1.%2.%3 resolution)
        foreach (var lvl in abstractNum.Elements<Level>())
        {
            var lvlIdx = lvl.LevelIndex?.Value ?? 0;
            var fmt = lvl.NumberingFormat?.Val?.Value ?? NumberFormatValues.Decimal;
            info.AllLevelFormats[lvlIdx] = fmt;
        }

        _levelInfoCache[key] = info;
        return info;
    }

    /// <summary>
    /// Resolve the LevelText format to actual text by:
    /// 1. Updating counters for this level
    /// 2. Substituting %1, %2, etc. with formatted numbers
    /// </summary>
    public string? ResolveNumberingPrefix(int numId, int level, LevelInfo info, SegmentType segmentType = SegmentType.ListItem)
    {
        if (string.IsNullOrEmpty(info.LevelText))
            return null;

        var counterKey = GetCounterKey(info.AbstractNumId);

        // Update counters using counterKey (nsid or abstractNumId)
        // This ensures continuous numbering across sections that share the same numbering style
        UpdateCounters(counterKey, level, info.StartValue);

        var counters = _levelCounters[counterKey];

        // Replace placeholders with formatted numbers
        // %1 refers to level 0, %2 to level 1, etc.
        var result = Regex.Replace(info.LevelText, @"%(\d+)", match =>
        {
            var placeholderLevel = int.Parse(match.Groups[1].Value) - 1; // %1 = level 0
            if (counters.TryGetValue(placeholderLevel, out var count))
            {
                var format = info.AllLevelFormats.GetValueOrDefault(placeholderLevel, NumberFormatValues.Decimal);
                return FormatNumber(count, format);
            }
            return match.Value; // Keep placeholder if level not found
        });

        return result;
    }

    /// <summary>
    /// Get the current count for a specific numId and level
    /// </summary>
    public int GetCurrentCount(int numId, int level)
    {
        // Get the abstractNumId for this numId
        if (!_numIdToAbstractNumId.TryGetValue(numId, out var abstractNumId))
            return 1;

        var counterKey = GetCounterKey(abstractNumId);

        if (_levelCounters.TryGetValue(counterKey, out var counters) &&
            counters.TryGetValue(level, out var count))
        {
            return count;
        }
        return 1;
    }

    private void UpdateCounters(string counterKey, int level, int startValue)
    {
        // Initialize counters for this counterKey if needed
        if (!_levelCounters.ContainsKey(counterKey))
        {
            _levelCounters[counterKey] = new Dictionary<int, int>();
        }

        var counters = _levelCounters[counterKey];

        // NEVER reset counters - Word documents have explicit control over numbering
        // Using nsid (numbering style ID) ensures continuous numbering across sections
        // that share the same underlying numbering definition

        if (!counters.ContainsKey(level))
        {
            // First occurrence at this level - start at startValue
            counters[level] = startValue;

            // Also initialize all parent levels if they don't exist
            for (int parentLevel = 0; parentLevel < level; parentLevel++)
            {
                if (!counters.ContainsKey(parentLevel))
                {
                    counters[parentLevel] = 1; // Default to 1 for parent levels
                }
            }
        }
        else
        {
            // Same numbering scheme, same level - increment
            counters[level]++;
        }

        // Reset deeper levels when moving to a shallower level
        // (e.g., going from 1.2.3 to 2 should reset levels 1 and 2)
        foreach (var deeperLevel in counters.Keys.Where(k => k > level).ToList())
        {
            counters.Remove(deeperLevel);
        }
    }

    /// <summary>
    /// Format a number according to NumberingFormat
    /// </summary>
    public static string FormatNumber(int number, NumberFormatValues format)
    {
        // NumberFormatValues is an EnumValue<T>, so we use if-else instead of switch
        if (format == NumberFormatValues.Decimal)
            return number.ToString();
        if (format == NumberFormatValues.DecimalZero)
            return number.ToString("00");
        if (format == NumberFormatValues.LowerLetter)
            return ToLetter(number, lowercase: true);
        if (format == NumberFormatValues.UpperLetter)
            return ToLetter(number, lowercase: false);
        if (format == NumberFormatValues.LowerRoman)
            return ToRoman(number, lowercase: true);
        if (format == NumberFormatValues.UpperRoman)
            return ToRoman(number, lowercase: false);
        if (format == NumberFormatValues.Ordinal)
            return ToOrdinal(number);
        if (format == NumberFormatValues.CardinalText)
            return ToCardinalText(number);
        if (format == NumberFormatValues.OrdinalText)
            return ToOrdinalText(number);

        return number.ToString();
    }

    private static string ToLetter(int number, bool lowercase)
    {
        // 1 -> A/a, 2 -> B/b, ..., 26 -> Z/z, 27 -> AA/aa, etc.
        var result = "";
        while (number > 0)
        {
            number--; // Make 0-indexed
            result = (char)((lowercase ? 'a' : 'A') + (number % 26)) + result;
            number /= 26;
        }
        return result;
    }

    private static string ToRoman(int number, bool lowercase)
    {
        if (number <= 0 || number > 3999)
            return number.ToString();

        var romanNumerals = new (int value, string numeral)[]
        {
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
            (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
            (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
        };

        var result = "";
        foreach (var (value, numeral) in romanNumerals)
        {
            while (number >= value)
            {
                result += numeral;
                number -= value;
            }
        }

        return lowercase ? result.ToLowerInvariant() : result;
    }

    private static string ToOrdinal(int number)
    {
        // 1st, 2nd, 3rd, 4th, etc.
        var suffix = (number % 100) switch
        {
            11 or 12 or 13 => "th",
            _ => (number % 10) switch
            {
                1 => "st",
                2 => "nd",
                3 => "rd",
                _ => "th"
            }
        };
        return number + suffix;
    }

    private static string ToCardinalText(int number)
    {
        // Basic implementation for common numbers
        var ones = new[] { "", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine",
                          "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen",
                          "seventeen", "eighteen", "nineteen" };
        var tens = new[] { "", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety" };

        if (number < 20)
            return ones[number];
        if (number < 100)
            return tens[number / 10] + (number % 10 > 0 ? "-" + ones[number % 10] : "");
        if (number < 1000)
            return ones[number / 100] + " hundred" + (number % 100 > 0 ? " " + ToCardinalText(number % 100) : "");

        return number.ToString(); // Fallback for large numbers
    }

    private static string ToOrdinalText(int number)
    {
        // Basic implementation
        var ordinals = new[] { "", "first", "second", "third", "fourth", "fifth", "sixth", "seventh",
                              "eighth", "ninth", "tenth", "eleventh", "twelfth", "thirteenth",
                              "fourteenth", "fifteenth", "sixteenth", "seventeenth", "eighteenth", "nineteenth" };

        if (number < 20)
            return ordinals[number];

        // For larger numbers, just add "th" to cardinal
        var cardinal = ToCardinalText(number);
        if (cardinal.EndsWith("y"))
            return cardinal[..^1] + "ieth";
        return cardinal + "th";
    }
}

/// <summary>
/// Information about a numbering level
/// </summary>
public class LevelInfo
{
    /// <summary>
    /// The LevelText format string (e.g., "%1.", "Article %1", "%1.%2.")
    /// </summary>
    public string? LevelText { get; set; }

    /// <summary>
    /// The numbering format for this level (Decimal, LowerRoman, etc.)
    /// </summary>
    public NumberFormatValues Format { get; set; }

    /// <summary>
    /// The start value for this level (default 1)
    /// </summary>
    public int StartValue { get; set; } = 1;

    /// <summary>
    /// The level index (0-based)
    /// </summary>
    public int LevelIndex { get; set; }

    /// <summary>
    /// The abstract numbering definition ID
    /// </summary>
    public int AbstractNumId { get; set; }

    /// <summary>
    /// Format info for ALL levels (needed for %1.%2.%3 resolution)
    /// </summary>
    public Dictionary<int, NumberFormatValues> AllLevelFormats { get; set; } = new();
}
