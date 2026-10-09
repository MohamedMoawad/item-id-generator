using System.Globalization;
using System.Text.RegularExpressions;

namespace ItemIdGenerator.Numbering;

public static class Formula
{
    private static readonly Regex Token = new(@"\{([^{}]+)\}", RegexOptions.Compiled);

    public static string NormalizeResetPeriod(string? value)
    {
        var text = (value ?? string.Empty).Trim().ToLowerInvariant();
        return text is "day" or "month" or "year" ? text : "none";
    }

    public static string PeriodKeyFor(string? resetPeriod, DateTimeOffset date)
    {
        var period = NormalizeResetPeriod(resetPeriod);
        var utc = date.ToUniversalTime();
        return period switch
        {
            "day" => utc.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            "month" => utc.ToString("yyyy-MM", CultureInfo.InvariantCulture),
            "year" => utc.ToString("yyyy", CultureInfo.InvariantCulture),
            _ => string.Empty
        };
    }

    public static bool PeriodRolledOver(string? resetPeriod, string? lastResetDate, DateTimeOffset now)
    {
        var period = NormalizeResetPeriod(resetPeriod);
        if (period == "none" || string.IsNullOrWhiteSpace(lastResetDate))
        {
            return false;
        }

        if (!DateTimeOffset.TryParse(lastResetDate, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var last))
        {
            return false;
        }

        return PeriodKeyFor(period, last) != PeriodKeyFor(period, now);
    }

    public static string Apply(
        string? formula,
        long sequence,
        DateTimeOffset date,
        int padLength,
        IReadOnlyDictionary<string, string>? columns = null)
    {
        if (string.IsNullOrEmpty(formula) ||
            (formula.IndexOf("{seq", StringComparison.OrdinalIgnoreCase) < 0 &&
             formula.IndexOf("{counter", StringComparison.OrdinalIgnoreCase) < 0))
        {
            throw new InvalidOperationException("Formula must include {counter} or {counter:n} so each item gets a distinct code.");
        }

        if (padLength < 0 || padLength > 12)
        {
            throw new InvalidOperationException("PadLength must be from 0 to 12.");
        }

        var utc = date.ToUniversalTime();
        var values = new Dictionary<string, string>
        {
            ["yyyy"] = utc.ToString("yyyy", CultureInfo.InvariantCulture),
            ["yy"] = utc.ToString("yy", CultureInfo.InvariantCulture),
            ["MM"] = utc.ToString("MM", CultureInfo.InvariantCulture),
            ["dd"] = utc.ToString("dd", CultureInfo.InvariantCulture),
            ["HH"] = utc.ToString("HH", CultureInfo.InvariantCulture),
            ["mm"] = utc.ToString("mm", CultureInfo.InvariantCulture)
        };

        return Token.Replace(formula, match =>
        {
            var token = match.Groups[1].Value.Trim();
            if (token.Equals("seq", StringComparison.OrdinalIgnoreCase) || token.Equals("counter", StringComparison.OrdinalIgnoreCase))
            {
                return padLength > 0 ? sequence.ToString(CultureInfo.InvariantCulture).PadLeft(padLength, '0') : sequence.ToString(CultureInfo.InvariantCulture);
            }

            var widthToken = token.StartsWith("seq:", StringComparison.OrdinalIgnoreCase) || token.StartsWith("counter:", StringComparison.OrdinalIgnoreCase)
                ? token[(token.IndexOf(':') + 1)..]
                : null;
            if (widthToken is not null)
            {
                if (!int.TryParse(widthToken, NumberStyles.None, CultureInfo.InvariantCulture, out var width) || width < 1 || width > 12)
                {
                    throw new InvalidOperationException("Counter width must be from 1 to 12.");
                }

                return sequence.ToString(CultureInfo.InvariantCulture).PadLeft(width, '0');
            }

            if (values.TryGetValue(token, out var value))
            {
                return value;
            }

            if (columns is null)
            {
                return match.Value;
            }

            foreach (var pair in columns)
            {
                if (pair.Key.Equals(token, StringComparison.OrdinalIgnoreCase))
                {
                    return pair.Value ?? string.Empty;
                }
            }

            return string.Empty;
        });
    }
}
