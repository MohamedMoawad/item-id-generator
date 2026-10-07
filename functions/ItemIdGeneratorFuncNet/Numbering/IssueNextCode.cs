namespace ItemIdGenerator.Numbering;

public static class IssueNextCode
{
    public static async Task<IssuedCode> ReserveAsync(
        IConfigStore store,
        DateTimeOffset now,
        int maxAttempts = 200,
        Func<int, Task>? sleep = null)
    {
        sleep ??= milliseconds => Task.Delay(milliseconds);

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            var current = await store.ReadAsync();
            if (current is null)
            {
                throw new InvalidOperationException("Numbering config row was not found.");
            }

            if (!current.IsActive)
            {
                throw new InvalidOperationException("Numbering config row is inactive.");
            }

            var rolledOver = Formula.PeriodRolledOver(current.ResetPeriod, current.LastResetDate, now);
            var sequence = (rolledOver ? 0 : current.CurrentCount) + 1;
            var code = Formula.Apply(current.Formula, sequence, now, current.PadLength);
            var lastReset = ToIso(now);
            var swapped = await store.CompareAndSwapAsync(current.Etag, sequence, lastReset);
            if (swapped)
            {
                return new IssuedCode
                {
                    Code = code,
                    Sequence = sequence,
                    LastResetDate = lastReset
                };
            }

            if (attempt < maxAttempts)
            {
                await sleep(Math.Min(500, 15 * attempt));
            }
        }

        throw new InvalidOperationException("Could not reserve a unique sequence after retrying the config row.");
    }

    public static string ToIso(DateTimeOffset value)
    {
        return value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
    }
}
