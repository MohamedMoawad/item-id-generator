using Azure;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Blobs.Specialized;

namespace ItemIdGenerator.Lease;

public sealed class LeaseHold
{
    public LeaseHold(bool acquired, Func<Task>? release = null)
    {
        Acquired = acquired;
        Release = release ?? (() => Task.CompletedTask);
    }

    public bool Acquired { get; }
    public Func<Task> Release { get; }
}

public interface ILeaseStore
{
    Task<LeaseHold> AcquireAsync(string listId);
}

public sealed class BlobLeaseStore : ILeaseStore
{
    public const string ContainerName = "numbering-locks";
    private readonly string _connectionString;

    public BlobLeaseStore(string connectionString)
    {
        _connectionString = connectionString ?? string.Empty;
    }

    public static string BlobNameForList(string listId)
    {
        var safe = new string((listId ?? string.Empty).Trim().ToLowerInvariant().Where(ch => char.IsAsciiLetterOrDigit(ch) || ch == '-').ToArray());
        if (safe.Length == 0)
        {
            throw new InvalidOperationException("Cannot lease an empty list id.");
        }

        return "list-" + safe;
    }

    public async Task<LeaseHold> AcquireAsync(string listId)
    {
        if (string.IsNullOrWhiteSpace(_connectionString) || _connectionString.Contains('<', StringComparison.Ordinal))
        {
            throw new InvalidOperationException("AzureWebJobsStorage must be set to a storage connection string before numbering can lease a list.");
        }

        var service = new BlobServiceClient(_connectionString);
        var container = service.GetBlobContainerClient(ContainerName);
        await container.CreateIfNotExistsAsync();
        var blob = container.GetBlobClient(BlobNameForList(listId));
        return await AcquireBlobLeaseAsync(blob);
    }

    public static async Task<LeaseHold> AcquireBlobLeaseAsync(BlobClient blob)
    {
        try
        {
            await blob.UploadAsync(BinaryData.FromBytes(Array.Empty<byte>()), new BlobUploadOptions
            {
                Conditions = new BlobRequestConditions { IfNoneMatch = ETag.All }
            });
        }
        catch (RequestFailedException error) when (IsConflict(error))
        {
        }

        var lease = blob.GetBlobLeaseClient();
        try
        {
            var acquired = await lease.AcquireAsync(TimeSpan.FromSeconds(60));
            var leaseId = acquired.Value.LeaseId;
            return new LeaseHold(true, () => blob.GetBlobLeaseClient(leaseId).ReleaseAsync());
        }
        catch (RequestFailedException error) when (IsConflict(error))
        {
            return new LeaseHold(false);
        }
    }

    public static bool IsConflict(RequestFailedException error)
    {
        return error.Status == 409
            || string.Equals(error.ErrorCode, "BlobAlreadyExists", StringComparison.Ordinal)
            || string.Equals(error.ErrorCode, "LeaseAlreadyPresent", StringComparison.Ordinal);
    }
}
