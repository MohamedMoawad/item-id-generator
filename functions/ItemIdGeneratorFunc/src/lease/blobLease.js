'use strict';

const CONTAINER_NAME = 'numbering-locks';

function blobNameForList(listId) {
  const safe = String(listId).trim().toLowerCase().replace(/[^a-z0-9-]/g, '');
  if (!safe) {
    throw new Error('Cannot lease an empty list id.');
  }
  return `list-${safe}`;
}

function isConflict(error) {
  const status = error && (error.statusCode || error.status);
  const code = error && error.code;
  return status === 409 || code === 'BlobAlreadyExists' || code === 'LeaseAlreadyPresent';
}

/**
 * Lease one append/block blob per list GUID for 60 seconds.
 * A second worker gets { acquired: false } and must let the queue retry.
 */
async function acquireBlobLease(blob) {
  try {
    await blob.upload('', 0, { conditions: { ifNoneMatch: '*' } });
  } catch (error) {
    if (!isConflict(error)) {
      throw error;
    }
  }

  const leaseClient = blob.getBlobLeaseClient();
  try {
    const acquired = await leaseClient.acquireLease(60);
    const leaseId = acquired.leaseId;
    return {
      acquired: true,
      async release() {
        await blob.getBlobLeaseClient(leaseId).releaseLease();
      }
    };
  } catch (error) {
    if (isConflict(error)) {
      return { acquired: false };
    }
    throw error;
  }
}

function createBlobLeaseStore(connectionString, options = {}) {
  if (!connectionString || connectionString.indexOf('<') >= 0) {
    throw new Error('AzureWebJobsStorage must be set to a storage connection string before numbering can lease a list.');
  }
  const sdk = options.blobServiceClient || require('@azure/storage-blob').BlobServiceClient;
  const service = typeof sdk.fromConnectionString === 'function'
    ? sdk.fromConnectionString(connectionString)
    : sdk;
  const container = service.getContainerClient(CONTAINER_NAME);

  return {
    async acquire(listId) {
      await container.createIfNotExists();
      const blob = container.getBlockBlobClient(blobNameForList(listId));
      return acquireBlobLease(blob);
    }
  };
}

module.exports = {
  CONTAINER_NAME,
  acquireBlobLease,
  blobNameForList,
  createBlobLeaseStore
};
