'use strict';

const assert = require('node:assert/strict');
const { describe, it } = require('node:test');
const { acquireBlobLease, blobNameForList } = require('./blobLease');

function conflict(code) {
  const error = new Error(code);
  error.statusCode = 409;
  error.code = code;
  return error;
}

describe('blob lease', () => {
  it('names the lock blob from the list GUID', () => {
    assert.equal(
      blobNameForList('{AAAAAAAA-BBBB-4CCC-8DDD-EEEEEEEEEEEE}'),
      'list-aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee'
    );
  });

  it('acquires a 60 second lease and releases that lease id', async () => {
    const releases = [];
    const blob = {
      async upload() {
        throw conflict('BlobAlreadyExists');
      },
      getBlobLeaseClient(leaseId) {
        if (!leaseId) {
          return {
            async acquireLease(duration) {
              assert.equal(duration, 60);
              return { leaseId: 'lease-1' };
            }
          };
        }
        return {
          async releaseLease() {
            releases.push(leaseId);
          }
        };
      }
    };

    const lease = await acquireBlobLease(blob);
    assert.equal(lease.acquired, true);
    await lease.release();
    assert.deepEqual(releases, ['lease-1']);
  });

  it('reports a held lease instead of throwing', async () => {
    const blob = {
      async upload() {},
      getBlobLeaseClient() {
        return {
          async acquireLease() {
            throw conflict('LeaseAlreadyPresent');
          }
        };
      }
    };
    const lease = await acquireBlobLease(blob);
    assert.equal(lease.acquired, false);
  });
});
