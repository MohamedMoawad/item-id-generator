'use strict';

const { app } = require('@azure/functions');
const { createBlobLeaseStore } = require('../lease/blobLease');
const { processNumberingMessage } = require('../queue/processNumberingMessage');
const { createNumberingGateway } = require('../sharepoint/numberingGateway');

app.storageQueue('processRequestNumber', {
  queueName: 'request-numbers',
  connection: 'AzureWebJobsStorage',
  handler: async (message, context) => {
    const payload = typeof message === 'string' ? JSON.parse(message) : message;
    const entries = Array.isArray(payload) ? payload : [payload];
    const leaseStore = createBlobLeaseStore(process.env.AzureWebJobsStorage);
    const gateway = createNumberingGateway(process.env);
    for (const entry of entries) {
      await processNumberingMessage(entry, context, { leaseStore, gateway });
    }
  }
});
