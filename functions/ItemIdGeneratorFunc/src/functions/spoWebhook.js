'use strict';

const { app, output } = require('@azure/functions');
const { handleSharePointWebhook } = require('../webhook/handleSharePointWebhook');

const queueOutput = output.storageQueue({
  queueName: 'request-numbers',
  connection: 'AzureWebJobsStorage'
});

app.http('spoWebhook', {
  methods: ['POST'],
  authLevel: 'function',
  extraOutputs: [queueOutput],
  handler: async (request, context) => handleSharePointWebhook(request, context, {
    enqueue: async (messages) => {
      context.extraOutputs.set(queueOutput, messages);
    }
  })
});
