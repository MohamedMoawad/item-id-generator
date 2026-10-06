'use strict';

const { app } = require('@azure/functions');
const { handleRegisterWebhook } = require('../webhook/handleRegisterWebhook');

app.http('RegisterWebhook', {
  methods: ['POST'],
  authLevel: 'function',
  handler: async (request, context) => handleRegisterWebhook(request, context)
});
