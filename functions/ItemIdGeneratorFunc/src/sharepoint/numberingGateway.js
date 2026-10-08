'use strict';

const { issueNextCode } = require('../numbering/issueNextCode');
const {
  createConfigStore,
  createSharePointDeps,
  getEnabledConfig,
  listUnnumberedItems,
  resolveSiteCollectionUrl,
  writeNumberIfBlank
} = require('./sharePointClient');

function createNumberingGateway(env, fetchImpl = globalThis.fetch) {
  const deps = createSharePointDeps(env, fetchImpl);
  return {
    async getEnabledConfig(siteUrl, listId) {
      const siteCollection = await resolveSiteCollectionUrl(deps, siteUrl);
      return getEnabledConfig(deps, listId, siteCollection);
    },
    async listUnnumbered(config, itemId) {
      return listUnnumberedItems(deps, config, itemId);
    },
    async reserveNextCode(config, now) {
      return issueNextCode(createConfigStore(deps, config), now);
    },
    async writeNumberIfBlank(config, itemId, code) {
      return writeNumberIfBlank(deps, config, itemId, code);
    }
  };
}

module.exports = {
  createNumberingGateway
};
