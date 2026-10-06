'use strict';

/**
 * Condition rules for generated item IDs.
 *
 * First matching rule wins. registerRule() inserts at the front so a new
 * rule takes priority over the built-in VIP rule. To run a rule only when
 * VIP did not match, splice it into `rules` after the VIP entry instead.
 *
 * Context passed to matches() and format():
 *   { listName, condition, itemId, metadata, timestamp }
 *
 * listName is the trimmed title with whitespace collapsed to hyphens.
 * timestamp is Unix epoch milliseconds. metadata holds any JSON fields
 * other than listName, condition, and itemId (department, role, column values).
 *
 * Example department rule:
 *
 *   registerRule({
 *     id: 'department',
 *     matches: (ctx) => typeof ctx.metadata.department === 'string' && ctx.metadata.department.trim() !== '',
 *     format: (ctx) => `${ctx.metadata.department.trim()}-${ctx.listName}-${ctx.timestamp}`
 *   });
 */

const rules = [
  {
    id: 'vip',
    matches: (ctx) => normalizeCondition(ctx.condition) === 'vip',
    format: (ctx) => `VIP-${ctx.listName}-${ctx.timestamp}`
  }
];

function normalizeCondition(condition) {
  return String(condition ?? '').trim().toLowerCase();
}

function normalizeListName(listName) {
  return String(listName ?? '').trim().replace(/\s+/g, '-');
}

function isPlainObject(value) {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function registerRule(rule) {
  if (!rule || typeof rule.id !== 'string' || rule.id.trim() === '') {
    throw new TypeError('Rule id is required.');
  }
  if (typeof rule.matches !== 'function' || typeof rule.format !== 'function') {
    throw new TypeError('Rule needs matches(ctx) and format(ctx).');
  }
  if (rules.some((existing) => existing.id === rule.id)) {
    throw new Error(`Rule "${rule.id}" is already registered.`);
  }
  rules.unshift(rule);
  return function unregister() {
    const index = rules.indexOf(rule);
    if (index >= 0) {
      rules.splice(index, 1);
    }
  };
}

function generateItemId(input, now = new Date()) {
  if (!input || typeof input.listName !== 'string' || input.listName.trim() === '') {
    throw new TypeError('listName is required.');
  }

  const ctx = {
    listName: normalizeListName(input.listName),
    condition: input.condition,
    itemId: input.itemId,
    metadata: isPlainObject(input.metadata) ? input.metadata : {},
    timestamp: String(now.getTime())
  };

  for (const rule of rules) {
    if (rule.matches(ctx)) {
      return rule.format(ctx);
    }
  }

  return `${ctx.listName}-${ctx.timestamp}`;
}

module.exports = {
  rules,
  registerRule,
  generateItemId,
  normalizeCondition,
  normalizeListName
};
