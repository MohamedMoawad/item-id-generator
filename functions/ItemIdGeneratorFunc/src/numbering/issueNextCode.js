'use strict';

const { applyFormula, periodRolledOver } = require('./formatCode');

/**
 * Reserves the next sequence with a compare-and-swap on the config row.
 * Callers serialize a list with a lease. This loop is the backstop when two
 * writers still read the same ETag: only one MERGE with that ETag succeeds.
 */
async function issueNextCode(store, now, options = {}) {
  const maxAttempts = options.maxAttempts ?? 200;
  const sleep = options.sleep ?? defaultSleep;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    const current = await store.read();
    if (!current) {
      throw new Error('Numbering config row was not found.');
    }
    if (current.isActive === false) {
      throw new Error('Numbering config row is inactive.');
    }

    const rolledOver = periodRolledOver(current.resetPeriod, current.lastResetDate, now);
    const sequence = (rolledOver ? 0 : Number(current.currentCount) || 0) + 1;
    const code = applyFormula(current.formula, {
      sequence,
      date: now,
      padLength: current.padLength
    });
    const swapped = await store.compareAndSwap(current.etag, {
      currentCount: sequence,
      lastResetDate: now.toISOString()
    });

    if (swapped.ok) {
      return { code, sequence, lastResetDate: now.toISOString() };
    }

    if (attempt < maxAttempts) {
      await sleep(Math.min(500, 15 * attempt));
    }
  }

  throw new Error('Could not reserve a unique sequence after retrying the config row.');
}

function defaultSleep(milliseconds) {
  return new Promise((resolve) => {
    setTimeout(resolve, milliseconds);
  });
}

module.exports = {
  issueNextCode
};
