import {
  batchCreateOutcome,
  clearAssignedPaints,
  columnValuesFromCreateBody,
  createdItemId,
  createSaveFailed,
  embedCreateField,
  isNewItemRequest,
  newItemJson,
  paintNumbers,
  readCreateField,
  refreshOpenList,
  revealAssignedNumbers,
  rewriteNewItemBody,
  watchNewListItems
} from './listCreateWatch';

const listGuid = 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee';
const listPath = '/sites/ops/Lists/Requests';

function jsonResponse(payload: unknown, status: number): Response {
  const text = JSON.stringify(payload);
  return {
    ok: status >= 200 && status < 300,
    status,
    clone: () => ({
      text: async () => text
    })
  } as unknown as Response;
}

describe('list create watch', () => {
  it('recognizes a new item save for this list', () => {
    expect(isNewItemRequest(
      `https://contoso.sharepoint.com/sites/ops/_api/web/lists(guid'${listGuid}')/items`,
      'POST',
      listGuid,
      listPath
    )).toBe(true);
    expect(isNewItemRequest(
      "https://contoso.sharepoint.com/sites/ops/_api/web/GetList(@a1)/AddValidateUpdateItemUsingPath?@a1='%2Fsites%2Fops%2FLists%2FRequests'",
      'POST',
      listGuid,
      listPath
    )).toBe(true);
    expect(isNewItemRequest(
      `https://contoso.sharepoint.com/sites/ops/_api/web/lists(guid'${listGuid}')/items(12)`,
      'POST',
      listGuid,
      listPath
    )).toBe(false);
  });

  it('reads the new item id from a form save', () => {
    expect(createdItemId({ Id: 12 })).toBe(12);
    expect(createdItemId({
      value: [
        { FieldName: 'Title', FieldValue: 'Desk' },
        { FieldName: 'Id', FieldValue: '18' }
      ]
    })).toBe(18);
    expect(createdItemId({
      value: [
        { FieldName: 'Title', FieldValue: 'Desk', HasException: false, ItemId: 21 }
      ]
    })).toBe(21);
    expect(batchCreateOutcome(JSON.stringify({
      value: [
        { FieldName: 'Title', FieldValue: 'Desk', HasException: false, ItemId: 21 }
      ]
    }))).toBe('saved');
    expect(createdItemId({ Title: 'Desk' })).toBeUndefined();
  });

  it('shows a generated number in a blank list cell and leaves a filled cell alone', () => {
    document.body.innerHTML = [
      '<div role="row" data-id="18">',
      '<span data-automation-key="RequestNumber"></span>',
      '</div>',
      '<div role="row" data-id="19">',
      '<span data-automation-key="RequestNumber">KEEP</span>',
      '</div>',
      '<div class="ms-Panel-main"><input id="RequestNumber" value="" /></div>'
    ].join('');
    paintNumbers([
      { itemId: 18, fieldName: 'RequestNumber', code: 'REQ-202610-0007' },
      { itemId: 19, fieldName: 'RequestNumber', code: 'REQ-202610-0008' }
    ]);
    revealAssignedNumbers([18, 19]);
    const cells = document.querySelectorAll('[data-automation-key="RequestNumber"]');
    expect(cells[0].textContent).toBe('REQ-202610-0007');
    expect(cells[1].textContent).toBe('KEEP');
    const input = document.querySelector('input');
    expect(input instanceof HTMLInputElement ? input.value : '').toBe('');
    clearAssignedPaints();
    document.body.innerHTML = '';
  });

  it('refreshes the list in place after the item form closes', () => {
    jest.useFakeTimers();
    const rects = jest.spyOn(HTMLElement.prototype, 'getClientRects').mockReturnValue({
      length: 1,
      item: () => null,
      0: { width: 32, height: 32 }
    } as unknown as DOMRectList);
    document.body.innerHTML = [
      '<div role="row" data-id="18"><span data-automation-key="RequestNumber"></span></div>',
      '<button data-automationid="refreshCommand" style="display:inline-block;width:32px;height:32px">Refresh</button>'
    ].join('');
    const button = document.querySelector('button');
    const click = jest.spyOn(button as HTMLButtonElement, 'click');
    paintNumbers([{ itemId: 18, fieldName: 'RequestNumber', code: 'REQ-1' }]);
    revealAssignedNumbers([18]);
    jest.advanceTimersByTime(500);
    expect(click).toHaveBeenCalled();
    clearAssignedPaints();
    document.body.innerHTML = '<button data-automationid="refreshCommand" style="display:inline-block;width:32px;height:32px">Refresh</button>';
    const again = document.querySelector('button');
    const refreshClick = jest.spyOn(again as HTMLButtonElement, 'click');
    refreshOpenList();
    jest.advanceTimersByTime(500);
    expect(refreshClick).toHaveBeenCalled();
    clearAssignedPaints();
    rects.mockRestore();
    jest.useRealTimers();
    document.body.innerHTML = '';
  });

  it('reads and writes the number column on a new-item form body', () => {
    const form = JSON.stringify({
      formValues: [
        { FieldName: 'Title', FieldValue: 'Desk' },
        { FieldName: 'RequestNumber', FieldValue: '' }
      ],
      bNewDocumentUpdate: false
    });
    expect(readCreateField(form, 'RequestNumber')).toEqual({ found: true, value: '' });
    expect(readCreateField(form, 'Department')).toEqual({ found: false, value: '' });
    const embedded = embedCreateField(form, 'RequestNumber', 'REQ-202610-0007');
    expect(embedded && JSON.parse(embedded).formValues).toEqual([
      { FieldName: 'Title', FieldValue: 'Desk' },
      { FieldName: 'RequestNumber', FieldValue: 'REQ-202610-0007' }
    ]);
    const direct = embedCreateField(JSON.stringify({
      __metadata: { type: 'SP.Data.RequestsListItem' },
      Title: 'Desk'
    }), 'RequestNumber', 'REQ-2');
    expect(direct && JSON.parse(direct)).toMatchObject({ Title: 'Desk', RequestNumber: 'REQ-2' });
    const wrapped = embedCreateField(JSON.stringify({ d: { Title: 'Desk' } }), 'RequestNumber', 'REQ-3');
    expect(wrapped && JSON.parse(wrapped).d.RequestNumber).toBe('REQ-3');
    expect(embedCreateField('not-json', 'RequestNumber', 'REQ-4')).toBeUndefined();
  });

  it('uses lookup titles from the save and skips a lookup that is only an id', () => {
    const columns = [
      { internalName: 'Department', title: 'Department', type: 'Lookup', token: '{Department.title}' },
      { internalName: 'Title', title: 'Title', type: 'Text', token: '{Title}' }
    ];
    const titled = JSON.stringify({
      formValues: [
        { FieldName: 'Title', FieldValue: 'Desk' },
        { FieldName: 'Department', FieldValue: '12;#Finance' }
      ]
    });
    expect(columnValuesFromCreateBody(titled, ['Title', 'Department'], columns)).toMatchObject({
      complete: true,
      values: { Title: 'Desk', 'Department.title': 'Finance' }
    });
    const idOnly = JSON.stringify({
      formValues: [
        { FieldName: 'Title', FieldValue: 'Desk' },
        { FieldName: 'Department', FieldValue: '12' }
      ]
    });
    expect(columnValuesFromCreateBody(idOnly, ['Department'], columns)?.complete).toBe(false);
    const missing = JSON.stringify({ formValues: [{ FieldName: 'Title', FieldValue: 'Desk' }] });
    expect(columnValuesFromCreateBody(missing, ['Department'], columns)?.complete).toBe(false);
    expect(columnValuesFromCreateBody(missing, [], columns)).toEqual({ values: {}, complete: true });
  });

  it('puts the number inside a SharePoint batch create and updates the part length', () => {
    const json = JSON.stringify({
      formValues: [{ FieldName: 'Title', FieldValue: 'Desk' }],
      bNewDocumentUpdate: false
    });
    const embedded = embedCreateField(json, 'RequestNumber', 'REQ-1');
    expect(embedded).toBeTruthy();
    const body = [
      '--batch_abc123',
      'Content-Type: multipart/mixed; boundary=changeset_1',
      '',
      '--changeset_1',
      'Content-Type: application/http',
      'Content-Transfer-Encoding: binary',
      '',
      `POST https://contoso.sharepoint.com/sites/ops/_api/web/GetList(@a1)/AddValidateUpdateItemUsingPath?@a1='%2Fsites%2Fops%2FLists%2FRequests' HTTP/1.1`,
      'Content-Type: application/json;odata=verbose',
      `Content-Length: ${json.length}`,
      '',
      json,
      '--changeset_1--',
      '--batch_abc123--',
      ''
    ].join('\r\n');
    const next = rewriteNewItemBody(body, listGuid, listPath, (part) => embedCreateField(part, 'RequestNumber', 'REQ-1'));
    expect(next).toContain(embedded || '');
    expect(next).toContain(`Content-Length: ${(embedded || '').length}`);
    expect(newItemJson(body, listGuid, listPath)).toBe(json);
    expect(next && next.indexOf('Desk')).toBeGreaterThan(0);
  });

  it('treats a form error as a failed save', () => {
    expect(createSaveFailed({
      value: [{ FieldName: 'RequestNumber', FieldValue: 'REQ-1', HasException: true, ErrorMessage: 'Invalid field' }]
    })).toBe(true);
    expect(createSaveFailed({ Id: 18 })).toBe(false);
    expect(batchCreateOutcome(JSON.stringify({ Id: 18 }))).toBe('saved');
    expect(batchCreateOutcome([
      '--batchresponse',
      '',
      'HTTP/1.1 200 OK',
      '',
      JSON.stringify({ value: [{ FieldName: 'Id', FieldValue: '0', HasException: true, ErrorMessage: 'Invalid field' }] }),
      '--batchresponse--'
    ].join('\r\n'))).toBe('failed');
  });

  it('writes the number into the create request before SharePoint receives it', async () => {
    const calls: string[] = [];
    const realFetch = window.fetch;
    window.fetch = ((input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
      calls.push(typeof init?.body === 'string' ? init.body : '');
      return Promise.resolve(jsonResponse({
        value: [{ FieldName: 'Id', FieldValue: '18', HasException: false }]
      }, 200));
    }) as typeof window.fetch;
    let created = 0;
    const stop = watchNewListItems(listGuid, listPath, async (body) => {
      const next = embedCreateField(body, 'RequestNumber', 'REQ-202610-0007');
      return next ? { body: next } : undefined;
    }, async () => {
      created += 1;
    });
    try {
      await window.fetch(
        `https://contoso.sharepoint.com/sites/ops/_api/web/lists(guid'${listGuid}')/AddValidateUpdateItemUsingPath`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ formValues: [{ FieldName: 'Title', FieldValue: 'Desk' }] })
        }
      );
      const sent = JSON.parse(calls[0]);
      expect(sent.formValues).toEqual([
        { FieldName: 'Title', FieldValue: 'Desk' },
        { FieldName: 'RequestNumber', FieldValue: 'REQ-202610-0007' }
      ]);
      expect(created).toBe(1);
    } finally {
      stop();
      window.fetch = realFetch;
    }
  });

  it('refreshes after a create whose body cannot be rewritten', async () => {
    const realFetch = window.fetch;
    window.fetch = (() => Promise.resolve(jsonResponse({
      value: [{ FieldName: 'Title', FieldValue: 'Desk', HasException: false, ItemId: 22 }]
    }, 200))) as typeof window.fetch;
    let createdId = 0;
    const stop = watchNewListItems(listGuid, listPath, async () => undefined, async (itemId) => {
      createdId = itemId || 0;
    });
    try {
      await window.fetch(
        `https://contoso.sharepoint.com/sites/ops/_api/web/lists(guid'${listGuid}')/AddValidateUpdateItemUsingPath`,
        { method: 'POST', body: new Blob(['{"formValues":[]}']) }
      );
      expect(createdId).toBe(22);
    } finally {
      stop();
      window.fetch = realFetch;
    }
  });

  it('fills a blank number after the form save returns only ItemId', async () => {
    const realFetch = window.fetch;
    window.fetch = (() => Promise.resolve(jsonResponse({
      value: [{ FieldName: 'Title', FieldValue: 'Desk', HasException: false, ItemId: 21 }]
    }, 200))) as typeof window.fetch;
    let createdId = 0;
    const stop = watchNewListItems(listGuid, listPath, async () => undefined, async (itemId) => {
      createdId = itemId || 0;
    });
    try {
      await window.fetch(
        `https://contoso.sharepoint.com/sites/ops/_api/web/lists(guid'${listGuid}')/AddValidateUpdateItemUsingPath`,
        {
          method: 'POST',
          body: JSON.stringify({ formValues: [{ FieldName: 'Title', FieldValue: 'Desk' }] })
        }
      );
      expect(createdId).toBe(21);
    } finally {
      stop();
      window.fetch = realFetch;
    }
  });

  it('returns the reserved number when the create does not succeed', async () => {
    const realFetch = window.fetch;
    window.fetch = (() => Promise.resolve(jsonResponse({
      value: [{ FieldName: 'RequestNumber', FieldValue: 'REQ-1', HasException: true, ErrorMessage: 'Invalid field' }]
    }, 200))) as typeof window.fetch;
    let rolledBack = 0;
    let created = 0;
    const stop = watchNewListItems(listGuid, listPath, async (body) => {
      const next = embedCreateField(body, 'RequestNumber', 'REQ-1');
      return next ? {
        body: next,
        rollback: async () => {
          rolledBack += 1;
        }
      } : undefined;
    }, async () => {
      created += 1;
    });
    try {
      await window.fetch(
        `https://contoso.sharepoint.com/sites/ops/_api/web/lists(guid'${listGuid}')/items`,
        { method: 'POST', body: JSON.stringify({ Title: 'Desk' }) }
      );
      expect(rolledBack).toBe(1);
      expect(created).toBe(0);
    } finally {
      stop();
      window.fetch = realFetch;
    }
  });

  it('fills the open item form when that save is the only new number', () => {
    document.body.innerHTML = '<div class="ms-Panel-main"><input id="RequestNumber" value="" /></div>';
    paintNumbers([{ itemId: 21, fieldName: 'RequestNumber', code: 'REQ-202610-0009' }]);
    const input = document.querySelector('input');
    expect(input instanceof HTMLInputElement ? input.value : '').toBe('REQ-202610-0009');
    clearAssignedPaints();
    document.body.innerHTML = '';
  });
});
