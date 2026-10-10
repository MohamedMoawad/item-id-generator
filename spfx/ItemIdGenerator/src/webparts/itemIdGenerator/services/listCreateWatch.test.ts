import { clearAssignedPaints, createdItemId, isNewItemRequest, paintNumbers, revealAssignedNumbers } from './listCreateWatch';

const listGuid = 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee';
const listPath = '/sites/ops/Lists/Requests';

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
    rects.mockRestore();
    jest.useRealTimers();
    document.body.innerHTML = '';
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
