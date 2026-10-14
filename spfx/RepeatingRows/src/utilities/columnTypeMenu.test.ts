import { addRepeatingColumnType, REPEATING_TYPE_LABEL } from './columnTypeMenu';

function columnMenu(): HTMLDivElement {
  const menu = document.createElement('div');
  menu.setAttribute('role', 'menu');
  menu.innerHTML = [
    '<button role="menuitem" type="button"><span>Single line of text</span></button>',
    '<button role="menuitem" type="button"><span>Choice</span></button>',
    '<button role="menuitem" type="button"><span>Number</span></button>'
  ].join('');
  document.body.appendChild(menu);
  return menu;
}

describe('column type menu', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('adds Repeating Rows to the create-column type list', () => {
    const menu = columnMenu();
    let chosen = 0;
    expect(addRepeatingColumnType(document, () => { chosen += 1; })).toBe(1);
    const added = menu.querySelector('[data-repeating-rows-type="true"]') as HTMLButtonElement;
    expect(added.textContent).toContain(REPEATING_TYPE_LABEL);
    added.click();
    expect(chosen).toBe(1);
    expect(addRepeatingColumnType(document, () => undefined)).toBe(0);
  });

  it('leaves an unrelated menu unchanged', () => {
    const menu = document.createElement('div');
    menu.setAttribute('role', 'menu');
    menu.innerHTML = '<button role="menuitem" type="button"><span>Rename</span></button>';
    document.body.appendChild(menu);
    expect(addRepeatingColumnType(document, () => undefined)).toBe(0);
    expect(menu.textContent).toBe('Rename');
  });
});