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

  it('ignores labels that only start with a column type name', () => {
    const menu = document.createElement('div');
    menu.setAttribute('role', 'menu');
    menu.innerHTML = [
      '<button role="menuitem" type="button"><span>Number column</span></button>',
      '<button role="menuitem" type="button"><span>Choice field</span></button>',
      '<button role="menuitem" type="button"><span>Lookup with cascade delete</span></button>'
    ].join('');
    document.body.appendChild(menu);
    expect(addRepeatingColumnType(document, () => undefined)).toBe(0);
    expect(menu.querySelector('[data-repeating-rows-type]')).toBeNull();
  });

  it('requires three exact column types', () => {
    const menu = document.createElement('div');
    menu.setAttribute('role', 'menu');
    menu.innerHTML = [
      '<button role="menuitem" type="button"><span>Choice</span></button>',
      '<button role="menuitem" type="button"><span>Number</span></button>'
    ].join('');
    document.body.appendChild(menu);
    expect(addRepeatingColumnType(document, () => undefined)).toBe(0);
  });

  it('does not add an option to a select', () => {
    const select = document.createElement('select');
    select.innerHTML = [
      '<option>Single line of text</option>',
      '<option>Choice</option>',
      '<option>Number</option>',
      '<option>Date and time</option>'
    ].join('');
    document.body.appendChild(select);
    let chosen = 0;
    expect(addRepeatingColumnType(document, () => { chosen += 1; })).toBe(0);
    expect(select.querySelector('[data-repeating-rows-type]')).toBeNull();
    select.value = select.options[0].value;
    select.dispatchEvent(new Event('change'));
    expect(chosen).toBe(0);
  });

  it('does not inject into dialogs, dropdowns, the command bar, or the setup host', () => {
    const hosts: Array<{ role?: string; className?: string; automation?: string; id?: string }> = [
      { role: 'dialog' },
      { className: 'ms-Dialog' },
      { className: 'ms-Dropdown-callout' },
      { className: 'ms-CommandBar' },
      { automation: 'CommandBar' },
      { automation: 'ListCommandBar' },
      { id: 'repeating-rows-setup-host' }
    ];
    hosts.forEach((host): void => {
      const wrapper = document.createElement('div');
      if (host.role) {
        wrapper.setAttribute('role', host.role);
      }
      if (host.className) {
        wrapper.className = host.className;
      }
      if (host.automation) {
        wrapper.setAttribute('data-automationid', host.automation);
      }
      if (host.id) {
        wrapper.id = host.id;
      }
      const menu = document.createElement('div');
      menu.setAttribute('role', 'menu');
      menu.innerHTML = [
        '<button role="menuitem" type="button"><span>Single line of text</span></button>',
        '<button role="menuitem" type="button"><span>Choice</span></button>',
        '<button role="menuitem" type="button"><span>Number</span></button>'
      ].join('');
      wrapper.appendChild(menu);
      document.body.appendChild(wrapper);
    });
    expect(addRepeatingColumnType(document, () => undefined)).toBe(0);
  });
});
