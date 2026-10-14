export const REPEATING_TYPE_LABEL = 'Repeating Rows';
const MARKER = 'data-repeating-rows-type';

const KNOWN_TYPES = [
  'single line of text',
  'multiple lines of text',
  'choice',
  'number',
  'yes/no',
  'person',
  'date and time',
  'lookup',
  'currency',
  'hyperlink'
];

export function addRepeatingColumnType(root: ParentNode, onChoose: () => void): number {
  let added = 0;
  const groups = root.querySelectorAll('[role="menu"], [role="listbox"], [role="radiogroup"]');
  for (let index = 0; index < groups.length; index += 1) {
    const group = groups[index];
    if (group.querySelector(`[${MARKER}]`) || !isColumnTypeList(group)) {
      continue;
    }
    const sample = group.querySelector('[role="menuitem"], [role="option"], [role="radio"]');
    if (!sample) {
      continue;
    }
    group.appendChild(cloneTypeChoice(sample, onChoose));
    added += 1;
  }
  const selects = root.querySelectorAll('select');
  for (let index = 0; index < selects.length; index += 1) {
    if (addSelectOption(selects[index], onChoose)) {
      added += 1;
    }
  }
  return added;
}

export function isColumnTypeList(root: ParentNode): boolean {
  const labels = readLabels(root);
  let hits = 0;
  KNOWN_TYPES.forEach((name) => {
    if (labels.some((label) => label === name || label.indexOf(name) === 0)) {
      hits += 1;
    }
  });
  return hits >= 2;
}

function addSelectOption(select: HTMLSelectElement, onChoose: () => void): boolean {
  if (select.querySelector(`[${MARKER}]`) || !isColumnTypeList(select)) {
    return false;
  }
  const option = select.ownerDocument.createElement('option');
  option.value = 'RepeatingRows';
  option.text = REPEATING_TYPE_LABEL;
  option.setAttribute(MARKER, 'true');
  select.appendChild(option);
  select.addEventListener('change', () => {
    if (select.value !== 'RepeatingRows') {
      return;
    }
    select.selectedIndex = 0;
    onChoose();
  });
  return true;
}

function cloneTypeChoice(sample: Element, onChoose: () => void): HTMLElement {
  const copy = sample.cloneNode(true) as HTMLElement;
  copy.setAttribute(MARKER, 'true');
  copy.removeAttribute('id');
  const identified = copy.querySelectorAll('[id]');
  for (let index = 0; index < identified.length; index += 1) {
    identified[index].removeAttribute('id');
  }
  replaceLabel(copy);
  copy.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    onChoose();
  });
  return copy;
}

function replaceLabel(root: HTMLElement): void {
  const spans = root.querySelectorAll('span');
  let replaced = false;
  for (let index = 0; index < spans.length; index += 1) {
    const text = normalize(spans[index].textContent || '');
    if (KNOWN_TYPES.some((name) => text === name || text.indexOf(name) === 0)) {
      spans[index].textContent = REPEATING_TYPE_LABEL;
      replaced = true;
    }
  }
  if (!replaced) {
    root.textContent = REPEATING_TYPE_LABEL;
  }
  root.setAttribute('aria-label', REPEATING_TYPE_LABEL);
  root.setAttribute('title', REPEATING_TYPE_LABEL);
  root.setAttribute('name', REPEATING_TYPE_LABEL);
}

function readLabels(root: ParentNode): string[] {
  const nodes = root.querySelectorAll('[role="menuitem"], [role="option"], [role="radio"], option');
  const labels: string[] = [];
  for (let index = 0; index < nodes.length; index += 1) {
    labels.push(normalize(nodes[index].textContent || ''));
  }
  return labels;
}

function normalize(value: string): string {
  return value.replace(/\s+/g, ' ').trim().toLowerCase();
}
