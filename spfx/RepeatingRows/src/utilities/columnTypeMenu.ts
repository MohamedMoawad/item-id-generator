export const REPEATING_TYPE_LABEL = 'Repeating Rows';
const MARKER = 'data-repeating-rows-type';

const KNOWN_TYPES = [
  'single line of text',
  'multiple lines of text',
  'choice',
  'number',
  'yes/no',
  'person',
  'person or group',
  'date and time',
  'lookup',
  'currency',
  'hyperlink'
];

const IGNORE_ANCESTOR = [
  '#repeating-rows-setup-host',
  '[role="dialog"]',
  '.ms-Dialog',
  '.ms-Dropdown-callout',
  '.ms-CommandBar',
  '[data-automationid="CommandBar"]',
  '[data-automationid="ListCommandBar"]'
].join(', ');

export function addRepeatingColumnType(root: ParentNode, onChoose: () => void): number {
  let added = 0;
  const groups = root.querySelectorAll('[role="menu"], [role="listbox"], [role="radiogroup"]');
  for (let index = 0; index < groups.length; index += 1) {
    const group = groups[index];
    if (group.querySelector(`[${MARKER}]`) || isIgnored(group) || !isColumnTypeList(group)) {
      continue;
    }
    const sample = group.querySelector('[role="menuitem"], [role="option"], [role="radio"]');
    if (!sample) {
      continue;
    }
    group.appendChild(cloneTypeChoice(sample, onChoose));
    added += 1;
  }
  return added;
}

export function isColumnTypeList(root: ParentNode): boolean {
  const labels = readLabels(root);
  let hits = 0;
  KNOWN_TYPES.forEach((name) => {
    if (labels.indexOf(name) !== -1) {
      hits += 1;
    }
  });
  return hits >= 3;
}

function isIgnored(element: Element): boolean {
  return !!element.closest(IGNORE_ANCESTOR);
}

function cloneTypeChoice(sample: Element, onChoose: () => void): HTMLElement {
  const copy = sample.cloneNode(true) as HTMLElement;
  copy.setAttribute(MARKER, 'true');
  copy.removeAttribute('id');
  copy.removeAttribute('onclick');
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
    if (KNOWN_TYPES.indexOf(text) !== -1) {
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
  const nodes = root.querySelectorAll('[role="menuitem"], [role="option"], [role="radio"]');
  const labels: string[] = [];
  for (let index = 0; index < nodes.length; index += 1) {
    labels.push(normalize(nodes[index].textContent || ''));
  }
  return labels;
}

function normalize(value: string): string {
  return value.replace(/\s+/g, ' ').trim().toLowerCase();
}
