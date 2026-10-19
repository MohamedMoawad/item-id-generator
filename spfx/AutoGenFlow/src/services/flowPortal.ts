const SDK_URLS = [
  'https://flow.microsoft.com/Content/msflowsdk-1.1.js',
  'https://content.powerapps.com/resource/makerx/static/js/msflowsdk-1.1.1e587b1a.js'
];

interface IFlowEvent {
  flowId?: string;
  flowUrl?: string;
  message?: string;
  error?: string | { message?: string };
}

interface IFlowWidgetDone {
  (error: undefined, result: { token: string } | { [key: string]: string }): void;
}

interface IFlowWidget {
  listen(eventName: string, handler: (data: IFlowEvent | undefined, done?: IFlowWidgetDone) => void): void;
  notify(action: string, payload: { templateDefinition: string }): Promise<unknown>;
}

interface IMsFlowSdk {
  renderWidget(kind: string, options: {
    container: HTMLElement;
    environmentId: string;
    templatesSettings: { useServerSideProvisioning: boolean };
  }): IFlowWidget;
}

interface ISdkWindow {
  MsFlowSdk?: new (options: { hostName: string; locale: string }) => IMsFlowSdk;
}

export async function createFlowThroughPortal(token: string, environmentId: string, templateDefinition: string): Promise<string> {
  await loadFlowSdk();
  const sdkCtor = (window as unknown as ISdkWindow).MsFlowSdk;
  if (!sdkCtor) {
    throw new Error('Power Automate did not load the numbering flow creator.');
  }
  const portal = ensureHost();
  const sdk = new sdkCtor({
    hostName: 'https://make.powerautomate.com',
    locale: 'en-US'
  });
  let widget: IFlowWidget;
  try {
    widget = sdk.renderWidget('flowCreation', {
      container: portal.container,
      environmentId,
      templatesSettings: { useServerSideProvisioning: true }
    });
  } catch (error) {
    portal.host.remove();
    throw error;
  }
  return new Promise<string>((resolve, reject) => {
    let settled = false;
    let timer = 0;
    const finish = (error: Error | undefined, flowId: string): void => {
      if (settled) {
        return;
      }
      settled = true;
      window.clearTimeout(timer);
      portal.host.remove();
      if (error || !flowId) {
        reject(error || new Error('Power Automate did not create the numbering flow.'));
        return;
      }
      resolve(flowId);
    };
    timer = window.setTimeout(() => {
      finish(new Error('Power Automate did not finish creating the numbering flow.'), '');
    }, 180000);
    const close = portal.host.querySelector('button');
    if (close) {
      close.addEventListener('click', () => {
        finish(new Error('The numbering flow was not created.'), '');
      });
    }
    widget.listen('GET_ACCESS_TOKEN', (_data, done) => {
      if (done) {
        done(undefined, { token });
      }
    });
    widget.listen('GET_STRINGS', (_data, done) => {
      if (done) {
        done(undefined, {
          FLOW_CREATION_HEADER: 'AutoGen numbering flow',
          FLOW_CREATION_CREATE_BUTTON: 'Create numbering flow'
        });
      }
    });
    widget.listen('FLOW_CREATION_SUCCEEDED', (data) => {
      finish(undefined, flowIdFromEvent(data));
    });
    widget.listen('FLOW_CREATION_FAILED', (data) => {
      finish(new Error(eventMessage(data) || 'Power Automate did not create the numbering flow.'), '');
    });
    widget.listen('TEMPLATE_LOAD_FAILED', (data) => {
      finish(new Error(eventMessage(data) || 'Power Automate did not create the numbering flow.'), '');
    });
    widget.listen('WIDGET_READY', () => {
      widget.notify('createFlowFromTemplateDefinition', { templateDefinition }).catch((error: unknown) => {
        finish(error instanceof Error ? error : new Error('Power Automate did not create the numbering flow.'), '');
      });
    });
  });
}

function sdkReady(): boolean {
  return !!(window as unknown as ISdkWindow).MsFlowSdk;
}

function loadFlowSdk(): Promise<void> {
  if (sdkReady()) {
    return Promise.resolve();
  }
  installBundledSdk();
  if (sdkReady()) {
    return Promise.resolve();
  }
  return loadScript(0);
}

function installBundledSdk(): void {
  if (sdkReady()) {
    return;
  }
  try {
    // Resolved from lib/services after the TypeScript emit.
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    require('../../thirdparty/msflowsdk.js');
  } catch {
    return;
  }
}

function loadScript(index: number): Promise<void> {
  if (sdkReady()) {
    return Promise.resolve();
  }
  if (index >= SDK_URLS.length) {
    return Promise.reject(new Error('Power Automate did not load the numbering flow creator.'));
  }
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = SDK_URLS[index];
    script.async = true;
    const next = (): void => {
      script.remove();
      loadScript(index + 1).then(resolve).catch(reject);
    };
    script.onload = (): void => {
      if (sdkReady()) {
        resolve();
        return;
      }
      next();
    };
    script.onerror = (): void => {
      next();
    };
    document.head.appendChild(script);
  });
}

function ensureHost(): { host: HTMLElement; container: HTMLElement } {
  const previous = document.getElementById('autogen-flow-portal-host');
  if (previous) {
    previous.remove();
  }
  const host = document.createElement('section');
  host.id = 'autogen-flow-portal-host';
  host.setAttribute('aria-label', 'Create the numbering flow');
  host.style.position = 'fixed';
  host.style.right = '16px';
  host.style.bottom = '16px';
  host.style.width = 'min(520px, calc(100vw - 24px))';
  host.style.height = 'min(640px, calc(100vh - 24px))';
  host.style.zIndex = '2000000';
  host.style.background = '#ffffff';
  host.style.border = '1px solid #e1dfdd';
  host.style.boxShadow = '0 8px 28px rgba(0, 0, 0, 0.22)';
  host.style.display = 'flex';
  host.style.flexDirection = 'column';
  const bar = document.createElement('div');
  bar.style.display = 'flex';
  bar.style.alignItems = 'center';
  bar.style.justifyContent = 'space-between';
  bar.style.padding = '8px 12px';
  bar.style.borderBottom = '1px solid #edebe9';
  const title = document.createElement('strong');
  title.textContent = 'Creating the numbering flow';
  const close = document.createElement('button');
  close.type = 'button';
  close.textContent = 'Close';
  bar.appendChild(title);
  bar.appendChild(close);
  const frame = document.createElement('div');
  frame.id = 'autogen-flow-portal';
  frame.style.flex = '1';
  frame.style.minHeight = '480px';
  host.appendChild(bar);
  host.appendChild(frame);
  document.body.appendChild(host);
  return { host, container: frame };
}

function flowIdFromEvent(data: IFlowEvent | undefined): string {
  if (!data) {
    return '';
  }
  if (data.flowId) {
    return data.flowId;
  }
  const match = /\/flows\/([0-9a-f-]{36})/i.exec(data.flowUrl || '');
  return match ? match[1] : '';
}

function eventMessage(data: IFlowEvent | undefined): string {
  if (!data) {
    return '';
  }
  if (typeof data.error === 'string') {
    return data.error;
  }
  if (data.error && data.error.message) {
    return data.error.message;
  }
  return data.message || '';
}
