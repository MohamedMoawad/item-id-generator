import { FormDisplayMode } from '@microsoft/sp-core-library';
import { BaseFormCustomizer } from '@microsoft/sp-listview-extensibility';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { IRepeatingColumn, IRepeatingRowsConfig } from '../models/repeatingRows';
import { HttpChildRepository } from '../repositories/httpChildRepository';
import { HttpParentRepository } from '../repositories/httpParentRepository';
import { IRepeatingStore } from '../repositories/contracts';
import { pnpDeleteRows } from '../services/pnpFactory';
import { ensurePerson, searchPeople } from '../services/peopleSearch';
import { saveRepeatingItem } from '../services/repeatingRowsService';
import { SpHttpGateway } from '../services/sharePointGateway';
import { parseRepeatingConfig } from '../utilities/configParser';
import { planChildChanges } from '../utilities/savePlan';
import { parentPayload } from '../utilities/parentValues';
import { toUserMessage } from '../utilities/errors';
import { FormHost } from './FormHost';
import { ILoadedForm, loadFormData } from './loadFormData';

export interface IRepeatingRowsFormProperties {
  childListTitle?: string;
  relationshipField?: string;
  relationshipKind?: 'lookup' | 'number';
  columns?: IRepeatingColumn[];
  minRows?: number;
  maxRows?: number;
  allowDelete?: boolean;
  allowEdit?: boolean;
}

export default class RepeatingRowsFormCustomizer extends BaseFormCustomizer<IRepeatingRowsFormProperties> {
  private _data: ILoadedForm | undefined;
  private _error = '';

  public onInit(): Promise<void> {
    return loadFormData(this._gateway(), this._listId(), this._itemId(), this._config()).then((data) => {
      this._data = data;
    }).catch((error: unknown) => {
      this._error = toUserMessage(error);
    });
  }

  public render(): void {
    const data = this._data;
    const element: React.ReactElement = data
      ? React.createElement(FormHost, {
        displayMode: this.displayMode,
        config: this._config(),
        fields: data.fields,
        initialValues: data.values,
        initialRows: data.rows,
        etag: data.etag,
        parentId: this._itemId(),
        lookups: data.lookups,
        hasMore: data.hasMore,
        onLoadMore: (afterSortOrder: number) => this._loadMore(afterSortOrder),
        onSearchPeople: (text: string) => searchPeople(this._gateway(), text),
        onResolvePerson: (person) => this._resolvePerson(person),
        onSave: (values, rows, attachments, etag) => this._save(values, rows, attachments, etag),
        onClose: () => this.formClosed()
      })
      : React.createElement('p', undefined, this._error || 'Loading line items');
    ReactDOM.render(element, this.domElement);
  }

  public onDispose(): void {
    ReactDOM.unmountComponentAtNode(this.domElement);
  }

  private _save(
    values: { [key: string]: string | number | boolean },
    rows: ILoadedForm['rows'],
    attachments: { rowKey: string; fileName: string; content: ArrayBuffer }[],
    etag?: string
  ): Promise<void> {
    const data = this._data;
    if (!data) {
      return Promise.resolve();
    }
    const config = this._config();
    const gateway = this._gateway();
    const parentId = this._itemId();
    const stamped = rows.map((row) => ({ ...row, parentId: parentId || row.parentId }));
    const removed = planChildChanges(data.rows, stamped).deleteIds;
    const kept = data.rows.filter((row) => row.id === undefined || removed.indexOf(row.id) < 0);
    const store: IRepeatingStore = {
      config,
      parent: new HttpParentRepository(gateway, this._listId(), [], []),
      children: new HttpChildRepository(gateway, config)
    };
    return saveRepeatingItem(store, {
      parentId,
      etag,
      parentValues: parentPayload(data.fields, values),
      rows,
      originalRows: kept,
      attachments
    }).then(() => pnpDeleteRows(this.context, config, removed)).then(() => {
      this.formSaved();
    });
  }

  private _loadMore(afterSortOrder: number): Promise<{ rows: ILoadedForm['rows']; hasMore: boolean }> {
    const itemId = this._itemId();
    if (!itemId) {
      return Promise.resolve({ rows: [], hasMore: false });
    }
    return new HttpChildRepository(this._gateway(), this._config()).loadPage(itemId, afterSortOrder);
  }

  private async _resolvePerson(person: { id: string; login: string; text: string }): Promise<string> {
    if (person.id) {
      return `${person.id}|${person.text}`;
    }
    const id = await ensurePerson(this._gateway(), person.login);
    return `${id}|${person.text}`;
  }

  private _config(): IRepeatingRowsConfig {
    const list = this.context.pageContext.list;
    const fallback = list ? `${list.title} Details` : 'Repeating Row Details';
    return parseRepeatingConfig(this.properties, fallback);
  }

  private _gateway(): SpHttpGateway {
    return new SpHttpGateway(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl);
  }

  private _listId(): string {
    const list = this.context.pageContext.list;
    if (!list) {
      return '';
    }
    return list.id.toString();
  }

  private _itemId(): number | undefined {
    const itemId = this.context.itemId;
    return itemId && itemId > 0 ? itemId : undefined;
  }
}

export function formModeLabel(mode: FormDisplayMode): string {
  if (mode === FormDisplayMode.New) {
    return 'New';
  }
  if (mode === FormDisplayMode.Edit) {
    return 'Edit';
  }
  return 'Display';
}
