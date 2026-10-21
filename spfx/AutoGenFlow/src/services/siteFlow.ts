import { normalizeGuid } from './contract';

export interface ISiteFlowDefinition {
  $schema: string;
  contentVersion: string;
  parameters: { [name: string]: { defaultValue: object; type: string } };
  triggers: { [name: string]: object };
  actions: { [name: string]: object };
}

export interface IListFlowTarget {
  webUrl: string;
  siteUrl: string;
  listGuid: string;
  listTitle: string;
  configListTitle: string;
}

export interface ICreateFlowPayload {
  properties: {
    displayName: string;
    definition: ISiteFlowDefinition;
    connectionReferences: {
      shared_sharepointonline?: {
        connectionName: string;
        source: string;
        id: string;
        tier: string;
      };
    };
    state: 'Started' | 'Stopped';
    flowFailureAlertSubscribed: boolean;
  };
}

export function flowDisplayName(listTitle: string): string {
  const leaf = listTitle.trim() || 'list';
  return `AutoGen Flow · ${leaf}`.slice(0, 80);
}

export function buildCreateFlowPayload(target: IListFlowTarget, connectionName: string, state: 'Started' | 'Stopped' = 'Stopped'): ICreateFlowPayload {
  const connectionReferences: ICreateFlowPayload['properties']['connectionReferences'] = {};
  if (connectionName.trim()) {
    connectionReferences.shared_sharepointonline = {
      connectionName,
      source: 'Embedded',
      id: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline',
      tier: 'NotSpecified'
    };
  }
  return {
    properties: {
      displayName: flowDisplayName(target.listTitle),
      definition: buildSiteFlowDefinition(target),
      connectionReferences,
      state,
      flowFailureAlertSubscribed: false
    }
  };
}

export function buildPortalTemplate(target: IListFlowTarget, connectionName: string): string {
  const flow = buildCreateFlowPayload(target, connectionName, 'Started');
  return JSON.stringify({
    displayName: flow.properties.displayName,
    description: 'AutoGen numbering flow for this site collection.',
    properties: {
      definition: flow.properties.definition,
      connectionReferences: flow.properties.connectionReferences,
      flowFailureAlertSubscribed: false
    }
  });
}

export function buildDataverseWorkflowBody(target: IListFlowTarget, connectionName: string): {
  category: 5;
  name: string;
  type: 1;
  description: string;
  primaryentity: 'none';
  clientdata: string;
} {
  const flow = buildCreateFlowPayload(target, connectionName);
  const connection = flow.properties.connectionReferences.shared_sharepointonline;
  const connectionReferences = connection ? {
    shared_sharepointonline: {
      runtimeSource: 'embedded',
      connection: { name: connection.connectionName },
      api: { name: 'shared_sharepointonline' }
    }
  } : {};
  return {
    category: 5,
    name: flow.properties.displayName,
    type: 1,
    description: 'AutoGen numbering flow for this site collection.',
    primaryentity: 'none',
    clientdata: JSON.stringify({
      properties: {
        connectionReferences,
        definition: flow.properties.definition
      },
      schemaVersion: '1.0.0.0'
    })
  };
}

export function buildShellFlowPayload(listTitle: string): ICreateFlowPayload {
  return {
    properties: {
      displayName: flowDisplayName(listTitle),
      definition: {
        $schema: 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#',
        contentVersion: '1.0.0.0',
        parameters: {
          $connections: { defaultValue: {}, type: 'Object' },
          $authentication: { defaultValue: {}, type: 'SecureObject' }
        },
        triggers: {
          Recurrence: {
            type: 'Recurrence',
            recurrence: { frequency: 'Day', interval: 1 }
          }
        },
        actions: {
          Compose: {
            type: 'Compose',
            inputs: 'AutoGen',
            runAfter: {}
          }
        }
      },
      connectionReferences: {},
      state: 'Stopped',
      flowFailureAlertSubscribed: false
    }
  };
}

export function buildSiteFlowDefinition(target: IListFlowTarget): ISiteFlowDefinition {
  const site = target.siteUrl.replace(/\/+$/, '');
  const web = target.webUrl.replace(/\/+$/, '');
  const guid = normalizeGuid(target.listGuid);
  const bakedTitle = target.configListTitle.replace(/'/g, "''''");
  return {
    $schema: 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#',
    contentVersion: '1.0.0.0',
    parameters: {
      $connections: { defaultValue: {}, type: 'Object' },
      $authentication: { defaultValue: {}, type: 'SecureObject' }
    },
    triggers: {
      When_an_item_is_created: {
        type: 'OpenApiConnection',
        inputs: {
          host: {
            connectionName: 'shared_sharepointonline',
            operationId: 'GetOnNewItems',
            apiId: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline'
          },
          parameters: {
            dataset: web,
            table: guid
          },
          authentication: "@parameters('$connections')['shared_sharepointonline']['connectionId']"
        },
        recurrence: { frequency: 'Minute', interval: 1 },
        splitOn: "@triggerBody()?['value']",
        runtimeConfiguration: { concurrency: { runs: 1 } }
      }
    },
    actions: {
      Initialize_next: variable('Next', 'integer', 0),
      Initialize_code: variable('Code', 'string', ''),
      Initialize_previous: variable('Previous', 'integer', 0),
      Get_rule: http(site, 'GET', oneRuleUri(target.configListTitle, guid), undefined, {
        Initialize_next: ['Succeeded'],
        Initialize_code: ['Succeeded'],
        Initialize_previous: ['Succeeded']
      }),
      Rule_row: {
        type: 'Compose',
        inputs: "@first(coalesce(body('Get_rule')?['value'], body('Get_rule')?['d']?['results'], json('[]')))",
        runAfter: { Get_rule: ['Succeeded'] }
      },
      Has_rule: {
        type: 'If',
        expression: {
          not: {
            equals: ["@coalesce(outputs('Rule_row')?['Id'], outputs('Rule_row')?['ID'], '')", '']
          }
        },
        runAfter: { Rule_row: ['Succeeded'] },
        else: { actions: {} },
        actions: createdActions(site, web, bakedTitle)
      }
    }
  };
}

function oneRuleUri(title: string, guid: string): string {
  const select = [
    'Id',
    'TargetListUrl',
    'TargetListGuid',
    'NumberColumnInternalName',
    'Formula',
    'FormulaColumns',
    'CounterToken',
    'CounterWidth',
    'CurrentCount',
    'ResetPeriod',
    'LastResetDate',
    'ItemQuery',
    'IsActive'
  ].join(',');
  const listTitle = title.replace(/'/g, "''");
  const filter = encodeURIComponent(`IsActive eq 1 and TargetListGuid eq '${guid}'`);
  return `_api/web/lists/getbytitle('${listTitle}')/items?$select=${select}&$filter=${filter}&$top=1`;
}

function createdActions(site: string, web: string, bakedTitle: string): { [name: string]: object } {
  return {
    Config: {
      type: 'Compose',
      inputs: "@outputs('Rule_row')"
    },
    Column_name: {
      type: 'Compose',
      inputs: "@outputs('Config')?['NumberColumnInternalName']",
      runAfter: { Config: ['Succeeded'] }
    },
    Web: {
      type: 'Compose',
      inputs: web,
      runAfter: { Column_name: ['Succeeded'] }
    },
    Seed_rolled: {
      type: 'Compose',
      inputs: `@${rolledExpression()}`,
      runAfter: { Web: ['Succeeded'] }
    },
    Set_previous: setVariable('Previous', "@if(outputs('Seed_rolled'), 0, int(coalesce(outputs('Config')?['CurrentCount'], 0)))", {
      Seed_rolled: ['Succeeded']
    }),
    ...itemActions(site, bakedTitle, { Set_previous: ['Succeeded'] })
  };
}

function itemActions(
  site: string,
  bakedTitle: string,
  itemRunAfter: { [action: string]: string[] }
): { [name: string]: object } {
  const triggerId = "coalesce(triggerBody()?['ID'], triggerBody()?['Id'])";
  const itemId = "coalesce(outputs('Item')?['Id'], outputs('Item')?['ID'], triggerBody()?['ID'], triggerBody()?['Id'])";
  const configId = "coalesce(outputs('Config')?['Id'], outputs('Config')?['ID'])";
  const configUri = `@concat('_api/web/lists/getbytitle(''${bakedTitle}'')/items(', string(${configId}), ')')`;
  const listGuid = "replace(replace(string(coalesce(outputs('Config')?['TargetListGuid'], '')), '{', ''), '}', '')";
  const itemFieldsUri = `@concat('_api/web/lists(guid''', ${listGuid}, ''')/items(', string(${triggerId}), ')?', coalesce(outputs('Config')?['ItemQuery'], '$select=Id,Created'))`;
  const itemUri = `@concat('_api/web/lists(guid''', ${listGuid}, ''')/items(', string(${itemId}), ')')`;
  const stamp = "utcNow('yyyy-MM-ddTHH:mm:ssZ')";
  return {
    Get_item: http('@outputs(\'Web\')', 'GET', itemFieldsUri, undefined, itemRunAfter),
    Item: {
      type: 'Compose',
      inputs: "@coalesce(body('Get_item'), triggerBody())",
      runAfter: { Get_item: ['Succeeded', 'Failed', 'TimedOut'] }
    },
    Current_value: {
      type: 'Compose',
      inputs: "@outputs('Item')?[outputs('Column_name')]",
      runAfter: { Item: ['Succeeded'] }
    },
    Is_blank: {
      type: 'If',
      expression: {
        and: [{
          equals: ["@empty(outputs('Current_value'))", true]
        }]
      },
      runAfter: { Current_value: ['Succeeded'] },
      else: { actions: {} },
      actions: {
        Set_next: setVariable('Next', "@add(variables('Previous'), 1)", {}),
        Set_code: setVariable('Code', codeExpression(), { Set_next: ['Succeeded'] }),
        For_each_column: {
          type: 'Foreach',
          foreach: "@split(if(equals(coalesce(outputs('Config')?['FormulaColumns'], ''), ''), '__none__', outputs('Config')?['FormulaColumns']), ',')",
          runtimeConfiguration: { concurrency: { repetitions: 1 } },
          runAfter: { Set_code: ['Succeeded'] },
          actions: {
            Apply_column: setVariable('Code', columnExpression(), {})
          }
        },
        Reserve_count: http(
          site,
          'POST',
          configUri,
          `@concat('{"CurrentCount":', string(variables('Next')), ',"LastResetDate":"', ${stamp}, '"}')`,
          { For_each_column: ['Succeeded'] }
        ),
        Write_item: http(
          '@outputs(\'Web\')',
          'POST',
          itemUri,
          "@concat('{\"', outputs('Column_name'), '\":\"', replace(variables('Code'), '\"', ''), '\"}')",
          { Reserve_count: ['Succeeded'] }
        ),
        Remember_count: setVariable('Previous', "@variables('Next')", { Write_item: ['Succeeded'] }),
        Restore_count: http(
          site,
          'POST',
          configUri,
          "@concat('{\"CurrentCount\":', string(variables('Previous')), '}')",
          { Write_item: ['Failed', 'TimedOut'] }
        )
      }
    }
  };
}

function rolledExpression(): string {
  const period = "outputs('Config')?['ResetPeriod']";
  const last = "outputs('Config')?['LastResetDate']";
  const key = (source: string): string => {
    return `if(equals(${period}, 'Day'), formatDateTime(${source}, 'yyyy-MM-dd'), if(equals(${period}, 'Month'), formatDateTime(${source}, 'yyyy-MM'), if(equals(${period}, 'Year'), formatDateTime(${source}, 'yyyy'), '')))`;
  };
  return `if(or(equals(${period}, 'None'), empty(${last})), false, not(equals(${key(last)}, ${key('utcNow()')})))`;
}

function codeExpression(): string {
  const width = "int(coalesce(outputs('Config')?['CounterWidth'], 0))";
  const padded = `if(or(equals(${width}, 0), greater(length(string(variables('Next'))), ${width})), string(variables('Next')), substring(concat('000000000000', string(variables('Next'))), sub(12, sub(${width}, length(string(variables('Next'))))), ${width}))`;
  const created = "coalesce(outputs('Item')?['Created'], triggerOutputs()?['body/Created'], utcNow())";
  return `@replace(replace(replace(replace(replace(replace(replace(outputs('Config')?['Formula'], outputs('Config')?['CounterToken'], ${padded}), '{yyyy}', formatDateTime(${created}, 'yyyy')), '{yy}', formatDateTime(${created}, 'yy')), '{MM}', formatDateTime(${created}, 'MM')), '{dd}', formatDateTime(${created}, 'dd')), '{HH}', formatDateTime(${created}, 'HH')), '{mm}', formatDateTime(${created}, 'mm'))`;
}

function columnExpression(): string {
  const field = "outputs('Item')?[first(split(item(), '.'))]";
  const title = `coalesce(${field}?['Title'], ${field}?['LookupValue'], ${field}?['EMail'], ${field}?['Email'], ${field}?['Value'], string(${field}))`;
  return `@if(equals(item(), '__none__'), variables('Code'), replace(variables('Code'), concat('{', item(), '}'), ${title}))`;
}

function variable(name: string, type: string, value: number | string): object {
  return {
    type: 'InitializeVariable',
    inputs: { variables: [{ name, type, value }] }
  };
}

function setVariable(name: string, value: string, runAfter: { [action: string]: string[] }): object {
  const action: { type: string; inputs: { name: string; value: string }; runAfter?: { [actionName: string]: string[] } } = {
    type: 'SetVariable',
    inputs: { name, value }
  };
  if (Object.keys(runAfter).length > 0) {
    action.runAfter = runAfter;
  }
  return action;
}

function http(
  site: string,
  method: string,
  uri: string,
  body: string | undefined,
  runAfter: { [action: string]: string[] }
): object {
  const headers: { [name: string]: string } = method === 'GET'
    ? { Accept: 'application/json;odata=nometadata' }
    : {
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE'
    };
  const parameters: { [name: string]: unknown } = {
    dataset: site,
    'parameters/method': method,
    'parameters/uri': uri,
    'parameters/headers': headers
  };
  if (body !== undefined) {
    parameters['parameters/body'] = body;
  }
  const action: { type: string; inputs: object; runAfter?: { [actionName: string]: string[] } } = {
    type: 'OpenApiConnection',
    inputs: {
      host: {
        connectionName: 'shared_sharepointonline',
        operationId: 'HttpRequest',
        apiId: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline'
      },
      parameters,
      authentication: "@parameters('$connections')['shared_sharepointonline']['connectionId']"
    }
  };
  if (Object.keys(runAfter).length > 0) {
    action.runAfter = runAfter;
  }
  return action;
}
