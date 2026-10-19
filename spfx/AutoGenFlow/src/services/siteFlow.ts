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

export function flowDisplayName(siteUrl: string): string {
  const trimmed = siteUrl.trim().replace(/\/+$/, '');
  const parts = trimmed.split('/');
  const leaf = parts[parts.length - 1] || 'site';
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
      displayName: flowDisplayName(target.siteUrl),
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
  const bakedTitle = target.configListTitle.replace(/'/g, "''''");
  return {
    $schema: 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#',
    contentVersion: '1.0.0.0',
    parameters: {
      $connections: { defaultValue: {}, type: 'Object' },
      $authentication: { defaultValue: {}, type: 'SecureObject' }
    },
    triggers: {
      Recurrence: {
        type: 'Recurrence',
        recurrence: { frequency: 'Minute', interval: 1 },
        runtimeConfiguration: { concurrency: { runs: 1 } }
      }
    },
    actions: {
      Initialize_next: variable('Next', 'integer', 0),
      Initialize_code: variable('Code', 'string', ''),
      Initialize_previous: variable('Previous', 'integer', 0),
      Get_rules: http(site, 'GET', rulesUri(target.configListTitle), undefined, {
        Initialize_next: ['Succeeded'],
        Initialize_code: ['Succeeded'],
        Initialize_previous: ['Succeeded']
      }),
      Rule_rows: {
        type: 'Compose',
        inputs: "@coalesce(body('Get_rules')?['value'], body('Get_rules')?['d']?['results'], json('[]'))",
        runAfter: { Get_rules: ['Succeeded'] }
      },
      For_each_rule: {
        type: 'Foreach',
        foreach: "@outputs('Rule_rows')",
        runtimeConfiguration: { concurrency: { repetitions: 1 } },
        runAfter: { Rule_rows: ['Succeeded'] },
        actions: {
          Scope_rule: {
            type: 'Scope',
            actions: ruleActions(site, bakedTitle)
          },
          Rule_done: {
            type: 'Compose',
            inputs: 'ok',
            runAfter: { Scope_rule: ['Succeeded', 'Failed', 'Skipped', 'TimedOut'] }
          }
        }
      }
    }
  };
}

function rulesUri(title: string): string {
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
  const filter = encodeURIComponent("IsActive eq 1 and TargetListGuid ne 'site'");
  return `_api/web/lists/getbytitle('${listTitle}')/items?$select=${select}&$filter=${filter}&$top=200`;
}

function ruleActions(site: string, bakedTitle: string): { [name: string]: object } {
  return {
    Config: {
      type: 'Compose',
      inputs: '@item()'
    },
    Column_name: {
      type: 'Compose',
      inputs: "@outputs('Config')?['NumberColumnInternalName']",
      runAfter: { Config: ['Succeeded'] }
    },
    Is_active: {
      type: 'If',
      expression: activeExpression(),
      runAfter: { Column_name: ['Succeeded'] },
      else: { actions: {} },
      actions: {
        Web: {
          type: 'Compose',
          inputs: webExpression()
        },
        Seed_rolled: {
          type: 'Compose',
          inputs: `@${rolledExpression()}`
        },
        Set_previous: setVariable('Previous', "@if(outputs('Seed_rolled'), 0, int(coalesce(outputs('Config')?['CurrentCount'], 0)))", {
          Seed_rolled: ['Succeeded']
        }),
        Get_items: http('@outputs(\'Web\')', 'GET', recentItemsUri(), undefined, {
          Web: ['Succeeded'],
          Set_previous: ['Succeeded']
        }),
        Item_rows: {
          type: 'Compose',
          inputs: "@reverse(coalesce(body('Get_items')?['value'], body('Get_items')?['d']?['results'], json('[]')))",
          runAfter: { Get_items: ['Succeeded'] }
        },
        For_each_item: {
          type: 'Foreach',
          foreach: "@outputs('Item_rows')",
          runtimeConfiguration: { concurrency: { repetitions: 1 } },
          runAfter: { Item_rows: ['Succeeded'] },
          actions: {
            Scope_item: {
              type: 'Scope',
              actions: itemActions(site, bakedTitle)
            },
            Item_done: {
              type: 'Compose',
              inputs: 'ok',
              runAfter: { Scope_item: ['Succeeded', 'Failed', 'Skipped', 'TimedOut'] }
            }
          }
        }
      }
    }
  };
}

function itemActions(site: string, bakedTitle: string): { [name: string]: object } {
  const itemId = "coalesce(outputs('Item')?['Id'], outputs('Item')?['ID'])";
  const configId = "coalesce(outputs('Config')?['Id'], outputs('Config')?['ID'])";
  const configUri = `@concat('_api/web/lists/getbytitle(''${bakedTitle}'')/items(', string(${configId}), ')')`;
  const itemUri = `@concat('_api/web/lists(guid''', coalesce(outputs('Config')?['TargetListGuid'], ''), ''')/items(', string(${itemId}), ')')`;
  const stamp = "utcNow('yyyy-MM-ddTHH:mm:ssZ')";
  return {
    Item: {
      type: 'Compose',
      inputs: '@item()'
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

function webExpression(): string {
  const url = "coalesce(outputs('Config')?['TargetListUrl'], '')";
  const lower = `toLower(${url})`;
  const listsAt = `indexOf(${lower}, '/lists/')`;
  const slashAt = `lastIndexOf(${url}, '/')`;
  return `@if(contains(${lower}, '/lists/'), substring(${url}, 0, ${listsAt}), if(greater(${slashAt}, 0), substring(${url}, 0, ${slashAt}), ${url}))`;
}

function recentItemsUri(): string {
  return "@concat('_api/web/lists(guid''', coalesce(outputs('Config')?['TargetListGuid'], ''), ''')/items?', coalesce(outputs('Config')?['ItemQuery'], ''), '&$orderby=Modified%20desc&$top=25')";
}

function activeExpression(): { and: object[] } {
  const value = "@outputs('Config')?['IsActive']";
  return {
    and: [
      {
        not: {
          or: [
            { equals: [value, false] },
            { equals: ["@toLower(string(outputs('Config')?['IsActive']))", 'false'] },
            { equals: [value, 0] },
            { equals: ["@string(outputs('Config')?['IsActive'])", '0'] }
          ]
        }
      },
      {
        not: {
          equals: ["@toLower(coalesce(outputs('Config')?['TargetListGuid'], ''))", 'site']
        }
      }
    ]
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
  return `@if(equals(item(), '__none__'), variables('Code'), replace(variables('Code'), concat('{', item(), '}'), coalesce(${field}?['Title'], ${field}?['Value'], string(${field}))))`;
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
