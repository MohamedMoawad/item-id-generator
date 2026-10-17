export interface ISiteFlowDefinition {
  $schema: string;
  contentVersion: string;
  parameters: { [name: string]: { defaultValue: object; type: string } };
  triggers: { [name: string]: object };
  actions: { [name: string]: object };
}

export interface ICreateFlowPayload {
  properties: {
    displayName: string;
    definition: ISiteFlowDefinition;
    connectionReferences: {
      shared_sharepointonline: {
        connectionName: string;
        source: string;
        id: string;
        tier: string;
      };
    };
    flowFailureAlertSubscribed: boolean;
  };
}

export function flowDisplayName(siteUrl: string): string {
  try {
    const path = new URL(siteUrl).pathname.replace(/\/+$/, '');
    const leaf = path.split('/').filter((part) => part.length > 0).pop() || 'site';
    return `AutoGen · ${decodeURIComponent(leaf)}`.slice(0, 80);
  } catch {
    return 'AutoGen numbering';
  }
}

export function buildCreateFlowPayload(siteUrl: string, listTitle: string, connectionName: string): ICreateFlowPayload {
  return {
    properties: {
      displayName: flowDisplayName(siteUrl),
      definition: buildSiteFlowDefinition(siteUrl, listTitle),
      connectionReferences: {
        shared_sharepointonline: {
          connectionName,
          source: 'Embedded',
          id: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline',
          tier: 'NotSpecified'
        }
      },
      flowFailureAlertSubscribed: false
    }
  };
}

export function buildSiteFlowDefinition(siteUrl: string, listTitle: string): ISiteFlowDefinition {
  const site = siteUrl.replace(/\/+$/, '');
  const title = listTitle.replace(/'/g, "''");
  const bakedTitle = listTitle.replace(/'/g, "''''");
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
      Get_rules: http(site, 'GET', rulesUri(title), undefined, {
        Initialize_next: ['Succeeded'],
        Initialize_code: ['Succeeded'],
        Initialize_previous: ['Succeeded']
      }),
      For_each_rule: {
        type: 'Foreach',
        foreach: "@coalesce(body('Get_rules')?['value'], body('Get_rules')?['d']?['results'], json('[]'))",
        runtimeConfiguration: { concurrency: { repetitions: 1 } },
        runAfter: { Get_rules: ['Succeeded'] },
        actions: ruleActions(site, bakedTitle)
      }
    }
  };
}

function rulesUri(title: string): string {
  const select = [
    'Id',
    'TargetListGuid',
    'NumberColumnInternalName',
    'Formula',
    'FormulaColumns',
    'CounterToken',
    'CounterWidth',
    'CurrentCount',
    'ResetPeriod',
    'LastResetDate',
    'ItemQuery'
  ].join(',');
  return `_api/web/lists/getbytitle('${title}')/items?$select=${select}&$filter=IsActive+eq+1+and+TargetListGuid+ne+'site'&$top=100`;
}

function ruleActions(site: string, bakedTitle: string): { [name: string]: object } {
  return {
    Rolled: {
      type: 'Compose',
      inputs: `@${rolledExpression()}`
    },
    Set_next: setVariable('Next', "@if(outputs('Rolled'), 1, add(int(coalesce(items('For_each_rule')?['CurrentCount'], 0)), 1))", {
      Rolled: ['Succeeded']
    }),
    Set_previous: setVariable('Previous', "@if(outputs('Rolled'), 0, int(coalesce(items('For_each_rule')?['CurrentCount'], 0)))", {
      Set_next: ['Succeeded']
    }),
    Has_query: {
      type: 'If',
      expression: {
        and: [
          {
            not: {
              equals: [
                "@coalesce(items('For_each_rule')?['ItemQuery'], '')",
                ''
              ]
            }
          }
        ]
      },
      runAfter: { Set_previous: ['Succeeded'] },
      else: { actions: {} },
      actions: {
        Get_blanks: http(site, 'GET', "@items('For_each_rule')?['ItemQuery']", undefined, {}),
        For_each_item: {
          type: 'Foreach',
          foreach: "@coalesce(body('Get_blanks')?['value'], body('Get_blanks')?['d']?['results'], json('[]'))",
          runtimeConfiguration: { concurrency: { repetitions: 1 } },
          runAfter: { Get_blanks: ['Succeeded'] },
          actions: itemActions(site, bakedTitle)
        }
      }
    }
  };
}

function itemActions(site: string, bakedTitle: string): { [name: string]: object } {
  const ruleItem = `@concat('_api/web/lists/getbytitle(''${bakedTitle}'')/items(', string(items('For_each_rule')?['Id']), ')')`;
  const targetItem = "@concat('_api/web/lists(guid''', items('For_each_rule')?['TargetListGuid'], ''')/items(', string(items('For_each_item')?['Id']), ')')";
  return {
    Set_code: setVariable('Code', codeExpression(), {}),
    For_each_column: {
      type: 'Foreach',
      foreach: "@split(if(equals(coalesce(items('For_each_rule')?['FormulaColumns'], ''), ''), '__none__', items('For_each_rule')?['FormulaColumns']), ',')",
      runtimeConfiguration: { concurrency: { repetitions: 1 } },
      runAfter: { Set_code: ['Succeeded'] },
      actions: {
        Apply_column: setVariable('Code', columnExpression(), {})
      }
    },
    Reserve_count: http(
      site,
      'POST',
      ruleItem,
      "@concat('{\"CurrentCount\":', string(variables('Next')), ',\"LastResetDate\":\"', utcNow(), '\"}')",
      { For_each_column: ['Succeeded'] }
    ),
    Write_item: http(
      site,
      'POST',
      targetItem,
      "@concat('{\"', items('For_each_rule')?['NumberColumnInternalName'], '\":\"', replace(variables('Code'), '\"', ''), '\"}')",
      { Reserve_count: ['Succeeded'] }
    ),
    Remember_count: setVariable('Previous', "@variables('Next')", { Write_item: ['Succeeded'] }),
    Advance_count: setVariable('Next', "@add(variables('Next'), 1)", { Remember_count: ['Succeeded'] }),
    Restore_count: http(
      site,
      'POST',
      ruleItem,
      "@concat('{\"CurrentCount\":', string(variables('Previous')), '}')",
      { Write_item: ['Failed', 'TimedOut'] }
    ),
    Stop_run: {
      type: 'Terminate',
      inputs: { runStatus: 'Failed' },
      runAfter: { Restore_count: ['Succeeded'] }
    }
  };
}

function rolledExpression(): string {
  const period = "items('For_each_rule')?['ResetPeriod']";
  const last = "items('For_each_rule')?['LastResetDate']";
  const key = (source: string): string => {
    return `if(equals(${period}, 'Day'), formatDateTime(${source}, 'yyyy-MM-dd'), if(equals(${period}, 'Month'), formatDateTime(${source}, 'yyyy-MM'), if(equals(${period}, 'Year'), formatDateTime(${source}, 'yyyy'), '')))`;
  };
  return `if(or(equals(${period}, 'None'), empty(${last})), false, not(equals(${key(last)}, ${key('utcNow()')})))`;
}

function codeExpression(): string {
  const width = "int(coalesce(items('For_each_rule')?['CounterWidth'], 0))";
  const padded = `if(or(equals(${width}, 0), greater(length(string(variables('Next'))), ${width})), string(variables('Next')), substring(concat('000000000000', string(variables('Next'))), sub(12, sub(${width}, length(string(variables('Next'))))), ${width}))`;
  const created = "items('For_each_item')?['Created']";
  return `@replace(replace(replace(replace(replace(replace(replace(items('For_each_rule')?['Formula'], items('For_each_rule')?['CounterToken'], ${padded}), '{yyyy}', formatDateTime(${created}, 'yyyy')), '{yy}', formatDateTime(${created}, 'yy')), '{MM}', formatDateTime(${created}, 'MM')), '{dd}', formatDateTime(${created}, 'dd')), '{HH}', formatDateTime(${created}, 'HH')), '{mm}', formatDateTime(${created}, 'mm'))`;
}

function columnExpression(): string {
  const field = "items('For_each_item')?[first(split(item(), '.'))]";
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
