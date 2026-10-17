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

export function flowDisplayName(listTitle: string): string {
  const leaf = listTitle.trim() || 'list';
  return `AutoGen · ${leaf}`.slice(0, 80);
}

export function buildCreateFlowPayload(target: IListFlowTarget, connectionName: string): ICreateFlowPayload {
  return {
    properties: {
      displayName: flowDisplayName(target.listTitle),
      definition: buildSiteFlowDefinition(target),
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
        type: 'OpenApiConnectionWebhook',
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
        splitOn: "@triggerOutputs()?['body/value']",
        runtimeConfiguration: { concurrency: { runs: 1 } }
      }
    },
    actions: {
      Initialize_next: variable('Next', 'integer', 0),
      Initialize_code: variable('Code', 'string', ''),
      Initialize_previous: variable('Previous', 'integer', 0),
      Get_rule: http(site, 'GET', rulesUri(target.configListTitle, guid), undefined, {
        Initialize_next: ['Succeeded'],
        Initialize_code: ['Succeeded'],
        Initialize_previous: ['Succeeded']
      }),
      Has_rule: {
        type: 'If',
        expression: {
          and: [{
            greater: [
              "@length(coalesce(body('Get_rule')?['value'], json('[]')))",
              0
            ]
          }]
        },
        runAfter: { Get_rule: ['Succeeded'] },
        else: { actions: {} },
        actions: numberingActions(web, site, guid, bakedTitle)
      }
    }
  };
}

function rulesUri(title: string, guid: string): string {
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
  const listTitle = title.replace(/'/g, "''");
  return `_api/web/lists/getbytitle('${listTitle}')/items?$select=${select}&$filter=TargetListGuid+eq+'${guid}'+and+IsActive+eq+1&$top=1`;
}

function numberingActions(web: string, site: string, guid: string, bakedTitle: string): { [name: string]: object } {
  const triggerId = "coalesce(triggerOutputs()?['body/ID'], triggerBody()?['ID'])";
  const itemId = `coalesce(${triggerId}, body('Get_item')?['Id'], body('Get_item')?['ID'])`;
  const configId = "coalesce(outputs('Config')?['Id'], outputs('Config')?['ID'])";
  const configUri = `@concat('_api/web/lists/getbytitle(''${bakedTitle}'')/items(', string(${configId}), ')')`;
  const readUri = `@concat('_api/web/lists(guid''${guid}'')/items(', string(${triggerId}), ')?', coalesce(outputs('Config')?['ItemQuery'], ''))`;
  const itemUri = `@concat('_api/web/lists(guid''${guid}'')/items(', string(${itemId}), ')')`;
  return {
    Config: {
      type: 'Compose',
      inputs: "@first(coalesce(body('Get_rule')?['value'], json('[]')))"
    },
    Get_item: http(web, 'GET', readUri, undefined, { Config: ['Succeeded'] }),
    Is_blank: {
      type: 'If',
      expression: {
        and: [{
          equals: [
            "@empty(body('Get_item')?[outputs('Config')?['NumberColumnInternalName']])",
            true
          ]
        }]
      },
      runAfter: { Get_item: ['Succeeded'] },
      else: { actions: {} },
      actions: {
        Rolled: {
          type: 'Compose',
          inputs: `@${rolledExpression()}`
        },
        Set_next: setVariable('Next', "@if(outputs('Rolled'), 1, add(int(coalesce(outputs('Config')?['CurrentCount'], 0)), 1))", {
          Rolled: ['Succeeded']
        }),
        Set_previous: setVariable('Previous', "@if(outputs('Rolled'), 0, int(coalesce(outputs('Config')?['CurrentCount'], 0)))", {
          Set_next: ['Succeeded']
        }),
        Set_code: setVariable('Code', codeExpression(), { Set_previous: ['Succeeded'] }),
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
          "@concat('{\"CurrentCount\":', string(variables('Next')), ',\"LastResetDate\":\"', utcNow(), '\"}')",
          { For_each_column: ['Succeeded'] }
        ),
        Write_item: http(
          web,
          'POST',
          itemUri,
          "@concat('{\"', outputs('Config')?['NumberColumnInternalName'], '\":\"', replace(variables('Code'), '\"', ''), '\"}')",
          { Reserve_count: ['Succeeded'] }
        ),
        Remember_count: setVariable('Previous', "@variables('Next')", { Write_item: ['Succeeded'] }),
        Advance_count: setVariable('Next', "@add(variables('Next'), 1)", { Remember_count: ['Succeeded'] }),
        Restore_count: http(
          site,
          'POST',
          configUri,
          "@concat('{\"CurrentCount\":', string(variables('Previous')), '}')",
          { Write_item: ['Failed', 'TimedOut'] }
        ),
        Stop_run: {
          type: 'Terminate',
          inputs: { runStatus: 'Failed' },
          runAfter: { Restore_count: ['Succeeded'] }
        }
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
  const created = "coalesce(body('Get_item')?['Created'], triggerOutputs()?['body/Created'], utcNow())";
  return `@replace(replace(replace(replace(replace(replace(replace(outputs('Config')?['Formula'], outputs('Config')?['CounterToken'], ${padded}), '{yyyy}', formatDateTime(${created}, 'yyyy')), '{yy}', formatDateTime(${created}, 'yy')), '{MM}', formatDateTime(${created}, 'MM')), '{dd}', formatDateTime(${created}, 'dd')), '{HH}', formatDateTime(${created}, 'HH')), '{mm}', formatDateTime(${created}, 'mm'))`;
}

function columnExpression(): string {
  const field = "body('Get_item')?[first(split(item(), '.'))]";
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
