export const FIELD_CREATION_OPTIONS = 25;

export interface IConfigFieldDefinition {
  internalName: string;
  displayName: string;
  schemaXml: string;
}

export const CONFIG_FIELD_DEFINITIONS: IConfigFieldDefinition[] = [
  {
    internalName: 'TargetListUrl',
    displayName: 'Target list URL',
    schemaXml: '<Field Type="Text" Name="TargetListUrl" StaticName="TargetListUrl" DisplayName="Target list URL" Required="TRUE" />'
  },
  {
    internalName: 'TargetListGuid',
    displayName: 'Target list GUID',
    schemaXml: '<Field Type="Text" Name="TargetListGuid" StaticName="TargetListGuid" DisplayName="Target list GUID" Required="TRUE" />'
  },
  {
    internalName: 'NumberColumnInternalName',
    displayName: 'Number column internal name',
    schemaXml: '<Field Type="Text" Name="NumberColumnInternalName" StaticName="NumberColumnInternalName" DisplayName="Number column internal name" Required="TRUE" />'
  },
  {
    internalName: 'Formula',
    displayName: 'Formula',
    schemaXml: '<Field Type="Text" Name="Formula" StaticName="Formula" DisplayName="Formula" Required="TRUE" />'
  },
  {
    internalName: 'CurrentCount',
    displayName: 'Current count',
    schemaXml: '<Field Type="Number" Name="CurrentCount" StaticName="CurrentCount" DisplayName="Current count" Decimals="0" Min="0" />'
  },
  {
    internalName: 'ResetPeriod',
    displayName: 'Reset period',
    schemaXml: '<Field Type="Choice" Name="ResetPeriod" StaticName="ResetPeriod" DisplayName="Reset period" Format="Dropdown"><CHOICES><CHOICE>None</CHOICE><CHOICE>Day</CHOICE><CHOICE>Month</CHOICE><CHOICE>Year</CHOICE></CHOICES><Default>None</Default></Field>'
  },
  {
    internalName: 'LastResetDate',
    displayName: 'Last reset date',
    schemaXml: '<Field Type="DateTime" Name="LastResetDate" StaticName="LastResetDate" DisplayName="Last reset date" Format="DateTime" />'
  },
  {
    internalName: 'IsActive',
    displayName: 'Is active',
    schemaXml: '<Field Type="Boolean" Name="IsActive" StaticName="IsActive" DisplayName="Is active"><Default>1</Default></Field>'
  },
  {
    internalName: 'PadLength',
    displayName: 'Pad length',
    schemaXml: '<Field Type="Number" Name="PadLength" StaticName="PadLength" DisplayName="Pad length" Decimals="0" Min="0" Max="12"><Default>4</Default></Field>'
  },
  {
    internalName: 'WebhookSubscriptionId',
    displayName: 'Webhook subscription ID',
    schemaXml: '<Field Type="Text" Name="WebhookSubscriptionId" StaticName="WebhookSubscriptionId" DisplayName="Webhook subscription ID" />'
  }
];

export function buildCreateFieldBody(schemaXml: string): {
  parameters: {
    SchemaXml: string;
    Options: number;
  };
} {
  return {
    parameters: {
      SchemaXml: schemaXml,
      Options: FIELD_CREATION_OPTIONS
    }
  };
}
