import { SPHttpClient } from '@microsoft/sp-http';

export interface IItemIdGeneratorProps {
  functionUrl: string;
  listName: string;
  targetFieldInternalName: string;
  condition: string;
  webAbsoluteUrl: string;
  spHttpClient: SPHttpClient;
}
