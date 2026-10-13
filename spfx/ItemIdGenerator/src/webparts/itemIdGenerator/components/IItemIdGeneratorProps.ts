import { SPHttpClient } from '@microsoft/sp-http';

export interface IItemIdGeneratorProps {
  configListTitle: string;
  siteAbsoluteUrl: string;
  spHttpClient: SPHttpClient;
}
