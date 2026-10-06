import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ItemIdGeneratorWebPartStrings';
import ItemIdGenerator from './components/ItemIdGenerator';
import { IItemIdGeneratorProps } from './components/IItemIdGeneratorProps';
import { DEFAULT_CONFIG_LIST_TITLE } from './services/configContract';

export interface IItemIdGeneratorWebPartProps {
  configListTitle: string;
  registerWebhookUrl: string;
}

export default class ItemIdGeneratorWebPart extends BaseClientSideWebPart<IItemIdGeneratorWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IItemIdGeneratorProps> = React.createElement(
      ItemIdGenerator,
      {
        configListTitle: this.properties.configListTitle || DEFAULT_CONFIG_LIST_TITLE,
        siteAbsoluteUrl: this.context.pageContext.site.absoluteUrl,
        registerWebhookUrl: this.properties.registerWebhookUrl || '',
        spHttpClient: this.context.spHttpClient
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('configListTitle', {
                  label: strings.ConfigListTitleFieldLabel,
                  description: strings.ConfigListTitleFieldDescription,
                  placeholder: DEFAULT_CONFIG_LIST_TITLE
                }),
                PropertyPaneTextField('registerWebhookUrl', {
                  label: strings.RegisterWebhookUrlFieldLabel,
                  description: strings.RegisterWebhookUrlFieldDescription,
                  placeholder: 'https://<function-app>.azurewebsites.net/api/RegisterWebhook?code=<function-key>'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
