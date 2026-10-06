import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ItemIdGeneratorWebPartStrings';
import ItemIdGenerator from './components/ItemIdGenerator';
import { IItemIdGeneratorProps } from './components/IItemIdGeneratorProps';

export interface IItemIdGeneratorWebPartProps {
  functionUrl: string;
  listName: string;
  targetFieldInternalName: string;
  condition: string;
}

export default class ItemIdGeneratorWebPart extends BaseClientSideWebPart<IItemIdGeneratorWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IItemIdGeneratorProps> = React.createElement(
      ItemIdGenerator,
      {
        functionUrl: this.properties.functionUrl || '',
        listName: this.properties.listName || '',
        targetFieldInternalName: this.properties.targetFieldInternalName || '',
        condition: this.properties.condition || 'default',
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
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
                PropertyPaneTextField('functionUrl', {
                  label: strings.FunctionUrlFieldLabel,
                  description: strings.FunctionUrlFieldDescription,
                  placeholder: 'https://<function-app>.azurewebsites.net/api/GenerateItemId'
                }),
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel,
                  description: strings.ListNameFieldDescription,
                  placeholder: 'Requests'
                }),
                PropertyPaneTextField('targetFieldInternalName', {
                  label: strings.TargetFieldFieldLabel,
                  description: strings.TargetFieldFieldDescription,
                  placeholder: 'GeneratedItemId'
                }),
                PropertyPaneDropdown('condition', {
                  label: strings.ConditionFieldLabel,
                  selectedKey: this.properties.condition || 'default',
                  options: [
                    { key: 'default', text: 'Default ({list}-{timestamp})' },
                    { key: 'VIP', text: 'VIP (VIP-{list}-{timestamp})' }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
