import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'GetSpListItemsWebPartStrings';
import GetSpListItems from './components/GetSpListItems';
import { IGetSpListItemsProps } from './components/IGetSpListItemsProps';

export interface IGetSpListItemsWebPartProps {
  description: string;
}

export default class GetSpListItemsWebPart extends BaseClientSideWebPart<IGetSpListItemsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IGetSpListItemsProps> = React.createElement(
      GetSpListItems,
      {
        description: this.properties.description,
        webURL:this.context.pageContext.web.absoluteUrl,

      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
