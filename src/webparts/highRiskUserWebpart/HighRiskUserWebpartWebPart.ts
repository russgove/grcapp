import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'HighRiskUserWebpartWebPartStrings';
import HighRiskUserWebpart from './components/HighRiskUserWebpart';
import { IHighRiskUserWebpartProps } from './components/IHighRiskUserWebpartProps';

export interface IHighRiskUserWebpartWebPartProps {
  description: string;
}

export default class HighRiskUserWebpartWebPart extends BaseClientSideWebPart<IHighRiskUserWebpartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IHighRiskUserWebpartProps > = React.createElement(
      HighRiskUserWebpart,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
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
