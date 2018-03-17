import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'HighRiskSiteSetupWebPartStrings';
import HighRiskSiteSetup from './components/HighRiskSiteSetup';
import { IHighRiskSiteSetupProps } from './components/IHighRiskSiteSetupProps';

export interface IHighRiskSiteSetupWebPartProps {
  description: string;
}

export default class HighRiskSiteSetupWebPart extends BaseClientSideWebPart<IHighRiskSiteSetupWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IHighRiskSiteSetupProps > = React.createElement(
      HighRiskSiteSetup,
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
