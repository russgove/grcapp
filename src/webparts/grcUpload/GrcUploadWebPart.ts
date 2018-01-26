import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import {} from "office-ui-fabric-react/"
import * as strings from 'GrcUploadWebPartStrings';
import GrcUpload from './components/GrcUpload';
import { IGrcUploadProps } from './components/IGrcUploadProps';

export interface IGrcUploadWebPartProps {
  description: string;
}

export default class GrcUploadWebPart extends BaseClientSideWebPart<IGrcUploadWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IGrcUploadProps > = React.createElement(
      GrcUpload,
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
