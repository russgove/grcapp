import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'RoleToTCodeWebPartStrings';
import RoleToTCode from './components/RoleToTCode';
import { IRoleToTCodeProps } from './components/IRoleToTCodeProps';

export interface IRoleToTCodeWebPartProps {
  azureFunctionUrl: string;

}
import { sp, EmailProperties, Items, Item } from "@pnp/sp";
import { AadHttpClient } from "@microsoft/sp-http/dist/index-internal";

export default class RoleToTCodeWebPart extends BaseClientSideWebPart<IRoleToTCodeWebPartProps> {

  private reactElement: React.ReactElement<IRoleToTCodeProps>;
  private formComponent: RoleToTCode;
  private aadHttpClient: AadHttpClient

  public async onInit(): Promise<void> {
    await super.onInit().then(() => {
      sp.setup({
        spfxContext: this.context,
      });
      return this.context.aadHttpClientFactory.getClient(this.properties.azureFunctionUrl).then((client) => {
        this.aadHttpClient = client;
      }).catch((err) => {
        debugger;
      });
      
    });
  }
  public render(): void {
    debugger;
    this.reactElement = React.createElement(
      RoleToTCode,
      {
        user: this.context.pageContext.user,
        azureFunctionUrl: this.properties.azureFunctionUrl,
        domElement: this.domElement,
        aadHttpClient: this.aadHttpClient
      }
    );
    this.formComponent = ReactDom.render(this.reactElement, this.domElement) as RoleToTCode;
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
                PropertyPaneTextField('webApiUrl', {
                  label: "Web API Url"
                }),
                PropertyPaneTextField('primaryApproverController', {
                  label: "Primary Approvers Controller"
                }),
                PropertyPaneTextField('roleReviewController', {
                  label: "Role Review Controller"
                }),
                PropertyPaneTextField('roleToTcodeController', {
                  label: "RoleToTcode Controller"
                }),

              ]
            }
          ]
        }
      ]
    };
  }
}
