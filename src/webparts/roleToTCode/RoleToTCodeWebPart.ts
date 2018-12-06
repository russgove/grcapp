import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';


import { AadHttpClient } from '@microsoft/sp-http';
import * as strings from 'RoleToTCodeWebPartStrings';
import RoleToTCode from './components/RoleToTCode';
import { IRoleToTCodeProps } from './components/IRoleToTCodeProps';

export interface IRoleToTCodeWebPartProps {
  azureFunctionUrl: string;
  accessCode:string;
  roleReviewsPath:string;
  primaryApproversPath:string;
  roleToTransactionsPath:string;
  enableUncomplete:boolean;

}
import { sp, EmailProperties, Items, Item } from "@pnp/sp";


export default class RoleToTCodeWebPart extends BaseClientSideWebPart<IRoleToTCodeWebPartProps> {

  private reactElement: React.ReactElement<IRoleToTCodeProps>;
  private formComponent: RoleToTCode;

  

  public async onInit(): Promise<void> {
    // return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
    //   this.context.AadHttpClientFactory
    //     .getClient('594e83da-9618-438f-a40a-4a977c03bc16')
    //     .then((client: AadHttpClient): void => {
    //       this.ordersClient = client;
    //       resolve();
    //     }, err => reject(err));
    // });
  }
  public render(): void {
    
    this.reactElement = React.createElement(
      RoleToTCode,
      {
        user: this.context.pageContext.user,
        azureFunctionUrl: this.properties.azureFunctionUrl,
        domElement: this.domElement,
        httpClient: this.context.httpClient,
        accessCode:this.properties.accessCode,
        roleReviewsPath:this.properties.roleReviewsPath,
        primaryApproversPath:this.properties.primaryApproversPath,
        roleToTransactionsPath:this.properties.roleToTransactionsPath,
        enableUncomplete:this.properties.enableUncomplete
      
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
                PropertyPaneTextField('azureFunctionUrl', {
                  label: "azureFunctionUrl"
                }),
                PropertyPaneTextField('accessCode', {
                  label: "accress code"
                }),
                PropertyPaneTextField('roleReviewsPath', {
                  label: "Path to roleReviews in Azure function (RoleReviews or EPXROleReviews)"
                }),
                PropertyPaneTextField('primaryApproversPath', {
                  label: "Path to Primary Approvers  in Azure function (PrimaryApprovers or EPXPrimaryApprovers)"
                }),
                PropertyPaneTextField('roleToTransactionsPath', {
                  label: "Path to Transactions   in Azure function (RoleToTransaction or EPXRoleToTranaction)"
                }),
                PropertyPaneToggle('enableUncomplete', {
                  label: "Enable UnCompleting review (useful for testing. Turn Off when Live"
                }),

              ]
            }
          ]
        }
      ]
    };
  }
}
