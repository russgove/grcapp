import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

import * as strings from 'HighRiskFunctionsWebPartStrings';
import HighRiskFunctions from './components/HighRiskFunctions';
import { IHighRiskFunctionsProps } from './components/IHighRiskFunctionsProps';

export interface IHighRiskFunctionsWebPartProps {
  azureFunctionUrl: string;
  accessCode:string;
  highRiskFunctionsPath:string;
  primaryApproversPath:string;
  roleToTransactionsPath:string;
  enableUncomplete:boolean;

}

export default class HighRiskFunctionsWebPart extends BaseClientSideWebPart<IHighRiskFunctionsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IHighRiskFunctionsProps > = React.createElement(
      HighRiskFunctions,
      {
        user: this.context.pageContext.user,
        azureFunctionUrl: this.properties.azureFunctionUrl,
        domElement:this.domElement,
        accessCode:this.properties.accessCode,
        highRiskFunctionsPath:this.properties.highRiskFunctionsPath,
        primaryApproversPath:this.properties.primaryApproversPath,
        roleToTransactionsPath:this.properties.roleToTransactionsPath,
        enableUncomplete:this.properties.enableUncomplete,
        httpClient:
        this.context.httpClient
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
