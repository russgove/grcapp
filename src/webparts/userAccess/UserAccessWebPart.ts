import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { HttpClient } from '@microsoft/sp-http';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

import * as strings from 'UserAccessWebPartStrings';
import UserAccess from './components/UserAccess';
import { IUserAccessProps } from './components/IUserAccessProps';
import { UserAccessItem, RoleToTransaction, PrimaryApproverItem } from "./datamodel";

import { find, filter } from "lodash";
import { sp, EmailProperties, Items, Item } from "@pnp/sp";

export interface IUserAccessWebPartProps {
azureFunctionUrl: string;
  accessCode: string;
  userAccessReviewPath: string;
  primaryApproversPath: string;
  roleToTransactionsPath: string;
  enableUncomplete: boolean;
  system:string;

}

export default class UserAccessWebPart extends BaseClientSideWebPart<IUserAccessWebPartProps> {

  private reactElement: React.ReactElement<IUserAccessProps>;
  private formComponent: UserAccess;
  
  public async onInit(): Promise<void> {
    await super.onInit().then(() => {
      sp.setup({
        spfxContext: this.context,
      });
      return;
    });
  }

  public render(): void {
    this.reactElement = React.createElement(
      UserAccess,
      {
        user: this.context.pageContext.user,
        azureFunctionUrl: this.properties.azureFunctionUrl,
        domElement: this.domElement,
        httpClient: this.context.httpClient,
        accessCode:this.properties.accessCode,
        userAccessReviewPath:this.properties.userAccessReviewPath,
        primaryApproversPath:this.properties.primaryApproversPath,
        roleToTransactionsPath:this.properties.roleToTransactionsPath,
        enableUncomplete:this.properties.enableUncomplete,
        system:this.properties.system
      }
    );
    this.formComponent = ReactDom.render(this.reactElement, this.domElement) as UserAccess;
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
                PropertyPaneTextField('userAccessReviewPath', {
                  label: "Path to User acess function (RoleReviews or EPXROleReviews)"
                }),
                PropertyPaneTextField('primaryApproversPath', {
                  label: "Path to Primary Approvers  in Azure function (PrimaryApprovers or EPXPrimaryApprovers)"
                }),
                PropertyPaneTextField('roleToTransactionsPath', {
                  label: "Path to Transactions   in Azure function (RoleToTransaction or EPXRoleToTranaction)"
                }),
                PropertyPaneTextField('system', {
                  label: "EPA/EPX/GRP"
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
