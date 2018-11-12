import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { AadHttpClient } from '@microsoft/sp-http';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'RoleToTCodeWebPartStrings';
import RoleToTCode from './components/RoleToTCode';
import { IRoleToTCodeProps } from './components/IRoleToTCodeProps';

export interface IRoleToTCodeWebPartProps {

  GetPrimaryApproverByEmailPath: string;
  UpdatePrimaryApproversPath: string;
  GetRoleReviewsForApproverPath: string;
  UpdateRoleReviewsForApproverPath: string;
  GetRoleToTransactionsForRoleNamePath: string;
}
import { sp, EmailProperties, Items, Item } from "@pnp/sp";

export default class RoleToTCodeWebPart extends BaseClientSideWebPart<IRoleToTCodeWebPartProps> {

  private reactElement: React.ReactElement<IRoleToTCodeProps>;
  private formComponent: RoleToTCode;
 

  public async onInit(): Promise<void> {
    await super.onInit().then(() => {
      sp.setup({
        spfxContext: this.context,
      });
     
    });
  }
  public render(): void {

    this.reactElement = React.createElement(
      RoleToTCode,
      {
        GetPrimaryApproverByEmailPath: this.properties.GetPrimaryApproverByEmailPath,
        UpdatePrimaryApproversPath: this.properties.UpdatePrimaryApproversPath,
        GetRoleReviewsForApproverPath: this.properties.GetRoleReviewsForApproverPath,
        UpdateRoleReviewsForApproverPath: this.properties.UpdateRoleReviewsForApproverPath,
        GetRoleToTransactionsForRoleNamePath: this.properties.GetRoleToTransactionsForRoleNamePath,
        user: this.context.pageContext.user,




        domElement: this.domElement,
        httpClient: this.context.httpClient
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
