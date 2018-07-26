import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { HttpClient } from '@microsoft/sp-http';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'UserAccessWebPartStrings';
import UserAccess from './components/UserAccess';
import { IUserAccessProps } from './components/IUserAccessProps';
import { UserAccessItem, RoleToTransaction, PrimaryApproverItem } from "./datamodel";

import { find, filter } from "lodash";
import { sp, EmailProperties, Items, Item } from "@pnp/sp";

export interface IUserAccessWebPartProps {
  webApiUrl: string;
  roleToTcodeController: string;
  primaryApproverController: string;
  userAccessController: string;
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
        webApiUrl: this.properties.webApiUrl,
        roleToTcodeController: this.properties.roleToTcodeController,
        primaryApproverController: this.properties.primaryApproverController,
        userAccessController: this.properties.userAccessController,
        domElement: this.domElement,
        httpClient: this.context.httpClient
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
                PropertyPaneTextField('webApiUrl', {
                  label: "Web API Url"
                }),
                PropertyPaneTextField('primaryApproverController', {
                  label: "Primary Approvers Controller"
                }),
                PropertyPaneTextField('userAccessController', {
                  label: "User Access Controller"
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
