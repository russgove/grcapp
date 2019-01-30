import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider, IPropertyPaneSliderProps, PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import * as strings from 'BusinessRoleReviewWebPartStrings';
import BusinessRoleReview from './components/BusinessRoleReview';
import { IBusinessRoleReviewProps } from './components/IBusinessRoleReviewProps';
import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
import { sp } from "@pnp/sp";
import { find, filter } from "lodash";
import { BusinessRoleReviewItem, PrimaryApproverItem, HelpLink } from "./dataModel";
import { SPUser } from '@microsoft/sp-page-context';
export interface IBusinessRoleReviewWebPartProps {
  azureFunctionUrl: string;
  accessCode: string;
  user: SPUser; // the sharepoint user accessing the webpart
  domElement: any; // needed to disable button postback after render on classic pages
  businessRoleOwnersPath: string;
  primaryApproversPath: string;
  enableUncomplete: boolean; // should we show menu item to uncomplete(for use in testing)
  roleNameWidth: number;
  // compositeRoleWidth:number;
  approverWidth: number;
  altApproverWidth: number;
  approvalDecisionWidth: number;
  commentsWidth: number;
  helpLinksListName: string;

}

export default class BusinessRoleReviewWebPart extends BaseClientSideWebPart<IBusinessRoleReviewWebPartProps> {
  private primaryApproverLists: Array<any>;
  private highRisks: Array<any>;
  private reactElement: React.ReactElement<IBusinessRoleReviewProps>;
  private formComponent: BusinessRoleReview;
  private helpLinks: Array<HelpLink>;
  public async onInit(): Promise<void> {



  }

  public render(): void {
    this.reactElement = React.createElement(
      BusinessRoleReview,
      {
        azureFunctionUrl: this.properties.azureFunctionUrl,
        accessCode: this.properties.accessCode,
        user: this.context.pageContext.user,
        domElement: this.domElement,
        httpClient: this.context.httpClient,
        businessRoleOwnersPath: this.properties.businessRoleOwnersPath,
        primaryApproversPath: this.properties.primaryApproversPath,
        enableUncomplete: this.properties.enableUncomplete,
        roleNameWidth: this.properties.roleNameWidth,
        approverWidth: this.properties.approverWidth,
        altApproverWidth: this.properties.altApproverWidth,
        approvalDecisionWidth: this.properties.approvalDecisionWidth,
        commentsWidth: this.properties.commentsWidth,
        helpLinks: this.helpLinks

      }
    );

    this.formComponent = ReactDom.render(this.reactElement, this.domElement) as BusinessRoleReview;

    if (Environment.type === EnvironmentType.ClassicSharePoint) {
      const buttons: NodeListOf<HTMLButtonElement> = this.domElement.getElementsByTagName('button');
      if (buttons && buttons.length) {
        for (let i: number = 0; i < buttons.length; i++) {
          if (buttons[i]) {
            /* tslint:disable */
            // Disable the button onclick postback
            buttons[i].onclick = function () { return false; };
            /* tslint:enable */
          }
        }
      }
    }
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
                PropertyPaneTextField('businessRoleOwnersPath', {
                  label: "Path to businessRoleOwnersPath in Azure function (RoleReviews or EPXROleReviews)"
                }),
                PropertyPaneTextField('primaryApproversPath', {
                  label: "Path to Primary Approvers  in Azure function (PrimaryApprovers or EPXPrimaryApprovers)"
                }),
                PropertyPaneToggle('enableUncomplete', {
                  label: "Enable UnCompleting review (useful for testing. Turn Off when Live"
                }),
                PropertyPaneSlider("roleNameWidth", {
                  min: 10,
                  max: 1000,
                  label: "Width of Role Name column",
                  showValue: true
                }),
                // PropertyPaneSlider("compositeRoleWidth",{
                //   min:10,
                //   max:1000,
                //   label:"Width of Composite Role Name column",
                //   showValue:true
                // }),
                PropertyPaneSlider("approverWidth", {
                  min: 10,
                  max: 1000,
                  label: "Width of Approver column",
                  showValue: true
                }),
                PropertyPaneSlider("altApproverWidth", {
                  min: 10,
                  max: 1000,
                  label: "Width of Alternate Approver column",
                  showValue: true
                }),
                PropertyPaneSlider("approvalDecisionWidth", {
                  min: 10,
                  max: 1000,
                  label: "Width of Approval Decision column",
                  showValue: true
                }),
                PropertyPaneSlider("commentsWidth", {
                  min: 10,
                  max: 1000,
                  label: "Width of Comments column",
                  showValue: true
                }),

              ]
            }
          ]
        }
      ]
    };
  }
}
