import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'BusinessRoleReviewSiteSetupWebPartStrings';
import BusinessRoleReviewSiteSetup from './components/BusinessRoleReviewSiteSetup';
import { IBusinessRoleReviewSiteSetupProps } from './components/IBusinessRoleReviewSiteSetupProps';
import { PropertyFieldCodeEditor,PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
import {sp} from "@pnp/sp";
export interface IBusinessRoleReviewSiteSetupWebPartProps {
  businessRoleReviewListName:string;
  primaryApproversListName:string;
  webPartXml:string;
}

export default class BusinessRoleReviewSiteSetupWebPart extends BaseClientSideWebPart<IBusinessRoleReviewSiteSetupWebPartProps> {
  public async onInit(): Promise<any> {
    await super.onInit().then(() => {
      sp.setup({
        spfxContext: this.context,
      });
      return;
    });
  
    return Promise.resolve();

  }
  public render(): void {
    const element: React.ReactElement<IBusinessRoleReviewSiteSetupProps > = React.createElement(
      BusinessRoleReviewSiteSetup,
      {
        businessRoleReviewListName:this.properties.businessRoleReviewListName,
        primaryApproversListName:this.properties.primaryApproversListName,
        siteUrl:this.context.pageContext.site.absoluteUrl,
        webPartXml:this.properties.webPartXml
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
                PropertyPaneTextField('businessRoleReviewListName', {
                  label: "Business Role Review List Name"
                }),
                PropertyPaneTextField('primaryApproversListName', {
                  label: "Primary Approvers List Name"
                }),
                PropertyFieldCodeEditor('webPartXml', {
                  key:"webPartXml",
                  properties:this.properties,
                  label: "Webpart XML to add to site homepage",
                  panelTitle:"WebpartXML",
                  language: PropertyFieldCodeEditorLanguages.XML,
                  onPropertyChange: (propertyPath: string, oldValue: any, newValue: any)=>{
                    debugger;
                  }
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
