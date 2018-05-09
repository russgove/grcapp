import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import {sp} from "@pnp/sp";
import * as strings from 'UserAccessSiteSetupWebPartStrings';
import UserAccessSiteSetup from './components/UserAccessSiteSetup';
import { IUserAccessSiteSetupProps } from './components/IUserAccessSiteSetupProps';
import { PropertyFieldCodeEditor,PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';

export interface IUserAccessSiteSetupWebPartProps {
  userAccessListName:string;
  primaryApproversListName:string;
  tcodeListName:string;
  webPartXml:string;
}

export default class UserAccessSiteSetupWebPart extends BaseClientSideWebPart<IUserAccessSiteSetupWebPartProps> {
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
    const element: React.ReactElement<IUserAccessSiteSetupProps > = React.createElement(
      UserAccessSiteSetup,
      {
        userAccessListName:this.properties.userAccessListName,
        primaryApproversListName:this.properties.primaryApproversListName,
        tcodeListName:this.properties.tcodeListName,
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
                PropertyPaneTextField('userAccessListName', {
                  label: "User Access List Name"
                }),
                PropertyPaneTextField('primaryApproversListName', {
                  label: "Primary Approvers List Name"
                }),
                PropertyPaneTextField('tcodeListName', {
                  label: "tCODE List Name"
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
