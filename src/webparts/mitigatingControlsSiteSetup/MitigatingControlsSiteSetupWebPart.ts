import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'MitigatingControlsSiteSetupWebPartStrings';
import MitigatingControlsSiteSetup from './components/MitigatingControlsSiteSetup';
import { IMitigatingControlsSiteSetupProps } from './components/IMitigatingControlsSiteSetupProps';
import { PropertyFieldCodeEditor,PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
import pnp from "sp-pnp-js";
export interface IMitigatingControlsSiteSetupWebPartProps {
  mitigatingControlsListName:string;
  primaryApproversListName:string;
  webPartXml:string;
}

export default class MitigatingControlsSiteSetupWebPart extends BaseClientSideWebPart<IMitigatingControlsSiteSetupWebPartProps> {
  public async onInit(): Promise<any> {
    await super.onInit().then(() => {
      pnp.setup({
        spfxContext: this.context,
      });
      return;
    });
  
    return Promise.resolve();

  }
  public render(): void {
    const element: React.ReactElement<IMitigatingControlsSiteSetupProps > = React.createElement(
      MitigatingControlsSiteSetup,
      {
        mitigatingControlsListName:this.properties.mitigatingControlsListName,
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
                PropertyPaneTextField('mitigatingControlsListName', {
                  label: "Mitigating Controls List Name"
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
