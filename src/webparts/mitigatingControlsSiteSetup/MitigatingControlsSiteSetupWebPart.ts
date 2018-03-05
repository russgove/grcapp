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
import pnp from "sp-pnp-js";
export interface IMitigatingControlsSiteSetupWebPartProps {
  mitigatingControlsListName:string;
  primaryApproversListName:string;
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
        primaryApproversListName:this.properties.primaryApproversListName
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
