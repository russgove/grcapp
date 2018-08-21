import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import * as strings from 'MitigatingControlsWebPartStrings';
import MitigatingControls from './components/MitigatingControls';
import { IMitigatingControlsProps } from './components/IMitigatingControlsProps';
import  { sp,EmailProperties, Items, Item } from "@pnp/sp";
import { find, filter } from "lodash";
//import fieldNames from "./fieldNames"
import { MitigatingControlsItem, PrimaryApproverItem, HelpLink } from "./dataModel";
export interface IMitigatingControlsWebPartProps {
  webApiUrl:string;
  mitigatngControlsController:string;
  primaryApproverController:string;
  effectiveLabel: string;
  continuesLabel: string;
  correctPersonLabel: string;
  helpLinksListName: string;
}

export default class MitigatingControlsWebPart extends BaseClientSideWebPart<IMitigatingControlsWebPartProps> {
  private primaryApproverLists: Array<any>;
  private highRisks: Array<any>;
  private helpLinks: Array<HelpLink>;
  private reactElement: React.ReactElement<IMitigatingControlsProps>;
  private formComponent: MitigatingControls;
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
      MitigatingControls,
      {
        webApiUrl:this.properties.webApiUrl,
        user: this.context.pageContext.user,
     
     
        primaryApproverController: this.properties.primaryApproverController,
        mitigatngControlsController: this.properties.mitigatngControlsController,
        domElement: this.domElement,
        httpClient: this.context.httpClient,
   
        effectiveLabel: this.properties.effectiveLabel,
        continuesLabel: this.properties.continuesLabel,
        correctPersonLabel: this.properties.correctPersonLabel,
        helpLinks: this.helpLinks

      }
    );

    this.formComponent = ReactDom.render(this.reactElement, this.domElement) as MitigatingControls;

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
                PropertyPaneTextField('primaryApproversListName', {
                  label: "Primary Approvers List"
                }),
                PropertyPaneTextField('mitigatingControlsListName', {
                  label: "Label For:High Risk with Mitigating Controls List"
                }),
                PropertyPaneTextField('effectiveLabel', {
                  label: "Label For:Does the mitigating control effectively remediate the assiciated risk?"
                }),
                PropertyPaneTextField('correctPersonLabel', {
                  label: "Label For:Is the mitigating control monitor the correct person to perform control?"
                })

              ]
            }
          ]
        }
      ]
    };
  }
}
