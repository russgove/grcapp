import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import pnp from "sp-pnp-js";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';

import * as strings from 'HighRiskAdminWebpartWebPartStrings';
import HighRiskAdminWebpart from './components/HighRiskAdminWebpart';
import { IHighRiskAdminWebpartProps } from './components/IHighRiskAdminWebpartProps';

export interface IHighRiskAdminWebpartWebPartProps {
  templateName:string;
  azureHighRiskUrl:string;
  azurePrimaryApproverUrl:string;// url to initiate processing of a Primary Approver  File
  azureRoleToCodeUrl:string;// url to initiate processing of a roleto transaction File
  webPartXml:string;
  batchSize:number;
  pauseBeforeBatchExecution:number;
}

export default class HighRiskAdminWebpartWebPart extends BaseClientSideWebPart<IHighRiskAdminWebpartWebPartProps> {

  public async onInit(): Promise<any> {
    await super.onInit().then(() => {
      pnp.setup({
        spfxContext: this.context,
      });
      return;
    });
    debugger;
    return Promise.resolve();

  }
  public render(): void {
    debugger;
    // const element: React.ReactElement<IHighRiskAdminWebpartProps > = React.createElement(
    //   HighRiskAdminWebpart,
    //   {
    //     siteUrl:this.context.pageContext.site.serverRelativeUrl,
    //     siteAbsoluteUrl:this.context.pageContext.site.absoluteUrl,
    //     templateName:this.properties.templateName,
    //     webPartXml:this.properties.webPartXml,
    //     httpClient:this.context.httpClient,
    //     azureRoleToCodeUrl: this.properties.azureRoleToCodeUrl,
    //     azurePrimaryApproverUrl:this.properties.azurePrimaryApproverUrl,
    //     azureHighRiskUrl:this.properties.azureHighRiskUrl,
    //     batchSize:this.properties.batchSize,
    //     pauseBeforeBatchExecution:this.properties.pauseBeforeBatchExecution
        
    //   }
    // );

   // ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('templateName', {
                  label: "Template user to create subsite (STS#0)"
                }),
                PropertyPaneTextField('webPartXml', {
                  label: "Webpart xml for custom webpart on home page"
                }),
                PropertyPaneTextField('azureFunctionUrl', {
                  label: "url to call to start processing batch files"
                }),
                PropertyPaneTextField('azureRoleToTransactionFunctionUrl', {
                  label: "url to call to process rtc"
                }),
                PropertyPaneSlider('batchSize', {
                  min:1,
                  max:100


                }),
                PropertyPaneSlider('pauseBeforeBatchExecution', {
                  min:0,
                  max:60000,
                  step:1000
                })




              ]
            }
          ]
        }
      ]
    };
  }
}
