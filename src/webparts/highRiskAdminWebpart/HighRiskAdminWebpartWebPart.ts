import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import pnp from "sp-pnp-js";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'HighRiskAdminWebpartWebPartStrings';
import HighRiskAdminWebpart from './components/HighRiskAdminWebpart';
import { IHighRiskAdminWebpartProps } from './components/IHighRiskAdminWebpartProps';

export interface IHighRiskAdminWebpartWebPartProps {
  templateName:string;
  primaryApproverContentTypeId:string;
  roleToTransactionContentTypeId:string;
  highRiskContentTypeId:string;
  webPartXml:string;
}

export default class HighRiskAdminWebpartWebPart extends BaseClientSideWebPart<IHighRiskAdminWebpartWebPartProps> {

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
    const element: React.ReactElement<IHighRiskAdminWebpartProps > = React.createElement(
      HighRiskAdminWebpart,
      {
        siteUrl:this.context.pageContext.site.serverRelativeUrl,
        siteAbsoluteUrl:this.context.pageContext.site.absoluteUrl,
        templateName:this.properties.templateName,
        primaryApproverContentTypeId: this.properties.primaryApproverContentTypeId,
        roleToTransactionContentTypeId: this.properties.roleToTransactionContentTypeId,
        highRiskContentTypeId: this.properties.highRiskContentTypeId,
        webPartXml:this.properties.webPartXml,
        azureFunctionUrl:"https://grctest.azurewebsites.net/api/HttpTriggerCSharp1?code=HBM82bnia7nKPC/nqVTbaCmfPaFyubQa8iY22lb0r88fdQH370CRUg==",
        httpClient:this.context.httpClient
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
