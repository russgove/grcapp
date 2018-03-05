// import * as React from 'react';
// import * as ReactDom from 'react-dom';
// import { Version } from '@microsoft/sp-core-library';
// import {
//   BaseClientSideWebPart,
//   IPropertyPaneConfiguration,
//   PropertyPaneTextField
// } from '@microsoft/sp-webpart-base';

// import * as strings from 'GrcAzureManagementWebPartStrings';
// import GrcAzureManagement from './components/GrcAzureManagement';
// import { IGrcAzureManagementProps } from './components/IGrcAzureManagementProps';

// import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';

// export interface IGrcAzureManagementWebPartProps {
//   description: string;
//   storageAccountName:string;
// }

// export default class GrcAzureManagementWebPart extends BaseClientSideWebPart<IGrcAzureManagementWebPartProps> {

//   public render(): void {

//     const requestHeaders: Headers = new Headers();
//     requestHeaders.append("Content-type", "application/json");
//     requestHeaders.append("Cache-Control", "no-cache");

//     const postOptions: IHttpClientOptions = {
//       headers: requestHeaders,
//     };
//     debugger;
//     let url=`https://${this.properties.storageAccountName}.queue.core.windows.net/roletotransactionbatch/messages?peekonly=true`
//     this.context.httpClient.get(url, HttpClient.configurations.v1, postOptions)
//       .then((response: HttpClientResponse) => {
//         alert('Request queued');
//           debugger;
//         return;
//       })
//       .catch((error) => {
//         alert('error queuing request');
//         return;
//       });

//     const element: React.ReactElement<IGrcAzureManagementProps> = React.createElement(
//       GrcAzureManagement,
//       {
//         description: this.properties.description
//       }
//     );

//     ReactDom.render(element, this.domElement);
//   }

//   protected get dataVersion(): Version {
//     return Version.parse('1.0');
//   }

//   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
//     return {
//       pages: [
//         {
//           header: {
//             description: strings.PropertyPaneDescription
//           },
//           groups: [
//             {
//               groupName: strings.BasicGroupName,
//               groupFields: [
//                 PropertyPaneTextField('description', {
//                   label: strings.DescriptionFieldLabel
//                 }),
              
//                   PropertyPaneTextField('storageAccountName', {
//                     label: strings.DescriptionFieldLabel
//                   })
//               ]
//             }
//           ]
//         }
//       ]
//     };
//   }
// }
