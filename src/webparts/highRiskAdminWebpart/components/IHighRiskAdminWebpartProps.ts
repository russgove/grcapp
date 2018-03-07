import { HttpClient } from '@microsoft/sp-http';
export interface IHighRiskAdminWebpartProps {
  siteUrl: string;
  siteAbsoluteUrl: string;
  templateName: string;
  webPartXml: string; // the webpart to be added to the Home page of the subsite
  httpClient: HttpClient;
  azureHighRiskUrl:string;// url to initiate processing of a roletotransaction File
  azurePrimaryApproverUrl:string;// url to initiate processing of a Primary Approver  File
  azureRoleToCodeUrl:string;// url to initiate processing of a roleto transaction File
  batchSize:number;
  pauseBeforeBatchExecution:number;
  messages?: Array<string>;


}
