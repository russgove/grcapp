import { HttpClient} from '@microsoft/sp-http';
export interface IGrcUploadProps {
  siteUrl:string;
  siteAbsoluteUrl:string;
  templateName:string;
  primaryApproverContentTypeId:string;
  roleToTransactionContentTypeId:string;
  roleReviewContentTypeId:string;
  webPartXml: string; // the webpart to be added to the Home page of the subsite
  azureFunctionUrl:string;// the url of the azure function we post to to kick off the webjob
 httpClient:HttpClient;
}