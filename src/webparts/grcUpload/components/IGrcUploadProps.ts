export interface IGrcUploadProps {
  siteUrl:string;
  templateName:string;
  primaryApproverContentTypeId:string;
  roleToTransactionContentTypeId:string;
  roleReviewContentTypeId:string;
  webPartXml: string; // the webpart to be added to the Home page of the subsite
 
}