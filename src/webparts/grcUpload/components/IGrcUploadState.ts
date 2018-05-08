import { Web } from "@pnp/sp";
export interface IGrcUploadState {

  siteName: string; // the name of the web to create
  newWeb: Web; // the web that we created
  newWebUrl:string;
  messages: Array<string>;


  roleToTransactionTotalRows: number;
  roleToTransactionRowsUploaded: number;
  roleToTransactionStatus: string;
  roleToTransactionFile: File;

  roleReviewTotalRows: number;
  roleReviewRowsUploaded: number;
  roleReviewStatus: string;
  roleReviewFile: File;

  primaryApproversTotalRows: number;
  primaryApproversRowsUploaded: number;
  primaryApproversStatus: string;
  primaryApproversFile: File;

}
