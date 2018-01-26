import { Web } from "sp-pnp-js";
export interface IGrcUploadState {

  siteName: string; // the name of the web to create
  newWeb: Web; // the web that we created
  messages: Array<string>;
  process: "" | "Uploading" | "Validating";

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
