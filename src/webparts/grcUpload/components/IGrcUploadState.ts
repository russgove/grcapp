export interface IGrcUploadState {

  siteName:string;
  messages:Array<string>;
  process:""|"Uploading"|"Validating";

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
