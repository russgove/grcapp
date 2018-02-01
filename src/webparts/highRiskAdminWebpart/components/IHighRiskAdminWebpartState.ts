import { Web } from "sp-pnp-js";
export interface IHighRiskAdminWebpartState {

  siteName: string; // the name of the web to create
  newWeb: Web; // the web that we created
  newWebUrl:string;
  messages: Array<string>;


  roleToTransactionTotalRows: number;
  roleToTransactionRowsUploaded: number;
  roleToTransactionStatus: string;
  roleToTransactionFile: File;

  highRiskTotalRows: number;
  highRiskRowsUploaded: number;
  highRiskStatus: string;
  highRiskFile: File;

  primaryApproversTotalRows: number;
  primaryApproversRowsUploaded: number;
  primaryApproversStatus: string;
  primaryApproversFile: File;

}
