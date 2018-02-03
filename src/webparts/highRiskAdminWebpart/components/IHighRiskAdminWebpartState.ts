import { Web } from "sp-pnp-js";
export interface IHighRiskAdminWebpartState {

  siteName: string; // the name of the web to create
  newWeb: Web; // the web that we created
  newWebUrl:string;
  messages: Array<string>;


  roleToTransactionTotalRows: number;
  roleToTransactionStatus: string;
  roleToTransactionFile: File;

  highRiskTotalRows: number;
  highRiskStatus: string;
  highRiskFile: File;

  primaryApproversTotalRows: number;
  primaryApproversStatus: string;
  primaryApproversFile: File;

}
