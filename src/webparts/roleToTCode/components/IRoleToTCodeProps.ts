import {PrimaryApproverItem,RoleReviewItem,RoleToTransaction} from "../datamodel";
import { HttpClient } from '@microsoft/sp-http';
import { SPUser } from '@microsoft/sp-page-context';
export interface IRoleToTCodeProps {
  GetPrimaryApproverByEmailPath:string;
  UpdatePrimaryApproversPath: string;
  GetRoleReviewsForApproverPath: string;
  UpdateRoleReviewsForApproverPath: string;
  GetRoleToTransactionsForRoleNamePath: string;

  user:SPUser; // the sharepoint user accessing the webpart
  domElement: any; // needed to disable button postback after render on classic pages
  httpClient:HttpClient;
}
