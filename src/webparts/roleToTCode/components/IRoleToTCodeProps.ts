import {PrimaryApproverItem,RoleReviewItem,RoleToTransaction} from "../datamodel";
import { HttpClient } from '@microsoft/sp-http';
import { SPUser } from '@microsoft/sp-page-context';
export interface IRoleToTCodeProps {
  webApiUrl:string;
  roleToTcodeController:string;
  primaryApproverController:string;
  roleReviewController:string;
  user:SPUser; // the sharepoint user accessing the webpart
  domElement: any; // needed to disable button postback after render on classic pages
  httpClient:HttpClient;
}
