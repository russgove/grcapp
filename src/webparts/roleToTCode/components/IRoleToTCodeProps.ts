import {PrimaryApproverItem,RoleReviewItem,RoleToTransaction} from "../datamodel";
import { AadHttpClient } from '@microsoft/sp-http';
import { SPUser } from '@microsoft/sp-page-context';
export interface IRoleToTCodeProps {
  azureFunctionUrl:string;
  user:SPUser; // the sharepoint user accessing the webpart
  domElement: any; // needed to disable button postback after render on classic pages
  aadHttpClient:AadHttpClient;
}
