import {BusinessRoleReviewItem, PrimaryApproverItem ,HelpLink} from "../dataModel";
import { HttpClient } from '@microsoft/sp-http';
import { SPUser } from '@microsoft/sp-page-context';
export interface IBusinessRoleReviewProps {
  azureFunctionUrl:string;
  accessCode:string;
  user:SPUser; // the sharepoint user accessing the webpart
  domElement: any; // needed to disable button postback after render on classic pages
  httpClient:HttpClient;
  businessRoleOwnersPath:string;
  primaryApproversPath:string;

  enableUncomplete:boolean; // should we show menu item to uncomplete(for use in testing)
  roleNameWidth:number;
  approverWidth:number;
  altApproverWidth:number;
  approvalDecisionWidth:number;
  commentsWidth:number;
  helpLinks:Array<HelpLink>;

}
