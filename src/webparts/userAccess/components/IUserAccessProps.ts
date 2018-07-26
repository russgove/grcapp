import {PrimaryApproverItem,UserAccessItem,RoleToTransaction} from "../datamodel";
import { HttpClient } from '@microsoft/sp-http';
import { SPUser } from '@microsoft/sp-page-context';
export interface IUserAccessProps {
  user:SPUser; // the sharepoint user accessing the webpart
  webApiUrl:string;
  roleToTcodeController:string;
  primaryApproverController:string;
  userAccessController:string;
  // save: (userAccess: Array<UserAccessItem>) => Promise<any>;
  // getRoleToTransaction: (role: string) => Promise<Array<RoleToTransaction>>;
  // fetchUserAccess: () => Promise<Array<UserAccessItem>>;
  // setComplete: (PrimaryApproverList: PrimaryApproverItem) => Promise<any>;
  domElement: any; // needed to disable button postback after render on classic pages
  httpClient:HttpClient;
}
