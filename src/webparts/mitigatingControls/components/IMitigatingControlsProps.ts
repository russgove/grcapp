import { MitigatingControlsItem, PrimaryApproverItem,HelpLink } from "../dataModel";
import { HttpClient } from '@microsoft/sp-http';
import { SPUser } from '@microsoft/sp-page-context';
export interface IMitigatingControlsProps {
  webApiUrl:string;
  mitigatngControlsController:string;
  primaryApproverController:string;
  user:SPUser;
  domElement: any; // needed to disable button postback after render on classic pages
  effectiveLabel:string;
  continuesLabel:string;
  correctPersonLabel:string;
  helpLinks:Array<HelpLink>;
  httpClient:HttpClient;

}
