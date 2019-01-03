
import { HttpClient } from '@microsoft/sp-http';
import { SPUser } from '@microsoft/sp-page-context';
export interface IHighRiskFunctionsProps {
  azureFunctionUrl:string;
  accessCode:string;
  user:SPUser; // the sharepoint user accessing the webpart
  domElement: any; // needed to disable button postback after render on classic pages
  httpClient:HttpClient;
  highRiskFunctionsPath:string;

  primaryApproversPath:string;
  roleToTransactionsPath:string;
  enableUncomplete:boolean; // should we show menu item to uncomplete(for use in testing)
}
