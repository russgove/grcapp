import { Dropdown, IDropdownOption, IDropdownProps } from "office-ui-fabric-react/lib/Dropdown";
import {List} from "@pnp/sp";
export interface IUserAccessSiteSetupState{
  webName: string;
  webUrl: string;
  siteDropDownOptions: Array<IDropdownOption>;
  
  tcodeListExists: boolean;
  tcodeList?: List;
  tcodeCount?: number;
  tcodeFieldsFound?: boolean;


  userAccessListExists: boolean;
  userAccessList?: List;
  userAccessCount?: number;
  userAccessFieldsFound?: boolean;

  primaryApproversListExists: boolean; //does the list exist?
  primaryApproversList?: List; //does the list exist?
  primaryApproversCount?: number; // the number of rows in the list
  primaryApproversFieldsFound?: boolean;// are the required fields present in the list
  messages: Array<string>;
}
