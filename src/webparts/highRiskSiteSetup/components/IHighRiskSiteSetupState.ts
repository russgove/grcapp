import { Dropdown, IDropdownOption, IDropdownProps } from "office-ui-fabric-react/lib/Dropdown";
import {List} from "@pnp/sp";
export interface IHighRiskSiteSetupState {
  webName: string;
  webUrl: string;
  siteDropDownOptions: Array<IDropdownOption>;
  highRiskListExists: boolean;
  highRiskList?: List;
  highRiskCount?: number;
  highRiskFieldsFound?: boolean;

  primaryApproversListExists: boolean; //does the list exist?
  primaryApproversList?: List; //does the list exist?
  primaryApproversCount?: number; // the number of rows in the list
  primaryApproversFieldsFound?: boolean;// are the required fields present in the list
  messages: Array<string>;
}
