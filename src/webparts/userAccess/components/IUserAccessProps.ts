import {PrimaryApproverItem,UserAccessItem,RoleToTransaction} from "../datamodel";
export interface IUserAccessProps {
  primaryApproverList: Array<PrimaryApproverItem>;
  save: (userAccess: Array<UserAccessItem>) => Promise<any>;
  getRoleToTransaction: (role: string) => Promise<Array<RoleToTransaction>>;
  fetchUserAccess: () => Promise<Array<UserAccessItem>>;
  setComplete: (PrimaryApproverList: PrimaryApproverItem) => Promise<any>;
  domElement: any; // needed to disable button postback after render on classic pages
}
