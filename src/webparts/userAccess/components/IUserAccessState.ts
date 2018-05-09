import {PrimaryApproverItem,UserAccessItem,RoleToTransaction} from "../datamodel"

export interface IUserAccessState {
  primaryApproverList: Array<PrimaryApproverItem>;
  userAccessItems: Array<UserAccessItem>;
  roleToTransaction?: Array<RoleToTransaction>;
  showPopup:boolean;
}
