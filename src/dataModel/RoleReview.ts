import { Item } from "@pnp/sp";
// main list for grid
export default class RoleReview extends Item {

    public Id:number;
    public GRCRoleName:string;    
    public GRCApproverId:string;
    public GRCApprover: {
        Title: string;
    };
    public GRCAlternateApproverId:string;
    public GRCAlternateApprover: {
        Title: string;
    };
    public GRCApproval:string;
    public GRCApprovedBy:string;
    public GRCDateReviewed:string;
    public GRCComments:string;
    public GRCRemediation:string;
    public hasBeenUpdated:boolean;
   
}