import { Item } from "sp-pnp-js";
// main list for grid
export default class HighRisk extends Item {

    public Id: number;
    public GRCUserId: string;
    public GRCUserFullName: string;
    public GRCRoleName: string;
    public GRCApproverId: string;
    public GRCApprover: {
        Title: string;
    };
    public GRCAlternateApproverId: string;
    public GRCAlternateApprover: {
        Title: string;
    };
    public GRCApproval: string;
    public GRCApprovedBy: string;
    public GRCDateReviewed: string;
    public GRCComments: string;
    public GRCRemediation: string;
    public hasBeenUpdated: boolean;

}