import { Item } from "sp-pnp-js";
// tassks to be approved
export default class PrimaryApproverList extends Item {

        public Id:number;
    public GRCCompleted: string;
    public GRCApproverId: number;
    public GRCApprover: {
        Title: string;

    };
}