import { MitigatingControlsItem, PrimaryApproverItem,HelpLink } from "../dataModel";
export interface IMitigatingControlsProps {
  primaryApprover: Array<PrimaryApproverItem>;
  save: (mitigatingControls: Array<MitigatingControlsItem>) => Promise<any>;
  fetchMitigatingControls: () => Promise<Array<MitigatingControlsItem>>;
  setComplete: (PrimaryApprover: PrimaryApproverItem) => Promise<any>;
  domElement: any; // needed to disable button postback after render on classic pages
  effectiveLabel:string;
  continuesLabel:string;
  correctPersonLabel:string;
  helpLinks:Array<HelpLink>;

}
