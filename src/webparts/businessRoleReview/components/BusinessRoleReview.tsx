import * as React from 'react';
import styles from './BusinessRoleReview.module.scss';
import { IBusinessRoleReviewProps } from './IBusinessRoleReviewProps';
import { IBusinessRoleReviewState } from './IBusinessRoleReviewState';
import { escape } from '@microsoft/sp-lodash-subset';
import { BusinessRoleReviewItem, PrimaryApproverItem } from "../dataModel";
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import {
  DetailsList, DetailsListLayoutMode, IColumn, SelectionMode, Selection,
  ColumnActionsMode
} from "office-ui-fabric-react/lib/DetailsList";
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { Dropdown, IDropdownOption, IDropdownProps } from "office-ui-fabric-react/lib/Dropdown";
import { Modal, IModalProps } from "office-ui-fabric-react/lib/Modal";
import { Panel, IPanelProps, PanelType } from "office-ui-fabric-react/lib/Panel";
import { CommandBar } from "office-ui-fabric-react/lib/CommandBar";
import { IContextualMenuItem } from "office-ui-fabric-react/lib/ContextualMenu";

import { PrimaryButton, ButtonType, Button, DefaultButton, ActionButton, IconButton } from "office-ui-fabric-react/lib/Button";
import { Dialog, DialogType, DialogContent, DialogFooter } from "office-ui-fabric-react/lib/Dialog";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { find, map, clone, filter } from "lodash";


export default class BusinessRoleReview extends React.Component<IBusinessRoleReviewProps, IBusinessRoleReviewState> {
  private selection: Selection = new Selection();
  public constructor(props: IBusinessRoleReviewProps) {
    super();
    console.log("in Construrctor");
    initializeIcons();
    this.selection.getKey = (item => { return item["Id"]; });
    this.save = this.save.bind(this);
    this.setComplete = this.setComplete.bind(this);
    this.updateSelected = this.updateSelected.bind(this);
    this.fetchBusinessRoleReview = this.fetchBusinessRoleReview.bind(this);
    this.state = {
      primaryApprover: props.primaryApprover,
      businessRoleReview: [],
      showPopup: false

    };
  }
  public componentDidMount() {
    this.fetchBusinessRoleReview();
  }
  public fetchBusinessRoleReview(): Promise<any> {

    return this.props.fetchBusinessRoleReview().then((items) => {
      debugger;
      this.setState((current) => ({ ...current, businessRoleReview: items }));
    }).catch((err) => {
      debugger;
      console.error(err);
      alert(err);
    });
  }
  public showUpdatedSelectedPopup(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {

    if (this.selection.count > 0) {
      this.setState((current) => ({ ...current, showPopup: true }));
    }
  }
  public updateSelected(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {
    debugger;
    var tempArray = map(this.state.businessRoleReview, (rr) => {
      if (this.selection.isKeySelected(rr.Id.toString()) === this.state.changeSelected) {
        return {
          ...rr,
          Approval: this.state.popupValueApproval?this.state.popupValueApproval:rr.Approval,
          Comments: this.state.popupValueComments ? this.state.popupValueComments : rr.Comments,
          hasBeenUpdated: true,
        };
      }
      else {
        return {
          ...rr
        };
      }
    });
    this.setState((current) => ({
      ...current,
      businessRoleReview: tempArray,
      popupValueApproval: null,
      popupValueComments: null,
      showPopup: false
    }));
  }


  public setComplete(): Promise<any> {

    return this.props.save(this.state.businessRoleReview).then(() => {
      var tempArray = map(this.state.businessRoleReview, (rr) => {
        return { ...rr, hasBeenUpdated: false };
      });
      this.setState((current) => ({ ...current, mitigatingControls: tempArray }));
      return this.props.setComplete(this.props.primaryApprover[0]).then(() => {
        alert("Completed");
      }).catch((err) => {
        debugger;
        console.error(err);
        alert(err);
      });
    }).catch((err) => {
      debugger;
      console.error(err);
      alert(err);
    });

  }
  public save( ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem) : void{
     this.props.save(this.state.businessRoleReview).then(() => {
      var tempArray = map(this.state.businessRoleReview, (rr) => {
        return { ...rr, hasBeenUpdated: false };
      });
      this.setState((current) => ({ ...current, mitigatingControls: tempArray }));
      alert("Saved");
    }).catch((err) => {
      debugger;
      console.error(err);
      alert(err);
    });
  }
  public RenderComments(item?: BusinessRoleReviewItem, index?: number, column?: IColumn): JSX.Element {
    if (this.props.primaryApprover[0].Completed === "Yes") {
      return (
        <div>
          {item.Comments}
        </div>
      );
    }
    else {
      return (
        <TextField key="Comments"

          value={item.Comments?item.Comments:""}
          onChanged={(newValue) => {
            let temp = this.state.businessRoleReview;
            let rtc = find(temp, (x) => {
              return x.Id === item.Id;
            });
            rtc.Comments = newValue;
            rtc.hasBeenUpdated = true;
            this.setState((current) => ({ ...current, businessRoleReview: temp, changesHaveBeenMade: true }));

          }}

        >

        </TextField>
      );
    }
  }
  public RenderApproval(item?: BusinessRoleReviewItem, index?: number, column?: IColumn): JSX.Element {
    let options = [
      { key: "1", text: "Yes" },
      { key: "2", text: "No" },

    ];
    if (this.props.primaryApprover[0].Completed === "Yes") {
      return (
        <div>
          {find(options, (o) => { return o.key === item[column.fieldName]; }).text}
        </div>
      );
    }
    else {
      return (
        <ChoiceGroup label="" 
        options={[
          {
            key: '1',
            text: 'Yes'
          },
          {
            key: '2',
            text: 'No',
          },
        ]}
        selectedKey={item[column.fieldName]}

           onChanged={(option: IChoiceGroupOption, event: any) => {
             let tempTable = this.state.businessRoleReview;

             let rtc = find(tempTable, (x) => {
               return x.Id === item.Id;
             });
             rtc[column.fieldName] = option.key as string;
             rtc.hasBeenUpdated = true;
             this.setState((current) => ({ ...current, mitigatingControls: tempTable, changesHaveBeenMade: true }));
           }}
      />
        // <Dropdown options={options}
        //   selectedKey={item[column.fieldName]}
        //   onChanged={(option: IDropdownOption, idx?: number) => {
        //     let tempTable = this.state.businessRoleReview;

        //     let rtc = find(tempTable, (x) => {
        //       return x.Id === item.Id;
        //     });
        //     rtc[column.fieldName] = option.key as string;
        //     rtc.hasBeenUpdated = true;
        //     this.setState((current) => ({ ...current, mitigatingControls: tempTable, changesHaveBeenMade: true }));
        //   }}
        // >
        // </Dropdown>
      );
    }
  }
  private areAllQuestionsAnswered(): boolean {
    for (let mitigatingControl of this.state.businessRoleReview) {
  
      if (mitigatingControl.Approval === "3") {
        return false;
      }
      if (!mitigatingControl.Comments) {
        return false;
      }
      if (mitigatingControl.Comments.length === 0) {
        return false;
      }

    }

    return true;
  }
  public render(): React.ReactElement<IBusinessRoleReviewProps> {
   
    let itemsNonFocusable: IContextualMenuItem[] = [
      {
        key: "Update Selected",
        name: "Update Selected",
        icon: "TriggerApproval",
        onClick: (e) => {
          if (this.selection.count > 0) {
            this.setState((current) => ({
              ...current,
              showPopup: true,
              changeSelected: true
            }));
          }
        },
        disabled: this.props.primaryApprover[0].Completed === "Yes"

      },
      {
        key: "Update Unselected",
        name: "Update Unselected",
        icon: "TriggerAuto",
        disabled: this.props.primaryApprover[0].Completed === "Yes",
        onClick: (e) => {
          debugger;
          if (!this.selection.count || this.selection.count < this.state.businessRoleReview.length) {
            this.setState((current) => ({
              ...current,
              showPopup: true,
              changeSelected: false // change UNSELECTED Items
            }));
          }
        },

      },

      {
        key: "Undo", name: "Undo", icon: "Undo", onClick: this.fetchBusinessRoleReview,
        disabled: !(filter(this.state.businessRoleReview, (rr) => {
          return rr
            .hasBeenUpdated;
        }).length > 0)
      },
      { // if the item has been comleted OR there are items with noo approvasl, diable
        key: "Done", name: "Task Complete", icon: "Completed", onClick: this.setComplete,
        disabled: this.props.primaryApprover[0].Completed === "Yes" || !(this.areAllQuestionsAnswered())

      }

    ];
    let farItemsNonFocusable: IContextualMenuItem[] = [
      {

        key: "Save", name: "Save", icon: "Save", onClick: this.save,
        disabled: !(filter(this.state.businessRoleReview, (rr) => { return rr.hasBeenUpdated; }).length > 0)
        || this.props.primaryApprover[0].Completed === "Yes"

      },
      {
        key: "helpLinks", name: "Help", icon: "help",
        items: map(this.props.helpLinks,(hl):IContextualMenuItem=>{
          debugger;
          return{
            key:hl.Id.toString(), // this is the id of the listitem
            href:hl.Url.Url, 
            title:hl.Url.Description,
            icon:hl.IconName,
            name:hl.Title,
            target:hl.Target

          };
        })
      }
    ];


    return (
      <div className={styles.businessRoleReview}>
        <Dialog isBlocking={true}
          hidden={!this.state.showPopup}
          onDismiss={(e) => { this.setState((current) => ({ ...current, showPopup: false })); }}
          dialogContentProps={{
            type: DialogType.close,
            title: this.state.changeSelected
              ? `Update ${this.selection.count} Selected Items`
              : this.selection.count
                ? `Update ${this.state.businessRoleReview.length - this.selection.count} Unselected Items`
                : `Update ${this.state.businessRoleReview.length} Unselected Items`,
            subText: 'All selected items will be updated with the following values'
          }} >
          <ChoiceGroup label="Approval Decision" 
            options={[
              {
                key: '1',
                text: 'Yes'
              },
              {
                key: '2',
                text: 'No',
              },
            ]}
            selectedKey={this.state.popupValueApproval}

            onChange={(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption) => {

              this.setState((current) => ({ ...current, popupValueApproval: option.key }));
            }}
          />
      
          <TextField label="Comments" onChanged={(e) => {

            this.setState((current) => ({ ...current, popupValueComments: e }));
          }}

          />



          <DialogFooter>
            <PrimaryButton text='Save' onClick={this.updateSelected.bind(this)} />
            <DefaultButton text='Cancel' onClick={(e) => {
              this.setState((current) => ({ ...current, showPopup: false }));
            }} />
          </DialogFooter>
        </Dialog>
        <CommandBar
          isSearchBoxVisible={false}
          items={itemsNonFocusable}
          farItems={farItemsNonFocusable}

        />
        <DetailsList
          items={this.state.businessRoleReview}
          selectionMode={SelectionMode.multiple}
          selection={this.selection}
          key="Risk_x0020_ID"
          layoutMode={DetailsListLayoutMode.justified}
          columns={[
            {
              key: "Role_x0020_Name", name: "Role Name / Composite Role",
              fieldName: "Role_x0020_Name", 
              minWidth: this.props.roleNameWidth,
              isResizable:true,
              onRender: (item?: BusinessRoleReviewItem, index?: number, column?: IColumn) => {
                return (
                  <div  >
                    {item.Role_x0020_Name}
                    <br />
                    {item.Composite_x0020_Role}
                  </div>
                );
              },
            },
           
            // {
            //   key: "Composite_x0020_Role", name: "Composite Role",
            //   fieldName: "Composite_x0020_Role", minWidth: this.props.compositeRoleWidth,
            // },
            {
              key: "Approver", name: "Approver ID/Name ",
              fieldName: "Approver", 
              minWidth: this.props.approverWidth,
              isResizable:true,
              onRender: (item?: BusinessRoleReviewItem, index?: number, column?: IColumn) => {
                return (
                  <div  >
                    {item.Approver}
                    <br />
                    {item.Approver_x0020_Name}
                  </div>
                );
              },
            },
            {
              key: "AlternateApprover", name: "Alt. Apprv. ID/Name",
              fieldName: "Composite_x0020_Role",
              minWidth: this.props.altApproverWidth,
              isResizable:true,
              onRender: (item?: BusinessRoleReviewItem, index?: number, column?: IColumn) => {
                return (
                  <div  >
                    {item.Alt_x0020_Apprv}
                    <br />
                    {item.Alternate_x0020_Approver}
                  </div>
                );
              },
            },
            
            {
              key: "Approval", name: "Approval Decision",
              fieldName: "Approval",
              minWidth: this.props.approvalDecisionWidth,
              isResizable:true,
              onRender: (item?: any, index?: number, column?: IColumn) => {
                return this.RenderApproval(item, index, column);
              },

            },
            {
              key: "Comments", name: "Comments",
              fieldName: "Comments",
              minWidth: this.props.commentsWidth,
              isResizable:true,
              onRender: (item?: any, index?: number, column?: IColumn) => { 
                return this.RenderComments(item, index, column);
               },
            },


          ]}
        />



      </div>
    );
  }
}
