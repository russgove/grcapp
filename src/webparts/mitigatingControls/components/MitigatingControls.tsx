import * as React from 'react';
import styles from './MitigatingControls.module.scss';
import { IMitigatingControlsProps } from './IMitigatingControlsProps';
import { IMitigatingControlsState } from './IMitigatingControlsState';
import { escape } from '@microsoft/sp-lodash-subset';
import { MitigatingControlsItem, PrimaryApproverItem } from "../dataModel";
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

export default class MitigatingControls extends React.Component<IMitigatingControlsProps, IMitigatingControlsState> {
  private selection: Selection = new Selection();
  public constructor(props: IMitigatingControlsProps) {
    super();
    console.log("in Construrctor");
    initializeIcons();
    this.selection.getKey = (item => { return item["Id"]; });
    this.save = this.save.bind(this);
    this.setComplete = this.setComplete.bind(this);
    this.updateSelected = this.updateSelected.bind(this);
    this.fetchMitigatingContols = this.fetchMitigatingContols.bind(this);
    this.state = {
      primaryApprover: props.primaryApprover,
      mitigatingControls: [],
      showPopup: false

    };
  }
  public componentDidMount() {
    this.fetchMitigatingContols();
  }
  public fetchMitigatingContols(): Promise<any> {

    return this.props.fetchMitigatingControls().then((mitigatingControls) => {

      this.setState((current) => ({ ...current, mitigatingControls: mitigatingControls }));
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

    var tempArray = map(this.state.mitigatingControls, (rr) => {
      if (this.selection.isKeySelected(rr.Id.toString()) === this.state.changeSelected) {
        return {
          ...rr,
          Effective: this.state.popupValueEffective ? this.state.popupValueEffective : rr.Effective,
          Continues: this.state.popupValueContinues ? this.state.popupValueContinues : rr.Continues,
          Right_x0020_Monitor_x003f_: this.state.popupValueCorrectPerson ? this.state.popupValueCorrectPerson : rr.Right_x0020_Monitor_x003f_,
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
      mitigatingControls: tempArray,
      popupValueEffective: null,
      popupValueContinues: null,
      popupValueCorrectPerson: null,
      popupValueComments: null,
      showPopup: false
    }));
  }


  public setComplete(): Promise<any> {

    return this.props.save(this.state.mitigatingControls).then(() => {
      var tempArray = map(this.state.mitigatingControls, (rr) => {
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
  public save(): Promise<any> {
    return this.props.save(this.state.mitigatingControls).then(() => {
      var tempArray = map(this.state.mitigatingControls, (rr) => {
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
  public RenderComments(item?: MitigatingControlsItem, index?: number, column?: IColumn): JSX.Element {
    if (this.props.primaryApprover[0].Completed === "Yes") {
      return (
        <div>
          {item.Comments}
        </div>
      );
    }
    else {
      return (
        <TextField
          value={item.Comments}
          onChanged={(newValue) => {
            let tempRoleToTCodeReview = this.state.mitigatingControls;
            let rtc = find(tempRoleToTCodeReview, (x) => {
              return x.Id === item.Id;
            });
            rtc.Comments = newValue;
            rtc.hasBeenUpdated = true;
            this.setState((current) => ({ ...current, roleToTCodeReview: tempRoleToTCodeReview, changesHaveBeenMade: true }));

          }}

        >

        </TextField>
      );
    }
  }
  public RenderApproval(item?: MitigatingControlsItem, index?: number, column?: IColumn): JSX.Element {
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
        <Dropdown options={options}
          selectedKey={item[column.fieldName]}
          onChanged={(option: IDropdownOption, idx?: number) => {
            let tempTable = this.state.mitigatingControls;

            let rtc = find(tempTable, (x) => {
              return x.Id === item.Id;
            });
            rtc[column.fieldName] = option.key as string;
            rtc.hasBeenUpdated = true;
            this.setState((current) => ({ ...current, mitigatingControls: tempTable, changesHaveBeenMade: true }));
          }}
        >
        </Dropdown>
      );
    }
  }
  private areAllQuestionsAnswered(): boolean {
    for (let mitigatingControl of this.state.mitigatingControls) {
      if (mitigatingControl.Continues === "3") {
        return false;
      }
      if (mitigatingControl.Effective === "3") {
        return false;
      }
      if (mitigatingControl.Right_x0020_Monitor_x003f_ === "3") {
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

  public render(): React.ReactElement<IMitigatingControlsProps> {

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
          if (!this.selection.count || this.selection.count < this.state.mitigatingControls.length) {
            this.setState((current) => ({
              ...current,
              showPopup: true,
              changeSelected: false // change UNSELECTED Items
            }));
          }
        },

      },

      {
        key: "Undo", name: "Undo", icon: "Undo", onClick: this.fetchMitigatingContols,
        disabled: !(filter(this.state.mitigatingControls, (rr) => {
          return rr
            .hasBeenUpdated;
        }).length > 0)
      },
      { // if the item has been comleted OR there are items with noo approvasl, diable
        key: "Done", name: "Complete", icon: "Completed", onClick: this.setComplete,
        disabled: this.props.primaryApprover[0].Completed === "Yes" || !(this.areAllQuestionsAnswered())

      }

    ];
    let farItemsNonFocusable: IContextualMenuItem[] = [
      {

        key: "Save", name: "Save", icon: "Save", onClick: this.save,
        disabled: !(filter(this.state.mitigatingControls, (rr) => { return rr.hasBeenUpdated; }).length > 0)
        || this.props.primaryApprover[0].Completed === "Yes"

      }
    ];


    return (
      <div className={styles.mitigatingControls}>
        <Dialog isBlocking={true}
          hidden={!this.state.showPopup}
          onDismiss={(e) => { this.setState((current) => ({ ...current, showPopup: false })); }}
          dialogContentProps={{
            type: DialogType.close,
            title: this.state.changeSelected
              ? `Update ${this.selection.count} Selected Items`
              : this.selection.count
                ? `Update ${this.state.mitigatingControls.length - this.selection.count} Unselected Items`
                : `Update ${this.state.mitigatingControls.length} Unselected Items`,
            subText: 'All selected items will be updated with the following values'
          }} >
          <ChoiceGroup label={this.props.effectiveLabel}
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
            selectedKey={this.state.popupValueEffective}

            onChange={(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption) => {

              this.setState((current) => ({ ...current, popupValueEffective: option.key }));
            }}
          />
          <ChoiceGroup label={this.props.continuesLabel}
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
            selectedKey={this.state.popupValueContinues}
            onChange={(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption) => {

              this.setState((current) => ({ ...current, popupValueContinues: option.key }));
            }}
          />
          <ChoiceGroup label={this.props.correctPersonLabel}
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
            selectedKey={this.state.popupValueCorrectPerson}
            onChange={(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption) => {

              this.setState((current) => ({ ...current, popupValueCorrectPerson: option.key }));
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
          items={this.state.mitigatingControls}
          selectionMode={SelectionMode.multiple}
          selection={this.selection}
          key="Risk_x0020_ID"
          layoutMode={DetailsListLayoutMode.fixedColumns}

          columns={[
            {
              key: "Risk_x0020_ID", name: "Risk ID",
              fieldName: "Risk_x0020_ID", minWidth: 50,
            },
            {
              key: "Risk_x0020_Description", name: "Risk Description",
              fieldName: "Risk_x0020_Description", minWidth: 170,
              onRender: (item?: MitigatingControlsItem, index?: number, column?: IColumn) => {
                return (
                  <div className={styles.riskDesription}>
                    {item.Risk_x0020_Description}
                  </div>
                );
              },
            },
            {
              key: "Control_x0020_ID", name: "Control ID",
              fieldName: "Control_x0020_ID", minWidth: 65,
            },
            {
              key: "Description", name: "Control Description",
              fieldName: "Description", minWidth: 170,
              onRender: (item?: MitigatingControlsItem, index?: number, column?: IColumn) => {
                return (
                  <div className={styles.controlDesription}>
                    {item.Description}
                  </div>
                );
              },

            },
            {
              key: "Owner_x0020_ID", name: "Owner ID/Name",
              fieldName: "Owner_x0020_ID", minWidth: 100,
              onRender: (item?: MitigatingControlsItem, index?: number, column?: IColumn) => {

                return (
                  <div  >
                    {item.Owner_x0020_ID}
                    <br />
                    {item.Control_x0020_Owner_x0020_Name}
                  </div>
                );
              },


            },
            {
              key: "Control_x0020_Monitor_x0020_ID", name: "Monitor ID/Name",
              fieldName: "Control_x0020_Monitor_x0020_ID", minWidth: 100,
              onRender: (item?: MitigatingControlsItem, index?: number, column?: IColumn) => {

                return (
                  <div  >
                    {item.Control_x0020_Monitor_x0020_ID}
                    <br />
                    {item.Control_x0020_Monitor_x0020_Name}
                  </div>
                );
              },

            },
            {
              key: "Effective", name: this.props.effectiveLabel,
              fieldName: "Effective", minWidth: 150,
              onRender: (item?: any, index?: number, column?: IColumn) => {
                return this.RenderApproval(item, index, column);
              },


            },
            {
              key: "Continues", name: this.props.continuesLabel,
              fieldName: "Continues", minWidth: 150,
              onRender: (item?: any, index?: number, column?: IColumn) => {
                return this.RenderApproval(item, index, column);
              },

            },
            {
              key: "Right_x0020_Monitor_x003f_", name: this.props.correctPersonLabel,
              fieldName: "Right_x0020_Monitor_x003f_", minWidth: 150,
              onRender: (item?: any, index?: number, column?: IColumn) => {
                return this.RenderApproval(item, index, column);
              },

            },

            {
              key: "Comments", name: "Comments",
              fieldName: "Comments", minWidth: 150,
              onRender: (item?: any, index?: number, column?: IColumn) => { return this.RenderComments(item, index, column); },
            },


          ]}
        />



      </div>
    );
  }
}
