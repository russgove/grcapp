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
import { Dropdown, IDropdownOption, IDropdownProps } from "office-ui-fabric-react/lib/Dropdown";
import { Modal, IModalProps } from "office-ui-fabric-react/lib/Modal";
import { Panel, IPanelProps, PanelType } from "office-ui-fabric-react/lib/Panel";
import { CommandBar } from "office-ui-fabric-react/lib/CommandBar";
import { IContextualMenuItem } from "office-ui-fabric-react/lib/ContextualMenu";

import { PrimaryButton, ButtonType, Button, DefaultButton, ActionButton, IconButton } from "office-ui-fabric-react/lib/Button";
import { Dialog } from "office-ui-fabric-react/lib/Dialog";
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
    this.changeUnSelected = this.changeUnSelected.bind(this);
    this.changeSelected = this.changeSelected.bind(this);
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
      alert(err);
    });
  }

  public changeSelected(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {
    debugger;
    var tempArray = map(this.state.mitigatingControls, (rr) => {
      if (this.selection.isKeySelected(rr.Id.toString())) {
        return {
          ...rr,
          Effective: item.data,
          Continues: item.data,
          Right_x0020_Monitor_x003f_: item.data,
          hasBeenUpdated: true,
        };
      }
      else {
        return {
          ...rr
        };
      }
    });
    this.setState((current) => ({ ...current, mitigatingControls: tempArray }));
  }
  public changeUnSelected(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {
    var tempArray = map(this.state.mitigatingControls, (rr) => {
      if (!this.selection.isKeySelected(rr.Id.toString())) {
        return {
          ...rr,
          Effective: item.data,
          Continues: item.data,
          Right_x0020_Monitor_x003f_: item.data,
          hasBeenUpdated: true
        };
      }
      else {
        return {
          ...rr
        };
      }
    });
    this.setState((current) => ({ ...current, mitigatingControls: tempArray }));
  }
  public setComplete(): Promise<any> {
    debugger;
    return this.props.setComplete(this.props.primaryApprover[0]).then(() => {

      alert("Completed");
    }).catch((err) => {
      debugger;
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
        return false
      }
      if (mitigatingControl.Effective === "3") {
        return false
      }
      if (mitigatingControl.Right_x0020_Monitor_x003f_ === "3") {
        return false
      }
      if (!mitigatingControl.Comments) {
        return false
      }
      if (mitigatingControl.Comments.length < 8) {
        return false
      }

    }

    return true;
  }

  public render(): React.ReactElement<IMitigatingControlsProps> {
    debugger;
    let itemsNonFocusable: IContextualMenuItem[] = [
      {
        key: "Change Selected",
        name: "Change Selected",
        icon: "TriggerApproval",
        disabled: this.props.primaryApprover[0].Completed === "Yes",
        subMenuProps: {
          items: [
            {
              key: 'Yes',
              name: 'Yes',
              data: "1",
              onClick: this.changeSelected,
              disabled: this.props.primaryApprover[0].Completed === "Yes"|| this.selection.count < 1
            },
            {
              key: 'No',
              name: 'No',
              data: "2",
              onClick: this.changeSelected,
              disabled: this.props.primaryApprover[0].Completed === "Yes" || this.selection.count < 1
            }
            

          ]
        }
      },
      {
        key: "Change Unselected",
        name: "Change Unselected",
        icon: "TriggerAuto",
        disabled: this.props.primaryApprover[0].Completed === "Yes",

        subMenuProps: {
          items: [
            {
              key: 'Yes',
              name: 'Yes',
              data: "1",
              onClick: this.changeUnSelected,
              disabled: this.props.primaryApprover[0].Completed === "Yes"

            },
            {
              key: 'No',
              name: 'No',
              data: "2",
              onClick: this.changeUnSelected
            },
         

          ]
        }
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

          columns={[
            {
              key: "Risk_x0020_ID", name: "Risk ID",
              fieldName: "Risk_x0020_ID", minWidth: 50,
            },
            {
              key: "Risk_x0020_Description", name: "Risk Description",
              fieldName: "Risk_x0020_Description", minWidth: 100,
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
              fieldName: "Control_x0020_ID", minWidth: 50,
            },
            {
              key: "Description", name: "Control Description",
              fieldName: "Description", minWidth: 100,
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
              key: "Effective", name: "Does the mitigating control effectively remediate the assiciated risk?",
              fieldName: "Effective", minWidth: 150,
              onRender: (item?: any, index?: number, column?: IColumn) => {
                return this.RenderApproval(item, index, column);
              },


            },
            {
              key: "Continues", name: "Does the mitigating control continue to be performed? ",
              fieldName: "Continues", minWidth: 150,
              onRender: (item?: any, index?: number, column?: IColumn) => {
                return this.RenderApproval(item, index, column);
              },

            },
            {
              key: "Right_x0020_Monitor_x003f_", name: "Is the mitigating control monitor the correct person to perform control?",
              fieldName: "Right_x0020_Monitor_x003f_", minWidth: 150,
              onRender: (item?: any, index?: number, column?: IColumn) => {
                return this.RenderApproval(item, index, column);
              },

            },

            {
              key: "Comments", name: "Comments",
              fieldName: "Comments", minWidth: 150, maxWidth: 150,
              onRender: (item?: any, index?: number, column?: IColumn) => { return this.RenderComments(item, index, column); },
            },


          ]}
        />



      </div>
    );
  }
}
