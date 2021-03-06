import * as React from 'react';
import styles from './GrcTest.module.scss';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { IGrcTestProps } from './IGrcTestProps';
import { IGrcTestState } from './IGrcTestState';
import { escape } from '@microsoft/sp-lodash-subset';
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
import PrimaryApproverList from "../../../dataModel/PrimaryApproverList";
import RoleReview from "../../../dataModel/RoleReview";
import RoleToTransaction from "../../../dataModel/RoleToTransaction";
import { find, map, clone, filter } from "lodash";
export default class GrcTest extends React.Component<IGrcTestProps, IGrcTestState> {
  private selection: Selection = new Selection();

  public constructor(props: IGrcTestProps) {
    super();
    console.log("in Construrctor");

    initializeIcons();
    this.selection.getKey = (item => { return item["Id"]; });
    this.save = this.save.bind(this);
    this.setComplete = this.setComplete.bind(this);
    this.changeAll = this.changeAll.bind(this);
    this.changeSelected = this.changeSelected.bind(this);
    this.fetchRoleReviews = this.fetchRoleReviews.bind(this);
    this.state = {
      primaryApproverList: props.primaryApproverList,
      roleReview: [],
      showPopup: false

    };
  }
  public componentDidMount() {
    this.props.fetchRoleReviews().then(reviews => {
      debugger;
      this.setState((current) => ({ ...current, roleReview: reviews, changesHaveBeenMade: false }));

    });
  }
  public componentDidUpdate(): void {
    // disable postback of buttons. see https://github.com/SharePoint/sp-dev-docs/issues/492
    if (Environment.type === EnvironmentType.ClassicSharePoint) {
      const buttons: NodeListOf<HTMLButtonElement> = this.props.domElement.getElementsByTagName('button');
      for (let i: number = 0; i < buttons.length; i++) {
        if (buttons[i]) {
          // Disable the button onclick postback
          buttons[i].onclick = ()=> {
            return false;
          };
        }
      }
    }
  }
  public RenderApproval(item?: RoleReview, index?: number, column?: IColumn): JSX.Element {

    let options = [
      { key: "0", text: "yup" },
      { key: "1", text: "nope" },
      { key: "2", text: "no f'in way" }
    ];
    if (this.props.primaryApproverList[0].GRCCompleted === "Yes") {
      return (
        <div>
          {find(options, (o) => { return o.key === item.GRCApproval; }).text}
        </div>
      );
    }
    else {
      return (
        <Dropdown options={options}
          selectedKey={item.GRCApproval}
          onChanged={(option: IDropdownOption, idx?: number) => {
            let tempRoleToTCodeReview = this.state.roleReview;

            let rtc = find(tempRoleToTCodeReview, (x) => {
              return x.Id === item.Id;
            });
            rtc.GRCApproval = option.key as string;
            rtc.hasBeenUpdated = true;
            this.setState((current) => ({ ...current, roleToTCodeReview: tempRoleToTCodeReview, changesHaveBeenMade: true }));

          }}

        >

        </Dropdown>
      );
    }


  }
  public RenderComments(item?: RoleReview, index?: number, column?: IColumn): JSX.Element {
    if (this.props.primaryApproverList[0].GRCCompleted === "Yes") {
      return (
        <div>
          {item.GRCComments}
        </div>
      );
    }
    else {
      return (
        <TextField
          value={item.GRCComments}
          onChanged={(newValue) => {
            let tempRoleToTCodeReview = this.state.roleReview;
            let rtc = find(tempRoleToTCodeReview, (x) => {
              return x.Id === item.Id;
            });
            rtc.GRCComments = newValue;
            rtc.hasBeenUpdated = true;
            this.setState((current) => ({ ...current, roleToTCodeReview: tempRoleToTCodeReview, changesHaveBeenMade: true }));

          }}

        >

        </TextField>
      );
    }


  }
  public save( ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem) : void{
     this.props.save(this.state.roleReview).then(() => {
      debugger;
      var roleReview = map(this.state.roleReview, (rr) => {
        return { ...rr, hasBeenUpdated: false };
      });
      this.setState((current) => ({ ...current, roleReview: roleReview }));
      alert("Saved");
    }).catch((err) => {
      debugger;
      alert(err);
    });
  }
  public setComplete(): Promise<any> {
    return this.props.setComplete(this.props.primaryApproverList[0]).then(() => {

      alert("Completed");
    }).catch((err) => {
      debugger;
      alert(err);
    });
  }
  public changeSelected(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {

    var roleReview = map(this.state.roleReview, (rr) => {
      if (this.selection.isKeySelected(rr.Id.toString())) {
        return {
          ...rr,
          GRCApproval: item.data, hasBeenUpdated: true
        };

      }
      else {
        return {
          ...rr
        };
      }
    });
    this.setState((current) => ({ ...current, roleReview: roleReview }));



  }
  public changeAll(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {

    var roleReview = map(this.state.roleReview, (rr) => {
      return { ...rr, GRCApproval: item.data, hasBeenUpdated: true };
    });
    this.setState((current) => ({ ...current, roleReview: roleReview }));
    // for (let roleReview of copy){
    //   roleReview.GRCApproval=item.data;
    //   roleReview.hasBeenUpdated=true;
    // }


  }
  public showPopup(item: RoleReview) {

    this.props.getRoleToTransaction(item.GRCRoleName)
      .then((roleToTransactions) => {

        this.setState((current) => ({ ...current, roleToTransaction: roleToTransactions, showPopup: true }));

      }).catch((err) => {
        console.error(err.data.responseBody["odata.error"].message.value);
        alert(err.data.responseBody["odata.error"].message.value);

        debugger;
      });
  }

  public fetchRoleReviews(): Promise<any> {
    debugger;
    return this.props.fetchRoleReviews().then((roleReviews) => {
      debugger;
      this.setState((current) => ({ ...current, roleReview: roleReviews }));
    }).catch((err) => {
      debugger;
      alert(err);
    });
  }

  public render(): React.ReactElement<IGrcTestProps> {

    let itemsNonFocusable: IContextualMenuItem[] = [
      //{
      //   key: "Change All",
      //   name: "Change All",
      //   icon: "TriggerAuto",
      //   disabled: this.props.primaryApproverList[0].GRCCompleted === "Yes",
      //   subMenuProps: {
      //     items: [
      //       {
      //         key: 'yup',
      //         name: 'Yup',
      //         data: "0",
      //         onClick: this.changeAll,
      //         disabled: this.props.primaryApproverList[0].GRCCompleted === "Yes"
      //       },
      //       {
      //         key: 'Nope',
      //         name: 'Nope',
      //         data: "1",
      //         onClick: this.changeAll
      //       },
      //       {
      //         key: "no f'in way",
      //         name: "no f'in way",
      //         data: "2",
      //         onClick: this.changeAll

      //       }

      //     ]
      //   }
      // },
      // {
      //   key: "Change Selected",
      //   name: "Change Selected",
      //   icon: "TriggerApproval",
      //   disabled: this.props.primaryApproverList[0].GRCCompleted === "Yes",
      //   subMenuProps: {
      //     items: [
      //       {
      //         key: 'yup',
      //         name: 'Yup',
      //         data: "0",
      //         onClick: this.changeSelected,
      //         disabled: this.props.primaryApproverList[0].GRCCompleted === "Yes"
      //       },
      //       {
      //         key: 'Nope',
      //         name: 'Nope',
      //         data: "1",
      //         onClick: this.changeSelected,
      //         disabled: this.props.primaryApproverList[0].GRCCompleted === "Yes"
      //       },
      //       {
      //         key: "no f'in way",
      //         name: "no f'in way",
      //         data: "2",
      //         onClick: this.changeSelected,
      //         disabled: this.props.primaryApproverList[0].GRCCompleted === "Yes"

      //       }

      //     ]
      //   }
      // },
      // {
      //   key: "Undo", name: "Undo", icon: "Undo", onClick: this.fetchRoleReviews,
      //   disabled: !(filter(this.state.roleReview, (rr) => { return rr.hasBeenUpdated; }).length > 0)
      // },
      // { // if the item has been comleted OR there are items with noo approvasl, diable
      //   key: "Done", name: "Complete", icon: "Completed", onClick: this.setComplete,
      //   disabled: this.props.primaryApproverList[0].GRCCompleted === "Yes" ||
      //   (filter(this.state.roleReview, (rr) => { return rr.GRCApproval === null; }).length > 0)
      // }

    ];
    let farItemsNonFocusable: IContextualMenuItem[] = [
      {

        key: "Save", name: "Save", icon: "Save", onClick: this.save,
        disabled: !(filter(this.state.roleReview, (rr) => { return rr.hasBeenUpdated; }).length > 0)
        || this.props.primaryApproverList[0].GRCCompleted === "Yes"

      }
    ];


    return (
      <div className={styles.grcTest}>
        <CommandBar
          isSearchBoxVisible={false}
          items={itemsNonFocusable}
          farItems={farItemsNonFocusable}

        />
         <IconButton iconProps={{ iconName: "Info" }} onClick={(e) => { this.showPopup(null); }} />
        <DetailsList
          items={this.state.roleReview}
          selectionMode={SelectionMode.multiple}
          selection={this.selection}
          key="Id"
          columns={[
            {
              key: "id", name: "Id",
              fieldName: "Id", minWidth: 5, maxWidth: 5,

            },
            {
              key: "title", name: "Role Name",
              fieldName: "GRCRoleName", minWidth: 400, maxWidth: 400,

            },
            {
              key: "info", name: "Info",
              fieldName: "GRCRoleName", minWidth: 10, maxWidth: 10,
              onRender: (item?: any, index?: number, column?: IColumn) => {
                return (
                  <IconButton iconProps={{ iconName: "Info" }} onClick={(e) => { this.showPopup(item); }} />
                );
              }

            },
            {
              key: "Approval", name: "Approval",
              fieldName: "GRCApproval", minWidth: 90, maxWidth: 90,
              onRender: (item?: any, index?: number, column?: IColumn) => { return this.RenderApproval(item, index, column); },
            },
            {
              key: "Comments", name: "Comments",
              fieldName: "GRCComments", minWidth: 150, maxWidth: 150,
              onRender: (item?: any, index?: number, column?: IColumn) => { return this.RenderComments(item, index, column); },
            },
            {
              key: "Remediation", name: "Remediation",
              fieldName: "GRCRemediation", minWidth: 150, maxWidth: 150,
            },

          ]}
        />

        <Panel
          type={PanelType.custom | PanelType.smallFixedNear}
          customWidth='600px'
          isOpen={this.state.showPopup}
          onDismiss={
            () => {
              this.setState((current) => ({ ...current, roleToTransaction: [], showPopup: false }));
            }}
          isBlocking={true}
        >
          <DetailsList
            items={this.state.roleToTransaction}
            selectionMode={SelectionMode.none}
            columns={[
              {
                key: "Role", name: "Role",
                fieldName: "GRCRole", minWidth: 150, maxWidth: 150,
              },
              {
                key: "Comments", name: "TCode",
                fieldName: "GRCTCode", minWidth: 50, maxWidth: 50,

              },
              {
                key: "Remediation", name: "TransactipnText",
                fieldName: "GRCTransactionText", minWidth: 150, maxWidth: 150,

              },

            ]}
          />
        </Panel>


      </div>
    );
  }
}
