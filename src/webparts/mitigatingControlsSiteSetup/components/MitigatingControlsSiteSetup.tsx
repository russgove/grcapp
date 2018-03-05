import * as React from 'react';
import styles from './MitigatingControlsSiteSetup.module.scss';
import { IMitigatingControlsSiteSetupProps } from './IMitigatingControlsSiteSetupProps';
import { IMitigatingControlsSiteSetupState } from './IMitigatingControlsSiteSetupState';
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
import pnp, { TypedHash, ItemAddResult, ListAddResult, ContextInfo, Web, WebAddResult, List as PNPList } from "sp-pnp-js";
import { map } from "lodash";
import {
  addPeopleFieldToList, convertEmailColumnsToUser,
  CachedId, findId, uploadFile, esnureUsers, extractColumnHeaders, processUploadedFiles, getListFromWeb, ensureFieldsAreInList
  , setWebToUseSharedNavigation, fixUpLeftNav, addCustomListWithContentType, cleanupHomePage, getContentTypeByName
} from "../../../utilities/utilities";

export default class MitigatingControlsSiteSetup extends React.Component<IMitigatingControlsSiteSetupProps, IMitigatingControlsSiteSetupState> {
  constructor(props: IMitigatingControlsSiteSetupProps) {
    super(props);
    // this.processUploadedFiles = this.processUploadedFiles.bind(this);


    this.state = {
      webName: "",
      webUrl: "",
      siteDropDownOptions: [],
      mitigatingControlsListExists: false,
      primaryApproversListExists: false,
    };
  }
  public componentDidMount() {
    this.getSitesDropDownOptions().then(sites => {
      this.setState((current) => ({ ...current, siteDropDownOptions: sites }));

    });
  }
  private getSitesDropDownOptions(): Promise<Array<IDropdownOption>> {
    return pnp.sp.site.rootWeb.webinfos.get()
      .then((wi) => {

        return map(wi, web => { return { text: web["Title"], key: web["ServerRelativeUrl"] }; });
      })
      .catch((err) => {
        console.error(err);
        return [];
      });
  }
  private siteChanged(option: IDropdownOption, idx?: number) {
    this.setState((current) => ({ ...current, webUrl: option.key, webName: option.text }));

    // test the mitigating controls list, ensure the list exists and has required fields
    getListFromWeb(option.key as string, this.props.mitigatingControlsListName).then(list => {
      this.setState((current) => ({
        ...current,
        mitigatingControlsListExists: true,
        mitigatingControlsCount: list["ItemCount"],
        mitigatingControlsList: list
      }));
      let fieldsfound = ensureFieldsAreInList(list, ["Control_x0020_ID", "Risk_x0020_ID", "ApproverEmail"]);
      this.setState((current) => ({
        ...current,
        mitigatingControlsFieldsFound: fieldsfound
      }));
    }).catch(err => {
      alert(`List ${this.props.mitigatingControlsListName} was not found on that site`);
      this.setState((current) => ({
        ...current,
        mitigatingControlsListExists: false,
        mitigatingControlsCount: 0,
        mitigatingControlsList: null
      }));
    });
    debugger;

    // test the primary approverslist, ensure the list exists and has required fields
    getListFromWeb(option.key as string, this.props.primaryApproversListName).then(list => {
      debugger;// why no list???
      this.setState((current) => ({
        ...current,
        primaryApproversListExists: true,
        primaryApproversCount: list["ItemCount"],
        primaryApproversList: list
      }));
      let fieldsfound = ensureFieldsAreInList(list, ["ApproverEmail"]);
      this.setState((current) => ({
        ...current,
        primaryApproversFieldsFound: fieldsfound,


      }));
    }).catch(err => {
      alert(`List ${this.props.primaryApproversListName} was not found on that site`);
      this.setState((current) => ({
        ...current,
        primaryApproversListExists: false,
        primaryApproversCount: 0
      }));
    });


  }
  private async ConvertUsersInMitigatingControls() {
    // so we need to convert the  ApproverEmail to a real User column
    // first we'll check if the user COLUMN is not prensent and If not Add it. (nah, do this later)
    // then , we'll repeatedly get 100 rows wher the user COLUMN  is empty.
    // for each of those rows, we'll call ensureUser and then update the row with the users ID.
    debugger;
    if (!ensureFieldsAreInList(this.state.mitigatingControlsList, ["PrimaryApprover"])) {
      await addPeopleFieldToList(this.state.webUrl, this.props.mitigatingControlsListName, "PrimaryApprover", "PrimaryApprover").then(d => {


      }).catch(err => {

      });
    }

    await convertEmailColumnsToUser(this.state.webUrl, this.props.mitigatingControlsListName, [["ApproverEmail", "PrimaryApproverId"]]);




  }
  private async ConvertUsersInApproversList() {
    // so we need to convert the  ApproverEmail to a real User column
    // first we'll check if the user COLUMN is not prensent and If not Add it. (nah, do this later)
    // then , we'll repeatedly get 100 (well , getem all for now....) rows wher the user COLUMN  is empty.(user column CANNOT be indexed)
    // for each of those rows, we'll call ensureUser and then update the row with the users ID.
    debugger;
    if (!ensureFieldsAreInList(this.state.primaryApproversList, ["PrimaryApprover"])) {
      await addPeopleFieldToList(this.state.webUrl, this.props.primaryApproversListName, "PrimaryApprover", "PrimaryApprover").then(d => {
        debugger;

      }).catch(err => {
        debugger;
      });
    }
    debugger;
    await convertEmailColumnsToUser(this.state.webUrl, this.props.primaryApproversListName, [["ApproverEmail", "PrimaryApproverId"]]);




  }
  public render(): React.ReactElement<IMitigatingControlsSiteSetupProps> {
    return (
      <div className={styles.mitigatingControlsSiteSetup}>
        Site:   <Dropdown options={this.state.siteDropDownOptions}
          selectedKey={this.state.webUrl}
          onChanged={this.siteChanged.bind(this)}
        >
        </Dropdown>
        <table style={{ border: 1 }}>
          <thead>
            <th>
              SharePoint List
            </th>
            <th>
              List Exists
            </th>
            <th>
              Has Requied Columns
            </th>
            <th>
              Total Rows
            </th>
            <th>
              Rows UpDated
            </th>
            <th>
              Convert Users
            </th>

          </thead>
          <tr>
            <td>
              {this.props.primaryApproversListName}
            </td>

            <td>
              {this.state.primaryApproversListExists ? "Yes" : "No"}
            </td>

            <td>
              {this.state.primaryApproversFieldsFound ? "Yes" : "No"}
            </td>

            <td>
              {this.state.primaryApproversCount}
            </td>

            <td>
            </td>

            <td>
              <IconButton
                iconProps={{ iconName: "DocumentApproval" }}
                onClick={this.ConvertUsersInApproversList.bind(this)}>
                Convert users
              </IconButton>
            </td>

          </tr>
          <tr>
            <td>
              {this.props.mitigatingControlsListName}
            </td>

            <td>
              {this.state.mitigatingControlsListExists ? "Yes" : "No"}
            </td>

            <td>
              {this.state.mitigatingControlsFieldsFound ? "Yes" : "No"}
            </td>

            <td>
              {this.state.mitigatingControlsCount}
            </td>

            <td>
            </td>

            <td>
              <IconButton
                iconProps={{ iconName: "DocumentApproval" }}
                onClick={this.ConvertUsersInMitigatingControls.bind(this)} >
                Convert users
              </IconButton>
            </td>
          </tr>

        </table>
        <hr />
        <PrimaryButton title="Update Site">Update Site </PrimaryButton>

      </div>
    );
  }
}
