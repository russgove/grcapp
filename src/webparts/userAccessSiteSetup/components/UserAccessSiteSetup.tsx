import * as React from 'react';
import styles from './UserAccessSiteSetup.module.scss';
import { IUserAccessSiteSetupProps } from './IUserAccessSiteSetupProps';
import { IUserAccessSiteSetupState } from './IUserAccessSiteSetupState';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  DetailsList, DetailsListLayoutMode, IColumn, SelectionMode, Selection,
  ColumnActionsMode
} from "office-ui-fabric-react/lib/DetailsList";
import { List } from "office-ui-fabric-react/lib/List";
import { Dropdown, IDropdownOption, IDropdownProps } from "office-ui-fabric-react/lib/Dropdown";
import { Modal, IModalProps } from "office-ui-fabric-react/lib/Modal";
import { Panel, IPanelProps, PanelType } from "office-ui-fabric-react/lib/Panel";
import { CommandBar } from "office-ui-fabric-react/lib/CommandBar";
import { IContextualMenuItem } from "office-ui-fabric-react/lib/ContextualMenu";

import { PrimaryButton, ButtonType, Button, DefaultButton, ActionButton, IconButton } from "office-ui-fabric-react/lib/Button";
import { Dialog } from "office-ui-fabric-react/lib/Dialog";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { sp, ItemAddResult, ListAddResult, ContextInfo, Web, WebAddResult, List as PNPList } from "@pnp/sp";
import { map, clone } from "lodash";
import {
  addPeopleFieldToList, convertEmailColumnsToUser, AddQuickLaunchItem, RemoveQuickLaunchItem, AddUsersInListToGroup,
  CachedId, findId, uploadFile, esnureUsers, extractColumnHeaders, processUploadedFiles, getListFromWeb, ensureFieldsAreInList
  , setWebToUseSharedNavigation, addCustomListWithContentType, cleanupHomePage, getContentTypeByName
} from "../../../utilities/utilities";
require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');
require('sp-workflow');

export default class UserAccessSiteSetup extends React.Component<IUserAccessSiteSetupProps, IUserAccessSiteSetupState> {
  constructor(props: IUserAccessSiteSetupProps) {
    super(props);
    this.addMessage = this.addMessage.bind(this);


    this.state = {
      webName: "",
      webUrl: "",
      siteDropDownOptions: [],
      userAccessListExists: false,
      tcodeListExists: false,
      primaryApproversListExists: false,
      messages: []
    };
  }
  public componentDidMount() {
    this.getSitesDropDownOptions().then(sites => {
      this.setState((current) => ({ ...current, siteDropDownOptions: sites }));

    });
  }
  private getSitesDropDownOptions(): Promise<Array<IDropdownOption>> {
    return sp.site.rootWeb.webinfos.get()
      .then((wi) => {

        return map(wi, web => { return { text: web["Title"], key: web["ServerRelativeUrl"] }; });
      })
      .catch((err) => {
        console.error(err);
        this.addMessage(`<h1>Error in getSitesDropDownOptions was  ${err.data.responseBody["odata.error"].message.value} </h1>`);
        return [];
      });
  }
  private siteChanged(option: IDropdownOption, idx?: number) {
    this.setState((current) => ({ ...current, webUrl: option.key, webName: option.text }));

    // test the user access  list, ensure the list exists and has required fields
    getListFromWeb(option.key as string, this.props.userAccessListName).then(list => {
      this.addMessage(`Required List ${this.props.userAccessListName} was found on that site`);
      this.setState((current) => ({
        ...current,
        userAccessListExists: true,
        userAccessCount: list["ItemCount"],
        userAccessList: list
      }));
      let fieldsfound = ensureFieldsAreInList(list, [
        "Role",
        "Role_x0020_name",
        "Approval",
        "Date_x0020_Reviewed",
        "User_x0020_ID",
        "User_x0020_Full_x0020_Name",
        "Remediation",
        "Approver",
        "Approver_x0020_Name",
        "ApproverEmail",
        "Alternate_x0020_Approver",
        "AlternateApproverEmail",
        "Approval", //approval decision
        "Comments"


      ], this.addMessage);
      if (fieldsfound) {
        this.addMessage(`Required fields found on List ${this.props.userAccessListName} `);

      }
      this.setState((current) => ({
        ...current,
        userAccessFieldsFound: fieldsfound
      }));
    }).catch(err => {
      this.addMessage(`<h1>List ${this.props.userAccessListName} was not found on that site</h1>`);
      this.setState((current) => ({
        ...current,
        userAccessListExists: false,
        userAccessCount: 0,
        userAccessList: null
      }));
    });


    // test the primary approverslist, ensure the list exists and has required fields
    getListFromWeb(option.key as string, this.props.primaryApproversListName).then(list => {
      this.addMessage(`Required List ${this.props.primaryApproversListName} was found on that site`);
      this.setState((current) => ({
        ...current,
        primaryApproversListExists: true,
        primaryApproversCount: list["ItemCount"],
        primaryApproversList: list
      }));
      let fieldsfound = ensureFieldsAreInList(list, [
        "ApproverEmail",
        "Completed"
      ], this.addMessage);
      if (fieldsfound) {
        this.addMessage(`Required fields found on List ${this.props.primaryApproversListName} `);

      }
      this.setState((current) => ({
        ...current,
        primaryApproversFieldsFound: fieldsfound,
      }));
    }).catch(err => {
      this.addMessage(`<h1>List ${this.props.primaryApproversListName} was not found on that site</h1>`);
      this.setState((current) => ({
        ...current,
        primaryApproversListExists: false,
        primaryApproversCount: 0
      }));
    });

    // test the  tcode list, ensure the list exists and has required fields
    getListFromWeb(option.key as string, this.props.tcodeListName).then(list => {
      this.addMessage(`Required List ${this.props.tcodeListName} was found on that site`);
      this.setState((current) => ({
        ...current,
        tcodeListExists: true,
        tcodeCount: list["ItemCount"],
        tcodeList: list
      }));
      let fieldsfound = ensureFieldsAreInList(list, [
        "Role",
        "Composite_x0020_role",
        "TCode",
        "Transaction_x0020_Text"], this.addMessage);
      if (fieldsfound) {
        this.addMessage(`Required fields found on List ${this.props.tcodeListName} `);

      }
      this.setState((current) => ({
        ...current,
        primaryApproversFieldsFound: fieldsfound,
      }));
    }).catch(err => {
      this.addMessage(`<h1>List ${this.props.tcodeListName} was not found on that site</h1>`);
      this.setState((current) => ({
        ...current,
        tcodeListExists: false,
        tcodeCount: 0
      }));
    });


  }
  private async ConvertUsersInuserAccess() {
    debugger;
    // so we need to convert the  ApproverEmail to a real User column
    // first we'll check if the user COLUMN is not prensent and If not Add it. (nah, do this later)
    // then , we'll repeatedly get 100 rows wher the user COLUMN  is empty.
    // for each of those rows, we'll call ensureUser and then update the row with the users ID.
    if (!ensureFieldsAreInList(this.state.userAccessList, ["PrimaryApprover"], this.addMessage)) {
      this.addMessage(`Creating PrimaryApprover Column in '${this.state.userAccessList["Title"]}'`);
      await addPeopleFieldToList(this.state.webUrl, this.props.userAccessListName, "PrimaryApprover", "PrimaryApprover").then(d => {
        this.addMessage(`Created PrimaryApprover Column  in '${this.state.userAccessList["Title"]}'`);

      }).catch(err => {
        this.addMessage(`<h1>There was an error adding the PrimaryApprover column to the ${this.state.userAccessList["Title"]} list</h1>`);
      });

    } else {
      this.addMessage(`PrimaryApprover Column already exists in '${this.state.userAccessList["Title"]}'`);
    }

    this.addMessage(`Updating PrimaryApprover column from ApproverEmail  in '${this.state.userAccessList["Title"]}'`);
    debugger;
    await convertEmailColumnsToUser(this.state.webUrl, this.props.userAccessListName, [["ApproverEmail", "PrimaryApproverId"]], this.addMessage);
    this.addMessage(`Updated PrimaryApprover column from ApproverEmail  in '${this.state.userAccessList["Title"]}'`);



  }
  private async ConvertUsersInApproversList() {
    // so we need to convert the  ApproverEmail to a real User column
    // first we'll check if the user COLUMN is not prensent and If not Add it. (nah, do this later)
    // then , we'll repeatedly get 100 (well , getem all for now....) rows wher the user COLUMN  is empty.(user column CANNOT be indexed)
    // for each of those rows, we'll call ensureUser and then update the row with the users ID.
    debugger;
    if (!ensureFieldsAreInList(this.state.primaryApproversList, ["PrimaryApprover"], this.addMessage)) {
      this.addMessage(`Creating PrimaryApprover Column in '${this.state.primaryApproversList["Title"]}'`);
      await addPeopleFieldToList(this.state.webUrl, this.props.primaryApproversListName, "PrimaryApprover", "PrimaryApprover").then(d => {
        this.addMessage(`Created PrimaryApprover Column in '${this.state.primaryApproversList["Title"]}'`);
      }).catch(err => {
        this.addMessage(`<h1>Error Creating PrimaryApprover Column in '${this.state.primaryApproversList["Title"]}'</h1>`);
        this.addMessage(`<h1>Error was  ${err.data.responseBody["odata.error"].message.value} </h1>`);
        console.error(err);
        debugger;
      });
    } else {
      this.addMessage(`PrimaryApprover Column already exists in '${this.state.primaryApproversList["Title"]}'`);
    }

    this.addMessage(`Updating PrimaryApprover column from ApproverEmail  in '${this.state.primaryApproversList["Title"]}'`);
    await convertEmailColumnsToUser(this.state.webUrl, this.props.primaryApproversListName, [["ApproverEmail", "PrimaryApproverId"]], this.addMessage);
    this.addMessage(`Updated PrimaryApprover column from ApproverEmail  in '${this.state.primaryApproversList["Title"]}'`);




  }
  private addMessage(message: string) {
    let messages = this.state.messages;
    var copy = map(this.state.messages, clone);
    copy.push(message);
    this.setState((current) => ({ ...current, messages: copy }));
  }
  private displayMessages(): any {
    const messages = map(this.state.messages, (m) => {
      return "<div>" + m + "</div>";
    });
    return { __html: messages.join('') };
  }
  // #region Site creation

  public async setupWeb() {
    debugger;
    let newWeb = new Web(window.location.origin + this.state.webUrl);



    this.addMessage(`Origin is ${window.location.origin}`);
    this.addMessage(`WebUrl is  ${this.state.webUrl}`);
    this.addMessage(`SiteUrl  is  ${this.props.siteUrl}`);
    this.addMessage(`Updating Site at ${window.location.origin + this.state.webUrl}`);

    await setWebToUseSharedNavigation(window.location.origin + this.state.webUrl, this.addMessage);

    await AddQuickLaunchItem(this.state.webUrl, "GRC Home", this.props.siteUrl, true, this.addMessage);
    await RemoveQuickLaunchItem(this.state.webUrl, ["Pages", "Documents"], this.addMessage);

    // customize the home paged
    let welcomePageUrl: string;
    await newWeb.rootFolder.get<any>().then(rootFolder => {
      welcomePageUrl = rootFolder.ServerRelativeUrl + rootFolder.WelcomePage;
    });
    this.addMessage("Customizing Home Page");

    await cleanupHomePage(this.props.siteUrl, welcomePageUrl, this.props.webPartXml, this.addMessage);
    this.addMessage("Customized Home Page");
    // add all the approvers as members
    await newWeb.associatedMemberGroup.get().then(async (membersGroup) => {
      debugger;
      await AddUsersInListToGroup(window.location.origin + this.state.webUrl, this.props.primaryApproversListName, "PrimaryApprover", membersGroup, this.addMessage);

      debugger;
    }).catch(err => {
      this.addMessage(`<h1>Error Adding users to members group</h1>`);
      this.addMessage(`<h1>Error was  ${err.data.responseBody["odata.error"].message.value} </h1>`);
      console.error(err);
    });
    this.addMessage("DONE!!");

  }
  public render(): React.ReactElement<IUserAccessSiteSetupProps> {
    return (
      <div className={styles.userAccessSiteSetup}>
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
              {this.props.userAccessListName}
            </td>

            <td>
              {this.state.userAccessListExists ? "Yes" : "No"}
            </td>

            <td>
              {this.state.userAccessFieldsFound ? "Yes" : "No"}
            </td>

            <td>
              {this.state.userAccessCount}
            </td>

            <td>
            </td>

            <td>
              <IconButton
                iconProps={{ iconName: "DocumentApproval" }}
                onClick={this.ConvertUsersInuserAccess.bind(this)} >
                Convert users
              </IconButton>
            </td>
          </tr>
          <tr>
            <td>
              {this.props.tcodeListName}
            </td>

            <td>
              {this.state.tcodeListExists ? "Yes" : "No"}
            </td>

            <td>
              {this.state.tcodeFieldsFound ? "Yes" : "No"}
            </td>

            <td>
              {this.state.tcodeCount}
            </td>

            <td>
            </td>

            <td>
              {/* <IconButton
                iconProps={{ iconName: "DocumentApproval" }}
                onClick={this.ConvertUsersInuserAccess.bind(this)} >
                Convert users
              </IconButton> */}
            </td>
          </tr>


        </table>
        <hr />
        <PrimaryButton title="Update Site" onClick={this.setupWeb.bind(this)}>Update Site </PrimaryButton>
        <div style={{ border: '1px', borderStyle: "solid" }} >
          <IconButton iconProps={{ iconName: "Clear" }}
            onClick={
              () => { this.setState((current) => ({ ...current, messages: [] })); }
            }></IconButton>
          <div dangerouslySetInnerHTML={this.displayMessages()} />

        </div>

      </div>
    );
  }
}
