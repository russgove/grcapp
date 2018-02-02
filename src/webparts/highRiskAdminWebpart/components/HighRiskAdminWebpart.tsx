import * as React from 'react';
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import styles from './HighRiskAdminWebpart.module.scss';
import { IHighRiskAdminWebpartProps } from './IHighRiskAdminWebpartProps';
import { IHighRiskAdminWebpartState } from './IHighRiskAdminWebpartState';
import { escape } from '@microsoft/sp-lodash-subset';
const parse = require('csv-parse');
import pnp, { TypedHash, ItemAddResult, ListAddResult, ContextInfo, Web, WebAddResult, List as PNPList } from "sp-pnp-js";
import { List } from "office-ui-fabric-react/lib/List";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Image, ImageFit } from "office-ui-fabric-react/lib/Image";
import { Button, IconButton, PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { find, clone, map } from "lodash";
require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');
require('sp-workflow');
import {
  CachedId, findId, uploadFile, esnureUsers, extractColumnHeaders, processUploadedFiles
  , setWebToUseSharedNavigation, fixUpLeftNav, addCustomListWithContentType, cleanupHomePage
} from "../../../utilities/Utilities";
export default class HighRiskAdminWebpart extends React.Component<IHighRiskAdminWebpartProps, IHighRiskAdminWebpartState> {
  cachedIds: Array<CachedId> = []
  private reader;
  private counter = 0;
  constructor(props: IHighRiskAdminWebpartProps) {

    super(props);


    this.state = {
      siteName: "",
      newWeb: null,
      newWebUrl: null,


      messages: [],
      roleToTransactionRowsUploaded: 0,
      roleToTransactionTotalRows: 0,
      roleToTransactionStatus: "",
      roleToTransactionFile: null,

      highRiskTotalRows: 0,
      highRiskRowsUploaded: 0,
      highRiskStatus: "",
      highRiskFile: null,

      primaryApproversTotalRows: 0,
      primaryApproversRowsUploaded: 0,
      primaryApproversStatus: "",
      primaryApproversFile: null,
    };


  }
  private addMessage(message: string) {
    let messages = this.state.messages;
    var copy = map(this.state.messages, clone);
    copy.push(message);
    this.setState((current) => ({ ...current, messages: copy }));
  }
  // #region High Risk Users
  public async esnureHighRiskUsers(error, data: Array<any>): Promise<any> {
    debugger;
    this.setState((current) => ({ ...current, roleReviewStatus: "Validating users" }));
    esnureUsers(pnp.sp.web, this.cachedIds, data, "ApproverEmail", this.addMessage);
    this.addMessage("Done validating RoleReviewUsers");
    this.setState((current) => ({ ...current, roleReviewStatus: "Completed Validation" }));
    return Promise.resolve();
  }
  public extractColumnHeadersHighRiskUsers(headerRow: Array<String>): String[] {
    debugger;
    const requiredColumns = ["ApproverEmail", "AlternateApproverEmail", "Role Name"];
    for (let requiredColumn of requiredColumns) {
      if (headerRow.indexOf(requiredColumn) === -1) {
        this.addMessage(`Column ${requiredColumn} is missing on Role Review Data File`);
      }
    }

    return extractColumnHeaders(headerRow);

  }
  public onHighRiskUsersDataloaded() {
    this.setState((current) => ({ ...current, roleReviewStatus: "Parsing file" }));
    parse(this.reader.result, { delimiter: ',', columns: this.extractColumnHeadersHighRiskUsers }, this.esnureHighRiskUsers);
  }
  public parseHighRiskUsersFile(): Promise<any> {

    //https://stackoverflow.com/questions/14446447/how-to-read-a-local-text-file
    //let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, roleReviewStatus: "Reading file" }));
    this.reader = new FileReader();
    this.reader.onload = this.onHighRiskUsersDataloaded;
    this.reader.readAsText(this.state.highRiskFile);
    return Promise.resolve();

  }
  public uploadHighRiskUsersFile(): Promise<any> {
    debugger;
    this.setState((current) => ({ ...current, roleReviewStatus: "Uploading file" }));
    return uploadFile(this.state.newWeb, "Documents", this.state.highRiskFile, "High Risk", this.addMessage);

  }
  public saveHighRiskFile(e: any) {

    //https://stackoverflow.com/questions/14446447/how-to-read-a-local-text-file
    let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, highRiskFile: file }));
  }
  public async createSite() {
    let newWeb: Web;  // the web that gets created
    let libraryList: Array<any>; // the list of libraries we need to create in the new site. has the library name and the name of the group that should get access
    let foldersList: Array<string>; // the list of folders to create in each of the libraries.
    let roleDefinitions: Array<any>;// the roledefs for the site, we need to grant 'contribute no delete'
    let siteGroups: Array<any>;// all the sitegroups in the site
    let tasks: Array<any>; // the list of tasks in the TaskMaster list. We need to create on e task for each of these in tye EFRTasks list in the new site
    let taskList: List; // the task list we created  in the new site
    let taskListId: string; // the ID of task list we created  in the new site
    let webServerRelativeUrl: string; // the url of the subweb
    let contextInfo: ContextInfo;
    let editformurl: string;



    this.addMessage("CreatingSite");
    await pnp.sp.site.getContextInfo().then((context: ContextInfo) => {
      contextInfo = context;
    });
    // create the site
    await pnp.sp.web.webs.add(this.state.siteName, this.state.siteName, this.state.siteName, this.props.templateName).then((war: WebAddResult) => {
      this.addMessage("CreatedSite");
      // show the response from the server when adding the web
      webServerRelativeUrl = war.data.ServerRelativeUrl;
      console.log(war.data);
      newWeb = war.web;
      this.setState((current) => ({
        ...current,
        newWeb: newWeb,
        newWebUrl: webServerRelativeUrl
      })); /// save in state so file uploads can use it

      return;
    }).catch(error => {
      debugger;
      this.addMessage("<h1>error creating site</h1>");
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      console.error(error);
      return;
    });

    await setWebToUseSharedNavigation(webServerRelativeUrl);
    debugger;
    await fixUpLeftNav(webServerRelativeUrl, this.props.siteUrl);

    // create the lists and assign permissions
    let highRiskList: PNPList = await addCustomListWithContentType(newWeb, 'High Risk', "High Risk Transactions"
      , this.props.highRiskContentTypeId, this.addMessage);

    let primaryAppproversList: PNPList = await addCustomListWithContentType(newWeb, 'Primary Approvers', "Primary Approvers of High Risk Transactions"
      , this.props.primaryApproverContentTypeId, this.addMessage);

    let roleToTransactionList: PNPList = await addCustomListWithContentType(newWeb, 'Role To Transaction', "Role To Transaction DETAILS"
      , this.props.roleToTransactionContentTypeId, this.addMessage);
    await roleToTransactionList.fields.getByInternalNameOrTitle("GRCRoleName").update({
      Indexed: true
    }).then((resp) => {
      debugger;
    }).catch((error) => {
      debugger;
      console.error(error);
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      return;

    });
    // this is a list that the webjob can use to log messages
    let messageList: PNPList = await addCustomListWithContentType(newWeb, 'Messages', "Messages from the provsioning Web Job"
      , "0x01", this.addMessage);

    // customize the home paged
    let welcomePageUrl: string;
    await newWeb.rootFolder.getAs<any>().then(rootFolder => {
      debugger;
      welcomePageUrl = rootFolder.ServerRelativeUrl + rootFolder.WelcomePage;
    });
    this.addMessage("Customizing Home Page");
    await cleanupHomePage(webServerRelativeUrl, welcomePageUrl, this.props.webPartXml);
    this.addMessage("Customized Home Page");

    this.addMessage("DONE!!");

  }
  private processUploadedFiles(): void {
    debugger;

    let functionUrl = `${this.props.azureFunctionUrl}
       &siteUrl=${this.props.siteAbsoluteUrl + "/" + this.state.siteName}
       &siteType=Role To Ttansaction
       &PrimaryApproverList=Primary Approvers List
       &RoleReview=Role Review
       &RoleToTransaction=Role To Transaction`
    processUploadedFiles(this.props.httpClient, functionUrl);
  }
  public render(): React.ReactElement<IHighRiskAdminWebpartProps> {

    return (
      <div className={styles.grcUpload} >


        <table>
          <thead>
            <th>
              SharePoint List
          </th>
            <th>
              File
          </th>
            <th>
              Status
          </th>
            <th>
              Total Rows
          </th>
            <th>
              Rows Uploaded
          </th>

          </thead>
          <tr>
            <td>
              Role Review
            </td>
            <td>
              <input type="file" id="uploadrttfile" onChange={e => { this.saveHighRiskFile(e); }} />
            </td>
            <td>
              {this.state.highRiskStatus}
            </td>
            <td>
              {this.state.highRiskTotalRows}
            </td>
            <td>
              {this.state.highRiskRowsUploaded}
            </td>
            <td>
              <IconButton iconProps={{ iconName: "DocumentApproval" }} onClick={this.parseHighRiskUsersFile}>Upload</IconButton>
              <IconButton iconProps={{ iconName: "Save" }} onClick={this.uploadHighRiskUsersFile}>Upload</IconButton>
            </td>

          </tr>

        </table>
        <hr />
        <table>
          <tr>
            <td>
              New Site Name
            </td>
            <td>
              <TextField label="" onChanged={(e) => {
                this.setState((current) => ({ ...current, siteName: e }));
              }} />
            </td>
            <td>
              <PrimaryButton onClick={this.createSite.bind(this)} title="Create Site">Create Site</PrimaryButton>
            </td>
            <td>
              Active Site:
            </td>
            <td>
              {this.state.newWebUrl}
            </td>

          </tr>

        </table>
        <button onClick={this.processUploadedFiles}>Process Uploaded Files</button>
        <div style={{ border: '1px', borderStyle: "solid" }} >
          <IconButton iconProps={{ iconName: "Clear" }}
            onClick={
              () => { this.setState((current) => ({ ...current, messages: [] })); }
            }>Upload</IconButton>
          <List items={this.state.messages}
            onRenderCell={(item: any, index: number, isScrolling: boolean) => {
              return (
                <div>{item}</div>
              );
            }
            } />

        </div>
      </div >
    );
  }
}
