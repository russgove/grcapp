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
  , setWebToUseSharedNavigation, fixUpLeftNav, addCustomListWithContentType, cleanupHomePage, getContentTypeByName
} from "../../../utilities/utilities";
export default class HighRiskAdminWebpart extends React.Component<IHighRiskAdminWebpartProps, IHighRiskAdminWebpartState> {
  private cachedIds: Array<CachedId> = [];
  private reader;
  private counter = 0;
  constructor(props: IHighRiskAdminWebpartProps) {
    super(props);

    this.addMessage = this.addMessage.bind(this);
    // this.processUploadedFiles = this.processUploadedFiles.bind(this);

    this.parseHighRiskFile = this.parseHighRiskFile.bind(this);
    this.esnureHighRisk = this.esnureHighRisk.bind(this);
    this.extractColumnHeadersHighRisk = this.extractColumnHeadersHighRisk.bind(this);
    this.onHighRiskDataloaded = this.onHighRiskDataloaded.bind(this);
    this.uploadHighRiskFile = this.uploadHighRiskFile.bind(this);
    this.saveHighRiskFile = this.saveHighRiskFile.bind(this);


    this.parsePrimaryApproversFile = this.parsePrimaryApproversFile.bind(this);
    this.esnurePrimaryApprovers = this.esnurePrimaryApprovers.bind(this);
    this.extractColumnHeadersPrimaryApprovers = this.extractColumnHeadersPrimaryApprovers.bind(this);
    this.onPrimaryApproversDataloaded = this.onPrimaryApproversDataloaded.bind(this);
    this.uploadPrimaryApproversFile = this.uploadPrimaryApproversFile.bind(this);
    this.savePrimaryApproversFile = this.savePrimaryApproversFile.bind(this);


    this.parseRoleToTCodeFile = this.parseRoleToTCodeFile.bind(this);
    this.esnureRoleToTCode = this.esnureRoleToTCode.bind(this);
    this.extractColumnHeadersRoleToTCode = this.extractColumnHeadersRoleToTCode.bind(this);
    this.onRoleToTCodeDataloaded = this.onRoleToTCodeDataloaded.bind(this);
    this.uploadRoleToTCodeFile = this.uploadRoleToTCodeFile.bind(this);
    this.saveRoleToTCodeFile = this.saveRoleToTCodeFile.bind(this);




    this.state = {
      siteName: "",
      newWeb: null,
      newWebUrl: null,

      messages: [],

      roleToTransactionTotalRows: 0,
      roleToTransactionStatus: "",
      roleToTransactionFile: null,

      highRiskTotalRows: 0,
      highRiskStatus: "",
      highRiskFile: null,

      primaryApproversTotalRows: 0,
      primaryApproversStatus: "",
      primaryApproversFile: null,
    };
  }


  // #region High Risk Users
  public async esnureHighRisk(error, data: Array<any>): Promise<any> {

    this.setState((current) => ({ ...current, highRiskStatus: "Validating users", highRiskTotalRows: data.length }));
    await esnureUsers(pnp.sp.web, this.cachedIds, data, "ApproverEmail", this.addMessage);
    this.addMessage("Done validating High Risk Users");
    this.setState((current) => ({ ...current, highRiskStatus: "Completed Validation" }));
    return Promise.resolve();
  }
  public extractColumnHeadersHighRisk(headerRow: Array<String>): String[] {

    const requiredColumns = ["ApproverEmail", "AlternateApproverEmail", "Role Name", "User ID", "User Full Name"];
    for (let requiredColumn of requiredColumns) {
      if (headerRow.indexOf(requiredColumn) === -1) {
        this.addMessage(`Column ${requiredColumn} is missing on Role Review Data File`);
      }
    }
    return extractColumnHeaders(headerRow);
  }
  public onHighRiskDataloaded() {
    this.setState((current) => ({ ...current, highRiskStatus: "Parsing file" }));
    parse(this.reader.result, { delimiter: ',', columns: this.extractColumnHeadersHighRisk }, this.esnureHighRisk);
  }
  public parseHighRiskFile(): Promise<any> {
    this.setState((current) => ({ ...current, highRiskStatus: "Reading file" }));
    this.reader = new FileReader();
    this.reader.onload = this.onHighRiskDataloaded;
    this.reader.readAsText(this.state.highRiskFile);
    return Promise.resolve();
  }
  public uploadHighRiskFile(): void {
    debugger;
    this.setState((current) => ({ ...current, highRiskStatus: "Uploading file" }));
    uploadFile(this.state.newWeb, "Documents", this.state.highRiskFile, "High Risk", this.addMessage)
      .then((resp) => {
        this.setState((current) => ({ ...current, highRiskStatus: "Uploaded file" }));
        let functionUrl = `${this.props.azureHighRiskUrl}&siteUrl=${this.props.siteAbsoluteUrl}/${this.state.siteName}&listName=High Risk&fileName=High Risk&batchSize=${this.props.batchSize}&pauseBeforeBatchExecution=${this.props.pauseBeforeBatchExecution}`;
        this.addMessage(functionUrl);
        processUploadedFiles(this.props.httpClient, functionUrl).then(() => {
          this.setState((current) => ({ ...current, highRiskStatus: "Job Scheduled" }));
        }).catch(e => {
          this.setState((current) => ({ ...current, highRiskStatus: "Job Schedule Failed" }));
        });

      })
      .catch((error) => {
        this.addMessage(error.data.responseBody["odata.error"].message.value);
        this.setState((current) => ({ ...current, highRiskStatus: "Error Uploading file" }));
      });


  }
  public saveHighRiskFile(e: any) {
    debugger;
    let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, highRiskFile: file }));
  }
  // #endregion High Risk Users


  // #region Primary Approvers
  public async esnurePrimaryApprovers(error, data: Array<any>): Promise<any> {

    this.setState((current) => ({ ...current, primaryApproversStatus: "Validating users", primaryApproversTotalRows: data.length }));
    await esnureUsers(pnp.sp.web, this.cachedIds, data, "ApproverEmail", this.addMessage);
    this.addMessage("Done validating Primary Approvers Users");
    this.setState((current) => ({ ...current, primaryApproversStatus: "Completed Validation" }));
    return Promise.resolve();
  }
  public extractColumnHeadersPrimaryApprovers(headerRow: Array<String>): String[] {

    const requiredColumns = ["ApproverEmail"];
    for (let requiredColumn of requiredColumns) {
      if (headerRow.indexOf(requiredColumn) === -1) {
        this.addMessage(`Column ${requiredColumn} is missing on Primary Approvers File`);
      }
    }
    return extractColumnHeaders(headerRow);
  }
  public onPrimaryApproversDataloaded() {
    this.setState((current) => ({ ...current, primaryApproversStatus: "Parsing file" }));
    parse(this.reader.result, { delimiter: ',', columns: this.extractColumnHeadersPrimaryApprovers }, this.esnurePrimaryApprovers);
  }
  public parsePrimaryApproversFile(): Promise<any> {
    this.setState((current) => ({ ...current, primaryApproversStatus: "Reading file" }));
    this.reader = new FileReader();
    this.reader.onload = this.onPrimaryApproversDataloaded;
    this.reader.readAsText(this.state.primaryApproversFile);
    return Promise.resolve();
  }
  public uploadPrimaryApproversFile(): void {
    debugger;
    this.setState((current) => ({ ...current, primaryApproversStatus: "Uploading file" }));
    uploadFile(this.state.newWeb, "Documents", this.state.primaryApproversFile, "Primary Approvers", this.addMessage)
      .then((resp) => {
        this.setState((current) => ({ ...current, primaryApproversStatus: "Uploaded file" }));
        let functionUrl = `${this.props.azurePrimaryApproverUrl}&siteUrl=${this.props.siteAbsoluteUrl}/${this.state.siteName}&listName=Primary Approvers&fileName=Primary Approvers&batchSize=${this.props.batchSize}&pauseBeforeBatchExecution=${this.props.pauseBeforeBatchExecution}`;
        this.addMessage(functionUrl);
        processUploadedFiles(this.props.httpClient, functionUrl).then(() => {
          this.setState((current) => ({ ...current, primaryApproversStatus: "Job Scheduled" }));
        }).catch(e => {
          this.setState((current) => ({ ...current, primaryApproversStatus: "Job Schedule Failed" }));
        });
      })
      .catch((error) => {
        this.addMessage(error.data.responseBody["odata.error"].message.value);
        this.setState((current) => ({ ...current, primaryApproversStatus: "Error Uploading file" }));
      });


  }
  public savePrimaryApproversFile(e: any) {
    debugger;
    let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, primaryApproversFile: file }));
  }
  // #endregion Primary Approvers


  // #region RoleToTcode
  public async esnureRoleToTCode(error, data: Array<any>): Promise<any> {

    this.setState((current) => ({ ...current, roleToTransactionStatus: "Validating users", roleToTransactionTotalRows: data.length }));
    await esnureUsers(pnp.sp.web, this.cachedIds, data, "ApproverEmail", this.addMessage);
    this.addMessage("Done validating Primary Approvers Users");
    this.setState((current) => ({ ...current, roleToTransactionStatus: "Completed Validation" }));
    return Promise.resolve();
  }
  public extractColumnHeadersRoleToTCode(headerRow: Array<String>): String[] {

    const requiredColumns = ["Role", "Composite Role", "TCode", "Transaction Text"];
    for (let requiredColumn of requiredColumns) {
      if (headerRow.indexOf(requiredColumn) === -1) {
        this.addMessage(`Column ${requiredColumn} is missing on Role To Tcode File`);
      }
    }
    return extractColumnHeaders(headerRow);
  }
  public onRoleToTCodeDataloaded() {
    this.setState((current) => ({ ...current, roleToTransactionStatus: "Parsing file" }));
    parse(this.reader.result, { delimiter: ',', columns: this.extractColumnHeadersRoleToTCode }, this.esnureRoleToTCode);
  }
  public parseRoleToTCodeFile(): Promise<any> {
    this.setState((current) => ({ ...current, roleToTransactionStatus: "Reading file" }));
    this.reader = new FileReader();
    this.reader.onload = this.onRoleToTCodeDataloaded;
    this.reader.readAsText(this.state.roleToTransactionFile);
    return Promise.resolve();
  }
  public uploadRoleToTCodeFile(): void {
    debugger;
    this.setState((current) => ({ ...current, roleToTransactionStatus: "Uploading file" }));
    uploadFile(this.state.newWeb, "Documents", this.state.roleToTransactionFile, "Role To Transaction", this.addMessage)
      .then((resp) => {
        this.setState((current) => ({ ...current, roleToTransactionStatus: "Uploaded file" }));
        debugger;
        // the RoleToTransaction parameter name must match the parameter name in the azure function
        let functionUrl = `${this.props.azureRoleToCodeUrl}&siteUrl=${this.props.siteAbsoluteUrl}/${this.state.siteName}&listName=Role To Transaction&fileName=Role To Transaction&batchSize=${this.props.batchSize}&pauseBeforeBatchExecution=${this.props.pauseBeforeBatchExecution}`;
        this.addMessage(functionUrl);
        processUploadedFiles(this.props.httpClient, functionUrl).then(() => {
          this.setState((current) => ({ ...current, roleToTransactionStatus: "Job Scheduled" }));
        }).catch(e => {
          this.setState((current) => ({ ...current, roleToTransactionStatus: "Job Schedule Failed" }));
        });
      })
      .catch((error) => {
        this.addMessage(error.data.responseBody["odata.error"].message.value);
        this.setState((current) => ({ ...current, roleToTransactionStatus: "Error Uploading file" }));
      });


  }
  public saveRoleToTCodeFile(e: any) {
    debugger;
    let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, roleToTransactionFile: file }));
  }
  // #endregion roletotcode



  // #region Site creation

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

    await fixUpLeftNav(webServerRelativeUrl, this.props.siteUrl);

    // create the lists and assign permissions
    let highRiskList: PNPList = await addCustomListWithContentType(newWeb, pnp.sp.web, 'High Risk', "High Risk Transactions"
      , "High Risk", this.addMessage);

    let primaryAppproversList: PNPList = await addCustomListWithContentType(newWeb, pnp.sp.web, 'Primary Approvers', "Primary Approvers of High Risk Transactions"
      , "Primary Approver List", this.addMessage);

    let roleToTransactionList: PNPList = await addCustomListWithContentType(newWeb, pnp.sp.web, 'Role To Transaction', "Role To Transaction DETAILS"
      , "Role To Transaction", this.addMessage);
    await roleToTransactionList.fields.getByInternalNameOrTitle("GRCRole").update({
      Indexed: true
    }).then((resp) => {

    }).catch((error) => {
      debugger;
      console.error(error);
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      return;

    });
    // this is a list that the webjob can use to log messages
    await newWeb.lists.add("Messages", "Messages", 100, false).then(async (listResponse: ListAddResult) => {
      this.addMessage("Created List " + "Messages");
    }).catch(error => {
      debugger;
      console.error(error);
      this.addMessage("Error creating Messages List");
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      return;
    });


    // customize the home paged
    let welcomePageUrl: string;
    await newWeb.rootFolder.getAs<any>().then(rootFolder => {

      welcomePageUrl = rootFolder.ServerRelativeUrl + rootFolder.WelcomePage;
    });
    this.addMessage("Customizing Home Page");
    debugger;
    await cleanupHomePage(webServerRelativeUrl, welcomePageUrl, this.props.webPartXml);
    this.addMessage("Customized Home Page");

    this.addMessage("DONE!!");

  }
  // #endregion Site creation

  //   private processUploadedFiles(): void {
  //     debugger;
  //     //! Can't have spaces ini the URL!!!
  //     // the parameters are the file names we uploaded.
  //     let functionUrl = `${this.props.azureFunctionUrl}
  // &siteUrl=${this.props.siteAbsoluteUrl + "/" + this.state.siteName}
  // &siteType=High Risk
  // &PrimaryApproverList=Primary Approvers
  // &HighRisk=High Risk
  // &RoleToTransaction=Role To Transaction`;
  //     processUploadedFiles(this.props.httpClient, functionUrl);
  //   }
  private addMessage(message: string) {
    let messages = this.state.messages;
    var copy = map(this.state.messages, clone);
    copy.push(message);
    this.setState((current) => ({ ...current, messages: copy }));
  }
  public render(): React.ReactElement<IHighRiskAdminWebpartProps> {

    return (
      <div className={styles.highRiskAdminWebpart} >


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
              Validate
          </th>
            <th>
              Upload
          </th>
          </thead>
          <tr>
            <td>
              High Risk
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
              <IconButton iconProps={{ iconName: "DocumentApproval" }} onClick={this.parseHighRiskFile}>Upload</IconButton>
            </td>
            <td>
              <IconButton iconProps={{ iconName: "Save" }} onClick={this.uploadHighRiskFile}>Upload</IconButton>
            </td>

          </tr>
          <tr>
            <td>
              Primary Approvers
            </td>
            <td>
              <input type="file" id="uploadrttfile" onChange={e => { this.savePrimaryApproversFile(e); }} />
            </td>
            <td>
              {this.state.primaryApproversStatus}
            </td>
            <td>
              {this.state.primaryApproversTotalRows}
            </td>
            <td>
              <IconButton iconProps={{ iconName: "DocumentApproval" }} onClick={this.parsePrimaryApproversFile}></IconButton>
            </td>
            <td>
              <IconButton iconProps={{ iconName: "Save" }} onClick={this.uploadPrimaryApproversFile}></IconButton>
            </td>

          </tr>
          <tr>
            <td>
              Role To TCode
            </td>
            <td>
              <input type="file" id="uploadrttfile" onChange={e => { this.saveRoleToTCodeFile(e); }} />
            </td>
            <td>
              {this.state.roleToTransactionStatus}
            </td>
            <td>
              {this.state.roleToTransactionTotalRows}
            </td>
            <td>
              <IconButton iconProps={{ iconName: "DocumentApproval" }} onClick={this.parseRoleToTCodeFile}></IconButton>
            </td>
            <td>
              <IconButton iconProps={{ iconName: "Save" }} onClick={this.uploadRoleToTCodeFile}></IconButton>
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
