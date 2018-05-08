import * as React from 'react';
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import styles from './GrcUpload.module.scss';
import { IGrcUploadProps } from './IGrcUploadProps';
import { IGrcUploadState } from './IGrcUploadState';
import { escape } from '@microsoft/sp-lodash-subset';
const parse = require('csv-parse');
import  {sp,  ItemAddResult, ListAddResult, ContextInfo, Web, WebAddResult, List as PNPList } from "@pnp/sp";
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
class CachedId {
  public upn: string;
  public id: number | null;
}
export default class GrcUpload extends React.Component<IGrcUploadProps, IGrcUploadState> {
  private reader;
  private counter = 0;
  private cachedIds: Array<CachedId> = [];
  //#region common code
  constructor(props: IGrcUploadProps) {
    super(props);
    // ROlte To transaction
    this.onRoleToTransactionDataloaded = this.onRoleToTransactionDataloaded.bind(this);
    this.saveRoleToTransactionFile = this.saveRoleToTransactionFile.bind(this);
    this.extractColumnHeadersRoleToTransactionData = this.extractColumnHeadersRoleToTransactionData.bind(this);
    this.uploadRoleToTransactionFile = this.uploadRoleToTransactionFile.bind(this);
    this.parseRoleToTransactionFile = this.parseRoleToTransactionFile.bind(this);
    this.esnureRoleToTransactionUsers = this.esnureRoleToTransactionUsers.bind(this);

    // Primary Approveres
    this.onPrimaryApproversDataloaded = this.onPrimaryApproversDataloaded.bind(this);
    this.savePrimaryApproversFile = this.savePrimaryApproversFile.bind(this);
    this.extractColumnHeadersPrimaryApproversData = this.extractColumnHeadersPrimaryApproversData.bind(this);
    this.uploadPrimaryApproversFile = this.uploadPrimaryApproversFile.bind(this);
    this.parsePrimaryApproversFile = this.parsePrimaryApproversFile.bind(this);
    this.esnurePrimaryApproversUsers = this.esnurePrimaryApproversUsers.bind(this);


    // Role Review
    this.onRoleReviewDataloaded = this.onRoleReviewDataloaded.bind(this);
    this.saveRoleReviewFile = this.saveRoleReviewFile.bind(this);
    this.extractColumnHeadersRoleReviewData = this.extractColumnHeadersRoleReviewData.bind(this);
    this.uploadRoleReviewFile = this.uploadRoleReviewFile.bind(this);

    this.parseRoleReviewFile = this.parseRoleReviewFile.bind(this);
    this.esnureRoleReviewUsers = this.esnureRoleReviewUsers.bind(this);





    this.extractColumnHeaders = this.extractColumnHeaders.bind(this);
    this.processUploadedFiles = this.processUploadedFiles.bind(this);


    this.state = {
      siteName: "",
      newWeb: null,
      newWebUrl: null,


      messages: [],
      roleToTransactionRowsUploaded: 0,
      roleToTransactionTotalRows: 0,
      roleToTransactionStatus: "",
      roleToTransactionFile: null,

      roleReviewTotalRows: 0,
      roleReviewRowsUploaded: 0,
      roleReviewStatus: "",
      roleReviewFile: null,

      primaryApproversTotalRows: 0,
      primaryApproversRowsUploaded: 0,
      primaryApproversStatus: "",
      primaryApproversFile: null,
    };


  }
  public extractColumnHeaders(headerRow: Array<String>): String[] {

    let headings: Array<String> = [];
    for (let header of headerRow) {
      headings.push(header.replace(/\s/g, "")); // remove spaces
    }

    return headings;
  }
  public async findId(upn: string): Promise<number | null> {

    let id: number | null = null;
    let cached: CachedId = find(this.cachedIds, (cachedId) => { return cachedId.upn === upn; });
    if (cached) {
      return cached.id;
    }
    await sp.web.ensureUser(upn)
      .then((response) => {
        id = response.data.Id;
        this.cachedIds.push({ upn: upn, id: id });
        return;
      })
      .catch((err) => {
        this.cachedIds.push({ upn: upn, id: null });
        return;
      });
    return id;
  }

  private addMessage(message: string) {
    let messages = this.state.messages;
    var copy = map(this.state.messages, clone);
    copy.push(message);
    this.setState((current) => ({ ...current, messages: copy }));
  }

  //#endregion

  //#region Role Review 
  public async esnureRoleReviewUsers(error, data: Array<any>): Promise<any> {
    debugger;
    this.setState((current) => ({ ...current, roleReviewStatus: "Validating users" }));
    let rowNumber = 0;
    for (let row of data) {
      rowNumber++;
      let approverId: number = await this.findId(row.ApproverEmail);
      let alternateApproverId: number = await this.findId(row.AlternateApproverEmail);
      if (!approverId) {
        this.addMessage(`Approver  ${row.ApproverEmail} on row ${rowNumber} of the Role To Transaction File is invalid`);
      }
    }
    this.addMessage("Done validating RoleReviewUsers");
    this.setState((current) => ({ ...current, roleReviewStatus: "Completed Validation" }));
    return Promise.resolve();
  }

  public extractColumnHeadersRoleReviewData(headerRow: Array<String>): String[] {
    debugger;
    const requiredColumns = ["ApproverEmail", "AlternateApproverEmail", "Role Name"];
    for (let requiredColumn of requiredColumns) {
      if (headerRow.indexOf(requiredColumn) === -1) {
        this.addMessage(`Column ${requiredColumn} is missing on Role Review Data File`);
      }
    }

    return this.extractColumnHeaders(headerRow);

  } public onRoleReviewDataloaded() {
    this.setState((current) => ({ ...current, roleReviewStatus: "Parsing file" }));
    parse(this.reader.result, { delimiter: ',', columns: this.extractColumnHeadersRoleReviewData }, this.esnureRoleReviewUsers);
  }

  public parseRoleReviewFile(): Promise<any> {

    //https://stackoverflow.com/questions/14446447/how-to-read-a-local-text-file
    //let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, roleReviewStatus: "Reading file" }));
    this.reader = new FileReader();
    this.reader.onload = this.onRoleReviewDataloaded;
    this.reader.readAsText(this.state.roleReviewFile);
    return Promise.resolve();

  }
  public uploadRoleReviewFile(): Promise<any> {
    debugger;
    this.setState((current) => ({ ...current, roleReviewStatus: "Uploading file" }));
    return this.state.newWeb.lists.getByTitle("Documents").rootFolder.files
      .addChunked("Role Review", this.state.roleReviewFile, data => {
        console.log({ data: data, message: "progress" });
        this.addMessage(`(Stage ${data.stage}) Uploaded block ${data.blockNumber} of ${data.totalBlocks}`);

      }, true)
      .then((results) => {
        return results.file.getItem().then(item => {
          return item.update({ Title: "Role Review" }).then((r) => {
            this.setState((current) => ({ ...current, roleReviewStatus: "Upload complete" }));
            return;
          }).catch((err) => {
            debugger;
            this.addMessage(err.data.responseBody["odata.error"].message.value);
            this.setState((current) => ({ ...current, roleReviewStatus: "Upload  Error" }));
            console.log(err);
          });
        });
      })
      .catch((err) => {
        this.addMessage(err.data.responseBody["odata.error"].message.value);
        this.setState((current) => ({ ...current, roleReviewStatus: "Upload  Error" }));
        console.log(err);
      });
  }

  public saveRoleReviewFile(e: any) {

    //https://stackoverflow.com/questions/14446447/how-to-read-a-local-text-file
    let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, roleReviewFile: file }));
  }

  //#endregion


  //#region Role To Transaction
  public async esnureRoleToTransactionUsers(error, data: Array<any>): Promise<any> {
    debugger;
    this.setState((current) => ({ ...current, roleToTransactionStatus: "Validating users" }));
    let rowNumber = 0;
    for (let row of data) {
      rowNumber++;
      let approverId: number = await this.findId(row.ApproverEmail);
      let alternateApproverId: number = await this.findId(row.AlternateApproverEmail);
      if (!approverId) {
        this.addMessage(`Approver  ${row.ApproverEmail} on row {rowNumber} of the Role To Transaction File is invalid`);
      }
    }
    this.addMessage("Done validating RoleToTransactionUsers");
    this.setState((current) => ({ ...current, roleToTransactionStatus: "Completed Validation" }));
    return Promise.resolve();
  }

  public extractColumnHeadersRoleToTransactionData(headerRow: Array<String>): String[] {
    debugger;
    const requiredColumns = ["ApproverEmail", "AlternateApproverEmail", "Role", "Role Name", "TCode", "Transaction Text"];
    for (let requiredColumn of requiredColumns) {
      if (headerRow.indexOf(requiredColumn) === -1) {
        this.addMessage(`Column ${requiredColumn} is missing on Role To Transaction Data File`);
      }
    }

    return this.extractColumnHeaders(headerRow);

  }
  public onRoleToTransactionDataloaded() {
    this.setState((current) => ({ ...current, roleToTransactionStatus: "Parsing file" }));
    parse(this.reader.result, { delimiter: ',', columns: this.extractColumnHeadersRoleToTransactionData }, this.esnureRoleToTransactionUsers);
  }
  public parseRoleToTransactionFile(): Promise<any> {
    this.setState((current) => ({ ...current, roleToTransactionStatus: "Reading file" }));
    this.reader = new FileReader();
    this.reader.onload = this.onRoleToTransactionDataloaded;
    this.reader.readAsText(this.state.roleToTransactionFile);
    return Promise.resolve();
  }
  public uploadRoleToTransactionFile(): Promise<any> {
    debugger;
    this.setState((current) => ({ ...current, roleToTransactionStatus: "Uploading file" }));
    return this.state.newWeb.lists.getByTitle("Documents").rootFolder.files
      .addChunked("Role To Transaction", this.state.roleToTransactionFile, data => {
        console.log({ data: data, message: "progress" });
        this.addMessage(`(Stage ${data.stage}) Uploaded block ${data.blockNumber} of ${data.totalBlocks}`);

      }, true)
      .then((results) => {
        return results.file.getItem().then(item => {
          return item.update({ Title: "Role To Transaction" }).then((r) => {
            this.setState((current) => ({ ...current, roleToTransactionStatus: "Upload complete" }));
            return;
          }).catch((err) => {
            debugger;
            this.addMessage(err.data.responseBody["odata.error"].message.value);
            this.setState((current) => ({ ...current, roleToTransactionStatus: "Upload  Error" }));
            console.log(err);
          });
        });
      })
      .catch((err) => {
        this.addMessage(err.data.responseBody["odata.error"].message.value);
        this.setState((current) => ({ ...current, roleToTransactionStatus: "Upload  Error" }));
        console.log(err);
      });
  }

  public saveRoleToTransactionFile(e: any) {

    //https://stackoverflow.com/questions/14446447/how-to-read-a-local-text-file
    let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, roleToTransactionFile: file }));
  }
  //#endregion

  //#region PrimaryApprovers
  public async esnurePrimaryApproversUsers(error, data: Array<any>): Promise<any> {
    debugger;
    this.setState((current) => ({ ...current, primaryApproversStatus: "Validating Users" }));
    let rowNumber = 0;
    for (let row of data) {
      rowNumber++;
      let approverId: number = await this.findId(row.ApproverEmail);
      let alternateApproverId: number = await this.findId(row.AlternateApproverEmail);
      if (!approverId) {
        this.addMessage(`Approver  ${row.ApproverEmail} on row {rowNumber} of the Primary Approver File is invalid`);
      }
    }
    this.addMessage("Done validating PrimaryApproversUsers");
    this.setState((current) => ({ ...current, primaryApproversStatus: "Completed Validation" }));

    return Promise.resolve();
  }
  public extractColumnHeadersPrimaryApproversData(headerRow: Array<String>): String[] {
    debugger;
    const requiredColumns = ["ApproverEmail"];
    for (let requiredColumn of requiredColumns) {
      if (headerRow.indexOf(requiredColumn) === -1) {
        this.addMessage(`Column ${requiredColumn} is missing on Role To Transaction Data File`);
      }
    }

    return this.extractColumnHeaders(headerRow);

  }
  public onPrimaryApproversDataloaded() {
    this.setState((current) => ({ ...current, primaryApproversStatus: "Parsing file" }));
    parse(this.reader.result, { delimiter: ',', columns: this.extractColumnHeadersPrimaryApproversData }, this.esnurePrimaryApproversUsers);
  }

  public parsePrimaryApproversFile(): Promise<any> {

    //https://stackoverflow.com/questions/14446447/how-to-read-a-local-text-file
    //let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, primaryApproversStatus: "Reading file" }));
    this.reader = new FileReader();
    this.reader.onload = this.onPrimaryApproversDataloaded;
    this.reader.readAsText(this.state.primaryApproversFile);
    return Promise.resolve();

  }

  public uploadPrimaryApproversFile(): Promise<any> {
    debugger;
    this.setState((current) => ({ ...current, primaryApproversStatus: "Uploading file" }));
    return this.state.newWeb.lists.getByTitle("Documents").rootFolder.files
      .addChunked("Primary Approvers List", this.state.primaryApproversFile, data => {
        console.log({ data: data, message: "progress" });
        this.addMessage(`(Stage ${data.stage}) Uploaded block ${data.blockNumber} of ${data.totalBlocks}`);

      }, true)
      .then((results) => {
        return results.file.getItem().then(item => {
          return item.update({ Title: "Primary Approvers List" }).then((r) => {
            this.setState((current) => ({ ...current, primaryApproversStatus: "Upload complete" }));
            return;
          }).catch((err) => {
            debugger;
            this.addMessage(err.data.responseBody["odata.error"].message.value);
            this.setState((current) => ({ ...current, primaryApproversStatus: "Upload  Error" }));
            console.log(err);
          });
        });
      })
      .catch((err) => {
        this.addMessage(err.data.responseBody["odata.error"].message.value);
        this.setState((current) => ({ ...current, primaryApproversStatus: "Upload  Error" }));
        console.log(err);
      });
  }
  public savePrimaryApproversFile(e: any) {
    debugger;
    //https://stackoverflow.com/questions/14446447/how-to-read-a-local-text-file
    let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, primaryApproversFile: file }));
  }

  //#endregion

  //#region Site  Creation Methods
  /**
     *  Adds a custom webpart to the edit form located at editformUrl
     * 
     * @param {string} webRelativeUrl -- The web 
     * @param {any} homePageUrl -- the url of the  page
     * @param {string} webPartXml  -- the xml for the webpart to add
     * @memberof EfrAdmin
     */
  public async CleanupHomePage(webRelativeUrl: string, homePageUrl, webPartXml: string) {
    const clientContext: SP.ClientContext = new SP.ClientContext(webRelativeUrl);
    var oFile = clientContext.get_web().getFileByServerRelativeUrl(homePageUrl);

    var limitedWebPartManager = oFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
    let webparts = limitedWebPartManager.get_webParts();
    clientContext.load(webparts, 'Include(WebPart)');
    clientContext.load(limitedWebPartManager);
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });
    let count = webparts.get_count();
    for (let i = 0; i < count; i++) {
      let originalWebPartDef = webparts.get_item(i);
      originalWebPartDef.deleteWebPart();
    }
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        console.log("the webpartw were deleted hidden");
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });

    let oWebPartDefinition = limitedWebPartManager.importWebPart(webPartXml);
    let oWebPart = oWebPartDefinition.get_webPart();

    limitedWebPartManager.addWebPart(oWebPart, 'Main', 1);

    clientContext.load(oWebPart);

    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        console.log("the new webpart was added");
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });
  }

  public async SetWebToUseSharedNavigation(webRelativeUrl: string) {

    const clientContext: SP.ClientContext = new SP.ClientContext(webRelativeUrl);
    var currentWeb = clientContext.get_web();
    var navigation = currentWeb.get_navigation();
    navigation.set_useShared(true);
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        console.log("the web was set to use shared navigation");
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });
  }
  public async AddQuickLaunchItem(webUrl: string, title: string, url: string, isExternal: boolean) {
    let nnci: SP.NavigationNodeCreationInformation = new SP.NavigationNodeCreationInformation();
    nnci.set_title(title);
    nnci.set_url(url);
    nnci.set_isExternal(isExternal);
    const clientContext: SP.ClientContext = new SP.ClientContext(webUrl);
    const web = clientContext.get_web();
    web.get_navigation().get_quickLaunch().add(nnci);
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });

  }
  public async RemoveQuickLaunchItem(webUrl: string, titlesToRemove: string[]) {
    const clientContext: SP.ClientContext = new SP.ClientContext(webUrl);
    const ql: SP.NavigationNodeCollection = clientContext.get_web().get_navigation().get_quickLaunch();
    clientContext.load(ql);
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });
    debugger;
    let itemsToDelete = [];
    let itemCount = ql.get_count();
    for (let x = 0; x < itemCount; x++) {
      let item = ql.getItemAtIndex(x);
      let itemtitle = item.get_title();
      if (titlesToRemove.indexOf(itemtitle) !== -1) {
        itemsToDelete.push(item);
      }
    }
    for (let item of itemsToDelete) {
      item.deleteObject();
    }
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });
    debugger;

  }

  public async fixUpLeftNav(webUrl: string, homeUrl: string) {
    debugger;
    await this.AddQuickLaunchItem(webUrl, "EFR Home", homeUrl, true);
    await this.RemoveQuickLaunchItem(webUrl, ["Pages", "Documents"]);

  }
  /**
   *  Adds a custom webpart to the edit form located at editformUrl
   * 
   * @param {string} webRelativeUrl -- The web containing the list
   * @param {any} editformUrl -- the url of the editform page
   * @param {string} webPartXml  -- the xml for the webpart to add
   * @memberof EfrAdmin
   */
  public async AddWebPartToEditForm(webRelativeUrl: string, editformUrl, webPartXml: string) {
    const clientContext: SP.ClientContext = new SP.ClientContext(webRelativeUrl);
    var oFile = clientContext.get_web().getFileByServerRelativeUrl(editformUrl);

    var limitedWebPartManager = oFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
    let webparts = limitedWebPartManager.get_webParts();
    clientContext.load(webparts, 'Include(WebPart)');
    clientContext.load(limitedWebPartManager);
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });
    let originalWebPartDef = webparts.get_item(0);
    let originalWebPart = originalWebPartDef.get_webPart();
    originalWebPart.set_hidden(true);
    originalWebPartDef.saveWebPartChanges();
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        console.log("the webpart was hidden");
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });

    let oWebPartDefinition = limitedWebPartManager.importWebPart(webPartXml);
    let oWebPart = oWebPartDefinition.get_webPart();

    limitedWebPartManager.addWebPart(oWebPart, 'Main', 1);

    clientContext.load(oWebPart);

    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        console.log("the new webpart was added");
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });
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
    await sp.site.getContextInfo().then((context: ContextInfo) => {
      contextInfo = context;
    });
    // create the site
    await sp.web.webs.add(this.state.siteName, this.state.siteName, this.state.siteName, this.props.templateName).then((war: WebAddResult) => {
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

    await this.SetWebToUseSharedNavigation(webServerRelativeUrl);
    debugger;
    await this.fixUpLeftNav(webServerRelativeUrl, this.props.siteUrl);

    // create the libraries and assign permissions
    let primaryApproverList: PNPList;
    await newWeb.lists.add("Primary Approver", "Primary Approver", 100, false).then(async (listResponse: ListAddResult) => {
      this.addMessage("Created List " + "Primary Approver");
      primaryApproverList = listResponse.list;
    }).catch(error => {
      debugger;
      console.error(error);
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      return;
    });
    await primaryApproverList.contentTypes.addAvailableContentType(this.props.primaryApproverContentTypeId).then((ct) => {
      this.addMessage("Added Primary Approver content type");
      return;
    }).catch(error => {
      debugger;
      console.error(error);
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      return;
    });
    let roleReviewList: PNPList;
    await newWeb.lists.add("Role Review", "Role Review", 100, false).then(async (listResponse: ListAddResult) => {
      this.addMessage("Created List " + "Role Review");
      roleReviewList = listResponse.list;
    }).catch(error => {
      debugger;
      console.error(error);
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      return;
    });
    await roleReviewList.contentTypes.addAvailableContentType(this.props.roleReviewContentTypeId).then((ct) => {
      this.addMessage("Added roleReviewList content type");
      return;
    }).catch(error => {
      debugger;
      console.error(error);
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      return;
    });
    let roleToTransactionList: PNPList;
    await newWeb.lists.add("Role To Transaction", "Role To Transaction", 100, false).then(async (listResponse: ListAddResult) => {
      this.addMessage("Created List " + "Role To Transaction");
      roleToTransactionList = listResponse.list;
    }).catch(error => {
      debugger;
      console.error(error);
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      return;
    });
    await roleToTransactionList.contentTypes.addAvailableContentType(this.props.roleToTransactionContentTypeId).then((ct) => {
      this.addMessage("Added Role To Transaction content type");
      return;
    }).catch(error => {
      debugger;
      console.error(error);
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      return;
    });
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
    await newWeb.rootFolder.get<any>().then(rootFolder => {
      debugger;
      welcomePageUrl = rootFolder.ServerRelativeUrl + rootFolder.WelcomePage;
    });
    this.addMessage("Customizing Home Page");
    await this.CleanupHomePage(webServerRelativeUrl, welcomePageUrl, this.props.webPartXml);
    this.addMessage("Customized Home Page");

    this.addMessage("DONE!!");

  }
  //#endregion


  private _onRenderMessage(item: any, index: number, isScrolling: boolean): JSX.Element {

    return (
      <div>{item}</div>
    );
  }
  private processUploadedFiles(): void {
    debugger;
    // call the azure function to write the message to the queue, to start the webjob to process the files
    //https://grctest.azurewebsites.net/api/HttpTriggerCSharp1?code=HBM82bnia7nKPC/nqVTbaCmfPaFyubQa8iY22lb0r88fdQH370CRUg==&SiteType=%27Role%20to%20Tcode%20Review%27&SiteUrl=%27jwh%27&PrimaryApproverList=%27pal%27&RoleReview=%27rr%27&RoleToTransaction=%27rtt%27
    // query param is SiteType='Role to Tcode Review' SiteUrl='url to the new web' PrimaryApproverList='name of the file we updaded to Documents'  RoleReview='name of the file we updaded to Documents' RoleToTransaction='name of the file we updaded to Documents'
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");

    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
    };
    let functionUrl = `${this.props.azureFunctionUrl}&siteUrl=${this.props.siteAbsoluteUrl + "/"+this.state.siteName}
    &siteType=Role To Ttansaction&PrimaryApproverList=Primary Approvers List&RoleReview=Role Review&RoleToTransaction=Role To Transaction`;
    this.props.httpClient.get(functionUrl, HttpClient.configurations.v1, postOptions)
      .then((response: HttpClientResponse) => {
        alert('Request queued');
      })
      .catch((error) => {
        alert('error queuing request');
      });
  }
  public render(): React.ReactElement<IGrcUploadProps> {

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
              <input type="file" id="uploadrttfile" onChange={e => { this.saveRoleReviewFile(e); }} />
            </td>
            <td>
              {this.state.roleReviewStatus}
            </td>
            <td>
              {this.state.roleReviewTotalRows}
            </td>
            <td>
              {this.state.roleReviewRowsUploaded}
            </td>
            <td>
              <IconButton iconProps={{ iconName: "DocumentApproval" }} onClick={this.parseRoleReviewFile}>Upload</IconButton>
              <IconButton iconProps={{ iconName: "Save" }} onClick={this.uploadRoleReviewFile}>Upload</IconButton>
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
              {this.state.primaryApproversRowsUploaded}
            </td>
            <td>
              <IconButton iconProps={{ iconName: "DocumentApproval" }} onClick={this.parsePrimaryApproversFile}>Upload</IconButton>
              <IconButton iconProps={{ iconName: "Save" }} onClick={this.uploadPrimaryApproversFile}>Upload</IconButton>
            </td>

          </tr>
          <tr>
            <td>
              Role To Transaction
            </td>
            <td>
              <input type="file" id="uploadrttfile" onChange={e => { this.saveRoleToTransactionFile(e); }} />
            </td>
            <td>
              {this.state.roleToTransactionStatus}
            </td>
            <td>
              {this.state.roleToTransactionTotalRows}
            </td>
            <td>
              {this.state.roleToTransactionRowsUploaded}
            </td>
            <td>
              <IconButton iconProps={{ iconName: "DocumentApproval" }} onClick={this.parseRoleToTransactionFile}>Upload</IconButton>
              <IconButton iconProps={{ iconName: "Save" }} onClick={this.uploadRoleToTransactionFile}>Upload</IconButton>
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
            onRenderCell={this._onRenderMessage} />

        </div>
      </div >
    );
  }
}
