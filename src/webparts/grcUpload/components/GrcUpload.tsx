import * as React from 'react';
import styles from './GrcUpload.module.scss';
import { IGrcUploadProps } from './IGrcUploadProps';
import { IGrcUploadState } from './IGrcUploadState';
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
class CachedId {
  public upn: string;
  public id: number | null;
}
export default class GrcUpload extends React.Component<IGrcUploadProps, IGrcUploadState> {
  private reader;
  private cachedIds: Array<CachedId> = [];
  //#region common code
  constructor(props: IGrcUploadProps) {
    super(props);
    // ROlte To transaction
    this.uploadRoleToTransactionBatch = this.uploadRoleToTransactionBatch.bind(this);
    this.onRoleToTransactionDataParsed = this.onRoleToTransactionDataParsed.bind(this);
    this.onRoleToTransactionDataloaded = this.onRoleToTransactionDataloaded.bind(this);
    this.saveRoleToTransactionFile = this.saveRoleToTransactionFile.bind(this);
    this.extractColumnHeadersRoleToTransactionData = this.extractColumnHeadersRoleToTransactionData.bind(this);
    this.validateRoleToTransactionUsers = this.validateRoleToTransactionUsers.bind(this);
    this.processRoleToTransaction = this.processRoleToTransaction.bind(this);

    // Primar7y Approveres
    this.uploadPrimaryApproversBatch = this.uploadPrimaryApproversBatch.bind(this);
    this.onPrimaryApproversDataParsed = this.onPrimaryApproversDataParsed.bind(this);
    this.onPrimaryApproversDataloaded = this.onPrimaryApproversDataloaded.bind(this);
    this.savePrimaryApproversFile = this.savePrimaryApproversFile.bind(this);
    this.extractColumnHeadersPrimaryApproversData = this.extractColumnHeadersPrimaryApproversData.bind(this);
    this.validatePrimaryApproverUsers = this.validatePrimaryApproverUsers.bind(this);
    this.processPrimaryApprovers = this.processPrimaryApprovers.bind(this);


    // Role Review
    this.uploadRoleReviewBatch = this.uploadRoleReviewBatch.bind(this);
    this.onRoleReviewDataParsed = this.onRoleReviewDataParsed.bind(this);
    this.onRoleReviewDataloaded = this.onRoleReviewDataloaded.bind(this);
    this.saveRoleReviewFile = this.saveRoleReviewFile.bind(this);
    this.extractColumnHeadersRoleReviewData = this.extractColumnHeadersRoleReviewData.bind(this);
    this.validateRoleReviewUsers = this.validateRoleReviewUsers.bind(this);
    this.processRoleReview = this.processRoleReview.bind(this);





    this.extractColumnHeaders = this.extractColumnHeaders.bind(this);


    this.state = {
      siteName: "",
      newWeb: null,
      process: "",

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
    await pnp.sp.web.ensureUser(upn)
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
  public async test() {

    // Scoping to web, pnp.sp.createBatch() works as well
    const batch = pnp.sp.web.createBatch();
    const list = await pnp.sp.web.lists.getByTitle("CustomList").get();
    // It's better get entity type separetly and pass it into item add method, 
    // so no additional requests will be sent to get entity type for each item
    debugger;
    const entityType = await list.ListItemEntityTypeFullName;
    const list2 = await pnp.sp.web.getList('/sites/GRCTest/test/Lists/CustomList');
    console.log(list2);
    debugger;
    for (let i = 0, len = 2; i < len; i += 1) {
      list2.items.inBatch(batch).add({
        Title: `Item ${i}`
      }, entityType).catch((err) => {
        console.error(err);
      });
    }
    // Promise.all should not be used together with requests in batches
    debugger;
    await batch.execute(); // Batch execute response doesn't contain responses
    // responces can be received in a specific requests promises resolutions
    // even if no phisical requests were not send

    console.log('Done');



    // debugger;
    // var batch = pnp.sp.createBatch();
    // var promises = [];

    // for (var i = 0, len = 25; i < len; i += 1) {
    //   promises.push(pnp.sp.web.lists.getByTitle('CustomList').items.inBatch(batch).add({ "Title": "Title " + i }));
    // }

    // Promise.all(promises).then(function () {
    //   console.log("Batch items creation is completed");
    // });

    // await batch.execute().then((x) => {
    //   debugger;
    // }).catch((err) => {
    //   debugger;
    // });


    // // without promise
    // debugger;
    // var batch2 = pnp.sp.createBatch();


    // for (var i = 0, len = 25; i < len; i += 1) {
    //   pnp.sp.web.lists.getByTitle('CustomList').items.inBatch(batch2).add({ "Title": "Title " + i });
    // }


    // await batch2.execute().then((x) => {
    //   debugger;
    // }).catch((err) => {
    //   debugger;
    // });


    // pnp.sp.web.lists.getByTitle("CustomList").items.add({
    //   Title: "",
    //   ContentTypeId: "0x0100B6ECFC98573CF04EB8FF9C888965431000743C05350624064F94BF4711E93C3A7D",

    // }).then((iar: ItemAddResult) => {
    //   console.log(iar);
    // }).catch((err) => {
    //   debugger;
    //   return;

    // });
  }
  private addMessage(message: string) {
    let messages = this.state.messages;
    var copy = map(this.state.messages, clone);
    copy.push(message);
    this.setState((current) => ({ ...current, messages: copy }));
  }

  //#endregion

  //#region Role Review 
  public async esnureRoleReviewUsers(data: Array<any>): Promise<any> {
    debugger;
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
  public async uploadRoleReviewBatch(rows: Array<any>, entityTypeFullName): Promise<any> {

    let batch = pnp.sp.createBatch();
    for (let row of rows) {
      let approverId: number = await this.findId(row.ApproverEmail);
      let alternateApproverId: number = await this.findId(row.AlternateApproverEmail);

      let obj = {
        Title: "",
        GRCRoleName: row.RoleName,
      };
      if (approverId != null) {
        obj["GRCApproverId"] = approverId;
      }
      if (alternateApproverId != null) {
        obj["GRCAlternateApproverId"] = alternateApproverId;
      }

      //add an item to the list
      await this.state.newWeb.lists.getByTitle('Role Review').items.add(obj, entityTypeFullName)
        // pnp.sp.web.lists.getByTitle('Role Review').items.inBatch(batch).add(obj, entityTypeFullName)
        .then((iar: ItemAddResult) => {

          return;
        })
        .catch((err) => {
          debugger;
          console.error(err);
          console.error(obj);
          return;

        });
    }

    return batch.execute();

  }
  public async uploadRoleReviewData(data: Array<any>): Promise<any> {
    let entityTypeFullName = await this.state.newWeb.lists.getByTitle('Role Review').getListItemEntityTypeFullName();
    var batchSize = 100;
    var batches = Math.ceil(data.length / batchSize);
    for (var i = 0; i < batches; i++) {
      var thisBatchItems = data.slice(i * batchSize, ((i * batchSize) + batchSize));
      await this.uploadRoleReviewBatch(thisBatchItems, entityTypeFullName)
        .then(resp => {
          this.setState((current) => ({ ...current, roleReviewRowsUploaded: current.roleReviewRowsUploaded + thisBatchItems.length }));
        })
        .catch(err => {
          debugger;
        });

    }
    this.setState((current) => ({ ...current, roleReviewStatus: "Upload Complete" }));
    return Promise.resolve();
  }
  public onRoleReviewDataParsed(error, data: Array<any>) {
    debugger;

    switch (this.state.process) {
      case "Uploading":
        this.setState((current) => ({ ...current, roleReviewStatus: "Uploading file", roleReviewTotalRows: data.length }));
        this.uploadRoleReviewData(data);
        break;
      case "Validating":
        this.setState((current) => ({ ...current, roleReviewStatus: "Validating", roleReviewTotalRows: data.length }));

        this.esnureRoleReviewUsers(data);
        break;
      default:
        alert("Invalid Process");

    }

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

  }
  public onRoleReviewDataloaded() {
    this.setState((current) => ({ ...current, roleReviewStatus: "Parsing file" }));
    parse(this.reader.result, { delimiter: ',', columns: this.extractColumnHeadersRoleReviewData }, this.onRoleReviewDataParsed);
  }

  public uploadRoleReviewFile(): Promise<any> {

    //https://stackoverflow.com/questions/14446447/how-to-read-a-local-text-file
    //let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, roleReviewStatus: "Reading file" }));
    this.reader = new FileReader();
    this.reader.onload = this.onRoleReviewDataloaded;
    this.reader.readAsText(this.state.roleReviewFile);
    return Promise.resolve();

  }
  public saveRoleReviewFile(e: any) {

    //https://stackoverflow.com/questions/14446447/how-to-read-a-local-text-file
    let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, roleReviewFile: file }));
  }
  public validateRoleReviewUsers(e: any) {

    this.setState((current) => ({ ...current, process: "Validating" }));// this determinse what runs after we parse the data
    debugger;
    this.uploadRoleReviewFile();

  }
  public processRoleReview(e: any) {

    this.setState((current) => ({ ...current, process: "Uploading" }));// this determinse what runs after we parse the data
    debugger;
    this.uploadRoleReviewFile();
  }


  //#endregion


  //#region Role To Transaction
  public async esnureRoleToTransactionUsers(data: Array<any>): Promise<any> {
    debugger;
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
  public async uploadRoleToTransactionBatch(rows: Array<any>, entityTypeFullName): Promise<any> {

    let batch = pnp.sp.createBatch();
    for (let row of rows) {
      let approverId: number = await this.findId(row.ApproverEmail);
      let alternateApproverId: number = await this.findId(row.AlternateApproverEmail);

      let obj = {
        Title: "",
        GRCRole: row.Role,
        GRCRoleName: row.RoleName,
        GRCTCode: row.TCode,
        GRCTransactionText: row.TransactionText,
      };
      if (approverId != null) {
        obj["GRCApproverId"] = approverId;
      }
      if (alternateApproverId != null) {
        obj["GRCAlternateApproverId"] = alternateApproverId;
      }

      //add an item to the list
      await this.state.newWeb.lists.getByTitle('Role to Transaction').items.add(obj, entityTypeFullName)
        // pnp.sp.web.lists.getByTitle('Role to Transaction').items.inBatch(batch).add(obj, entityTypeFullName)
        .then((iar: ItemAddResult) => {

          return;
        })
        .catch((err) => {
          debugger;
          console.error(err);
          console.error(obj);
          return;

        });
    }

    return batch.execute();

  }
  public async uploadRoleToTransactionData(data: Array<any>): Promise<any> {
    let entityTypeFullName = await this.state.newWeb.lists.getByTitle('Role to Transaction').getListItemEntityTypeFullName();
    var batchSize = 100;
    var batches = Math.ceil(data.length / batchSize);
    batches = 5;
    this.addMessage("NOTE!!!! Role toi transaction currently limited to 500 items");

    for (var i = 0; i < batches; i++) {
      var thisBatchItems = data.slice(i * batchSize, ((i * batchSize) + batchSize));
      await this.uploadRoleToTransactionBatch(thisBatchItems, entityTypeFullName)
        .then(resp => {
          this.setState((current) => ({ ...current, roleToTransactionRowsUploaded: current.roleToTransactionRowsUploaded + thisBatchItems.length }));
        })
        .catch(err => {
          debugger;
        });

    }
    this.setState((current) => ({ ...current, roleToTransactionStatus: "Upload Complete" }));
    return Promise.resolve();
  }
  public onRoleToTransactionDataParsed(error, data: Array<any>) {
    debugger;
    switch (this.state.process) {
      case "Uploading":
        this.setState((current) => ({ ...current, roleToTransactionStatus: "Uploading file", roleToTransactionTotalRows: data.length }));

        this.uploadRoleToTransactionData(data);
        break;
      case "Validating":
        this.setState((current) => ({ ...current, roleToTransactionStatus: "Validating", roleToTransactionTotalRows: data.length }));

        this.esnureRoleToTransactionUsers(data);
        break;
      default:
        alert("Invalid Process");

    }

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
    parse(this.reader.result, { delimiter: ',', columns: this.extractColumnHeadersRoleToTransactionData }, this.onRoleToTransactionDataParsed);
  }

  public uploadRoleToTransactionFile(): Promise<any> {

    //https://stackoverflow.com/questions/14446447/how-to-read-a-local-text-file
    //let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, roleToTransactionStatus: "Reading file" }));
    this.reader = new FileReader();
    this.reader.onload = this.onRoleToTransactionDataloaded;
    this.reader.readAsText(this.state.roleToTransactionFile);
    return Promise.resolve();

  }
  public saveRoleToTransactionFile(e: any) {

    //https://stackoverflow.com/questions/14446447/how-to-read-a-local-text-file
    let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, roleToTransactionFile: file }));
  }
  public validateRoleToTransactionUsers(e: any) {

    this.setState((current) => ({ ...current, process: "Validating" }));// this determinse what runs after we parse the data
    debugger;
    this.uploadRoleToTransactionFile();

  }

  public processRoleToTransaction(e: any) {

    this.setState((current) => ({ ...current, process: "Uploading" }));// this determinse what runs after we parse the data
    debugger;
    this.uploadRoleToTransactionFile();
  }

  //#endregion




  //#region PrimaryApprovers
  public async esnurePrimaryApproversUsers(data: Array<any>): Promise<any> {
    debugger;
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
  public async uploadPrimaryApproversBatch(rows: Array<any>, entityTypeFullName): Promise<any> {

    let batch = pnp.sp.createBatch();
    for (let row of rows) {
      let approverId: number = await this.findId(row.ApproverEmail);


      let obj = {
        Title: "",
      };
      if (approverId != null) {
        obj["GRCApproverId"] = approverId;
      }

      //add an item to the list
      await this.state.newWeb.lists.getByTitle('Primary Approver').items.add(obj, entityTypeFullName)
        // pnp.sp.web.lists.getByTitle('EPA Role to Transaction').items.inBatch(batch).add(obj, entityTypeFullName)
        .then((iar: ItemAddResult) => {

          return;
        })
        .catch((err) => {
          debugger;
          console.error(err);
          console.error(obj);
          return;

        });
    }

    return batch.execute();

  }
  public async uploadPrimaryApproversData(data: Array<any>): Promise<any> {
    let entityTypeFullName = await this.state.newWeb.lists.getByTitle('Primary Approver').getListItemEntityTypeFullName();
    var batchSize = 100;
    var batches = Math.ceil(data.length / batchSize);
    for (var i = 0; i < batches; i++) {
      var thisBatchItems = data.slice(i * batchSize, ((i * batchSize) + batchSize));
      await this.uploadPrimaryApproversBatch(thisBatchItems, entityTypeFullName)
        .then(resp => {
          this.setState((current) => ({ ...current, primaryApproversRowsUploaded: current.primaryApproversRowsUploaded + thisBatchItems.length }));
        })
        .catch(err => {
          debugger;
        });

    }
    this.setState((current) => ({ ...current, primaryApproversStatus: "Upload Complete" }));

    return Promise.resolve();
  }
  public onPrimaryApproversDataParsed(error, data: Array<any>) {
    debugger;
    switch (this.state.process) {
      case "Uploading":
        this.setState((current) => ({ ...current, primaryApproversStatus: "Uploading file", primaryApproversTotalRows: data.length }));
        this.uploadPrimaryApproversData(data);
        break;
      case "Validating":
        this.setState((current) => ({ ...current, primaryApproversStatus: "Validating", primaryApproversTotalRows: data.length }));
        this.esnurePrimaryApproversUsers(data);
        break;
      default:
        alert("Invalid Process");

    }

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
    parse(this.reader.result, { delimiter: ',', columns: this.extractColumnHeadersPrimaryApproversData }, this.onPrimaryApproversDataParsed);
  }

  public uploadPrimaryApproversFile(): Promise<any> {

    //https://stackoverflow.com/questions/14446447/how-to-read-a-local-text-file
    //let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, primaryApproversStatus: "Reading file" }));
    this.reader = new FileReader();
    this.reader.onload = this.onPrimaryApproversDataloaded;
    this.reader.readAsText(this.state.primaryApproversFile);
    return Promise.resolve();

  }
  public savePrimaryApproversFile(e: any) {
    debugger;
    //https://stackoverflow.com/questions/14446447/how-to-read-a-local-text-file
    let file: File = e.target["files"][0];
    this.setState((current) => ({ ...current, primaryApproversFile: file }));
  }
  public validatePrimaryApproverUsers(e: any) {

    this.setState((current) => ({ ...current, process: "Validating" }));// this determinse what runs after we parse the data
    debugger;
    this.uploadPrimaryApproversFile();

  }
  public processPrimaryApprovers(e: any) {

    this.setState((current) => ({ ...current, process: "Uploading" }));// this determinse what runs after we parse the data
    debugger;
    this.uploadPrimaryApproversFile();
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
      this.setState((current) => ({ ...current, newWeb: newWeb })); /// save in state so file uploads can use it

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
    });
    await primaryApproverList.contentTypes.addAvailableContentType(this.props.primaryApproverContentTypeId).then((ct) => {
      this.addMessage("Added Primary Approver content type");
      return;
    }).catch(error => {
      debugger;
      return;
    });
    let roleReviewList: PNPList;
    await newWeb.lists.add("Role Review", "Role Review", 100, false).then(async (listResponse: ListAddResult) => {
      this.addMessage("Created List " + "Role Review");
      roleReviewList = listResponse.list;
    });
    await roleReviewList.contentTypes.addAvailableContentType(this.props.roleReviewContentTypeId).then((ct) => {
      this.addMessage("Added roleReviewList content type");
      return;
    }).catch(error => {
      debugger;
      return;
    });
    let roleToTransactionList: PNPList;
    await newWeb.lists.add("Role To Transaction", "Role To Transaction", 100, false).then(async (listResponse: ListAddResult) => {
      this.addMessage("Created List " + "Role To Transaction");
      roleToTransactionList = listResponse.list;
    });
    await roleToTransactionList.contentTypes.addAvailableContentType(this.props.roleToTransactionContentTypeId).then((ct) => {
      this.addMessage("Added Role To Transaction content type");
      return;
    }).catch(error => {
      debugger;
      return;
    });
    await roleToTransactionList.fields.getByInternalNameOrTitle("GRCRoleName").update({
      Indexed: true
    }).then(function (resp) {
      debugger;
    }).catch(function (err) {
      debugger;
    });
    // customize the home paged
    let welcomePageUrl: string;
    await newWeb.rootFolder.getAs<any>().then(rootFolder => {
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
              <IconButton iconProps={{ iconName: "DocumentApproval" }} onClick={this.validateRoleReviewUsers}>Upload</IconButton>
              <IconButton iconProps={{ iconName: "Save" }} onClick={this.processRoleReview}>Upload</IconButton>
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
              <IconButton iconProps={{ iconName: "DocumentApproval" }} onClick={this.validatePrimaryApproverUsers}>Upload</IconButton>
              <IconButton iconProps={{ iconName: "Save" }} onClick={this.processPrimaryApprovers}>Upload</IconButton>
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
              <IconButton iconProps={{ iconName: "DocumentApproval" }} onClick={this.validateRoleToTransactionUsers}>Upload</IconButton>
              <IconButton iconProps={{ iconName: "Save" }} onClick={this.processRoleToTransaction}>Upload</IconButton>
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
              {this.state.siteName}
            </td>

          </tr>

        </table>
        <button onClick={this.test}>test</button>
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
