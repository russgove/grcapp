import { find, clone, map } from "lodash";
import pnp, { Web, List, ListAddResult } from "sp-pnp-js";
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
export function addPeopleFieldToList(webUrl: string, listTitle: string, internalName: string, displayName): Promise<any> {
    let web = new Web(window.location.origin + webUrl);
    let fieldXMl = `<Field Type='User' StaticName='${internalName}'  Name='${internalName}'  DisplayName='${displayName}' Required='FALSE' EnforceUniqueValues='FALSE' />`;
    return web.lists.getByTitle(listTitle).fields.createFieldAsXml(fieldXMl);

}

export async function convertEmailColumnsToUser(webUrl: string, listTitle: string, columns: Array<[string, string]>) {
    debugger;
    let web = new Web(window.location.origin + webUrl);
    let fieldsToFetch = ["Id"];
    for (let column of columns) {
        fieldsToFetch.push(column[0]);
        fieldsToFetch.push(column[1]);
    }
    await web.lists.getByTitle(listTitle).items.top(4000).select(fieldsToFetch.join(",")).get()
        .then(async rows => {
            debugger;
            for (let row of rows) {
                for (let column of columns) {
                    let emailColumn: string = column[0];
                    let userColumn: string = column[1];
                    let user = await web.ensureUser(row[column[0]]);
                    let update= {[userColumn]:user.data.Id};// enclose in brakts to eval
                    web.lists.getByTitle(listTitle).items.getById(row["Id"]).update(update)
                        .then(x => {
                            console.log(`Updated user email ${row[column[0]]} To user ID ${user.data.Id} `);
                        })
                        .catch(err => {
                            console.error (`Error updateing user email ${row[column[0]]} To user ID ${user.data.Id} `);
                            debugger;
                        });

                }
            };
        })
        .catch(err => {
            debugger;
        });

    return Promise.resolve();
}
export function ensureFieldsAreInList(list: List, fieldInternalNames: Array<string>): boolean {
    // besure when you called pnp get list you expanded Fields/
  

    let allFound: boolean = true;
    for (let fieldName of fieldInternalNames) {
        if (!find(list["Fields"], f => { return f["InternalName"] === fieldName; })) {
            alert(`Field ${fieldName} was not found in list ${list["Title"]}`);
            allFound = false;
        }
    }
    return allFound;


}
export async function getListFromWeb(webUrl: string, listTitle: string): Promise<List> {
    let web = new Web(window.location.origin + webUrl);
    return web.lists.getByTitle(listTitle).expand("Fields").get();

}
export function extractColumnHeaders(headerRow: Array<String>): String[] {
    let headings: Array<String> = [];
    for (let header of headerRow) {
        headings.push(header.replace(/\s/g, "")); // remove spaces
    }
    return headings;
}
export class CachedId {
    public upn: string;
    public id: number | null;
}
export async function findId(web: Web, cachedIds: Array<CachedId>, upn: string): Promise<number | null> {
    let id: number | null = null;
    let cached: CachedId = find(cachedIds, (cachedId) => { return cachedId.upn === upn; });
    if (cached) {
        return cached.id;
    }
    await web.ensureUser(upn)
        .then((response) => {
            id = response.data.Id;
            cachedIds.push({ upn: upn, id: id });
            return;
        })
        .catch((err) => {
            cachedIds.push({ upn: upn, id: null });
            return;
        });
    return id;
}
export async function esnureUsers(web: Web, cachedIds: Array<CachedId>,
    data: Array<any>, columnName: string, addMessage: (message: string) => void): Promise<any> {
    let rowNumber = 0;
    for (let row of data) {
        rowNumber++;
        let approverId: number = await findId(web, cachedIds, row[columnName]);
        if (!approverId) {
            addMessage(`Approver  ${row.ApproverEmail} on row ${rowNumber} is invalid`);
        }
    }
    return Promise.resolve();
}
export function uploadFile(web: Web, libraryName: string, file: File, saveAsFileName: string, addMessage: (message: string) => void): Promise<any> {

    return web.lists.getByTitle(libraryName).rootFolder.files
        .addChunked(saveAsFileName, file, data => {
            addMessage(`(Stage ${data.stage}) Uploaded block ${data.blockNumber} of ${data.totalBlocks}`);
        }, true)
        .then((results) => {
            return results.file.getItem().then(item => {
                return item.update({ Title: saveAsFileName }).then((r) => {
                    return;
                }).catch((err) => {

                    addMessage(err.data.responseBody["odata.error"].message.value);
                    console.log(err);
                });
            });
        })
        .catch((err) => {
            addMessage(err.data.responseBody["odata.error"].message.value);
            console.log(err);
        });
}
export async function cleanupHomePage(webRelativeUrl: string, homePageUrl, webPartXml: string) {
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
export async function setWebToUseSharedNavigation(webRelativeUrl: string) {
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
export async function AddQuickLaunchItem(webUrl: string, title: string, url: string, isExternal: boolean) {
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
export async function RemoveQuickLaunchItem(webUrl: string, titlesToRemove: string[]) {
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

}
export async function fixUpLeftNav(webUrl: string, homeUrl: string) {

    await AddQuickLaunchItem(webUrl, "EFR Home", homeUrl, true);
    await RemoveQuickLaunchItem(webUrl, ["Pages", "Documents"]);
}
/**
 *  Adds a custom webpart to the page at editformUrl
 * 
 * @param {string} webRelativeUrl -- The web containing the list
 * @param {any} editformUrl -- the url of the editform page
 * @param {string} webPartXml  -- the xml for the webpart to add
 * @memberof EfrAdmin
 */
export async function AddWebPartToEditForm(webRelativeUrl: string, editformUrl, webPartXml: string) {
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
export function getContentTypeByName(web: Web, contentTypeName: string): Promise<any> {

    return web.contentTypes.filter(`Name eq '${contentTypeName}'`).get();
}
export async function addCustomListWithContentType(web: Web, rootweb: Web, listTitle: string,
    listdDescription: string, contentTypeName: string,
    addMessage: (message: string) => void): Promise<List> {

    // content types are on the rootweb
    let contentTypes = await getContentTypeByName(rootweb, contentTypeName);
    if (contentTypes.length !== 1) {
        addMessage(`Error.. content type ${contentTypeName} was not found`);
        addMessage(`List ${listTitle} could not be created`);
        return;
    }

    let list: List;
    await web.lists.add(listTitle, listdDescription, 100, false)
        .then(async (listResponse: ListAddResult) => {
            addMessage("Created List " + listTitle);
            list = listResponse.list;
        }).catch(error => {
  
            console.error(error);
            addMessage(error.data.responseBody["odata.error"].message.value);
            return;
        });
    await list.contentTypes.addAvailableContentType(contentTypes[0].Id.StringValue).then((ct) => {
        addMessage("Added content type" + contentTypeName + " to list ");
        return;
    }).catch(error => {
        debugger;
        console.error(error);
        addMessage(error.data.responseBody["odata.error"].message.value);
        return;
    });
    return Promise.resolve(list);
}
export function processUploadedFiles(httpClient: HttpClient, functionUrl: string): Promise<any> {
 
    // call the azure function to write the message to the queue, to start the webjob to process the files
    //https://grctest.azurewebsites.net/api/HttpTriggerCSharp1?code=HBM82bnia7nKPC/nqVTbaCmfPaFyubQa8iY22lb0r88fdQH370CRUg==&SiteType=%27Role%20to%20Tcode%20Review%27&SiteUrl=%27jwh%27&PrimaryApproverList=%27pal%27&RoleReview=%27rr%27&RoleToTransaction=%27rtt%27
    // query param is SiteType='Role to Tcode Review' SiteUrl='url to the new web' PrimaryApproverList='name of the file we updaded to Documents'  RoleReview='name of the file we updaded to Documents' RoleToTransaction='name of the file we updaded to Documents'
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");

    const postOptions: IHttpClientOptions = {
        headers: requestHeaders,
    };
    return httpClient.get(functionUrl, HttpClient.configurations.v1, postOptions)
        .then((response: HttpClientResponse) => {
            alert('Request queued');
            return;
        })
        .catch((error) => {
            alert('error queuing request');
            return;
        });
}