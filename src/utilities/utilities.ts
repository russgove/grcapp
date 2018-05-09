import { find, clone, map } from "lodash";
import { sp, Web, List, ListAddResult } from "@pnp/sp";
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
export function addPeopleFieldToList(webUrl: string, listTitle: string, internalName: string, displayName): Promise<any> {
    let web = new Web(window.location.origin + webUrl);
    let fieldXMl = `<Field Type='User' StaticName='${internalName}'  Name='${internalName}'  DisplayName='${displayName}' Required='FALSE' EnforceUniqueValues='FALSE' />`;
    return web.lists.getByTitle(listTitle).fields.createFieldAsXml(fieldXMl);

}

export async function convertEmailColumnsToUser(webUrl: string, listTitle: string, columns: Array<[string, string]>, addMessage: (message: string) => void) {
    debugger;
    let web = new Web(window.location.origin + webUrl);
    let fieldsToFetch = ["Id"];
    for (let column of columns) {
        fieldsToFetch.push(column[0]);
        fieldsToFetch.push(column[1]);
    }
    let moreRows: boolean = true;
    let lastId: number = 0;
    let batchsize: number = 200;
    while (moreRows) {
        await web.lists.getByTitle(listTitle).items.top(batchsize).filter(`ID gt ${lastId}`).select(fieldsToFetch.join(",")).get()
            .then(async rows => {
                debugger;
                if (rows.length === 0) {
                    moreRows = false;
                }
                for (let row of rows) {
                   
                    lastId = row.Id;
                    for (let column of columns) {
                        let emailColumn: string = column[0];
                        let userColumn: string = column[1];
                        let user = await web.ensureUser(row[column[0]])
                            .then(u => {
                                return u;
                            })
                            .catch(err => {
                                debugger;
                                addMessage(`<h2>User with an eMail/UPN of '${row[column[0]]}' could not be found</h2>`);
                                addMessage(`<h1>Error was  ${err.data.responseBody["odata.error"].message.value} </h1>`);
                                return null;
                            });
                        if (user) { // if the user was ensured!
                            let update = { [userColumn]: user.data.Id };// enclose in brakts to eval
                            web.lists.getByTitle(listTitle).items.getById(row["Id"]).update(update)
                                .then(x => {
                                })
                                .catch(err => {
                                    debugger;
                                    addMessage(`<h1>Error updating user</h1>`);
                                    addMessage(`<h1>Error updating user email ${row[column[0]]} To user ID ${user.data.Id}</h1>`);
                                    addMessage(`<h1>Error was  ${err.data.responseBody["odata.error"].message.value} </h1>`);

                                    debugger;
                                });


                        }
                    }
                }
            })
            .catch(err => {
                console.error(err);
                addMessage(`<h1>Error fetching items from ${listTitle} <h1>`);
                addMessage(`<h1>Error was  ${err.data.responseBody["odata.error"].message.value} </h1>`);
            });
    }
    addMessage(`Done converting users on list ${listTitle}.`);
    return Promise.resolve();
}
export function ensureFieldsAreInList(list: List, fieldInternalNames: Array<string>, addMessage: (message: string) => void): boolean {
    // besure when you called pnp get list you expanded Fields/


    let allFound: boolean = true;
    for (let fieldName of fieldInternalNames) {
        if (!find(list["Fields"], f => { return f["InternalName"] === fieldName; })) {
            if (addMessage) {
                addMessage(`<h1>Field ${fieldName} was not found in list ${list["Title"]}</h1>`);
            }
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
export async function cleanupHomePage(webRelativeUrl: string, homePageUrl, webPartXml: string, addMessage: (message: string) => void) {
    addMessage(`Home page Url is ${homePageUrl}`);
    addMessage(`Web relative url is  ${webRelativeUrl}`);
    const clientContext: SP.ClientContext = new SP.ClientContext(webRelativeUrl);
    var oFile = clientContext.get_web().getFileByServerRelativeUrl(homePageUrl);
    var limitedWebPartManager = oFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
    let webparts = limitedWebPartManager.get_webParts();
    clientContext.load(webparts, 'Include(WebPart)');
    clientContext.load(limitedWebPartManager);
    await new Promise((resolve, reject) => {
        clientContext.executeQueryAsync((req: SP.ClientRequest, args: SP.ClientRequestSucceededEventArgs) => {
            resolve();
        }, (req: SP.ClientRequest, args: SP.ClientRequestFailedEventArgs) => {
            addMessage(`Error getting the list of webparts on the page at '${homePageUrl}' in method cleanupHomePage`);
            addMessage(args.get_message());
            addMessage(args.get_errorDetails());
            addMessage(args.get_errorValue());
            console.error(args);
            reject();
        });
    });
    let count = webparts.get_count();
    addMessage(`There are ${count} webparts on the homme page to remove`);
    for (let i = 0; i < count; i++) {
        let originalWebPartDef = webparts.get_item(i);
        addMessage(`Removing webpart  ${i} `);
        originalWebPartDef.deleteWebPart();
        await new Promise((resolve, reject) => {
            clientContext.executeQueryAsync((req: SP.ClientRequest, args: SP.ClientRequestSucceededEventArgs) => {
                addMessage(`Webparts removed from page at '${homePageUrl}'`);
                resolve();
            }, (req: SP.ClientRequest, args: SP.ClientRequestFailedEventArgs) => {
                addMessage(`Error removing Webparts from page at '${homePageUrl}'`);
                addMessage(args.get_message());
                addMessage(args.get_errorDetails());
                addMessage(args.get_errorValue());
                console.error(args);
                reject();
            });
        });
    }

    let oWebPartDefinition = limitedWebPartManager.importWebPart(webPartXml);
    let oWebPart = oWebPartDefinition.get_webPart();
    limitedWebPartManager.addWebPart(oWebPart, 'Main', 1);
    clientContext.load(oWebPart);
    await new Promise((resolve, reject) => {
        clientContext.executeQueryAsync((req: SP.ClientRequest, args: SP.ClientRequestSucceededEventArgs) => {
            addMessage(`Custom webpart added to page at '${homePageUrl}'`);

            resolve();
        }, (req: SP.ClientRequest, args: SP.ClientRequestFailedEventArgs) => {
            addMessage(`Error adding custom webpart to page  at '${homePageUrl}'`);
            addMessage(args.get_message());
            addMessage(args.get_errorDetails());
            addMessage(args.get_errorValue());
            console.error(args);
            reject();
        });
    });
}
export async function setWebToUseSharedNavigation(webAbsoluteUrl: string, addMessage: (message: string) => void) {
    let clientContext: SP.ClientContext;
    try {
        clientContext = new SP.ClientContext(webAbsoluteUrl);
    }
    catch (err) {
        addMessage(`Error creating client context on web ${webAbsoluteUrl}`);
        console.log(err);
        debugger;
    }
    var currentWeb = clientContext.get_web();
    var navigation = currentWeb.get_navigation();
    navigation.set_useShared(true);
    await new Promise((resolve, reject) => {
        clientContext.executeQueryAsync((req: SP.ClientRequest, ars: SP.ClientRequestSucceededEventArgs) => {
            addMessage(`The web at  ${webAbsoluteUrl} was set to use shared navigation`);
            resolve();
        }, (req: SP.ClientRequest, args: SP.ClientRequestFailedEventArgs) => {
            addMessage(`<h1>Error setting web at  ${webAbsoluteUrl} to use shared navigation</h1>`);
            addMessage(args.get_message());
            addMessage(args.get_errorDetails());
            addMessage(args.get_errorValue());
            console.error(args);
            reject();
        });
    });
}
export async function AddUsersInListToGroup(webUrl: string, listName: string, userFieldName: string, membersGroup: any, addMessage: (message: string) => void) {
    debugger;
    addMessage(`Adding the people in the  ${listName} to the group ${membersGroup.Title}`);
    let web = new Web(webUrl);
    await web.lists.getByTitle(listName).items.expand(userFieldName).select(userFieldName + "/Name").get() // fld/Name gets u the loginname
        .then(async (listItems) => {
            debugger;
            for (const item of listItems) {
                if (item[userFieldName]) {

                    await sp.web.siteGroups.getByName(membersGroup.Title).users.add(item[userFieldName]["Name"])
                        .then(e => {
                            addMessage(`added ${item[userFieldName]["Name"]}`);
                        })
                        .catch((err) => {
                            debugger;
                            addMessage(`<h1>Error adding user ${item[userFieldName]["Name"]}</h1>`);
                            addMessage(err.data.responseBody["odata.error"].message.value);
                            return;
                        });


                }
                else {
                    debugger;
                    addMessage(`<h3>User is missing on row</h3>`);
                }
            }

        })
        .catch((err) => {
            addMessage(`<h1>Error fetching people from the list named  ${listName}</h1>`);
            addMessage(err.data.responseBody["odata.error"].message.value);
        });
}
export async function AddQuickLaunchItem(webUrl: string, title: string, url: string, isExternal: boolean, addMessage: (message: string) => void) {
    let nnci: SP.NavigationNodeCreationInformation = new SP.NavigationNodeCreationInformation();
    nnci.set_title(title);
    nnci.set_url(url);
    nnci.set_isExternal(isExternal);
    const clientContext: SP.ClientContext = new SP.ClientContext(webUrl);
    const web = clientContext.get_web();
    web.get_navigation().get_quickLaunch().add(nnci);
    await new Promise((resolve, reject) => {
        clientContext.executeQueryAsync((req: SP.ClientRequest, ars: SP.ClientRequestSucceededEventArgs) => {
            addMessage(`Added QuickLaunch Item ${title} to web at ${webUrl}`);
            resolve();
        }, (req: SP.ClientRequest, args: SP.ClientRequestFailedEventArgs) => {
            addMessage(`<h1>Error adding QuickLaunch Item ${title}</h1>`);
            addMessage(args.get_message());
            addMessage(args.get_errorDetails());
            addMessage(args.get_errorValue());
            console.error(args);
            reject();
        });
    });
}
export async function RemoveQuickLaunchItem(webUrl: string, titlesToRemove: string[], addMessage: (message: string) => void) {
    const clientContext: SP.ClientContext = new SP.ClientContext(webUrl);
    const ql: SP.NavigationNodeCollection = clientContext.get_web().get_navigation().get_quickLaunch();
    clientContext.load(ql);
    await new Promise((resolve, reject) => {
        clientContext.executeQueryAsync((req: SP.ClientRequest, ars: SP.ClientRequestSucceededEventArgs) => {
            resolve();
        }, (req: SP.ClientRequest, args: SP.ClientRequestFailedEventArgs) => {
            addMessage(`<h1>Error fetching quicklaunch items in method RemoveQuickLaunchItem</h1>`);
            addMessage(args.get_message());
            addMessage(args.get_errorDetails());
            addMessage(args.get_errorValue());
            console.error(args);
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
        clientContext.executeQueryAsync((req: SP.ClientRequest, ars: SP.ClientRequestSucceededEventArgs) => {
            addMessage(`Removed items ${titlesToRemove.join(",")} from QuickLaunch`);
            resolve();
        }, (req: SP.ClientRequest, args: SP.ClientRequestFailedEventArgs) => {
            addMessage(`<h1>Error Removing items from  quicklaunch in method RemoveQuickLaunchItem</h1>`);
            addMessage(args.get_message());
            addMessage(args.get_errorDetails());
            addMessage(args.get_errorValue());
            console.error(args);
            reject();
        });
    });

}
// export async function fixUpLeftNav(webUrl: string, homeUrl: string) {

//     await AddQuickLaunchItem(webUrl, "EFR Home", homeUrl, true);
//     await RemoveQuickLaunchItem(webUrl, ["Pages", "Documents"]);
// }
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