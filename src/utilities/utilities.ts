import { find, clone, map } from "lodash";
import { Web } from "sp-pnp-js";

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
            this.cachedIds.push({ upn: upn, id: id });
            return;
        })
        .catch((err) => {
            this.cachedIds.push({ upn: upn, id: null });
            return;
        });
    return id;
}

export async function esnureUsers(web: Web, cachedIds: Array<CachedId>,
    data: Array<any>, columnName: string, addMessage: (message: string) => void): Promise<any> {
    let rowNumber = 0;
    for (let row of data) {
        rowNumber++;
        let approverId: number = await findId(web, cachedIds, row["columnName"]);
        if (!approverId) {
            addMessage(`Approver  ${row.ApproverEmail} on row ${rowNumber} of the Role To Transaction File is invalid`);
        }
    }

    return Promise.resolve();
}
export function  uploadFile(web:Web,libraryName:string,file:File,saveAsFileName:string,addMessage: (message: string) => void): Promise<any> {
    debugger;
    return web.lists.getByTitle(libraryName).rootFolder.files
      .addChunked(saveAsFileName, this.state.file, data => {
        addMessage(`(Stage ${data.stage}) Uploaded block ${data.blockNumber} of ${data.totalBlocks}`);

      }, true)
      .then((results) => {
        return results.file.getItem().then(item => {
          return item.update({ Title: saveAsFileName }).then((r) => {
            return;
          }).catch((err) => {
            debugger;
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