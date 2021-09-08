// João Mendes
// March 2019

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp, Fields, Web, SearchResults, Field, PermissionKind, RegionalSettings, PagedItemCollection } from '@pnp/sp';
import { graph, } from "@pnp/graph";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions, HttpClient, MSGraphClient } from '@microsoft/sp-http';
import * as $ from 'jquery';

import { registerDefaultFontFaces } from "@uifabric/styling";
import * as moment from 'moment';
import { SiteUser } from "@pnp/sp/src/siteusers";
import { dateAdd } from "@pnp/common";
import { escape, update } from '@microsoft/sp-lodash-subset';


// Class Services
export default class spservices {
  private ServerRelativeUrl = [];
  private graphClient: MSGraphClient = null;
  private spHttpClient: SPHttpClient;
  constructor(private context: WebPartContext) {
    // Setuo Context to PnPjs and MSGraph
    sp.setup({
      spfxContext: this.context,

    });
    this.spHttpClient = this.context.spHttpClient;
    graph.setup({
      spfxContext: this.context
    });
    // Init
    this.onInit();
  }
  // OnInit Function
  private async onInit() {
  }

  public async getSiteLists(siteUrl: string) {
    let results: any[] = [];
    if (!siteUrl) {
      return [];
    }
    try {
      const web = new Web(siteUrl);
      results = await web.lists
        .select("Title", "ID")
        //.filter('BaseTemplate eq 109')
        .usingCaching()
        .get();

    } catch (error) {
      return Promise.reject(error);
    }
    return results;
  }
  

  private getPageListItems(siteUrl: string, listId: string, index: number, numberImages: number): Promise<any[]> {
    return new Promise<any[]>((resolve, reject): void => {
      let requestUrl = siteUrl
        + `/_api/web/Lists/GetById('` + listId + `')/items`
        + `?$skiptoken=Paged=TRUE%26p_ID=` + (index * numberImages + 1)
        + `&$top=` + numberImages
        + `&$select=*&$expand=File`;

      this.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          response.json().then((responseJSON: any) => {
            resolve(responseJSON.value);
          });
        });
    });
  }
  public getLatestItemId(siteUrl: string, listId: string): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      const web = new Web(siteUrl);
      web.lists.getById(listId)
        .items.orderBy('Id', false).top(1).select('Id').get()
        .then((items: { Id: number }[]): void => {
          if (items.length === 0) {
            resolve(-1);
          }
          else {
            resolve(items[0].Id);
          }
        }, (error: any): void => {
          reject(error);
        });
    });
  }
  public async getLargeListItems(siteUrl: string, listId: string, numberImages: number): Promise<any[]> {
    var largeListItems: any[] = [];
    var pageSize = 500;
    var item_count = 0;
    return new Promise<any[]>(async (resolve, reject) => {
      // Array to hold async calls  
      const asyncFunctions = [];

      this.getLatestItemId(siteUrl, listId).then(async (itemCount: number) => {
        for (let i = 0; i < Math.ceil(itemCount / pageSize); i++) {
          // Make multiple async calls  
          let resolvePagedListItems = () => {
            return new Promise(async (resolve1) => {
              let pagedItems: any[] = await this.getPageListItems(siteUrl, listId, i, pageSize);
              let sub_results: any[] = [];
              for (let j = 0; j < pagedItems.length; j++) {
                if ((pagedItems[j].status == 'show') && (item_count < numberImages)) {
                  item_count++;
                  sub_results = sub_results.concat(pagedItems[j]);
                }
              }
              resolve1(sub_results);
            });
          };
          asyncFunctions.push(resolvePagedListItems());
        }

        // Wait for all async calls to finish  
        const results: any = await Promise.all(asyncFunctions);
        for (let i = 0; i < results.length; i++) {

          largeListItems = largeListItems.concat(results[i]);

        }

        resolve(largeListItems);
      });
    });
  }
  public async getImagesQuery(siteUrl: string, listId: string, numberImages: number): Promise<any[]> {
    return new Promise<any>(async (resolve, reject) => {
      let query = "<View><Query>" +
        "<Where>" +
        "<Eq><FieldRef Name=\"status\"/><Value Type=\"Text\">show</Value></Eq>" +
        "</Where>" +
        "<OrderBy><FieldRef Name='ID' Ascending='False' /></OrderBy></Query><RowLimit>" + numberImages + "</RowLimit>" +
        "</Query></View>";
      let requestUrl = siteUrl + "api/web/Lists/GetByTitle('" + listId + "')/GetItems";
      let camlQueryPayLoad: any = {
        query: {
          __metadata: { type: "SP.CamlQuery" },
          ViewXml: query
        }
      };

      let spOpts = {
        body: JSON.stringify(camlQueryPayLoad)
      };


      this.context.spHttpClient.post(requestUrl, SPHttpClient.configurations.v1, spOpts)
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            response.json().then((responseJSON) => {
              if (responseJSON != null && responseJSON.value != null) {
                let items: any[] = responseJSON.value;
                console.log(items);
                resolve(items);
              }
            });
          }
        });
    });
  }
  public async getImages(siteUrl: string, listId: string, numberImages: number): Promise<any[]> {
    return new Promise<any>(async (resolve, reject) => {
      this.spHttpClient.get(`${siteUrl}/_api/web/lists/GetById('` + listId + `')/items?$select=*&$expand=File&$filter=(status eq 'show')&$top=` + numberImages + `&$orderby=ID`,
        SPHttpClient.configurations.v1
        /* ,{
             headers: {
                 'Accept': 'application/json;odata=nometadata',
                 'odata-version': '3.0'
             }
         }*/
      )
        .then((response: SPHttpClientResponse): Promise<any> => {
          //console.log(response);
          //console.log(response.ok);
          return response.json();
        }, (error: any): void => {
          reject(error);
        }).then((responseJSON) => {
          //console.log(responseJSON);
          //console.log("Start carusel");
          if (responseJSON != null && responseJSON.value != null) {
            let items: any[] = responseJSON.value;
            console.log('responseJSON', responseJSON);
            resolve(items);
          }

        });
      /*.then((response: { ListItemEntityTypeFullName: string }): void => {
          var listItemEntityTypeName = response.ListItemEntityTypeFullName;
          resolve(listItemEntityTypeName);
      })*/

    });
  }

}
