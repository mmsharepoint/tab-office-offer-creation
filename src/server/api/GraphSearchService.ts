import Axios, { AxiosRequestConfig } from "axios";
import * as debug from 'debug';
import { IOfferDocument } from "../../model/IOfferDcoument";

const log = debug('graphRouter');

export default class GraphSearchService {
  public async getFiles(token: string): Promise<IOfferDocument[]> {  
    const searchResponse = {
      requests: [
        { entityTypes: ['driveItem'],
          query: {
            queryString: 'ContentTypeID:0x0101003656A003937692408E62ADAA56A5AEEF*'
          }
        }
      ]};
    const requestUrl: string = `https://graph.microsoft.com/v1.0/search/microsoft.graph.query`;
    return Axios.post(requestUrl,
      searchResponse,
      {
        headers: {          
          Authorization: `Bearer ${token}`
      }})
      .then(response => {
        let docs: IOfferDocument[] = [];
        response.data.value[0].hitsContainers[0].hits.forEach(element => {
          docs.push({
            name: element.resource.name,
            description: element.summary,
            author: element.resource.createdBy.user.displayName,
            url: element.resource.webUrl,
            id: element.resource.parentReference.sharepointIds.listItemId,
            modified: new Date(element.resource.lastModifiedDateTime)
          });
        });
        return docs;
      })
      .catch(err => {
        log(err);
        return [];
      });
  }

  public async reviewItem(token: string, itemID: string, user: string): Promise<void> {
    let requestUrl: string = await this.getSiteAndListByPath(token, process.env.SiteUrl!);
    // Get user LookupID
    const userInfoListID = await this.getUserInfoListID(token, requestUrl);
    const userLookupID = await this.getUserLookupID(token, requestUrl, userInfoListID, user);
    requestUrl += `/${itemID}/fields`;
    const config: AxiosRequestConfig = {  headers: {      
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    }};
    const fieldValueSet = {
      OfferingReviewedDate: new Date().toISOString(),
      OfferingReviewerLookupId: userLookupID
    };  
    Axios.patch(requestUrl, 
                fieldValueSet,
                config
    )
    .then((response) => {
      console.log(response.data);
    })
    .catch((error) => {
      log(error);
    });
  }

  private async getSiteAndListByPath (accessToken: string, siteUrl: string): Promise<string> {
    const siteURL = new URL(siteUrl);
    const domain = siteURL.hostname;
    const path = siteURL.pathname;
    const apiSiteUrl =`https://graph.microsoft.com/v1.0/sites/${domain}:/${path}?$expand=drives`;
    try {
      const siteResponse = await Axios.get(apiSiteUrl, {
        headers: {
          Authorization: `Bearer ${accessToken}`
        }
      });
      const requestUrl = `https://graph.microsoft.com/v1.0/sites/${siteResponse.data.id}/lists/${siteResponse.data.drives[0].name}/items`;
      return requestUrl;
    }
    catch (error) {
      log("Error while retrieving siteID: ")
      log(error);
      return "";
    }
  }

  private async getUserInfoListID (accessToken: string, requestUrl: string): Promise<string> {
    let listRequestUrl = requestUrl.split('/lists')[0];
    listRequestUrl += "/lists?$select=name,webUrl,displayName,Id,system";
    try {
      const response = await Axios.get(listRequestUrl, {
        headers: {
          Authorization: `Bearer ${accessToken}`
        }
      });
      const lists: any[] = response.data.value;
      let listID = "";
      lists.forEach((l) => {
        if (l.webUrl.endsWith('/_catalogs/users')) {
          listID = l.id;
        }
      });
      return listID;
    }
    catch (error) {
      log("Error while retrieving listID: ")
      log(error);
      return "";
    }
  }

  private async getUserLookupID (accessToken: string, requestUrl: string, listID: string, userName: string): Promise<string> {
    let listRequestUrl = requestUrl.split('/lists')[0];
    listRequestUrl += `/lists/${listID}/items?$expand=fields&$filter=fields/UserName eq '${userName}'`;
    try {
      const response = await Axios.get(listRequestUrl, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Prefer': 'HonorNonIndexedQueriesWarningMayFailRandomly' // No chance to index User Information List
        }
      });
      return response.data.value[0].id;
    }
    catch (error) {
      log("Error while retrieving userID: ")
      log(error);
      return "";
    }
  }
}