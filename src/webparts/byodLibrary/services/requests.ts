import { WebPartContext } from "@microsoft/sp-webpart-base";
import {SPHttpClient, ISPHttpClientOptions} from "@microsoft/sp-http";
import { SPPermission } from "@microsoft/sp-page-context";

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/views";
import "@pnp/sp/items/get-all";

export const getGraphMemberOf = async (context: WebPartContext) =>{
    const graphResponse = await context.msGraphClientFactory.getClient('3');
    const graphUrl = '/me/transitiveMemberOf/microsoft.graph.group';
    const memberOfGraph = await graphResponse
        .api(graphUrl)
        .header('ConsistencyLevel', 'eventual')
        .count(true)
        .select('displayName,mail')
        .top(500)
        .get();

    const userGroups = [];
    for (const group of memberOfGraph.value){
        userGroups[group.displayName] = {displayName: group.displayName, email: group.mail};
    }

    return userGroups;
};

export const isFromTargetAudience = (userEmail: string, userGroups: any, targetAudience: any, targetAudienceKey: string) => {
    
    console.log("userGroups", userGroups);
    console.log("targetAudience", targetAudience);

    for (const audience of targetAudience){
        if (userEmail === audience.email) return true;
        if (userGroups[audience[targetAudienceKey]]) return true;
    }

    return false;
};

export const getListItems = async (context: WebPartContext, siteUrl: string, listName: string) =>{ 
    const responseUrl = `${siteUrl}/_api/web/Lists/GetByTitle('${listName}')/items`;
    const response = await context.spHttpClient.get(responseUrl, SPHttpClient.configurations.v1).then(r => r.json());
    console.log("response.value", response.value);
    return response.value;
};

export const getListItemsCamlQuery = async (context: WebPartContext, siteUrl: string, listName: string) =>{
    const sp = spfi(siteUrl).using(SPFx(context));  

    const viewFields = await sp.web.lists.getByTitle(listName).views.getByTitle('All Items').fields.getSchemaXml();
    console.log("viewFields", viewFields);

    //const xml = `<View><ViewFields><FieldRef Name='Title' /><FieldRef Name='AssignedTo' /></ViewFields><RowLimit>100</RowLimit></View>`;
    const xml = `<View><Query>${viewFields}</Query><RowLimit>100</RowLimit></View>`;
    const items = await sp.web.lists.getByTitle(listName).getItemsByCAMLQuery({ViewXml : xml});
    return items;

}   

const getSiteId = async (context: WebPartContext, siteUrl: string) =>{
    const responseUrl = `${siteUrl}/_api/site/id`;
    const response = await context.spHttpClient.get(responseUrl, SPHttpClient.configurations.v1).then(r => r.json());
    return response.value;
};

export const getListGuid  = async (context: WebPartContext, siteUrl: string, listName: string) => {
    const responseUrl = `${siteUrl}/_api/web/lists/getByTitle('${listName}')/Id`;
    const response = await context.spHttpClient.get(responseUrl, SPHttpClient.configurations.v1).then(r => r.json());
    return response.value;
};

export const getListItemsGraph = async (context: WebPartContext, siteUrl: string, listName: string) => {

    const siteId = await getSiteId(context, siteUrl);
    const listGuid = await getListGuid(context, siteUrl, listName);

    const graphClient = await context.msGraphClientFactory.getClient('3');
    const items = await graphClient.api(`sites/${siteId}/lists/${listGuid}/items?expand=fields(select=Title,link,Image,_ModernAudienceTargetUserField,Author,Id,login,pwd,LoginDisclaimer,NewTab,Category,ID,Created,Modified,Short_x0020_Description,Order)&orderby=fields/Order`).get();
    return items.value;
}

export const copyTextToClipboard = async (textToCopy: string) => {
    try {
      if (navigator?.clipboard?.writeText) {
        await navigator.clipboard.writeText(textToCopy);
      }
    } catch (err) {
      console.error(err);
    }
};


export const groupBy = (objArr: any, key: string) => {
    return objArr.reduce((rv: any, x: any) => {
        (rv[x[key]] = rv[x[key]] || []).push(x);
        return rv;
    }, {});
};


export const isUserManage = (context: WebPartContext) : boolean =>{
    const userPermissions = context.pageContext.web.permissions,
        permission = new SPPermission (userPermissions.value);
    
    return permission.hasPermission(SPPermission.manageWeb);
};

export const deleteItem = async (context: WebPartContext, siteUrl: string, listTitle: string, itemId: any) =>{
    const restUrl = `${siteUrl}/_api/web/lists/getByTitle('${listTitle}')/items(${itemId})`;
    const spOptions: ISPHttpClientOptions = {
        headers:{
            Accept: "application/json;odata=nometadata", 
            "Content-Type": "application/json;odata=nometadata",
            "odata-version": "",
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"         
        },
    };

    const _data = await context.spHttpClient.post(restUrl, SPHttpClient.configurations.v1, spOptions);
    if (_data.ok){
        console.log('Tile is deleted!');
    }
};