import { PageContext, SPPermission } from '@microsoft/sp-page-context';
import {SPHttpClient, ISPHttpClientOptions} from "@microsoft/sp-http";

export const deleteForm = async (sphttpClient: any ,listUrl: string, listTitle: string, itemId: any) =>{
    const restUrl = `${listUrl}/_api/web/lists/getByTitle('${listTitle}')/items(${itemId})/recycle`;
    let spOptions: ISPHttpClientOptions = {
        headers:{
            Accept: "application/json;odata=nometadata", 
            "Content-Type": "application/json;odata=nometadata",
            "odata-version": "",
            "IF-MATCH": "*",
            // "X-HTTP-Method": "DELETE"         
        },
    };

    const _data = await sphttpClient.post(restUrl, SPHttpClient.configurations.v1, spOptions);
    if (_data.ok){
        console.log('Form is deleted! Please check Recycle Bin to restore it.');
        return _data;
    }
};

export const getFollowed = async (msGraphClientFactory: any) => {
    const graphResponse = await msGraphClientFactory.getClient();
    const followedDocsResponse = await graphResponse.api(`/me/drive/following`).top(1000).get();
    console.log("My Followed documents", followedDocsResponse);
    return followedDocsResponse.value.map(item => {
        return {
            name: decodeURI(item.name),
            driveId: item.parentReference.driveId
        };
    });
};

const getDocDriveInfo = async (msGraphClientFactory: any, siteId: string, webId: string, listId: string, listItemId: string) => {
    const graphResponse = await msGraphClientFactory.getClient();
    const driveResponse = await graphResponse.api(`/sites/${siteId},${webId}/lists/${listId}/items/${listItemId}/driveItem`).get();
    return [driveResponse.parentReference.driveId, driveResponse.id];
};

export const followDocument = async (msGraphClientFactory: any, siteId: string, webId: string, listId: string, listItemId: string) => {

    const [driveId, driveItemId] =  await getDocDriveInfo(msGraphClientFactory, siteId, webId, listId, listItemId);

    console.log("driveId", driveId);
    console.log("driveItemId", driveItemId);

    const graphResponse = await msGraphClientFactory.getClient();
    const followResponse = await graphResponse.api(`/drives/${driveId}/items/${driveItemId}/follow`).post(JSON.stringify(''));

    console.log("followResponse", followResponse);
};

export const unFollowDocument = async (msGraphClientFactory: any, siteId: string, webId: string, listId: string, listItemId: string) => {
    const [driveId, driveItemId] =  await getDocDriveInfo(msGraphClientFactory, siteId, webId, listId, listItemId);

    console.log("driveId", driveId);
    console.log("driveItemId", driveItemId);

    const graphResponse = await msGraphClientFactory.getClient();
    const unfollowResponse = await graphResponse.api(`/drives/${driveId}/items/${driveItemId}/unfollow`).post(JSON.stringify(''));

    console.log("unfollowResponse", unfollowResponse);
};
 
export const isUserManage = (pageContext: PageContext) : boolean =>{
    const userPermissions = pageContext.web.permissions,
        permission = new SPPermission (userPermissions.value);
    
    return permission.hasPermission(SPPermission.manageWeb);
};

