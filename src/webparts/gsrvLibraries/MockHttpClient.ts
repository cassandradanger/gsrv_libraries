import { ISPList } from './GsrvLibrariesWebPart';

export default class MockHttpClient  {

   private static _items: ISPList[] = [{ Title: 'Mock List', Id: '1',
   AnncURL:"https://girlscoutsrv.sharepoint.com/gsrv_teams/sdgdev/Lists/Announcements",
   DeptURL:"https://girlscoutsrv.sharepoint.com/gsrv_teams/sdgdev",
   CalURL:"https://girlscoutsrv.sharepoint.com/gsrv_teams/sdgdev/Lists/Events",
   a85u:"https://girlscoutsrv.sharepoint.com/gsrv_teams/sdgdev/Lists/Helpful%20Links" }];

   public static get(): Promise<ISPList[]> {
   return new Promise<ISPList[]>((resolve) => {
           resolve(MockHttpClient._items);
       });
   }
}