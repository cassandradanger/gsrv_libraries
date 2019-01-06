import { Version } from '@microsoft/sp-core-library';
import { sp } from "@pnp/sp";
import {
  Environment,
  EnvironmentType
 } from '@microsoft/sp-core-library';

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GsrvLibrariesWebPart.module.scss';

import * as strings from 'GsrvLibrariesWebPartStrings';

import {
  SPHttpClient,
  SPHttpClientResponse   
 } from '@microsoft/sp-http';

export interface IMyLibrariesWebPartProps {
  description: string;
}


export interface ISPLists {
  value: ISPList[];
 }

 export interface ISPList {
  Title: string; // this is the department name in the List
  Id: string;
  AnncURL:string;
  DeptURL:string;
  CalURL:string;
  a85u:string; // this is the LINK URL
 }

 //global vars
 var userDept = "";


export default class MyLibrariesWebPart extends BaseClientSideWebPart<IMyLibrariesWebPartProps> {
public render(): void {
  this.domElement.innerHTML = `
  <h1 class=${styles.titleLI}>Libraries</h1>
    <div class=${styles.contentLI}>
      <ul class=${styles.ulLI} id="libraryList"/>
    </div>
    `;
}

getuser = new Promise((resolve,reject) => {
  // SharePoint PnP Rest Call to get the User Profile Properties
  return sp.profiles.myProperties.get().then(function(result) {
    var props = result.UserProfileProperties;

    props.forEach(function(prop) {
      //this call returns key/value pairs so we need to look for the Dept Key
      if(prop.Key == "Department"){
        // set our global var for the users Dept.
        userDept += prop.Value;
      }
    });
    return result;
  }).then((result) =>{
    this._getListData().then((response) =>{
      this._renderList(response.value);
    });
  });

});

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  
  // main REST Call to the list...passing in the deaprtment into the call to 
  //return a single list item
  public _getListData(): Promise<ISPLists> {  
    return this.context.spHttpClient.get(`https://girlscoutsrv.sharepoint.com/_api/web/lists/GetByTitle('TeamDashboardSettings')/Items?$filter=Title eq '`+ userDept +`'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
   }

 private _renderList(items: ISPList[]): void {
  var siteURL = "";

  items.forEach((item: ISPList) => {
    let html: string = '';
    siteURL = item.DeptURL;

    //SharePoint PnP Call to get all the document libraries for the site
    sp.site.getDocumentLibraries("https://girlscoutsrv.sharepoint.com" + siteURL).then((data) => {
        data.forEach((data) => {
          html += `<li class=${styles.liLI}><a href=${data.AbsoluteUrl}>${data.Title}</a></li>`;
        });
        const listContainer: Element = this.domElement.querySelector('#libraryList');
        listContainer.innerHTML = html;
    }).catch((err) => {
     console.log(err);
    });
  });
}
  
// this is required to use the SharePoint PnP shorthand REST CALLS
   public onInit():Promise<void> {
     return super.onInit().then (_=> {
       sp.setup({
         spfxContext:this.context
       });
     });
   }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
