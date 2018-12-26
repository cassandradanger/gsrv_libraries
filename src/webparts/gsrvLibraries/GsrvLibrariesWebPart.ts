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
import MockHttpClient from './MockHttpClient';

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
  Title: string; // this is the department
  Id: string;
  AnncURL:string;
  DeptURL:string;
  CalURL:string;
  a85u:string; // this is the LINK URL
 }

 //global vars
 var userDept;
 var listUrl;
 var siteURL;


export default class MyLibrariesWebPart extends BaseClientSideWebPart<IMyLibrariesWebPartProps> {
// just to get this mocked up I left in the generic code form the hello world build

  public render(): void {
    this.domElement.innerHTML = `
    <div class=${styles.mainHR}>
      <p class=${styles.titleHR}>
        Libraries
      </p>
      <ul class=${styles.contentHR}>
        <div id="libraryList" /></div>
      </ul>
    </div>`;
// get the user data 1st
     this._getUserInfo();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // main REST Call to the list...passing in the deaprtment into the call to return a single list item
  public _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get( `https://girlscoutsrv.sharepoint.com/_api/web/lists/GetByTitle('TeamDashboardSettings')/Items?$filter=Title eq '` + userDept + `'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
   }

   // this will get any and all user data, but we just need the department
   private _getUserInfo(){
    sp.profiles.myProperties.get()
    .then((result) => {
        var props = result.UserProfileProperties;
        var propValue = "";
      
        props.forEach((prop) => {
          if(prop.Key == "Department"){
            //get the department and set it to a global var
            userDept = prop.value;
          }
          // this is a test line to make sure profile data is coming back
          // propValue += userDept + "<br/>";
          this._getListData()
          .then((response) => {
            this._renderList(response.value);
          });
      });
      

        //document.getElementById("spUserProfileProperties").innerHTML = propValue;
      }).catch((err) => {
        console.log("Error: " + err);
      });
   }

//mock up for the list in the test enviroment 
private _getMockListData(){
   return MockHttpClient.get()
     .then((data: ISPList[]) => {
       var listData: ISPLists = { value: data };
       return listData;
     }) as Promise<ISPLists>;
 }

 // this is how you can call test vs prod data
 private _renderListAsync(): void {
  // Local environment
  if (Environment.type === EnvironmentType.Local) {
    this._getMockListData().then((response) => {
      this._renderList(response.value);
    });
  }
  else if (Environment.type == EnvironmentType.SharePoint || 
            Environment.type == EnvironmentType.ClassicSharePoint) {
    this._getListData()
      .then((response) => {
        this._renderList(response.value);
      });
  }
}


 private _renderList(items: ISPList[]): void {
  let html: string = '';
  console.log('hey', items);
  items.forEach((item: ISPList) => {
    // pnp call to get all the libraries
    sp.site.getDocumentLibraries("https://girlscoutsrv.sharepoint.com"+ item.DeptURL).then((data) => {
      var docLibNames = "";
      for (var i= 0; i < data.length; i++){
        docLibNames += data[i].Title + '\r\n';
      }
      // MOCK CALL: to list all the document libraries
      document.getElementById("LibraryNames").innerText = docLibNames;
    }).catch((err) => {
      console.log(err);
    });
    // you can use the get element by ID or build your own HTML and pass to the to the element below
    html += ``;
  });

  const listContainer: Element = this.domElement.querySelector('#libraryList');
  listContainer.innerHTML = html;
}
  
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
