import { Version } from '@microsoft/sp-core-library';
import { sp, Items, ItemVersion, Web } from "@pnp/sp";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {  
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';  

import styles from './GsrvAnnouncementsWebPart.module.scss';
import * as strings from 'GsrvAnnouncementsWebPartStrings';

export interface IGsrvAnnouncementsWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];
 }

export interface ISPList {
  Body: string;
  Title: string; // this is the department name in the List
  Id: string;
  AnncURL:string;
  DeptURL:string;
  CalURL:string;
  a85u:string; // this is the LINK URL
 }

//global vars
var userDept = "";

export interface IGsvrAnnouncementsWebPartProps {
  description: string;
}
export default class GsrvAnnouncementsWebPart extends BaseClientSideWebPart<IGsrvAnnouncementsWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class=${styles.mainAN}>
      <ul class=${styles.contentAN}>
        <div id="ListItems" /></div>
      </ul>
    </div>`;
  }

  getuser = new Promise((resolve,reject) => {
    // SharePoint PnP Rest Call to get the User Profile Properties
    return sp.profiles.myProperties.get().then(function(result) {
      var props = result.UserProfileProperties;
      var propValue = "";
      var userDepartment = "";
  
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

  public _getListData(): Promise<ISPLists> {  
    return this.context.spHttpClient.get(`https://girlscoutsrv.sharepoint.com/_api/web/lists/GetByTitle('TeamDashboardSettings')/Items?$filter=Title eq '`+ userDept +`'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
   }

   private _renderList(items: ISPList[]): void {
    let html: string = '';
    var siteURL = "";
    var announcementsListName =  "";
    var date = new Date();
    var strToday = "";
    var mm = date.getMonth()+1;
    var dd = date.getDate();
    var yyyy = date.getFullYear();
    if(dd < 10){
      dd = 0 + dd;
    }

    if(mm < 10){
      mm = 0 + mm;
    }

    strToday = mm + "/" + dd + "/" + yyyy;
    items.forEach((item: ISPList) => {
      siteURL = item.DeptURL;
      announcementsListName = item.AnncURL;

      const w = new Web("https://girlscoutsrv.sharepoint.com" + siteURL);
      
      // then use PnP to query the list
      w.lists.getByTitle(announcementsListName).items.filter("Expires ge '" + strToday + "'").top(1)
      .get()
      .then((data) => {
        console.log(data);
        html += `
        <div>${data[0].Body}</div`
        const listContainer: Element = this.domElement.querySelector('#ListItems');
        listContainer.innerHTML = html;
      }).catch(e => { console.error(e); });
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
