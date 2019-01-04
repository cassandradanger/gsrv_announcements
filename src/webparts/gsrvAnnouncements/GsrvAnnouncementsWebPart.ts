import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {  
  SPHttpClient  
} from '@microsoft/sp-http';  

import styles from './GsrvAnnouncementsWebPart.module.scss';
import * as strings from 'GsrvAnnouncementsWebPartStrings';

export interface IGsrvAnnouncementsWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];  
}

export interface ISPList{
  Title: string;
  Body: any;
}


export default class GsrvAnnouncementsWebPart extends BaseClientSideWebPart<IGsrvAnnouncementsWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class=${styles.mainAN}>
      <ul class=${styles.contentAN}>
        <div id="spListContainer" /></div>
      </ul>
    </div>`;
      this._firstGetList();
  }

  private _firstGetList() {
    this.context.spHttpClient.get('https://girlscoutsrv.sharepoint.com' + 
      `/gsrv_teams/sdgdev/_api/web/Lists/GetByTitle('Announcements')/Items`, SPHttpClient.configurations.v1)
      .then((response)=>{
        response.json().then((data)=>{
          console.log(data);
          this._renderList(data.value)
        })
      });
    }

  private _renderList(items: ISPList[]): void {
    let html: string = ``;
    items.forEach((item: ISPList) => {
      let announcement = item.Body;

      html += `
        <div>${announcement}</div>
        `;  
    });  
    const listContainer: Element = this.domElement.querySelector('#spListContainer');  
    listContainer.innerHTML = html;  
  } 

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
