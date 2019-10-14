import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import MockHttpClient from './MockHttpClient';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './Demo1WebPart.module.scss';
import * as strings from 'Demo1WebPartStrings';

export interface IDemo1WebPartProps {
  description: string;
  comments: string;
  isSPFx: boolean;
  version: string;
  isValid: boolean
}

export interface ISPList{
  Title: string;
  Id: string;
}
export interface ISPLists{
  value: ISPList[];
}

export default class Demo1WebPart extends BaseClientSideWebPart<IDemo1WebPartProps> {

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?filter=Hidden eq false`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse)=> {
      return response.json();
    });
  }
  public _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get()
    .then((data:ISPList[])=>{
      var listData: ISPLists={value:data};
      return listData;
    }) as Promise<ISPLists>;

  }

  private _RenderData(items:ISPList[]): void{
    let html:string='<ul>';
    items.forEach((item:ISPList)=>{
      html+=`
      
        <li>${item.Title}</li>
      
      `;
    })
html+='</ul>';
  const listContainer: Element = this.domElement.querySelector('#dtReport');
  listContainer.innerHTML=html;
  }


  private _renderListAsync(): void
{
  
  if(Environment.type === EnvironmentType.Local)
  {
    this._getMockListData().then((response) =>{
      this._RenderData(response.value);
    })
  }
  else if(Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint)
  {

    this._getListData().then((response) =>{
      console.log(response.value);
      this._RenderData(response.value);
    })
  }
}
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.demo1 }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <p class="${ styles.description }">${escape(this.properties.comments)}</p>
              <p class="${ styles.description }">${(this.properties.isSPFx)}</p>
              <p class="${ styles.description }">${escape(this.properties.version)}</p>
              <p class="${ styles.description }">${(this.properties.isValid)}</p>
              <p class="${ styles.description }">${escape(this.context.pageContext.web.title)}</p>
              <p class="${ styles.description }">${escape(this.context.pageContext.web.serverRelativeUrl)}</p>
              <p class="${ styles.description }">${escape(this.context.pageContext.user.displayName)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
        <div id="dtReport"></div>
      </div>`;
     this._renderListAsync();
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
                }),
                PropertyPaneTextField('comments',{
                  label: "Comments",
                  multiline: true
                }),
                PropertyPaneCheckbox('isSPFx',{
                  text: "IS SpFx"
                }),
                PropertyPaneDropdown('version',{
                  label: "Version",
                  options: [
                    {key:"1", text: 'One'},
                    {key:'2', text:'Two'},
                    {key: '3', text: 'Threee'},
                    {key: '4', text: 'Four'}
                  ]
                }),
                PropertyPaneToggle('isValid',{
                  label: "isValid",
                  onText: 'On',
                  offText: 'Off'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
