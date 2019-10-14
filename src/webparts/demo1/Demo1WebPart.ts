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
import * as jquery from 'jquery';
import { IDigestCache, DigestCache } from '@microsoft/sp-http';

export interface IDemo1WebPartProps {
  description: string;
  comments: string;
  isSPFx: boolean;
  version: string;
  isValid: boolean;
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
    });
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
    });
  }
  else if(Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint)
  {

    this._getListData().then((response) =>{
      console.log(response.value);
      this._RenderData(response.value);
    });
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
              <button class="ms-Button read-httpButton ${styles.button} readall-Button">
                <span class="ms-Button-label">Read httpClient All item</span>
              </button>
              <button class="ms-Button read-ajaxButton ${styles.button} readall-Button">
                <span class="ms-Button-label">Read Ajax All item</span>
              </button>
              <button class="ms-Button create-httpButton ${styles.button} readall-Button">
                <span class="ms-Button-label">Creare http item</span>
              </button>
              <button class="ms-Button create-ajaxButton ${styles.button} readall-Button">
                <span class="ms-Button-label">Create Ajax item</span>
              </button>
              <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
                <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <div class="status"></div>
                  <ul class="items"><ul>
                </div>
              </div>
              
            </div>
            
          </div>
        </div>
        <div id="dtReport"></div>
      </div>`;
      const webPart: Demo1WebPart = this;
  this.domElement.querySelector('button.read-httpButton').addEventListener('click', () => { webPart.readHttpItem(); });
  this.domElement.querySelector('button.read-ajaxButton').addEventListener('click', () => { webPart.readAjaxItem(); });
  this.domElement.querySelector('button.create-httpButton').addEventListener('click', () => { webPart.createHttpItem(); });
  this.domElement.querySelector('button.create-ajaxButton').addEventListener('click', () => { webPart.createAjaxItem(); });

     //this._renderListAsync();
  }

  //create ajax item
  private createAjaxItem(): void {
    alert("ajax");
    const digestCache: IDigestCache = this.context.serviceScope.consume(DigestCache.serviceKey);

  digestCache.fetchDigest(this.context.pageContext.web.serverRelativeUrl)
    .then((digest: string) => {

    jquery.ajax({    
      url: `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items`,    
      type: "POST",    
      headers:{
        "accept": "application/json;odata=verbose",  
        "content-type": "application/json;odata=verbose",
        "X-RequestDigest": digest
      },
      data: JSON.stringify({  
        '__metadata': {  
            'type': 'SP.Data.EmployeesListItem'  
        },  
        'Title': 'Name1'
    }),  
      success: (resultData) => {             
          //alert(resultData.d.results);
          this.updateStatus(`Item successfully created`);
      },    
      error : (errorThrown) => {  
        this.updateStatus('Loading all items failed with error: ' + JSON.stringify(errorThrown));
      }    
  });


});
  }
  //create http item
  private createHttpItem(): void {
    const body: string = JSON.stringify({  
      'Title': `Item ${new Date()}`  
    });  
    
    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items`,  
    SPHttpClient.configurations.v1,  
    {  
      headers: {  
        'Accept': 'application/json;odata=nometadata',  
        'Content-type': 'application/json;odata=nometadata',  
        'odata-version': ''  
      },  
      body: body  
    })  
    .then((response: SPHttpClientResponse): Promise<ISPList> => {  
      return response.json();  
    })  
    .then((item: ISPList): void => {       
      this.updateStatus(`Item '${item.Title}' (ID: ${item.Id}) successfully created`);
    }, (error: any): void => {  
      this.updateStatus('Loading all items failed with error: ' + error);
    });  
  }


  //read ietms by ajax
  private readAjaxItem(): void {
    var reactHandler = this;    
    jquery.ajax({    
        url: `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items?$select=Title,Id`,    
        type: "GET",    
        headers:{'Accept': 'application/json; odata=verbose;'},    
        success: (resultData) => {             
            //alert(resultData.d.results);
            this.updateStatus(`Successfully loaded ${resultData.d.results.length} items`, resultData.d.results);
        },    
        error : (jqXHR, textStatus, errorThrown) => {  
          this.updateStatus('Loading all items failed with error: ' + errorThrown);
        }    
    });    
  }

  //Read Item by spHttpClient
  private readHttpItem(): void {
    this.updateStatus('Loading all items...');
    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items?$select=Title,Id`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      })
      .then((response: SPHttpClientResponse): Promise<{ value: ISPList[] }> => {
        return response.json();
      })
      .then((response: { value: ISPList[] }): void => {
        this.updateStatus(`Successfully loaded ${response.value.length} items`, response.value);
      }, (error: any): void => {
        this.updateStatus('Loading all items failed with error: ' + error);
      });
  }

  private updateStatus(status: string, items: ISPList[] = []): void {
    this.domElement.querySelector('.status').innerHTML = status;
    this.updateItemsHtml(items);
  }

  private updateItemsHtml(items: ISPList[]): void {
    this.domElement.querySelector('.items').innerHTML = items.map(item => `<li>${item.Title} (${item.Id})</li>`).join("");
  }
  //

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
