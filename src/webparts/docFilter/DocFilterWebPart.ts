import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField, 
  PropertyPaneDropdown, IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';
import { 
  SPHttpClient, SPHttpClientResponse, 
  ODataVersion, ISPHttpClientConfiguration, SPHttpClientConfiguration
} from '@microsoft/sp-http';

import * as strings from 'DocFilterWebPartStrings';
import DocFilter from './components/DocFilter';
import Renderer from './components/Renderer';
import { IDocFilterProps } from './components/IDocFilterProps';

import {sp, Web, List} from 'sp-pnp-js';

export interface IDocFilterWebPartProps {
  description: string;
  webpartName: string;
  sharePointList: string;
  sharePointView: string;
  sharePointLink: string;
  sharePointColumn: string;
}

export interface ISPList {
  value: ISPListItem[];
}

export interface ISPListItem { 
  Title: string;
}

export default class DocFilterWebPart extends BaseClientSideWebPart<IDocFilterWebPartProps> {

  //private queryString: String = "http://sp2019server/sites/it/_api/web/lists/GetByTitle('HeaderText')/Items";
  //private siteUrl: string = this.context.pageContext.web.absoluteUrl;
  private siteUrl: string = "http://sp2019server/sites/it";

  private lists: IPropertyPaneDropdownOption[] = [];
  private listsDropDownDisabled: boolean = false;

  private views: IPropertyPaneDropdownOption[] = [];
  private viewsDropDownDisabled: boolean = false;

  private columns: any[] = [];
  private columnValues: IPropertyPaneDropdownOption[] = [
    {key: "Dokart", text: "Dokart"}
  ];

  private filter: string = 'all';

  private listName: string;
  private viewName: string;

  private web: Web = new Web(this.siteUrl);

  setButtonHandlers(): void {
    const spListContainer: HTMLElement = this.domElement.querySelector('[id^="content"]') as HTMLElement;

    const allButton = this.domElement.querySelector('[id^="all"]') as HTMLElement;
    allButton.addEventListener('click', () => {
      spListContainer.innerHTML = '';
      this.filter = 'all';
      this.render();
    });

    const spButton = this.domElement.querySelector('[id^="sp"]') as HTMLElement;
    spButton.addEventListener('click', () => {
      spListContainer.innerHTML = '';
      this.filter = 'SharePoint';
      this.render();
    });

    const officeButton = this.domElement.querySelector('[id^="office"]') as HTMLElement;
    officeButton.addEventListener('click', () => {
      spListContainer.innerHTML = '';
      this.filter = 'Office';
      this.render();
    });

    const confButton = this.domElement.querySelector('[id^="conf"]') as HTMLElement;
    confButton.addEventListener('click', () => {
      spListContainer.innerHTML = '';
      this.filter = 'Confidential';
      this.render();
    });
  }

  async loadColumns(listName: string, viewName: string): Promise<any> {
    const filter = `Hidden eq false and ReadOnlyField eq false`;
    
    return this.web.lists.getByTitle(listName).views.getByTitle(viewName).fields.select('Items').filter(filter).get().then((response) => {
      return response;
    }).catch(() => {
      console.log('List name or view name could not be found.');
    });
  }

  async getViewQuery(listName: string, viewName: string): Promise<any> {
    return this.web.lists.getByTitle(listName).views.getByTitle(viewName).select('ViewQuery').get().then((response) => {
      return response.ViewQuery;
    }).catch(() => {
      return console.log('List name or view name could not be found.');
    });
  }

  async loadItems(listName: string, query: string): Promise<any> {
    const viewFields = `<ViewFields><FieldRef Name='Alt_x0020_Text' />
      <FieldRef Name='AlternateThumbnailUrl' />
      <FieldRef Name='FileRef' />
      <FieldRef Name='ID' />
      <FieldRef Name='Title' />
    </ViewFields>`;
  
    const queryOptions = "<QueryOptions><ViewAttributes Scope='RecursiveAll'/><top=10000 /></QueryOptions>";

    const xml = '<View><Query>' + viewFields + query + queryOptions + '</Query></View>';

    return this.web.lists.getByTitle(listName).getItemsByCAMLQuery({'ViewXml': xml}, 'FileRef', 'File', 'File_x0020_Type', 'FieldValuesAsText').then((response: any) => {
      return response;
    }).catch(() => {
      return console.log('List name or view name could not be found.');
    });
  }

  async loadSPLists(): Promise<ISPList> {
    return this.context.spHttpClient.get(`${this.siteUrl}/_api/web/lists?$filter=BaseTemplate eq 101`, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      return response.json();
    });
  }

  async loadSPViews(listName: string): Promise<any> {
    return this.context.spHttpClient.get(`${this.siteUrl}/_api/web/lists/GetByTitle('${listName}')/Views`, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      return response.json();
    });
  }

  public render(): void {
    const element: React.ReactElement<IDocFilterProps> = React.createElement(
      DocFilter,
      {
        description: this.properties.description, 
        webpartName: this.properties.webpartName,
        sharePointList: this.properties.sharePointList,
        sharePointView: this.properties.sharePointView, 
        sharePointLink: this.properties.sharePointLink, 
        sharePointColumn: this.properties.sharePointColumn
      }
    );

    ReactDom.render(element, this.domElement);

    if (this.properties.sharePointList) {
      this.listName = this.properties.sharePointList;
      this.viewName = this.properties.sharePointView;
    }

    if (Environment.type === EnvironmentType.SharePoint) {
      if (this.listName && this.viewName) {
        this.loadColumns(this.listName, this.viewName).then((result) => {
          if (result) {
            this.columns = [];
            const fields = (result as any).Items.results || (result as any).Items;
            for (let f = 0; f < fields.length; f++) {
              this.columns.push({ key: fields[f], text: fields[f] });
            }
          }
        }).then(() => {
          this.getViewQuery(this.listName, this.viewName).then((result) => {
            this.loadItems(this.listName, result).then((items: any) => {
              Renderer.renderTitle(this.listName, this.domElement).then(() => {
                Renderer.renderList(items, this.columns, this.domElement, this.filter).then(() => {
                  this.setButtonHandlers();
                });
              });
            });
          });
        });
      }
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    //this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

    this.loadSPLists().then((result) => {
      result.value.forEach((item: ISPListItem) => {
        this.lists.push({
          key: item.Title, 
          text: item.Title
        });    
      });  
    }).then(() => {
      this.context.propertyPane.refresh();
      //this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      if (this.listName) this.onPropertyPaneFieldChanged();
    });
    
  }

  protected onPropertyPaneFieldChanged(): void {
    //this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');
    this.listName = this.properties.sharePointList;
    this.views = [];
    
    if (this.listName) {
      this.loadSPViews(this.listName).then((response) => {
        response.value.forEach((item: ISPListItem) => {
          this.views.push({
            key: item.Title, 
            text: item.Title
          });
        });
      }).then(() => {
        this.context.propertyPane.refresh();
        //this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      });
    }
    
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          /* header: {
            description: strings.PropertyPaneDescription
          }, */
          groups: [
            {
              groupName: 'Settings',
              groupFields: [
                /*PropertyPaneTextField('webpartName', {
                  label: strings.WebpartNameFieldLabel
                }),*/
                PropertyPaneDropdown('sharePointList', {
                  label: strings.SharePointListFieldLabel, 
                  options: this.lists, 
                  disabled: this.listsDropDownDisabled
                }), 
                PropertyPaneDropdown('sharePointView', {
                  label: strings.SharePointViewFieldLabel, 
                  options: this.views, 
                  disabled: this.viewsDropDownDisabled
                }), 
                PropertyPaneDropdown('sharePointColumn', {
                  label: strings.SharePointColumnFieldLabel, 
                  options: this.columnValues, 
                  disabled: this.viewsDropDownDisabled
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
