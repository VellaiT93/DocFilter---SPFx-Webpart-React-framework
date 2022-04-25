import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField, 
  PropertyPaneDropdown, IPropertyPaneDropdownOption, WebPartContext
} from '@microsoft/sp-webpart-base';
import { 
  SPHttpClient, SPHttpClientResponse, 
  ODataVersion, ISPHttpClientConfiguration, SPHttpClientConfiguration
} from '@microsoft/sp-http';

import * as strings from 'DocFilterWebPartStrings';
import DocFilter from './components/DocFilter';
import { IDocFilterProps } from './components/IDocFilterProps';

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

export interface ISPView {
  value: ISPViewItem[];
}

export interface ISPViewItem {
  Title: string;
}

export default class DocFilterWebPart extends BaseClientSideWebPart<IDocFilterWebPartProps> {

  //private queryString: String = "http://sp2019server/sites/it/_api/web/lists/GetByTitle('HeaderText')/Items";
  //private siteUrl: string = this.context.pageContext.web.absoluteUrl;
  private siteUrl: string = "http://sp2019server/sites/it";
  private listFilter: string = "BaseTemplate eq 101";

  private lists: IPropertyPaneDropdownOption[] = [];
  private listsDropDownDisabled: boolean = false;

  private views: IPropertyPaneDropdownOption[] = [];
  private viewsDropDownDisabled: boolean = false;

  private columnValues: IPropertyPaneDropdownOption[] = [
    {key: "Dokart", text: "Dokart"}
  ];

  private listName: string;

  // Load all SP libraries from site (only bibs --> Teplate 101)
  private async _loadSPLists(url: string): Promise<ISPList> {
    return this.context.spHttpClient.get(`${url}/_api/web/lists?$filter=${this.listFilter}`, 
    SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      return response.json();
    });
  }

  // Load SP views according to given list
  private async _loadSPViews(listName: string): Promise<ISPView> {
    return this.context.spHttpClient.get(`${this.siteUrl}/_api/web/lists/GetByTitle('${listName}')/Views`, 
    SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
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
        sharePointColumn: this.properties.sharePointColumn,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    //this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

    this._loadSPLists(this.siteUrl).then((result) => {
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
      this._loadSPViews(this.listName).then((response) => {
        response.value.forEach((item: ISPViewItem) => {
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
