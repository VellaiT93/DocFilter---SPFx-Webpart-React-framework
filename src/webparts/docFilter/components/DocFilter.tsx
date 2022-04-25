import * as React from 'react';
import styles from './DocFilter.module.scss';
import { IDocFilterProps } from './IDocFilterProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {sp, Web, List} from 'sp-pnp-js';
import DOMRenderer from './DOMRenderer';

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

export default class DocFilter extends React.Component<IDocFilterProps, {}> {
  private siteUrl: string = "http://sp2019server/sites/it";
  private web: Web = new Web(this.siteUrl);

  private columns: any[] = [];

  private filter: string = 'all';

  private _applyFilter(arg: string): void {
    const spListContainer: HTMLElement = this.props.context.domElement.querySelector('[id^="content"]') as HTMLElement;
    if (spListContainer.innerHTML !== '') spListContainer.innerHTML = '';

    this.filter = arg;
    this.componentDidUpdate();
  }

  private async _loadColumns(listName: string, viewName: string): Promise<any> {
    const filter = `Hidden eq false and ReadOnlyField eq false`;
    
    return this.web.lists.getByTitle(listName).views.getByTitle(viewName).fields.select('Items').filter(filter).get().then((response) => {
      return response;
    }).catch(() => {
      console.log('List name or view name could not be found.');
    });
  }

  private async _getViewQuery(listName: string, viewName: string): Promise<any> {
    return this.web.lists.getByTitle(listName).views.getByTitle(viewName).select('ViewQuery').get().then((response) => {
      return response.ViewQuery;
    }).catch(() => {
      return console.log('List name or view name could not be found.');
    });
  }

  private async _loadItems(listName: string, query: string): Promise<any> {
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

  private async _initiate(): Promise<void> {
    this._loadColumns(this.props.sharePointList, this.props.sharePointView).then((result) => {
      if (result) {
        this.columns = [];
        const fields = (result as any).Items.results || (result as any).Items;
        for (let f = 0; f < fields.length; f++) {
          this.columns.push({ key: fields[f], text: fields[f] });
        }
      }
    }).then(() => {
      this._getViewQuery(this.props.sharePointList, this.props.sharePointView).then((result) => {
        this._loadItems(this.props.sharePointList, result).then((items: any) => {
          DOMRenderer._renderTitle(this.props.sharePointList, this.props.context.domElement).then(() => {
            DOMRenderer._renderList(items, this.columns, this.props.context.domElement, this.filter);
          });
        });
      });
    });
  }

  public componentDidMount(): void {
    if (this.props.sharePointList && this.props.sharePointView) {
      this._initiate();
    }
  }

  public componentDidUpdate(): void {
    if (this.props.sharePointList && this.props.sharePointView) {
      this._initiate();
    }
  }

  public render(): React.ReactElement<IDocFilterProps> {
    return(
      <div id={styles.docFilter}>
        <div id={styles.library}>
          <div id={styles.title}>
            <a id='titleHolder' href='' target='_blank'></a>
          </div>
          <div id={styles.filterButtons}>
            <button id='alle' onClick={() => this._applyFilter('all')}>All</button>
            <button id='sp' onClick={() => this._applyFilter('SharePoint')}>SharePoint</button>
            <button id='office' onClick={() => this._applyFilter('Office')}>Office</button>
            <button id='conf' onClick={() => this._applyFilter('Confidential')}>Confidential</button>
          </div>
          <div className='spListContainer' id={styles.content}></div>
        </div>
      </div >
    );
  }
  
}
