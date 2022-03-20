import * as React from 'react';
import styles from './DocFilter.module.scss';
import { IDocFilterProps } from './IDocFilterProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class DocFilter extends React.Component<IDocFilterProps, {}> {
  public render(): React.ReactElement<IDocFilterProps> {
    return (
      <div id={styles.docFilter}>
        <div id={styles.library}>
          <div id={styles.title}>
            <a id='titleHolder' href='' target='_blank'></a>
          </div>
          <div id={styles.filterButtons}>
            <button id='alle'>All</button>
            <button id='sp'>SharePoint</button>
            <button id='office'>Office</button>
            <button id='conf'>Confidential</button>
          </div>
          <div className='spListContainer' id={styles.content}></div>
        </div>
      </div >
    );
  }
}
