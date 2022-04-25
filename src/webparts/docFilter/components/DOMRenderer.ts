import styles from './DocFilter.module.scss';

export default class Renderer {
  private static siteUrl: string = "http://sp2019server/sites/it";

  static async _renderTitle(title: string, dom: HTMLElement): Promise<void> {
    const titleHolder: HTMLLinkElement = dom.querySelector('[id^="titleHolder"]') as HTMLLinkElement;
    titleHolder.href = this.siteUrl + '/' + title;
    titleHolder.innerHTML = title;
  }

  static async _renderList(items: any, columns: any, dom: any, filter: string): Promise<void> {
    const spListContainer: HTMLElement = dom.querySelector('[class="spListContainer"]') as HTMLElement;
    let html: string = `<table id=${styles.listTable} style="border-collapse: collapse;">`;
    
    // Setup of columns
    html += `
      <thead>
        <tr>
          <td><img src='http://sp2019server/sites/it/_layouts/15/images/icgen.gif' /></td>
          <td>Name</td>
    `;

    for (let i = 3; i < columns.length; i++) {
      html += `<td>${columns[i].text}</td>`;
    }

    html += `
        </tr>
      </thead>
      <tbody>
    `;

    // Setup of table content
    items.forEach((item: any) => {
      let icon: string;
      let link: string;

      //icon = this.GetIconPath(item.File_x0020_Type);
      icon = Renderer.GetIconPath(item.File_x0020_Type, this.siteUrl);

      if (item.File_x0020_Type === 'doc' || item.File_x0020_Type === 'docx' 
        || item.File_x0020_Type === 'ppt' || item.File_x0020_Type === 'pptx' 
        || item.File_x0020_Type === 'xls' || item.File_x0020_Type === 'xlsx' 
        || item.File_x0020_Type === 'vsd' || item.File_x0020_Type === 'vsdx') {
          link = item.FieldValuesAsText.FileRef + '?web=1';
      } else {
        link = item.FieldValuesAsText.FileRef;
      }

      if (filter === item.Dokart || filter === 'all') {
        html += `
          <tr>
            <td><img src='${icon}' /></td>
            <td><a id='${styles.fileName}' href='${link}' target='_blank'>${item.File.Name}</a></td>
        `;

        for (let j = 3; j < columns.length; j++) {
          if (item[columns[j].text] === null) item[columns[j].text] = '';
          if (columns[j].text === 'Author') item[columns[j].text] = item.FieldValuesAsText.Author;
          if (columns[j].text === 'Editor') item[columns[j].text] = item.FieldValuesAsText.Editor;

          html += `<td>${item[columns[j].text]}</td>`;
        }

        html += `</tr>`;
      }
    });

    html += '</tbody></table>';

    spListContainer.innerHTML = html;
  }

  static GetIconPath(extn: string, siteUrl: string) {
    let imgPath: string = '';

    if (extn === 'pptx' || extn === 'ppt') {
      imgPath = '/_layouts/15/images/icpptx.png';
    } else if (extn === 'docx' || extn === 'doc') {
      imgPath = '/_layouts/15/images/icdocx.png';
    } else if (extn === 'xlsx' || extn === 'xls') {
      imgPath = '/_layouts/15/images/icxlsx.png';
    } else if (extn === 'vsdx' || extn === 'vsd') {
      imgPath = '/_layouts/15/images/icvsdx.png';
    } else if (extn === 'xml') {
      imgPath = '/_layouts/15/images/icxml.gif';
    } else if (extn === 'xlsm') {
      imgPath = '/_layouts/15/images/icxlsm.png';
    } else if (extn === 'csv') {
      imgPath = '/_layouts/15/images/icxls.png';
    } else if (extn === 'txt') {
      imgPath = '/_layouts/15/images/ictxt.gif';
    } else if (extn === 'pdf') {
      imgPath = '/_layouts/15/images/icpdf.png';
    } else {
      imgPath = '/_layouts/15/images/icgen.gif';
    }
    return siteUrl + imgPath;
  }
}