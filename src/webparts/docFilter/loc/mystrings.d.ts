declare interface IDocFilterWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  WebpartNameFieldLabel: string;
  SharePointListFieldLabel: string;
  SharePointViewFieldLabel: string;
  SharePointColumnFieldLabel: string;
}

declare module 'DocFilterWebPartStrings' {
  const strings: IDocFilterWebPartStrings;
  export = strings;
}
