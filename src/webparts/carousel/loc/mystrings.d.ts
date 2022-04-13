declare interface ICarouselWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  SiteUrlFieldLabel: string;
  ListFieldLabel: string;
  WebPartConfigButtonLabel: string;
  WebpartConfigDescription: string;
  WebpartConfigIconText: string;
  TitleLabel:string;
  showStatusLabel:string;
}

declare module 'CarouselWebPartStrings' {
  const strings: ICarouselWebPartStrings;
  export = strings;
}
