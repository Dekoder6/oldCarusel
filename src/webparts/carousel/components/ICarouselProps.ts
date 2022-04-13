import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from '@microsoft/sp-core-library';
export interface ICarouselProps {
  title: string;
  siteUrl: string;
  showStatus: string;
  list: string;
  context: WebPartContext;
  numberImages: number;
  updateProperty: (value: string) => void;
  displayMode: DisplayMode;
}
