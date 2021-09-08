import { ICarouselImages } from './ICarouselmages';
export interface ICarouselState {
  carouselImages: any[];
  files: any[];
  isLoading: boolean;
  errorMessage: string;
  hasError: boolean;
  teamsTheme: string;
  photoIndex: number;
  loadingImage: boolean;
  folderServerRelativeUrl: string;
}
