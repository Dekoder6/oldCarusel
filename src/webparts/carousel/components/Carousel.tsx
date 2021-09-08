import * as React from 'react';
import styles from './Carousel.module.scss';
import { ICarouselProps } from './ICarouselProps';
import { ICarouselState } from './ICarouselState';
import { escape } from '@microsoft/sp-lodash-subset';
import spservices from '../../../spservices/spservices';
import * as microsoftTeams from '@microsoft/teams-js';
import { ICarouselImages } from './ICarouselmages';
import 'video-react/dist/video-react.css'; // import css
import { Player, BigPlayButton } from 'video-react';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import * as $ from 'jquery';
import { FontSizes, } from '@uifabric/fluent-theme/lib/fluent/FluentType';
import { sp, Fields, Web, SearchResults, Field, PermissionKind, RegionalSettings, PagedItemCollection } from '@pnp/sp';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import * as strings from 'CarouselWebPartStrings';
import { DisplayMode } from '@microsoft/sp-core-library';
import { CommunicationColors } from '@uifabric/fluent-theme/lib/fluent/FluentColors';
//import parse from 'html-react-parser'
import {
	Spinner,
	SpinnerSize,
	MessageBar,
	MessageBarType,
	Label,
	Icon,
	ImageFit,
	Image,
	ImageLoadState,
} from 'office-ui-fabric-react';



export default class Carousel extends React.Component<ICarouselProps, ICarouselState> {
	private spService: spservices = null;
	private _teamsContext: microsoftTeams.Context = null;
	private Files: any = [];
	private arrayImages: any = [];
	public constructor(props: ICarouselProps) {
		super(props);
		this.state = {
			isLoading: true,
			errorMessage: '',
			hasError: false,
			teamsTheme: 'default',
			photoIndex: 0,
			carouselImages: [],
			files: [],
			loadingImage: true,
			folderServerRelativeUrl: null,
		};
	}

	private onConfigure() {
		this.props.context.propertyPane.open();
	}

	public async componentDidMount() {
		let getServerUrl = await this.getServerRelativeUrl();
		let getFile = await this.getFiles(getServerUrl);
		let checkedStatusImage = await this.checkedStatusImage(getFile);
		this.showImages(checkedStatusImage);
	}

	private getServerRelativeUrl(): Promise<any> {
		const web = new Web(this.props.siteUrl);
		return new Promise((resolve, reject) => {
			web.lists.getById(this.props.list)
				.select("Title", "RootFolder/ServerRelativeUrl").expand("RootFolder")
				.get().then((response) => {
					resolve(response.RootFolder.ServerRelativeUrl);
				});

		});
	}

	public getFiles(folderUrl): Promise<any> {
		return new Promise((resolve, reject) => {
			const web = new Web(this.props.siteUrl);
			web.getFolderByServerRelativeUrl(folderUrl)
				.expand("Folders, Files").select('Files,ListItemAllFields')
				.get().then((getFolder: any) => {
					getFolder.Folders.forEach(item => {
						return resolve(this.getFiles(item.ServerRelativeUrl));
					});
					getFolder.Files.forEach((file) => {
						if (file.Name.split('.')[1] == 'PDF' || file.Name.split('.')[1] == 'pdf' || file.Name.split('.')[1] == 'PNG' || file.Name.split('.')[1] == 'png' || file.Name.split('.')[1] == 'JPG' || file.Name.split('.')[1] == 'jpg' || file.Name.split('.')[1] == 'JPEG' || file.Name.split('.')[1] == 'jpeg' || file.Name.split('.')[1] == 'PPTX' || file.Name.split('.')[1] == 'pptx') {
							this.Files.push(file);
						}
					});
					return resolve(this.Files);
				}).catch((ex) => {
					reject(ex);
				});
		});
	}

	public checkedStatusImage(files): Promise<any> {
		const web = new Web(this.props.siteUrl);
		return new Promise((resolve, reject) => {
			Promise.all(files.map((file) => {
				return web.getFolderByServerRelativeUrl(file.ServerRelativeUrl).listItemAllFields.get().then((itemFiltred: any) => {
					if (itemFiltred.status == 'show') {
						this.arrayImages.push(file);
					}
					return this.arrayImages;
				});
			})).then(() => {
				resolve(this.arrayImages);
			});
		});

	}

	// public async componentDidUpdate(prevProps: ICarouselProps) {

	// 	if (!this.props.list || !this.props.siteUrl) return;
	// 	// Get  Properties change
	// 	if (prevProps.list !== this.props.list || prevProps.numberImages !== this.props.numberImages) {
	// 		/*
	// 		 this.galleryImages = [];
	// 		 this._carouselImages = [];
	// 		 this.setState({ images: this.galleryImages, carouselImages: t.his._carouselImages, isLoading: false });
	// 		 */
	// 		await this.loadPictures();
	// 	}
	// }

	public showImages(arrayWithImages): void {
		let galleryImages: ICarouselImages[] = [];
		let carouselImages: React.ReactElement<HTMLElement>[] = [];
		arrayWithImages.map((image) => {
			galleryImages.push(
				{
					serverRelativeUrl: image.ServerRelativeUrl.split(image.Name)[0],
					linkUrl: image.ServerRelativeUrl
				},
			);

		});

		carouselImages = galleryImages.map((galleryImage, i) => {
			return (
				<div className='slideLoading' >
					<div className='pointer' data-is-focusable={true} onClick={() => { this.clickImage(galleryImage.serverRelativeUrl); }}>
						<Image src={galleryImage.linkUrl}
							onLoadingStateChange={async (loadState: ImageLoadState) => {
								if (loadState == ImageLoadState.loaded) {
									this.setState({ loadingImage: false });
								}
							}}
							height={'400px'}
							imageFit={ImageFit.cover}
						/>
					</div>
				</div>
			);
		}
		);
		this.setState({ carouselImages: carouselImages, isLoading: false });
	}

	private clickImage(path) {
		location.assign(path);
	}

	public render(): React.ReactElement<ICarouselProps> {
		const sliderSettings = {
			dots: true,
			dotsClass: 'slick-dots',
			infinite: true,
			speed: 500,
			slidesToShow: 1,
			slidesToScroll: 1,
			lazyLoad: 'progressive',
			autoplaySpeed: 5000,
			initialSlide: this.state.photoIndex,
			arrows: true,
			draggable: true,
			adaptiveHeight: true,
			useCSS: true,
			useTransform: true,
		};

		return (
			<div className={styles.carousel} >
				<div>
				</div>
				{
					(!this.props.list) ?
						<Placeholder iconName='Edit'
							iconText={strings.WebpartConfigIconText}
							description={strings.WebpartConfigDescription}
							buttonLabel={strings.WebPartConfigButtonLabel}
							hideButton={this.props.displayMode === DisplayMode.Read}
							onConfigure={this.onConfigure.bind(this)} />
						:
						this.state.hasError ?
							<MessageBar messageBarType={MessageBarType.error}>
								{this.state.errorMessage}
							</MessageBar>
							:
							this.state.isLoading ?
								<Spinner size={SpinnerSize.large} label='loading images...' />
								:
								this.state.carouselImages.length == 0 ?
									<div style={{ width: '300px', margin: 'auto' }}>
										<Icon iconName="PhotoCollection"
											style={{ fontSize: '250px', color: '#d9d9d9' }} />
										<Label style={{ width: '250px', margin: 'auto', fontSize: FontSizes.size20 }}>No images in the library</Label>
									</div>
									:
									<div style={{ width: '100%', height: '100%' }}>

										<div style={{ width: '100%' }}>
											<Slider
												{...sliderSettings}
												autoplay={true}
												onReInit={() => {
													if (!this.state.loadingImage)
														$(".slideLoading").removeClass("slideLoading");
												}}>
												{
													this.state.carouselImages
												}
											</Slider>
										</div>
										{
											this.state.loadingImage &&
											<Spinner size={SpinnerSize.small} label={'Loading...'} style={{ verticalAlign: 'middle', right: '30%', top: 20, position: 'absolute', fontSize: FontSizes.size18, color: CommunicationColors.primary }}></Spinner>
										}
									</div>
				}
			</div >
		);
	}
}
