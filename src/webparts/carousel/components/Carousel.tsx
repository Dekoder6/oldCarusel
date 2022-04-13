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

import Carousel from 'react-multi-carousel';
import 'react-multi-carousel/lib/styles.css';


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



export default class Carousels extends React.Component<ICarouselProps, ICarouselState> {
	private spService: spservices = null;
	private _teamsContext: microsoftTeams.Context = null;
	private Files: any = [];
	private arrayImages: any = [];
	private responsive = {
		desktop: {
			breakpoint: { max: 2000, min: 1080 },
			items: 1
		},
		tablet: {
			breakpoint: { max: 1080, min: 464 },
			items: 1
		},
		mobile: {
			breakpoint: { max: 464, min: 0 },
			items: 1
		}
	};
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
		this.getFiles(getServerUrl);
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


	public getFiles(folderUrl): void {
		const web = new Web(this.props.siteUrl);



		web.getFolderByServerRelativeUrl(folderUrl)
			.expand("Folders, Files,Files/ListItemAllFields/FieldValuesAsText").select('Files,Files/ListItemAllFields/FieldValuesAsText')
			.get().then((getFolder: any) => {
				getFolder.Folders.forEach(item => {
					this.getFiles(item.ServerRelativeUrl);
				});
				getFolder.Files.forEach((file) => {				
					if (file.ListItemAllFields && file.ListItemAllFields.FieldValuesAsText && file.ListItemAllFields.FieldValuesAsText.status.toLowerCase() == this.props.showStatus.toLowerCase()) {
						this.Files.push(file);
					}
				});
				this.setState({
					carouselImages: this.Files,
					isLoading: false
				});
			}).catch((ex) => {
				console.error(ex);
			});
	}

	private clickImage(path) {
		location.assign(path);
	}

	public render(): React.ReactElement<ICarouselProps> {
		const { carouselImages, isLoading, hasError, errorMessage, loadingImage } = this.state;
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
				{
					(!this.props.list) ?
						<Placeholder iconName='Edit'
							iconText={strings.WebpartConfigIconText}
							description={strings.WebpartConfigDescription}
							buttonLabel={strings.WebPartConfigButtonLabel}
							hideButton={this.props.displayMode === DisplayMode.Read}
							onConfigure={this.onConfigure.bind(this)} />
						:
						hasError ?
							<MessageBar messageBarType={MessageBarType.error}>
								{errorMessage}
							</MessageBar>
							:
							isLoading ?
								<Spinner size={SpinnerSize.large} label='loading images...' />
								:
								carouselImages.length == 0 ?
									<div style={{ width: '300px', margin: 'auto' }}>
										<Icon iconName="PhotoCollection"
											style={{ fontSize: '250px', color: '#d9d9d9' }} />
										<Label style={{ width: '250px', margin: 'auto', fontSize: FontSizes.size20 }}>No images in the library</Label>
									</div>
									:
									<div style={{ width: '100%', height: '100%' }}>
										<Carousel
											//ssr
											responsive={this.responsive}
											infinite={true}
											showDots={true}
											keyBoardControl={true}
											transitionDuration={500}
											autoPlay={true}
											autoPlaySpeed={5000}
											customTransition='all 1s'
										>
											{/* <Slider
												{...sliderSettings}
												autoplay={true}
												onReInit={() => {
													if (!loadingImage)
														$(".slideLoading").removeClass("slideLoading");
												}}> */}

											{carouselImages && carouselImages.length > 0 && carouselImages.map((galleryImage, i) => {										
												return (
													<div className='pointer' data-is-focusable={true} onClick={() => { this.clickImage(galleryImage.ListItemAllFields.FieldValuesAsText.FileRef.split(galleryImage.ListItemAllFields.FieldValuesAsText.FileLeafRef)[0]); }}>
														<Image src={galleryImage.ListItemAllFields.FieldValuesAsText.FileRef}
															onLoadingStateChange={async (loadState: ImageLoadState) => {
																if (loadState == ImageLoadState.loaded) {
																	this.setState({ loadingImage: false });
																}
															}}
															width={'100%'}
															height={'400px'}
															imageFit={ImageFit.centerContain}
														/>
													</div>

												);
											})}
											{/* </Slider> */}
										</Carousel>
										{/* {
											loadingImage &&
											<Spinner size={SpinnerSize.small} label={'Loading...'} style={{ verticalAlign: 'middle', right: '30%', top: 20, position: 'absolute', fontSize: FontSizes.size18, color: CommunicationColors.primary }}></Spinner>
										} */}
									</div>
				}
			</div >
		);
	}
}
