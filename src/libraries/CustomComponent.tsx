import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { PageContext } from '@microsoft/sp-page-context';
import {MSGraphClientFactory, SPHttpClient} from "@microsoft/sp-http";
import { initializeFileTypeIcons, getFileTypeIconProps } from '@uifabric/file-type-icons';
import { Icon, IconButton, TooltipHost, MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { followDocument, unFollowDocument, getFollowed, isUserManage } from './Services/Requests';
import styles from './component.module.scss';
import toast, { Toaster } from 'react-hot-toast';
import ListControls from './components/ListControls/ListControls';

export interface IObjectParam {
    myProperty: string;
}

export interface ICustomComponentProps {
    pageUrlParam? : string;
    pageTitleParam? : string;
    pageFileTypeParam? : string;   
    pageId?: string;
    pageContext?: PageContext; 
    sphttpClient?: SPHttpClient;
    msGraphClientFactory?: MSGraphClientFactory;
    pages?: any;
}

export function CustomComponent (props: ICustomComponentProps){

    console.log("props.pages", props.pages);

    initializeFileTypeIcons();
    
    const [myFollowedItems, setMyfollowedItems] = React.useState([]);
    const [iFrameVisible, setIFrameVisible] = React.useState(false);
    const [iFrameUrl, setIFrameUrl] = React.useState(null);

    // Add & View All Controls
    const uploadDocumentHandler = () => {
        const docUrl = props.pages.items[0].Path;
        setIFrameUrl(`${docUrl.substring(0, docUrl.lastIndexOf('/'))}/Forms/Upload.aspx`);
        setIFrameVisible(true);
        // toast.custom((t) => (
        //     <div className={styles.toastMsg}>
        //       <Icon iconName='Accept' /> Item has been added to the library. Please allow few minutes for it to update.
        //     </div>
        // ));
    };

    const viewAllHandler = () => {
        const docUrl = props.pages.items[0].Path;
        window.open(`${docUrl.substring(0, docUrl.lastIndexOf('/'))}/Forms/Allitems.aspx`, '_blank');
    };

    React.useEffect(()=>{
        getFollowed(props.msGraphClientFactory).then(res => {
            console.log("setMyfollowedItems(res)", res);
            setMyfollowedItems(res);
        });
    }, []);
    React.useEffect(()=>{
    }, [myFollowedItems.toString()]);


    // Follow & Unfollow
    const followDocHandler = (page: any) => {
        console.log("followDocHandler", page);
        followDocument(props.msGraphClientFactory, page.SiteId, page.WebId, page.ListId, page.ListItemID).then(() => {
            setMyfollowedItems(prev => {
                const currentFollowedItems = [...prev];
                currentFollowedItems.push({name: decodeURI(page.Filename), driveId: page.DriveId});
                console.log("currentFollowedItems", currentFollowedItems);
                return currentFollowedItems;
            });
            toast.custom((t) => (
                <div className={styles.toastMsg}>
                  <Icon iconName='Accept' /> Added to <a target='_blank' href="https://www.office.com/mycontent">Favorites!</a>
                </div>
            ));
        });
    };
    const unFollowDocHandler = (page: any) => {
        console.log("followDocHandler", page);
        unFollowDocument(props.msGraphClientFactory, page.SiteId, page.WebId, page.ListId, page.ListItemID).then(()=>{
            setMyfollowedItems(prev => {
                const currentFollowedItems = prev.filter(item => !(item.name === decodeURI(page.Filename) && item.driveId === page.DriveId));
                console.log("currentFollowedItems", currentFollowedItems);
                return currentFollowedItems;
            });
            toast.custom((t) => (
                <div className={styles.toastMsg}>
                  <Icon iconName='Accept' /> Removed from <a target='_blank' href="https://www.office.com/mycontent">Favorites!</a>
                </div>
            ));
        });
    };


    return(
        <>
            <Toaster position='bottom-center' toastOptions={{custom:{duration: 4000}}}/>
            
            <div className={styles.listViewNoWrap}>
				<table className={styles.customTable} cellPadding='0' cellSpacing='0'>
                    <colgroup>
                        <col width={'5%'} />
                        <col width={'5%'} />
                        <col width={'33%'} />
                        <col width={'10%'} />
                        <col width={'22%'} />
                        <col width={'20%'} />
                        <col width={'20%'} />
                    </colgroup>
					<thead>
						<tr>
							<th></th>
                            <th>ID</th>
							<th>Name</th>							
							<th>Legacy ID</th>		
                            <th>Department</th>					
							<th>Category</th>							
							<th>Document Type</th>							
						</tr>
					</thead>
					<tbody>
                        {props.pages.items.map(page => {
                            return (
                                <tr key={page.ListItemID}>
                                    <td>
                                        <div className={styles.formItem}>
                                            <div className={styles.favIconBtns}>
                                                {myFollowedItems.find(item => item.name === decodeURI(page.Filename) && item.driveId === page.DriveId ) ? 
                                                    <IconButton title='Unfavorite' onClick={() => unFollowDocHandler(page)} iconProps={{iconName : 'FavoriteStarFill'}} />
                                                : 
                                                    <IconButton title='Favorite' onClick={() => followDocHandler(page)} iconProps={{iconName : 'FavoriteStar'}} />
                                                }
                                            </div>
                                            <div className={styles.cellDiv}> 
                                                {page.FileType !== 'SharePoint.Link' &&
                                                    <a className={styles.attachmentLinkDownload} href={`${page.DefaultEncodingURL}`} title='Download' download>
                                                        <Icon iconName='Download' />
                                                    </a>
                                                }                                             
                                            </div>
                                        </div>
                                    </td>
                                    <td>{page.RefinableString139}</td>
                                    <td>
                                        <div className={styles.formItem}>
                                            <div className={styles.cellDiv}>
                                                <TooltipHost content={`${page.FileType} file`}>
                                                    <Icon {...getFileTypeIconProps({extension: page.FileType, size: 16}) }/>
                                                </TooltipHost> 
                                                <a className={styles.defautlLink + ' ' + styles.docLink} target="_blank" data-interception="off" href={page.DefaultEncodingURL}>{page.Title}</a>
                                            </div>
                                        </div>
                                    </td>
                                    <td>{page.RefinableString103}</td>
                                    <td>{page.MMIntranetDepartment}</td> 
                                    <td>{page.RefinableString138 && page.RefinableString138.replace(/;/g, ', ')}</td>
                                    <td>{page.DocType}</td>
                                </tr>
                            );
                        })}
                    </tbody>
				</table>

                {/* MMIntranetDepartment */}
                {/* RefinableString10 */}
                {/* RefinableString148 */}

                {isUserManage(props.pageContext) &&
                    <>
                        <ListControls 
                            iFrameVisible = {iFrameVisible}
                            setIFrameVisible = {setIFrameVisible}
                            iFrameUrl = {iFrameUrl}
                            uploadDocumentHandler={uploadDocumentHandler}
                            viewAllHandler={viewAllHandler}    
                        />
                        <MessageBar isMultiline={true} messageBarType={MessageBarType.warning}>
                            Please allow few minutes to see the updates if you have uploaded new documents. You can manually refresh the page to see the changes.
                        </MessageBar>
                    </>
                }
			</div>
            
        </>
    );

}

export class MyCustomComponentWebComponent extends BaseWebComponent {
    
    private sphttpClient: SPHttpClient;
    private pageContext: PageContext;
    private msGraphClientFactory: MSGraphClientFactory;

    public constructor() {
        super(); 
        this._serviceScope.whenFinished(()=>{
            this.pageContext = this._serviceScope.consume(PageContext.serviceKey);
            this.sphttpClient = this._serviceScope.consume(SPHttpClient.serviceKey);
            this.msGraphClientFactory = this._serviceScope.consume(MSGraphClientFactory.serviceKey);
        });
    }
 
    public async connectedCallback() {
        let props = this.resolveAttributes();
        const customComponent = <CustomComponent pageContext={this.pageContext} sphttpClient={this.sphttpClient} msGraphClientFactory={this.msGraphClientFactory} {...props}/>;
        ReactDOM.render(customComponent, this);
    }    
}