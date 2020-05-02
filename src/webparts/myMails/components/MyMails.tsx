import * as React from 'react';
import styles from './MyMails.module.scss';
import { IMyMailsProps } from './IMyMailsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {AppInsights} from "applicationinsights-js";
import { PivotItem, IPivotItemProps, Pivot, PivotLinkFormat } from 'office-ui-fabric-react/lib/Pivot';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types'; //npm install @microsoft/microsoft-graph-types --save-dev
//import CommonUtils from '../../../common/CommonUtils';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { ISPMessage } from './ISPMessage';
import { IMyMailsState } from './IMyMailsState';
import * as strings from 'MyMailsWebPartStrings';
import { Facepile, IFacepilePersona, IFacepileProps } from 'office-ui-fabric-react/lib/Facepile';
import { PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { DisplayMails } from './DIsplayMails';

export default class MyMails extends React.Component<IMyMailsProps, IMyMailsState> {
  constructor(props: IMyMailsProps) {
    super(props);

    this.state = {
      unreadMailsCount: 0,
      allMailsCount: 0,
      allMails: [],
      unreadMails: [],
      selectedTab: "UnreadMail",
      readyToLoadAllMails: false,
      readyToLoadUnread: false, //Unread emails Data is not ready to load
      showUserPanel: false,
      userInfoForPanel: '',
      readyToLoadPanelData: false
    };

    this.onLinkClick = this.onLinkClick.bind(this);
  }

  public async componentDidMount(): Promise<void> {
    if(this.props.trackInsights) {
      if(this.props.teamsContext) {
        AppInsights.trackEvent("Component Did Mount - Teams",
          { 
            SW_PageUrl: this.props.context.pageContext.web.absoluteUrl,
            SW_UserName: this.props.context.pageContext.user.displayName,
            SW_UserEmail: this.props.context.pageContext.user.email,
            SW_Source: 'Teams'},
            
          { timeTaken: new Date() }
        );
      }
      else {
        AppInsights.trackEvent("Component Did Mount - SP",
          { SW_PageUrl: this.props.context.pageContext.web.absoluteUrl,
            SW_UserName: this.props.context.pageContext.user.displayName,
            SW_UserEmail: this.props.context.pageContext.user.email,
            SW_Source: 'SharePoint'},
          { timeTaken: new Date() }
        );
      }
    }

    this._getEmails("All");
  }

  //Gets emails using MS Graph
  private _getEmails(emailType: string): string {
    let graphQuery: string = '';

    (emailType == 'Unread') 
      ? graphQuery = 'me/mailFolders/Inbox/messages?$filter=isRead ne true&$count=true'
      : graphQuery = 'me/mailFolders/Inbox/messages?$count=true';

    this.props.context.msGraphClientFactory.getClient()
      .then((graphClient: MSGraphClient): void => {
        //Timer Start
        let _startTime = new Date();
        graphClient.api(graphQuery)
          .get((error: any, messages: MicrosoftGraph.Message[], rawResponse?: any) => {
            // end timer
            let _endTime = new Date();
            var timeTaken: number = _endTime.valueOf() - _startTime.valueOf();
            console.log(`GetEmails - ${ emailType } : Took ${ timeTaken } ms to call Graph.`);
            if(this.props.trackInsights) {
              if(this.props.teamsContext) {
                AppInsights.trackEvent("Get Emails using MS Graph for Teams",
                  { SW_MSGraphUrl: graphQuery,
                    SW_EmailType: emailType,
                    SW_PageUrl: this.props.context.pageContext.web.absoluteUrl,
                    SW_UserName: this.props.context.pageContext.user.displayName,
                    SW_UserEmail: this.props.context.pageContext.user.email,
                    SW_Source: 'Teams'},
                    
                  { timeTaken: timeTaken }
                );
                AppInsights.trackTrace({
                  message: 'MS Graph Query executed for Teams'
                });
              }
              else {
                console.log("SP Context");
                AppInsights.trackEvent("Get Emails using MS Graph for SP",
                  { SW_MSGraphUrl: graphQuery,
                    SW_EmailType: emailType,
                    SW_PageUrl: this.props.context.pageContext.web.absoluteUrl,
                    SW_UserName: this.props.context.pageContext.user.displayName,
                    SW_UserEmail: this.props.context.pageContext.user.email,
                    SW_Source: 'SharePoint'},
                  { timeTaken: timeTaken }
                );

                AppInsights.trackTrace({
                  message: 'MS Graph Query executed for SP'
                });
              }
            }
            
              let mailsList: ISPMessage[] = [];
              
              if(messages != null && typeof(messages) !== 'undefined' && messages["value"].length > 0) {
                messages["value"].map((messageItem) => 
                mailsList.push({
                  from_Email: messageItem.from.emailAddress.address as string,
                  from_Name: messageItem.from.emailAddress.name,
                  subject: messageItem.subject, 
                  webLink: messageItem.webLink, 
                  receivedDate: messageItem.receivedDateTime})
                );

                if(emailType == 'Unread') {
                  this.setState({
                    //unreadMailsCount: messages["@odata.count"],
                    unreadMails: mailsList,
                    readyToLoadUnread: true
                  });
                }
                else {
                  this.setState({
                    //allMailsCount: messages["@odata.count"],
                    allMails: mailsList,
                    readyToLoadAllMails: true
                  });
                }
              }
              else { //No Emails
                this.setState({
                  //unreadMailsCount: messages["@odata.count"],
                  unreadMails: mailsList,
                  readyToLoadUnread: true,
                  //allMailsCount: messages["@odata.count"],
                  allMails: mailsList,
                  readyToLoadAllMails: true
                });
              }
            
          });
        });
      return '';
  }

  //On tab click
  public onLinkClick(item: PivotItem): void {
    if(item.props.headerText === 'Unread Emails') {
      this._getEmails('Unread');
      
      if(this.props.trackInsights) {
        AppInsights.trackEvent("Email Tab - Unread",
          { SW_EmailTab: 'Unread Emails',
            SW_PageUrl: this.props.context.pageContext.web.absoluteUrl,
            SW_UserName: this.props.context.pageContext.user.displayName,
            SW_UserEmail: this.props.context.pageContext.user.email},
          {  }
        );
      }
    }
    else {
      this._getEmails('All');
      
      if(this.props.trackInsights) {
        AppInsights.trackEvent("Email Tab - All",
          { SW_EmailTab: 'All Emails',
            SW_PageUrl: this.props.context.pageContext.web.absoluteUrl,
            SW_UserName: this.props.context.pageContext.user.displayName,
            SW_UserEmail: this.props.context.pageContext.user.email},
          {  }
        );
      }
    }
  }

  //On clicking of 'New Mail' button or individual mail item
  private _showMail = (event: any): any => {
    window.open(event, '_blank', 'location=yes,height=570,width=1000, top=150, left=300, scrollbars=yes,status=yes');
  }

  private _facePileClick = async (ev?: React.MouseEvent<HTMLElement>, persona?: IFacepilePersona) => {
    //this._getUserInfo(persona.data);
    this.setState({
      showUserPanel: true,
      readyToLoadPanelData: false
    })
    this._userInfoMSGraph(persona.data);
  }

  private _hidePanel = () => {
    this.setState({ 
      showUserPanel: false,
      readyToLoadPanelData: false
     });
  }

  private _userInfoMSGraph(email: string): any {
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        // From https://github.com/microsoftgraph/msgraph-sdk-javascript
        client
          .api('/users/me')
          //.version("v1.0")
          //.select("displayName,mail,userPrincipalName")
          //.filter(`(givenName eq '${escape(this.state.searchFor)}') or (surname eq '${escape(this.state.searchFor)}') or (displayName eq '${escape(this.state.searchFor)}')`)
          .get((err, res) => {  
            if(res != null && res['userPrincipalName'].length > 0) {
              this.setState({ 
                showUserPanel: true,
                readyToLoadPanelData: true,
                userInfoForPanel: res 
              });
            }
            else {
              console.log(err);
              this.setState({ 
                showUserPanel: true,
                readyToLoadPanelData: true,
                userInfoForPanel: 'No Data Exists'
              });
            }
          });
      });
  }

  public render(): React.ReactElement<IMyMailsProps> {
    /*
    AppInsights.trackEvent("Component Did Mount",
      {},
      { timeTaken: timeTaken }
    );*/

    return (
      <div className={ styles.myMails }>
        <div className={styles.titleText}>MY EMAILS</div>
        <div className={styles.divider}></div>

        <div className={styles.groupNav}>
          <div className={styles.groupNavLeft}>
            <DefaultButton className={styles.button} secondaryText="New Mail Message" onClick={() => this._showMail('https://outlook.office.com/owa/?viewmodel=IMailComposeViewModelFactory&path=')} text="New Message" />
          </div>
          <div className={styles.rightBtn}><a href='#' className={styles.btnMore}>See More</a></div>
        </div>
        <div className={styles.divider}></div>
        <div>
          <Pivot className={styles.tabItem} linkFormat={PivotLinkFormat.tabs} onLinkClick={this.onLinkClick} selectedKey={this.state.selectedTab}>
            <PivotItem headerText={strings.AllEmailsLabel} itemIcon="News" key="AllMail">
              { this.state.readyToLoadAllMails
                ?
                  <DisplayMails mailsToDisplay={this.state.allMails} facePileClick={this._facePileClick} readyToLoad={this.state.readyToLoadAllMails} />
                :
                  <Spinner size={SpinnerSize.large} label="Loading..." ariaLive="assertive" />
              }
            </PivotItem>
            <PivotItem headerText={strings.UnreadEmailsLabel} itemIcon="Mail" key="UnreadMail">
              <DisplayMails mailsToDisplay={this.state.unreadMails} facePileClick={this._facePileClick} readyToLoad={this.state.readyToLoadUnread} />
            </PivotItem>
          </Pivot>
        </div>
        { 
          this.state.userInfoForPanel && this.state.userInfoForPanel['userPrincipalName'] != null && this.state.userInfoForPanel['userPrincipalName'].length > 0
          ?
            <Panel
              isOpen={this.state.showUserPanel}
              type={PanelType.smallFixedFar}
              onDismiss={this._hidePanel}
            >
              {this.state.readyToLoadPanelData
                ?
                  <div>
                    <h2>Name: {this.state.userInfoForPanel.displayName}</h2>
                    <h3>Email: {this.state.userInfoForPanel.mail}</h3>
                  </div>
                :
                  <Spinner size={SpinnerSize.large} label="Loading..." ariaLive="assertive" />
              }
          </Panel>
          :
            <Panel
              isOpen={this.state.showUserPanel}
              type={PanelType.smallFixedFar}
              onDismiss={this._hidePanel}
            >
              {this.state.readyToLoadPanelData
                ?
                  <h2>No Data</h2>
                :
                  <Spinner size={SpinnerSize.large} label="Loading..." ariaLive="assertive" />
              }
          </Panel>
        }
      </div>
    );
  }

}
