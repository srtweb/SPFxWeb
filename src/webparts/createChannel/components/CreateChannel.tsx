import * as React from 'react';
import styles from './CreateChannel.module.scss';
import { ICreateChannelProps, IMyTeams } from './ICreateChannelProps';
import { ICreateChannelState, CreationState } from './ICreateChannelState';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types'; //npm install @microsoft/microsoft-graph-types --save-dev
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import * as strings from 'CreateChannelWebPartStrings';
import { PeoplePicker, PrincipalType, IPeoplePickerUserItem } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

const ctOptions = [
  { key: 'Standard', text: 'Standard - Accessible to everyone on the team' },
  { key: 'Private', text: 'Private - Accessible only to a specific group of people within the team' }
];

export default class CreateChannel extends React.Component<ICreateChannelProps, ICreateChannelState> {
  constructor(props: ICreateChannelProps) {
    super(props);

    this.state = ({
      teamName: '', //Seleted Team Name
      creationState: CreationState.notStarted,
      myTeams: [], //All Teams
      channelType: ''
    });
  }

  public async componentDidMount() {
    await this._getMyTeams();
  }

  //Get Teams to load the DDL
  private async _getMyTeams() {
    this.props.context.msGraphClientFactory.getClient()
      .then((graphClient: MSGraphClient): void => {
        graphClient.api('/me/joinedTeams')
          .get((error: any, myTeams: any, rawResponse?: any) => {
            if(myTeams != null) {
              if(myTeams.value != null && myTeams.value.length > 0) {
                let ownedTeams: IMyTeams[] = [];
                for (let myTeam of myTeams.value) {
                  ownedTeams.push({
                    text: myTeam.displayName,
                    key: myTeam.id
                  });
                }

                this.setState({
                  myTeams: ownedTeams
                });
              }
              else {
                //No Teams
                  
              }
            }
            else if(error != null) {
              
            }
          });
      });
  }

  public render(): React.ReactElement<ICreateChannelProps> {
    const {
      teamName,
      channelName,
      channelDescription,
      channelType,
      spinnerText,
      creationState,
      channelUrl,
      messageToDisplay
    } = this.state;

    return (
      <div className={ styles.createChannel }>
        <h2>{strings.Welcome}</h2>
        <div className={styles.container}>
          {{
            0: <div>
              <div className={styles.channelSection}>
                <Dropdown
                    placeholder="Select a Team"
                    label="Select a Team"
                    options={this.state.myTeams}
                    required
                    onChange={async(ev, opt) => {
                      await this.setState({
                        teamName: opt
                      });
                    }}
                    id="ddlTeams"
                  />
                <TextField required={true} label={strings.ChannelName} value={channelName} onChanged={this._onChannelNameChange.bind(this)}></TextField>
                <TextField label={strings.ChannelDescription} value={channelDescription} onChanged={this._onChannelDescriptionChange.bind(this)}></TextField>
                <Dropdown
                    placeholder="Select Channel Type"
                    label="Channel Type"
                    options={ctOptions}
                    required
                    errorMessage={channelType == 'Required' ? 'Channel Type is required' : undefined}
                    onChange={async(ev, opt) => {
                      await this.setState({
                        channelType: opt
                      });
                    }}
                    id="ddlChannelType"
                  />
                { (channelType.key == 'Private') && <div>
                  <PeoplePicker
                  context={this.props.context}
                  titleText={strings.Owners}
                  personSelectionLimit={3}
                  showHiddenInUI={false}
                  selectedItems={this._onOwnersSelected.bind(this)}
                  isRequired={false} />
                <PeoplePicker
                  context={this.props.context}
                  titleText={strings.Members}
                  personSelectionLimit={3}
                  showHiddenInUI={false}
                  selectedItems={this._onMembersSelected.bind(this)} />
                  </div>}
              </div>
              <div className={styles.buttons}>
                <PrimaryButton text={strings.Create} className={styles.button} onClick={this._onCreateClick.bind(this)} />
                  {/*<DefaultButton text={strings.Clear} className={styles.button} onClick={this._onClearClick.bind(this)} />*/}
              </div>
            </div>,
            1: <div>
              <Spinner label={spinnerText} />
            </div>,
            2: <div>
              <div>{strings.Success}</div> <br />
              <PrimaryButton iconProps={{ iconName: 'TeamsLogo' }} href={this.state.channelUrl} target='_blank'>{strings.OpenTeams}</PrimaryButton>
              <DefaultButton onClick={this._onClearClick.bind(this)}>{strings.StartOver}</DefaultButton>
            </div>,
            4: <div>
              <div className={styles.error}>{this.state.messageToDisplay}</div>
              <DefaultButton onClick={this._onClearClick.bind(this)}>{strings.StartOver}</DefaultButton>
            </div>
          }[creationState]}
        </div>
      </div>
    );
  }

  private _onChannelNameChange(value: string) {
    this.setState({
      channelName: value
    });
  }

  private _onChannelDescriptionChange(value: string) {
    this.setState({
      channelDescription: value
    });
  }

  private _onOwnersSelected(owners: IPeoplePickerUserItem[]) {
    this.setState({
      owners: owners.map(o => o.id)
    });
  }

  private _onMembersSelected(members: IPeoplePickerUserItem[]) {
    this.setState({
      members: members.map(m => m.id)
    });
  }

  //Get User ID based on Email Address
  private async _getUserID(emailID: string) {
    const graphClient = await this.props.context.msGraphClientFactory.getClient();
    const userID = await graphClient.api(`users/${emailID}`).version('v1.0').get();
    if(userID != null && userID.id != null) {
      return userID.id;
    }
    return '';
  }

  //Create channel in the team
 /*  private async _createChannel(teamId: string, graphClient: MSGraphClient): Promise<string> {
    const {
      channelName,
      channelDescription
    } = this.state;

    try {
      const response = await graphClient.api(`teams/${teamId}/channels`).version('v1.0').post({
        displayName: channelName,
        description: channelDescription
      });

      return response.webUrl;
    }
    catch (error) {
      console.error(error);
      return '';
    }
  } */

  private async _createChannel(teamId: string, channelType: string, graphClient: MSGraphClient): Promise<string> {
    const {
      channelName,
      channelDescription,
      owners,
      members
    } = this.state;

    try {
      //Check for existing channels
      const existingChannels = await graphClient.api(`teams/${teamId}/channels`).version('v1.0').get();
      if(existingChannels != null) {
        for (let teamChannel of existingChannels.value) {
          if(teamChannel.displayName === channelName) {
            return 'Exists';
          }
        }
      }

      let channel: any = '';

      if(channelType == "Private") {
        channel = {
          '@odata.type': "#Microsoft.Teams.Core.channel",
          membershipType: "private",
          displayName: channelName,
          description: channelDescription,
          isFavoriteByDefault: true
        };

        const privateResponse = await graphClient.api(`teams/${teamId}/channels`).version('beta').post(channel);
        if(privateResponse != null) {
          //Add Members
          this.setState({
            spinnerText: strings.AddingMembers
          });

          const conversationMember = {
            '@odata.type': "#microsoft.graph.aadUserConversationMember",
            roles: ["owner"],
            'user@odata.bind': "https://graph.microsoft.com/beta/users('4d124ae6-86fc-43d7-ac76-e258eddd0790')"
          };

          const addMembers = await graphClient.api(`teams/${teamId}/channels/${privateResponse.id}/members`).version('beta').post(conversationMember);
          if(addMembers != null) {
            console.log('Added Members');
          }

        }
        return privateResponse.webUrl;
      }
      else {
        const standardResponse = await graphClient.api(`teams/${teamId}/channels`).version('v1.0').post({
          displayName: channelName,
          description: channelDescription
        });
  
        return standardResponse.webUrl;
      }
      

     /*  
     ,
        members: [
          {
            '@odata.type':"#microsoft.graph.aadUserConversationMember",
            'user@odata.bind':"https://graph.microsoft.com/beta/users('i:0#.f|membership|srikanth@srtm365.onmicrosoft.com')",
            roles:["owner"]
          }
        ]
        
      const response = await graphClient.api(`teams/${teamId}/channels`).version('beta').post({
        displayName: channelName,
        description: channelDescription,
        isFavoriteByDefault: true,
        membershipType: "private"
      }); */

      const response = await graphClient.api(`teams/${teamId}/channels`).version('beta').post(channel);

      return response.webUrl;
    }
    catch (error) {
      console.error(error);
      return `ERROR - ${error.message}`;
    }
  }


  private async _createPrivateChannel(teamId: string, graphClient: MSGraphClient): Promise<string> {
    const {
      channelName,
      channelDescription,
      owners,
      members
    } = this.state;

    try {
      //Check for existing channels
      const existingChannels = await graphClient.api(`teams/${teamId}/channels`).version('v1.0').get();
      if(existingChannels != null) {
        for (let teamChannel of existingChannels.value) {
          if(teamChannel.displayName === channelName) {
            return 'Exists';
          }
        }
      }

      const channel = {
        '@odata.type': "#Microsoft.Teams.Core.channel",
        membershipType: "private",
        displayName: channelName,
        description: channelDescription,
        isFavoriteByDefault: true
      };

     /*  
     ,
        members: [
          {
            '@odata.type':"#microsoft.graph.aadUserConversationMember",
            'user@odata.bind':"https://graph.microsoft.com/beta/users('i:0#.f|membership|srikanth@srtm365.onmicrosoft.com')",
            roles:["owner"]
          }
        ]
        
      const response = await graphClient.api(`teams/${teamId}/channels`).version('beta').post({
        displayName: channelName,
        description: channelDescription,
        isFavoriteByDefault: true,
        membershipType: "private"
      }); */

      const response = await graphClient.api(`teams/${teamId}/channels`).version('beta').post(channel);

      return response.webUrl;
    }
    catch (error) {
      console.error(error);
      return `ERROR - ${error.message}`;
    }
  }

  private _onError(message?: string): void {
    this.setState({
      creationState: CreationState.error
    });
    if(message != "" || message == undefined) {
      this.setState({
        messageToDisplay: message
      });
    }
    else {
      this.setState({
        messageToDisplay: strings.Error
      });
    }
  }

  private _onClearClick() {
    this.setState({
      teamName: 'Select a Team',
      members: [],
      owners: [],
      channelName: '',
      channelDescription: '',
      channelType: [],
      creationState: CreationState.notStarted,
      spinnerText: ''
    });
  }

  //On button click
  private async _onCreateClick() { 
    if(this.state.channelType == '') {
      this.setState({
        channelType: 'Required'
      });
    }
    if(this.state.teamName != "") {
      // initializing graph client
      const graphClient = await this.props.context.msGraphClientFactory.getClient();

      this.setState({
        creationState: CreationState.creating,
        spinnerText: strings.CreatingChannel
      });

      //Create channel
      //const channelId = await this._createChannel(this.state.teamName.key, graphClient);
      let channelUrl: any = '';
      if(this.state.channelType.key == "Private") {
        channelUrl = await this._createChannel(this.state.teamName.key, "Private", graphClient);
      }
      else {
        //Standard Channel
        channelUrl = await this._createChannel(this.state.teamName.key, "Standard", graphClient);
      }
      
      if (!channelUrl) {
        this._onError();
        return;
      }
      else if (channelUrl == "Exists") {
        this._onError(`'${this.state.channelName}' channel already exists.`);
        return;
      }
      else if (channelUrl.indexOf('ERROR -') >= 0) {
        this._onError(channelUrl);
        return;
      }
      else {
        this.setState({
          creationState: CreationState.created,
          channelUrl: channelUrl
        });
      }
    }
  }
}
