import * as React from 'react';
import styles from './CreateChannels.module.scss';
import { ICreateChannelsProps, IMyTeams } from './ICreateChannelsProps';
import { ICreateChannelsState } from './ICreateChannelsState';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types'; //npm install @microsoft/microsoft-graph-types --save-dev
import { Channel } from "@microsoft/microsoft-graph-types";
//import CommonUtils from '../../../common/CommonUtils';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import { DefaultButton } from 'office-ui-fabric-react';



export default class CreateChannels extends React.Component<ICreateChannelsProps, ICreateChannelsState> {
  constructor(props: ICreateChannelsProps) {
    super(props);

    this.state = ({
      messageToDisplay: '', //Message to display
      myTeams: [], //All Teams
      selectedTeam: '', //Drop down selected Team
      existingChannels: [], //Holds existing channel names
      newChannel: '', //New Channel Name
      createNewChannel: true //Create new channel if 'true'
    });

    //this._onTeamsSelect = this._onTeamsSelect.bind(this);
    this._createChannel = this._createChannel.bind(this);
  }

  public async componentDidMount() {
    //Get My Teams
    this.setState({
      messageToDisplay: 'Fetching Teams'
    });
    await this._getMyTeams();

  }

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
                  myTeams: ownedTeams,
                  messageToDisplay: ''
                });
              }
              else {
                //No Teams
                this.setState({
                  messageToDisplay: `You do not have access to any Teams.`
                });  
              }
            }
            else if(error != null) {
              this.setState({
                messageToDisplay: `Error fetching Teams - ${error}`
              });
            }

          });
      });
  }

  private async _createChannel() {
    //Check whether the Channel already exists
    await this.setState({
      messageToDisplay: 'Validating Channel..'
    });

    const channelQ = `/teams/${this.state.selectedTeam.key}/channels`;
    this.props.context.msGraphClientFactory.getClient()
      .then((graphClient: MSGraphClient): void => {
        graphClient.api(channelQ)
          .version("v1.0")
          .get(async(error: any, teamChannels: any, rawResponse?: any) => {
            if(teamChannels != null) {
              if(teamChannels.value != null && teamChannels.value.length > 0) {
                let allChannels: IMyTeams[] = [];
                let createNew: boolean = true;
                for (let teamChannel of teamChannels.value) {
                  if(teamChannel.displayName === this.state.newChannel) {
                    createNew = false;
                    break;
                  }
                  allChannels.push({
                    text: teamChannel.displayName,
                    key: teamChannel.id
                  });  
                }
                if(createNew) {
                  await this.setState({
                    messageToDisplay: 'Creating Channel..'
                  });
                  const newChannel: any = {
                    "displayName": this.state.newChannel,
                    "isFavoriteByDefault": true,
                    "membershipType": "private"
                  };

                  this.props.context.msGraphClientFactory.getClient()
                    .then((graphClient: MSGraphClient): void => {
                      graphClient.api(`/teams/${this.state.selectedTeam.key}/channels`)
                        .version('beta')
                        .post(newChannel)
                        .then((res) => {
                          console.log(res);
                          if(res != null && res.webUrl != null && res.webUrl != "") {
                            this.setState({
                              messageToDisplay: 'New Channel has been created - ' + res.webUrl
                            });
                          }
                        });
                    });
                }
                else {
                  //Channel already exists
                  this.setState({
                    messageToDisplay: `Channel '${this.state.newChannel}' already exists. Failed to create.`
                  });
                }
              }
              else {
                //No Channels. Create New
              }
            }
            else if(error != null) {
              await this.setState({
                messageToDisplay: `Error fetching Channels - ${error.code} - ${error.message}`
              });
            }

          });
      });
}

    public render(): React.ReactElement<ICreateChannelsProps> {
    return (
      <div className={ styles.createChannels }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Create New Channel</span>
              < br /> <br />
              <div>
                <Dropdown
                  placeholder="Select a Team"
                  label="Select a Team"
                  options={this.state.myTeams}
                  required
                  onChange={async(ev, opt) => {
                    await this.setState({
                      selectedTeam: opt
                    });
                  }}
                  id="ddlTeams"
                />
              </div>
              <div>
                <TextField label="Channel" placeholder="Enter Channel Name" value={this.state.newChannel} onChanged={e => this.setState({ newChannel: e })} />
              </div>
              <br />
              <div>
                <DefaultButton text="Create Channel" onClick={this._createChannel} />
              </div>
              <br />
              <div>
                <span><b>{this.state.messageToDisplay}</b></span>
              </div>
              
            </div>
          </div>
        </div>
      </div>
    );
  }
}
