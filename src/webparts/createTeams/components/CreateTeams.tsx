import * as React from 'react';
import styles from './CreateTeams.module.scss';
import { ICreateTeamsProps, IMyTeams } from './ICreateTeamsProps';
import { ICreateTeamsState, CreationState } from './ICreateTeamsState';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types'; //npm install @microsoft/microsoft-graph-types --save-dev
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import * as strings from 'CreateTeamsWebPartStrings';
import { PeoplePicker, PrincipalType, IPeoplePickerUserItem } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

const createOptions = [
  { key: 'Team', text: 'Create Microsoft Team' },
  { key: 'Channel', text: 'Create a Channel' }
];

const ctOptions = [
  { key: 'Standard', text: 'Standard - Accessible to everyone on the team' },
  { key: 'Private', text: 'Private - Accessible only to a specific group of people within the team' }
];

export default class CreateTeams extends React.Component<ICreateTeamsProps, ICreateTeamsState> {
  constructor(props: ICreateTeamsProps) {
    super(props);

    this.state = ({
      teamName: '', //New Team Name
      creationState: CreationState.notStarted,
      channelTeam: ''
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
                  cmyTeams: ownedTeams
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

  public render(): React.ReactElement<ICreateTeamsProps> {
    const {
      teamName,
      teamDescription,
      createChannel,
      cchannelName,
      cchannelDescription,
      spinnerText,
      creationState,
      cchannelType,
      Success,
      buttonText,
      messageToDisplay
    } = this.state;


    return (
      <div className={ styles.createTeams }>
        <h2>{strings.Welcome}</h2>
        <div className={styles.container}>
          {{
            0: <div>
                <Dropdown
                    placeholder="Select an option"
                    label="Select an option"
                    options={createOptions}
                    required
                    onChange={async(ev, opt) => {
                      await this.setState({
                        channelTeam: opt
                      });
                    }}
                    id="ddlChannelTeams"
                  />
                  {/*Create Team*/}
                  {this.state.channelTeam.key == "Team" && <div>
                    <div className={styles.teamSection}>
                      <TextField required={true} label={strings.TeamNameLabel} value={teamName} onChanged={this._onTeamNameChange.bind(this)}></TextField>
                      <TextField label={strings.TeamDescriptionLabel} value={teamDescription} onChanged={this._onTeamDescriptionChange.bind(this)}></TextField>
                      <PeoplePicker
                        context={this.props.context}
                        titleText={strings.Owners}
                        personSelectionLimit={3}
                        showHiddenInUI={false}
                        selectedItems={this._onTOwnersSelected.bind(this)}
                         />
                      <PeoplePicker
                        context={this.props.context}
                        titleText={strings.Members}
                        personSelectionLimit={3}
                        showHiddenInUI={false}
                        selectedItems={this._onTMembersSelected.bind(this)} />
                    </div>
                    </div>}
                    
                    {/*Create Channel*/}
                    {this.state.channelTeam.key == "Channel" && <div>
                      <Dropdown
                        placeholder="Select a Team"
                        label="Select a Team"
                        options={this.state.cmyTeams}
                        required
                        onChange={this._onTeamSelected.bind(this)}
                        id="ddlTeams"
                      />
                      <TextField required={true} label={strings.ChannelName} value={cchannelName} onChanged={this._onChannelNameChange.bind(this)}></TextField>
                      <TextField label={strings.ChannelDescription} value={cchannelDescription} onChanged={this._onChannelDescriptionChange.bind(this)}></TextField>
                      <Dropdown
                          placeholder="Select Channel Type"
                          label="Channel Type"
                          options={ctOptions}
                          required
                          errorMessage={cchannelType == 'Required' ? 'Channel Type is required' : undefined}
                          onChange={async(ev, opt) => {
                            await this.setState({
                              cchannelType: opt
                            });
                          }}
                          id="ddlChannelType"
                        />
                        { (cchannelType != undefined && cchannelType.key == 'Private') && <div>
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
                      </div>}
                
              

              {/* <Checkbox label={strings.CreateChannel} checked={createChannel} onChange={this._onCreateChannelChange.bind(this)} />
              {createChannel && <div>
                <div className={styles.channelSection}>
                  <TextField required={createChannel} label={strings.ChannelName} value={channelName} onChanged={this._onChannelNameChange.bind(this)}></TextField>
                  <TextField label={strings.ChannelDescription} value={channelDescription} onChanged={this._onChannelDescriptionChange.bind(this)}></TextField>
                </div>
              </div>} */}
              <div className={styles.buttons}>
                <PrimaryButton text={strings.Create} className={styles.button} onClick={this._onCreateClick.bind(this)} />
                {/* <DefaultButton text={strings.Clear} className={styles.button} onClick={this._onClearClick.bind(this)} /> */}
              </div>
            </div>,
            1: <div>
              <Spinner label={spinnerText} />
            </div>,
            2: <div>
              <div>{Success}</div><br /> <br />
              <PrimaryButton iconProps={{ iconName: 'TeamsLogo' }} href={this.state.channelTeamUrl} target='_blank'>{buttonText}</PrimaryButton>
              <DefaultButton onClick={this._onClearClick.bind(this)}>{strings.StartOver}</DefaultButton>
            </div>,
            4: <div>
              <div className={styles.error}>{messageToDisplay}</div>
              <DefaultButton onClick={this._onClearClick.bind(this)}>{strings.StartOver}</DefaultButton>
            </div>
          }[creationState]}
        </div>
      </div>
    );
  }

  private async _onTeamSelected(event, option) {
    this.setState({
      cselectedTeam: option
    });
    //Get Team Members
    const graphClient = await this.props.context.msGraphClientFactory.getClient();
    const tMembers = await graphClient.api(`groups/${option.key}/members`).version("v1.0", ).get();
    if(tMembers != null && tMembers.value.length > 0) {
      let teamMembers: string[] = [];
      tMembers.value.map(tMember => teamMembers.push(tMember.mail));
      this.setState({
        teamMembers: teamMembers
      });
    }
  }

  private _onTeamNameChange(value: string) {
    this.setState({
      teamName: value
    });
  }

  private _onTeamDescriptionChange(value: string) {
    this.setState({
      teamDescription: value
    });
  }

  //Channel Owners
  private _onOwnersSelected(owners: IPeoplePickerUserItem[]) {
    this.setState({
      cowners: owners.map(o => o.secondaryText)
    });
  }

  //Channel Members
  private _onMembersSelected(members: IPeoplePickerUserItem[]) {
    this.setState({
      cmembers: members.map(m => m.secondaryText)
    });
  }

  //Team Owners
  private _onTOwnersSelected(owners: IPeoplePickerUserItem[]) {
    this.setState({
      towners: owners.map(o => o.secondaryText)
    });
  }

  //Team Members
  private _onTMembersSelected(members: IPeoplePickerUserItem[]) {
    this.setState({
      tmembers: members.map(m => m.secondaryText)
    });
  }

  private _onCreateChannelChange(e: React.FormEvent<HTMLElement | HTMLInputElement>, checked: boolean) {
    this.setState({
      createChannel: checked
    });
  }

  private _onChannelNameChange(value: string) {
    this.setState({
      cchannelName: value
    });
  }

  private _onChannelDescriptionChange(value: string) {
    this.setState({
      cchannelDescription: value
    });
  }

  private async _wait(ms) {
    return new Promise(resolve => {
      setTimeout(resolve, ms);
    });
  }

  private async _onCreateClick() {
    const { towners, tmembers } = this.state;
    //this._processCreationRequest();
    //Team or Channel
    if(this.state.channelTeam != undefined && this.state.channelTeam.key == "Channel") {
      //Create Channel
      if(this.state.cchannelType == '' || this.state.cchannelType == undefined) {
        this.setState({
          cchannelType: 'Required'
        });
        return;
      }

      if(this.state.cselectedTeam != "" && this.state.cselectedTeam != undefined) {
        // initializing graph client
        const graphClient = await this.props.context.msGraphClientFactory.getClient();
  
        this.setState({
          creationState: CreationState.creating,
          spinnerText: strings.CreatingChannel
        });
        
        //Create channel
        //const channelId = await this._createChannel(this.state.teamName.key, graphClient);
        let channelUrl: any = '';
        if(this.state.cchannelType != null && this.state.cchannelType.key == "Private") {
          channelUrl = await this._ccreateChannel(this.state.cselectedTeam.key, "Private", graphClient);
        }
        else {
          //Standard Channel
          channelUrl = await this._ccreateChannel(this.state.cselectedTeam.key, "Standard", graphClient);
        }
      
        if (!channelUrl) {
          this._onError();
          return;
        }
        else if (channelUrl == "Exists") {
          this._onError(`'${this.state.cchannelName}' channel already exists.`);
          return;
        }
        else if (channelUrl.indexOf('ERROR -') >= 0) {
          this._onError(channelUrl);
          return;
        }
        else {
          this.setState({
            creationState: CreationState.created,
            Success: strings.cSuccess,
            buttonText: 'Open Channel',
            channelTeamUrl: channelUrl
          });
        }
      }
    }
    else if(this.state.channelTeam != undefined && this.state.channelTeam.key == "Team") {
      //Create Team
      if(this.state.teamName != undefined && this.state.teamName != "") {
        // initializing graph client
        const graphClient = await this.props.context.msGraphClientFactory.getClient();
  
        this.setState({
          creationState: CreationState.creating,
          spinnerText: strings.CreatingTeam
        });

        const teamUrl = await this._tCreateTeam(graphClient);

        if (!teamUrl) {
          this._onError();
          return;
        }
        else if (teamUrl == "Exists") {
          this._onError(`'${this.state.teamName}' team already exists.`);
          return;
        }
        else if (teamUrl.indexOf('ERROR -') >= 0) {
          this._onError(teamUrl);
          return;
        }
        else {
          this.setState({
            creationState: CreationState.created,
            Success: strings.tSuccess,
            buttonText: 'Open Team',
            channelTeamUrl: teamUrl
          });
        }
      }
    }
  }

  private async _processCreationRequest(): Promise<void> { 
    // initializing graph client
    const graphClient = await this.props.context.msGraphClientFactory.getClient();

    this.setState({
      creationState: CreationState.creating,
      spinnerText: strings.CreatingGroup
    });
    //this._createTeamWithBeta(graphClient);

    // Create a group
    const groupId = await this._createGroup(graphClient);
    if (!groupId) {
      this._onError();
      return;
    }

    this.setState({
      spinnerText: strings.CreatingTeam
    });

    //Create Team
    const teamId = await this._createTeamWithAttempts(groupId, graphClient);
    if (!teamId) {
      this._onError();
      return;
    }

    if (!this.state.createChannel) {
      this.setState({
        creationState: CreationState.created
      });
      return;
    }

    this.setState({
      spinnerText: strings.CreatingChannel
    });

    // Create channel
    const channelId = await this._createChannel(teamId, graphClient);
    if (!channelId) {
      this._onError();
      return;
    }
    else {
      this.setState({
        creationState: CreationState.created
      });
    }
  }

  private async _ccreateChannel(teamId: string, channelType: string, graphClient: MSGraphClient): Promise<string> {
    const {
      cchannelName,
      cchannelDescription,
      cowners,
      cmembers
    } = this.state;

    try {
      //Check for existing channels
      const existingChannels = await graphClient.api(`teams/${teamId}/channels`).version('v1.0').get();
      if(existingChannels != null) {
        for (let teamChannel of existingChannels.value) {
          if(teamChannel.displayName === cchannelName) {
            return 'Exists';
          }
        }
      }

      let channel: any = '';

      if(channelType == "Private") {
        channel = {
          '@odata.type': "#Microsoft.Teams.Core.channel",
          membershipType: "private",
          displayName: cchannelName,
          description: cchannelDescription,
          isFavoriteByDefault: true
        };

        const privateResponse = await graphClient.api(`teams/${teamId}/channels`).version('beta').post(channel);
        if(privateResponse != null) {
          //Add Members
          this.setState({
            spinnerText: strings.AddingMembers
          });

          //Owners
          if(cowners != null && cowners.length > 0) {
            for(let owner of cowners) {
              let cOwnerUserDetails = await this._getUserDetails(owner, graphClient);
              if(cOwnerUserDetails.id != null) {
                this.setState({
                  spinnerText: `Adding '${cOwnerUserDetails.displayName}' as Owner`
                });

                //Check whether the user is Team Member
                let isMember: boolean = await this._isUserTeamMember(cOwnerUserDetails.mail);
                try {
                  if(isMember) {
                    console.log(`${cOwnerUserDetails.mail} is a member`);
                    await this._addUserToChannel(cOwnerUserDetails.id, teamId, privateResponse.id, "owner", graphClient);
                  }
                  else {
                    console.log(`${cOwnerUserDetails.mail} is a not a member`);
                    await this._addUserToGroup(cOwnerUserDetails.id, graphClient);
                    console.log(`${cOwnerUserDetails.mail} has been added to Group`);
                    await this._wait(5000);
                    await this._addUserToChannel(cOwnerUserDetails.id, teamId, privateResponse.id, "owner", graphClient);
                  }
                }
                catch(ex) {
                  console.log('Error adding Owners - ' + cOwnerUserDetails.mail + " - " +  ex.message);
                }
               
                /* await graphClient.api(`teams/${teamId}/channels/${privateResponse.id}/members`).version('beta').post({
                  '@odata.type': "#microsoft.graph.aadUserConversationMember",
                  roles: ["owner"],
                  'user@odata.bind': `https://graph.microsoft.com/beta/users('${userDetails.id}')`
                }); */
              }
            }
          }

          //Members
          if(cmembers != null && cmembers.length > 0) {
            for(let member of cmembers) {
              let memUserDetails = await this._getUserDetails(member, graphClient);
              if(memUserDetails.id != null) {
                this.setState({
                  spinnerText: `Adding '${memUserDetails.displayName}' as Member`
                });

                //Check whether the user is Team Member
                let isMember: boolean = await this._isUserTeamMember(memUserDetails.mail);
                try {
                  if(isMember) {
                    console.log(`${memUserDetails.mail} is a member`);
                    await this._addUserToChannel(memUserDetails.id, teamId, privateResponse.id, "", graphClient);
                  }
                  else {
                    console.log(`${memUserDetails.mail} is a not a member`);
                    await this._addUserToGroup(memUserDetails.id, graphClient);
                    console.log(`${memUserDetails.mail} has been added to Group`);

                    await this._wait(5000);
                    //Pass "" for Members
                    await this._addUserToChannel(memUserDetails.id, teamId, privateResponse.id, "", graphClient);
                  }
                }
                catch(ex) {
                  console.log('Error adding Members - ' + memUserDetails.mail + " - " +  ex.message);
                }
                
              }
            }
          }
        }
        return privateResponse.webUrl;
      }
      else {
        const standardResponse = await graphClient.api(`teams/${teamId}/channels`).version('v1.0').post({
          displayName: cchannelName,
          description: cchannelDescription
        });
  
        return standardResponse.webUrl;
      }
    }
    catch (error) {
      console.error(error);
      return `ERROR - ${error.message}`;
    }
  }

  private async _tCreateTeam(graphClient: MSGraphClient): Promise<string> {
    await this._getMyTeams();

    const {
      teamName,
      teamDescription,
      towners,
      tmembers,
      cmyTeams
    } = this.state;

    try {
      //Check for existing Teams
      if(cmyTeams != null && cmyTeams != undefined && cmyTeams.length > 0) {
        for (let team of cmyTeams) {
          if(team.text == teamName) {
            return 'Exists';
          }
        }
      }

      this.setState({
        creationState: CreationState.creating,
        spinnerText: strings.CreatingTeam
      });
      const userDetails = await this._getUserDetails(this.props.context.pageContext.user.email, graphClient);
      let teamResult: any = 'Microsoft Team';

        await graphClient.api("teams").version('beta').post({
          "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
          "displayName": this.state.teamName,
          "description": this.state.teamDescription,
        }).then(async response => {
          await this._wait(5000);
          //Check for the Team
          const myTeams = await graphClient.api(`me/joinedTeams`).version("v1.0").get();
          if(myTeams != null && myTeams.value.length > 0) {
            for(let myTeam of myTeams.value) {
              let teamDesc: string = this.state.teamDescription;

              if(myTeam.displayName == this.state.teamName) {
                const teamDetails = await graphClient.api(`teams/${myTeam.id}`).version("v1.0").get();
                if(teamDetails != null && teamDetails.webUrl != null) {
                  //Add Members/Ownerrs
                  //Owners
                  if(towners != null && towners.length > 0) {
                    for(let owner of towners) {
                      let tOwnerUserDetails = await this._getUserDetails(owner, graphClient);
                      if(tOwnerUserDetails.id != null) {
                        this.setState({
                          spinnerText: `Adding '${tOwnerUserDetails.displayName}' as Owner`
                        });

                        try {
                          await this._addUserToTeam(tOwnerUserDetails.id.trim(), teamDetails.id.trim(), "owners", graphClient);
                          console.log(`${tOwnerUserDetails.mail} has been added to Group`);
                        }
                        catch(ex) {
                          console.log('Error adding Owners to Team - ' + this.state.teamName + " -" + tOwnerUserDetails.mail + " - " +  ex.message);
                        }
                      }
                    }
                  }

                  //Members
                  if(tmembers != null && tmembers.length > 0) {
                    for(let member of tmembers) {
                      let memberUserDetails = await this._getUserDetails(member, graphClient);
                      if(memberUserDetails.id != null) {
                        this.setState({
                          spinnerText: `Adding '${memberUserDetails.displayName}' as Member`
                        });

                        try {
                          await this._addUserToTeam(memberUserDetails.id.trim(), teamDetails.id.trim(), "members", graphClient);
                          console.log(`${memberUserDetails.mail} has been added to Group`);
                        }
                        catch(ex) {
                          console.log('Error adding Members to Team - ' + this.state.teamName + " -" + memberUserDetails.mail + " - " +  ex.message);
                        }
                      }
                    }
                  }

                  this.setState({
                    creationState: CreationState.created,
                    Success: strings.tSuccess,
                    buttonText: 'Open Team',
                    channelTeamUrl: teamDetails.webUrl
                  });
                  teamResult = teamDetails.webUrl;
                  //return teamDetails.webUrl;
                }
              }
            }
          }
          //return;
        }).catch(err => {
          console.log(err);
          teamResult = `ERROR - ${err.message}`;
          //return;
        });
      
      
        return teamResult;
      //const newTeam =  await this._createTeamWithBeta(userDetails.id, graphClient);
      //Get New Team Details
      

    }
    catch (error) {
      console.error(error);
      return `ERROR - ${error.message}`;
    }
  }

  private async _getUserDetails(email: string, graphClient: MSGraphClient): Promise<any> {
    const userDetails = await graphClient.api(`users/${email}`).version("v1.0", ).get();
    if(userDetails != undefined && userDetails.id != null) {
      return userDetails;
    }
    return;
  }

  //Add User to a Channel
  private async _addUserToChannel(userId: string, teamId: string, channelId: string, accessType: string, graphClient: MSGraphClient): Promise<any> {
    const userToChannel = await graphClient.api(`teams/${teamId}/channels/${channelId}/members`).version('beta').post({
      '@odata.type': "#microsoft.graph.aadUserConversationMember",
      roles: [`${accessType}`],
      'user@odata.bind': `https://graph.microsoft.com/beta/users('${userId}')`
    });
    if(userToChannel != undefined) {
      console.log('User To Channel - ' + userToChannel);
    }
    return;
  }

  //Add User to Group - Team
  private async _addUserToGroup(userId: string, graphClient: MSGraphClient): Promise<any> {
    const { cselectedTeam } = this.state;
    let timesQueryExecuted: number = 0;
    while(timesQueryExecuted <= 2) {
      if(cselectedTeam != undefined && cselectedTeam.key != null) {
        let addUserToGroup = await graphClient.api(`groups/${cselectedTeam.key}/members/$ref`).version("v1.0").post({
          '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`
        }); 
        await this._wait(2000);
  
        const groupMembers = await graphClient.api(`groups/${cselectedTeam.key}/members`).version("v1.0").get();
        if(groupMembers != null && groupMembers.value.length > 0) {
          for(let groupMember of groupMembers.value) {
            if(groupMember.id == userId) {
              return;
            }
          }
          timesQueryExecuted++;
          await this._wait(2000);
          continue;
        }
      }
     }
    return;
  }

  //Add User to Team - Based on Team ID
  private async _addUserToTeam(userId: string, teamId: string, accessType: string, graphClient: MSGraphClient): Promise<any> {
    await graphClient.api(`groups/${teamId}/${accessType}/$ref`).version("v1.0").post({	
      '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`	
    }); 
    return;
    /* let timesQueryExecuted: number = 0;
    while(timesQueryExecuted <= 2) {
      await graphClient.api(`groups/${teamId}/${accessType}/$ref`).version("v1.0").post({
        '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`
      });
      await this._wait(2000);

      const teamUsers = await graphClient.api(`groups/${teamId}/${accessType}`).version("v1.0", ).get();
      if(teamUsers != null && teamUsers.value.length > 0) {
        let teamMembers: string[] = [];
        for(let teamMember of teamUsers.value) {
          if(teamMember.id == userId) {
            return;
          }
        }
        timesQueryExecuted++;
        await this._wait(2000);
        continue;
      }
    }
    return; */
  }

  //Check whether the user is part of the Team
  private async _isUserTeamMember(email: string): Promise<boolean> {
    const { teamMembers } = this.state;
    if(teamMembers != null && teamMembers.length > 0) {
      for(let teamMember of teamMembers) {
        if(teamMember === email) {
          return true;
        }
      }
    }
    return false;
  }

  //Create O365 Group
  private async _createGroup(graphClient: MSGraphClient): Promise<string> {
    const displayName = this.state.teamName;
    const mailNickname = this._generateMailNickname(displayName);

    let {
      towners,
      tmembers
    } = this.state;

    const groupRequest = {
      displayName: displayName,
      description: this.state.teamDescription,
      groupTypes: [
        'Unified'
      ],
      mailEnabled: true,
      mailNickname: mailNickname,
      securityEnabled: false
    };

    if (towners && towners.length) {
      groupRequest['owners@data.bind'] = towners.map(owner => {
        return `https://graph.microsoft.com/v1.0/users/${owner}`;
      });
    }
    if (tmembers && tmembers.length) {
      groupRequest['members@data.bind'] = tmembers.map(member => {
        return `https://graph.microsoft.com/v1.0/users/${member}`;
      });
    }

    try {
      const response = await graphClient.api('groups').version('v1.0').post(groupRequest);
      return response.id;
    }
    catch (error) {
      console.error(error);
      return '';
    }
  }

  //Generates mail nick name by display name of the group
  private _generateMailNickname(displayName: string): string {
    return displayName.toLowerCase().replace(/\s/gmi, '-');
  }

  //Creates team. as mentioned in the documentation - we need to make 3 attempts if team creation request errored
  //https://docs.microsoft.com/en-us/graph/api/team-put-teams?view=graph-rest-1.0&tabs=http
  private async _createTeamWithAttempts(groupId: string, graphClient: MSGraphClient): Promise<string> {
    let attemptsCount = 0;
    let teamId: string = '';

    // From the documentation: If the group was created less than 15 minutes ago, it's possible for the Create team call to fail with a 404 error code due to replication delays. 
    // The recommended pattern is to retry the Create team call three times, with a 10 second delay between calls.
    //https://docs.microsoft.com/en-us/graph/api/team-put-teams?view=graph-rest-1.0&tabs=http

    do {
      teamId = await this._createTeam(groupId, graphClient);
      if (teamId) {
        attemptsCount = 3;
      }
      else {
        attemptsCount++;
      }
    } while (attemptsCount < 3);

    return teamId;
  }

  //Waits 10 seconds and tries to create a team
  private async _createTeam(groupId: string, graphClient: MSGraphClient): Promise<string> {
    return new Promise<string>(resolve => {
      setTimeout(() => {
        graphClient.api(`groups/${groupId}/team`).version('v1.0').put({
          memberSettings: {
            allowCreateUpdateChannels: true
          },
          messagingSettings: {
            allowUserEditMessages: true,
            allowUserDeleteMessages: true
          },
          funSettings: {
            allowGiphy: true,
            giphyContentRating: "strict"
          }
        }).then(response => {
          resolve(response.id);
        }, () => {
          resolve('');
        });
      }, 10000);
    });
  }

  private async _createTeamWithBeta(userID: string, graphClient: MSGraphClient): Promise<any> {
    const teamBody = {
      'template@odata.bind': "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      displayName: this.state.teamName,
      description: this.state.teamDescription,
    };
    const createTeam = await graphClient.api(`teams`).version("beta").post(teamBody);

    if(createTeam != null && createTeam != undefined) {
      return createTeam;
    }
    return createTeam;

    /* var team:any = {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": this.state.teamName,
      "description": "My Sample Team’s Description",
    };
    graphClient
      .api("https://graph.microsoft.com/beta/teams/")
      .post(team)
      .then((groupResponse) => {
         console.log(groupResponse);  
      }); */
    
  }

  //Create channel in the team
  private async _createChannel(teamId: string, graphClient: MSGraphClient): Promise<string> {
    const {
      cchannelName,
      cchannelDescription
    } = this.state;

    try {
      const response = await graphClient.api(`teams/${teamId}/channels`).version('v1.0').post({
        displayName: cchannelName,
        description: cchannelDescription
      });

      return response.id;
    }
    catch (error) {
      console.error(error);
      return '';
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
      teamName: '',
      teamDescription: '',
      cmembers: [],
      cowners: [],
      createChannel: false,
      cchannelName: '',
      cchannelDescription: '',
      creationState: CreationState.notStarted,
      spinnerText: '',
      cselectedTeam: '',
      cchannelType: '',
      channelTeam: []
    });
  }

}
