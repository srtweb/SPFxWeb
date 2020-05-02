import * as React from 'react';
import styles from './CreateTeams.module.scss';
import { ICreateTeamsProps } from './ICreateTeamsProps';
import { ICreateTeamsState, CreationState } from './ICreateTeamsState';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types'; //npm install @microsoft/microsoft-graph-types --save-dev
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import * as strings from 'CreateTeamsWebPartStrings';
import { PeoplePicker, PrincipalType, IPeoplePickerUserItem } from "@pnp/spfx-controls-react/lib/PeoplePicker";


export default class CreateTeams extends React.Component<ICreateTeamsProps, ICreateTeamsState> {
  constructor(props: ICreateTeamsProps) {
    super(props);

    this.state = ({
      teamName: '', //New Team Name
      creationState: CreationState.notStarted
    });
  }

  public render(): React.ReactElement<ICreateTeamsProps> {
    const {
      teamName,
      teamDescription,
      createChannel,
      channelName,
      channelDescription,
      spinnerText,
      creationState,
    } = this.state;


    return (
      <div className={ styles.createTeams }>
        <h2>{strings.Welcome}</h2>
        <div className={styles.container}>
          {{
            0: <div>
              <div className={styles.teamSection}>
                <TextField required={true} label={strings.TeamNameLabel} value={teamName} onChanged={this._onTeamNameChange.bind(this)}></TextField>
                <TextField label={strings.TeamDescriptionLabel} value={teamDescription} onChanged={this._onTeamDescriptionChange.bind(this)}></TextField>
                <PeoplePicker
                  context={this.props.context}
                  titleText={strings.Owners}
                  personSelectionLimit={3}
                  showHiddenInUI={false}
                  selectedItems={this._onOwnersSelected.bind(this)}
                  isRequired={true} />
                <PeoplePicker
                  context={this.props.context}
                  titleText={strings.Members}
                  personSelectionLimit={3}
                  showHiddenInUI={false}
                  selectedItems={this._onMembersSelected.bind(this)} />
              </div>
              <Checkbox label={strings.CreateChannel} checked={createChannel} onChange={this._onCreateChannelChange.bind(this)} />
              {createChannel && <div>
                <div className={styles.channelSection}>
                  <TextField required={createChannel} label={strings.ChannelName} value={channelName} onChanged={this._onChannelNameChange.bind(this)}></TextField>
                  <TextField label={strings.ChannelDescription} value={channelDescription} onChanged={this._onChannelDescriptionChange.bind(this)}></TextField>
                </div>
              </div>}
              <div className={styles.buttons}>
                <PrimaryButton text={strings.Create} className={styles.button} onClick={this._onCreateClick.bind(this)} />
                <DefaultButton text={strings.Clear} className={styles.button} onClick={this._onClearClick.bind(this)} />
              </div>
            </div>,
            1: <div>
              <Spinner label={spinnerText} />
            </div>,
            2: <div>
              <div>{strings.Success}</div>
              <PrimaryButton iconProps={{ iconName: 'TeamsLogo' }} href='https://aka.ms/mstfw' target='_blank'>{strings.OpenTeams}</PrimaryButton>
              <DefaultButton onClick={this._onClearClick.bind(this)}>{strings.StartOver}</DefaultButton>
            </div>,
            4: <div>
              <div className={styles.error}>{strings.Error}</div>
              <DefaultButton onClick={this._onClearClick.bind(this)}>{strings.StartOver}</DefaultButton>
            </div>
          }[creationState]}
        </div>
      </div>
    );
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

  private _onCreateChannelChange(e: React.FormEvent<HTMLElement | HTMLInputElement>, checked: boolean) {
    this.setState({
      createChannel: checked
    });
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

  private async _onCreateClick() {
    this._processCreationRequest();
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

  //Create O365 Group
  private async _createGroup(graphClient: MSGraphClient): Promise<string> {
    const displayName = this.state.teamName;
    const mailNickname = this._generateMailNickname(displayName);

    let {
      owners,
      members
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

    if (owners && owners.length) {
      groupRequest['owners@data.bind'] = owners.map(owner => {
        return `https://graph.microsoft.com/v1.0/users/${owner}`;
      });
    }
    if (members && members.length) {
      groupRequest['members@data.bind'] = members.map(member => {
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
        graphClient.api(`groups/b80ee56a-01f5-4aa1-a9f1-84155bc97ddd/team`).version('v1.0').put({
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

  private async _createTeamWithBeta(graphClient: MSGraphClient): Promise<string> {
    return new Promise<string>(resolve => {
      graphClient.api(`beta/teams`).version('beta').post({
        "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
        "displayName": "My Sample Team",
        "description": "My Sample Team’s Description"
      }).then(response => {
        resolve(response.id);
      }, () => {
        resolve('');
      });
    });
  }

  //Create channel in the team
  private async _createChannel(teamId: string, graphClient: MSGraphClient): Promise<string> {
    const {
      channelName,
      channelDescription
    } = this.state;

    try {
      const response = await graphClient.api(`teams/${teamId}/channels`).version('v1.0').post({
        displayName: channelName,
        description: channelDescription
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
  }

  private _onClearClick() {
    this.setState({
      teamName: '',
      teamDescription: '',
      members: [],
      owners: [],
      createChannel: false,
      channelName: '',
      channelDescription: '',
      creationState: CreationState.notStarted,
      spinnerText: ''
    });
  }

}
