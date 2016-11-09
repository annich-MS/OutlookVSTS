/// <reference path="../../../office.d.ts" />
import * as React from 'react';
import { Provider, connect } from 'react-redux';
import {updateTeamSettingsAction, ISettingsInfo} from '../../Redux/LogInActions';
import {updatePopulatingAction, updateNotificationAction, PopulationStage, NotificationType } from '../../Redux/FlowActions';
import {Rest, RestError, Team } from '../../RestHelpers/rest';
import { Dropdown, IDropdownOptions } from 'office-ui-fabric-react';

/**
 * Properties needed for the AreaDropdown component
 * @interface IAreaProps
 */
interface IAreaProps {
  /**
   * intermediate to dispatch actions to update the global store
   * @type {any}
   */
  dispatch?: any;
  /**
   * user's VSTS memberID
   * @type {string}
   */
  id?: string;
  /**
   * user's email address
   * @type {string}
   */
  email?: string;
  /**
   * currently selected account option
   * @type {string}
   */
  account?: string;
  /**
   * currently selected project option
   * @type {string}
   */
  project?: string;
  /**
   * currently selected team option
   * @type {string}
   */
  team?: string;
  /**
   * Represents the lists of teams for current project
   * @type {ISettingsInfo[]}
   */
  teams?: ISettingsInfo[];

  /**
   * Represents what tier is currently being populated
   * @type {number}
   */
  populationStage?: PopulationStage;
}

/**
 * maps state in application store to properties for the component
 * @param {any} state
 */
function mapStateToProps(state: any): IAreaProps {
  return ({
    account: state.currentSettings.settings.account,
    email: state.userProfile.email,
    id: state.userProfile.memberID,
    populationStage: state.controlState.populationStage,
    project: state.currentSettings.settings.project,
    team: state.currentSettings.settings.team,
    teams: state.currentSettings.lists.teamList,
  });
}

@connect(mapStateToProps)

/**
 * Smart component
 * Renders area dropdown
 * @class {AreaDropdown} 
 */
export class AreaDropdown extends React.Component<IAreaProps, any> {

  public constructor() {
    super();
    this.populateTeams = this.populateTeams.bind(this);
  }

  /** 
   * executed first time component renders, reads the default project
   * @return {void}
   */
  public componentWillMount(): void {
    /*let defaultTeam: string = Office.context.roamingSettings.get('default_team');
    if (defaultTeam !== undefined) {
      this.props.dispatch(updateTeamSettingsAction(defaultTeam, this.props.teams));
    }*/
  }

  /**
   * determines whether or not the component should re-render based on changes in state
   * @param {any} nextProps
   * @param {any} nextState
   */
  public shouldComponentUpdate(nextProps: any, nextState: any): boolean {
    console.log('shouldcomponentupdate: team');
    let projectChanged: boolean =  this.props.project !== nextProps.project;
    let teamChanged: boolean =  this.props.team !== nextProps.team;
    let teamListChanged: boolean = JSON.stringify(this.props.teams) !== JSON.stringify(nextProps.teams);
    let populationChanged: boolean = this.props.populationStage !== nextProps.poulationStage;
    return projectChanged || teamChanged || teamListChanged || populationChanged;
  }

  public componentWillUpdate(nextProps: any, nextState: any): void {
    console.log('willcomponentupdate: team');
    if ((this.props.project !== nextProps.project && nextProps.project !== '') ||
        nextProps.populationStage === PopulationStage.projectReady) {
      this.populateTeams(nextProps.account, nextProps.project);
    }
  }
  /**
   * Reaction to selection of team option from dropdown list
   * @param {any} option
   * @returns {void}
   */
  public onTeamSelect(option: any): void {
    let team: string;
    if (option.text) {
      team = option.text;
    } else {
      team = option;
    }
    this.props.dispatch(updateTeamSettingsAction(team, this.props.teams));
  }

  /**
   * Renders the react-select dropdown component
   */
  public render(): React.ReactElement<Provider> {
    let teams: IDropdownOptions[] = [];
    let containsTeam: boolean = false;
    this.props.teams.forEach((option: IDropdownOptions) => {
      let isSelected: boolean = false;
      if (option.text === this.props.team) {
        containsTeam = true;
        isSelected = true;
      }
      teams.push({
        isSelected: isSelected,
        key: option.key,
        text: option.text,
      });
    });

    return (
      <Dropdown
        label={'Team'}
        options={teams}
        onChanged={this.onTeamSelect.bind(this)}
        disabled={this.props.populationStage < PopulationStage.teamReady}
      />);
  }

  /**
   * Populates list of teams for given project from VSTS rest call
   * Updates the store for current sesttings and current options lists
   * @param {string} account, {string} project
   * @returns {void}
   */
  public populateTeams(account: string, project: string): void {
    this.props.dispatch(updatePopulatingAction(PopulationStage.teamPopulating));
    let teamOptions: ISettingsInfo[] = [];
    let teamNamesOnly: string[] = [];
    let selectedTeam: string = this.props.team;

    Rest.getTeams(project, account, (error: RestError, teams: Team[]) => {
      if (error) {
        this.props.dispatch(updateNotificationAction(NotificationType.Error, error.toString('populate teams')));
        return;
      }
      teams = teams.sort(Team.compare);
      teams.forEach(team => {
        teamOptions.push({ key: team.name, text: team.name });
        teamNamesOnly.push(team.name);
      });
      console.log('teamList: ' + JSON.stringify(teams));
      let defaultTeam: string = Office.context.roamingSettings.get('default_team');
      if (defaultTeam !== undefined && defaultTeam !== '') {
        selectedTeam = defaultTeam;
        console.log('setting default project:' + defaultTeam);
      }
      if (selectedTeam === '' || (teamNamesOnly.indexOf(selectedTeam) === -1)) { // very first time user
        selectedTeam = teamNamesOnly[0];
        console.log('setting first project:' + selectedTeam);
      }
      try {
        this.props.dispatch(updateTeamSettingsAction(selectedTeam, teamOptions));
        this.props.dispatch(updatePopulatingAction(PopulationStage.teamReady));
      } catch (e) {
        // bug in fabricReact requires this
        this.props.dispatch(updatePopulatingAction(PopulationStage.teamReady));
        this.props.dispatch(updateTeamSettingsAction(selectedTeam, teamOptions));
      }
    });
  }
}
