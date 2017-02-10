import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { SettingsInfo } from '../../Redux/LogInActions';
import { PageVisibility, updatePageAction, PopulationStage } from '../../Redux/FlowActions';
import { RoamingSettings } from '../RoamingSettings';
import { Button, ButtonType } from 'office-ui-fabric-react';

interface ISettingsProps {
  /**
   * intermediate to dispatch actions to update the global store
   * @type {any}
   */
  dispatch?: any;
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

  accounts?: SettingsInfo[];
  projects?: SettingsInfo[];
  teams?: SettingsInfo[];

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
function mapStateToProps(state: any): ISettingsProps {
  return ({
    account: state.currentSettings.settings.account,
    accounts: state.currentSettings.lists.accountList,
    populationStage: state.controlState.populationStage,
    project: state.currentSettings.settings.project,
    projects: state.currentSettings.lists.projectList,
    team: state.currentSettings.settings.team,
    teams: state.currentSettings.lists.teamList,
  });
}

@connect(mapStateToProps)

/**
 * Smart component
 * Renders area path dropdowns and save button
 * @class {Settings} 
 */
export class SaveDefaultsButton extends React.Component<ISettingsProps, any> {

  /**
   * saves current selected settings to Office Roaming Settings
   * updates page state to Create Work Item page
   * @returns {void}
   */
  public saveDefaults(): void {
    let rs: RoamingSettings = RoamingSettings.GetInstance();
    rs.account = this.props.account;
    rs.project = this.props.project;
    rs.team = this.props.team;
    rs.accounts = this.props.accounts;
    rs.projects = this.props.projects;
    rs.teams = this.props.teams;
    rs.save();

    this.props.dispatch(updatePageAction(PageVisibility.CreateItem));
  }

  /**
   * Renders the area path dropdowns and save button
   */
  public render(): React.ReactElement<Provider> {
    return (
      <div style={{ float: 'left' }}>
        <Button
          buttonType={ButtonType.command}
          icon='Save'
          onClick={this.saveDefaults.bind(this)}
          disabled={this.props.populationStage < PopulationStage.teamReady}>
          Save and continue
          </Button>
      </div>
    );
  }
}
