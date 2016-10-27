/// <reference path="../../../office.d.ts" />
import * as React from 'react';
import { Provider, connect } from 'react-redux';
import {ISettingsInfo, updateProjectSettingsAction } from '../../Redux/LogInActions';
import {updatePopulatingAction, updateErrorAction } from '../../Redux/FlowActions';
import {Rest, RestError, Project } from '../../RestHelpers/rest';
import { Dropdown, IDropdownOptions } from 'office-ui-fabric-react';

/**
 * Properties needed for the ProjectDropdown component
 * @interface IProjectProps
 */
interface IProjectProps {
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
  project?: string;
  /**
   * Represents the lists of projects for current account
   * @type {ISettingsInfo[]}
   */
  projects?: ISettingsInfo[];

  /**
   * Represents what tier is currently being populated
   * @type {number}
   */
  populationTier?: number;
}

/**
 * maps state in application store to properties for the component
 * @param {any} state
 */
function mapStateToProps(state: any): IProjectProps {
  return ({
    account: state.currentSettings.settings.account,
    email: state.userProfile.email,
    id: state.userProfile.memberID,
    populationTier: state.controlState.populationTier,
    project: state.currentSettings.settings.project,
    projects: state.currentSettings.lists.projectList,
  });
}

@connect(mapStateToProps)

/**
 * Smart component
 * Renders project dropdown
 * @class {ProjectDropdown} 
 */
export class ProjectDropdown extends React.Component<IProjectProps, any> {

  private POPULATION_TIER: number = 2;

  public constructor() {
    super();
    this.populateProjects = this.populateProjects.bind(this);
  }

  /** 
   * executed first time component renders, reads the default project
   * @return {void}
   */
  public componentWillMount(): void {
    /*let defaultProject: string = Office.context.roamingSettings.get('default_project');
    if (defaultProject !== undefined) {
      this.props.dispatch(updateProjectSettingsAction(defaultProject, this.props.projects));
    }*/
  }

  /**
   * determines whether or not the component should re-render based on changes in state
   * @param {any} nextProps
   * @param {any} nextState
   */
  public shouldComponentUpdate(nextProps: any, nextState: any): boolean {
    console.log('shouldcomponentupdate project');
    let accountChanged: boolean = this.props.account !== nextProps.account;
    let projectChanged: boolean =  this.props.project !== nextProps.project;
    let projectListChanged: boolean = JSON.stringify(this.props.projects) !== JSON.stringify(nextProps.projects);
    let populationChanged: boolean = this.props.populationTier !== nextProps.poulationTier;
    return accountChanged || projectChanged || projectListChanged || populationChanged;
  }

  public componentWillUpdate(nextProps: any, nextState: any): void {
    console.log('willcomponentupdate project');
    if (this.props.account !== nextProps.account && nextProps.account !== '') {
      this.populateProjects(nextProps.account);
    }
  }

  /**
   * Reaction to selection of project option from dropdown list
   * @param {any} option
   * @returns {void}
   */
  public onProjectSelect(option: any): void {
    let project: string;
    if (option.text) {
      project = option.text;
    } else {
      project = option;
    }
    this.props.dispatch(updateProjectSettingsAction(project, this.props.projects));
  }

  /**
   * Renders the react-select dropdown component
   */
  public render(): React.ReactElement<Provider> {

    let projects: IDropdownOptions[] = [];
    let containsProject: boolean = false;
    this.props.projects.forEach((option: IDropdownOptions) => {
      let isSelected: boolean = false;
      if (option.text === this.props.project) {
        containsProject = true;
        isSelected = true;
      }
      projects.push({
        isSelected: isSelected,
        key: option.key,
        text: option.text,
      });
    });

    return (
      <Dropdown
        label={'Project'}
        options={projects}
        onChanged={this.onProjectSelect.bind(this)}
        disabled={this.props.populationTier >= this.POPULATION_TIER}
      />);
  }

  /**
   * Populates list of projects for given account from VSTS rest call
   * Updates the store for current settings and current options lists
   * @param {string} account
   * @returns {void}
   */
  public populateProjects(account: string): void {
    this.props.dispatch(updatePopulatingAction(true, this.POPULATION_TIER));
    let projectOptions: ISettingsInfo[] = [];
    let projectNamesOnly: string[] = [];
    let selectedProject: string = this.props.project;
    console.log('populating projects');

    Rest.getProjects(account, (error: RestError, projects: Project[]) => {
      if (error) {
        this.props.dispatch(updateErrorAction(true, 'Projects failed to populate due to ' + error.type));
        return;
      }
      projects = projects.sort(Project.compare);
      projects.forEach(project => {
        projectOptions.push({ key: project.name, text: project.name });
        projectNamesOnly.push(project.name);
      });
      // console.log('ProjectList: ' + JSON.stringify(projects));
      let defaultProject: string = Office.context.roamingSettings.get('default_project');
      if (defaultProject !== undefined && defaultProject !== '') {
        selectedProject = defaultProject;
        console.log('setting default project:' + defaultProject);
      }
      if (selectedProject === '' || (projectNamesOnly.indexOf(selectedProject) === -1)) { // very first time user
        selectedProject = projectNamesOnly[0];
        console.log('setting first project:' + selectedProject);
      }
      try {
        this.props.dispatch(updateProjectSettingsAction(selectedProject, projectOptions));
      } catch (e) {
        // bug in fabric react requires this
        this.props.dispatch(updatePopulatingAction(false, this.POPULATION_TIER));
        this.props.dispatch(updateProjectSettingsAction(selectedProject, projectOptions));
      }
    });
  }

}
