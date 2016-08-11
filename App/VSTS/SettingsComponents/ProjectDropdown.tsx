/// <reference path="../../../office.d.ts" />
import * as React from 'react';
import { Provider, connect } from 'react-redux';
import {ISettingsInfo, updateProjectSettingsAction } from '../../Redux/LoginActions';
import {Rest, Project } from '../../RestHelpers/rest';

// other import statements don't work properly
require('react-select/dist/react-select.css');
let Select: any = require('react-select');

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

  public constructor() {
    super();
    this.populateProjects = this.populateProjects.bind(this);
  }

  /** 
   * executed first time component renders, reads the default project
   * @return {void}
   */
  public componentWillMount(): void {
    console.log('willcomponentmount');
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
    return this.props.account !== nextProps.account || this.props.project !== nextProps.project ||
      JSON.stringify(this.props.projects) !== JSON.stringify(nextProps.projects); // this.props.projects !== nextProps.projects;
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
    if (option.label) {
      project = option.label;
    } else {
      project = option;
    }
    this.props.dispatch(updateProjectSettingsAction(project, this.props.projects));
  }

  /**
   * Renders the react-select dropdown component
   */
  public render(): React.ReactElement<Provider> {
    return (
      <Select
        name='form-field-name'
        options={this.props.projects}
        value={this.props.project}
        onChange={this.onProjectSelect.bind(this) }
        />
    );
  }

  /**
   * Populates list of projects for given account from VSTS rest call
   * Updates the store for current settings and current options lists
   * @param {string} account
   * @returns {void}
   */
  public populateProjects(account: string): void {
    let projectOptions: ISettingsInfo[] = [];
    let projectNamesOnly: string[] = [];
    let selectedProject: string = this.props.project;
    console.log('populating projects');

    Rest.getProjects(this.props.email, account, (projects: Project[]) => {
      projects.forEach(project => {
        projectOptions.push({ label: project.name, value: project.name });
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
      this.props.dispatch(updateProjectSettingsAction(selectedProject, projectOptions));
    });
  }

}
