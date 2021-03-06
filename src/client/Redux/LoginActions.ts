import { IDropdownOption } from 'office-ui-fabric-react';

/**
 * Represents the data for area path information, duplicated values for display in dropdown
 * Aliased to IDropdownOption for conversion to fabric-react
 * TODO: Remove alias
 * @interface SettingsInfo
 */
export class SettingsInfo implements IDropdownOption {
  public key: string | number;
  public text: string;
  public isSelected?: boolean;

  public static convertStringArray(array: string[]): SettingsInfo[] {
    let ret: SettingsInfo[] = [];
    array.forEach((element: string) => {
      ret.push({ key: element, text: element });
    });
    return ret;
  }

}

/**
 * Represents the currently selected area path
 * @interface ISettingsAction
 */
export interface IAccountSettingsAction {
  /**
   * the name of the currently selected account
   * @type {string}
   */
  account: string;
  /**
   * list of accounts for user's profile
   * @type {SettingsInfo[]}
   */
  accountList: SettingsInfo[];
  /**
   * the type of the action
   * @type {string}
   */
  type: 'ACCOUNT_SETTINGS';
}

/**
 * action to update the area path and lists in the state
 * @param {string} accountNew
 * @param {SettingsInfo[]} accounts
 * @returns {IAccountSettingsAction}
 */
export function updateAccountSettingsAction(accountNew: string, accounts: SettingsInfo[]): IAccountSettingsAction {
  return {
    account: accountNew,
    accountList: accounts,
    type: 'ACCOUNT_SETTINGS',
  };
}

/**
 * Represents the currently selected area path
 * @interface ISettingsAction
 */
export interface IProjectSettingsAction {
  /**
   * the name of the currently selected project
   * @type {string}
   */
  project: string;
  /**
   * list of projects for currently selected account
   * @type {SettingsInfo[]}
   */
  projectList: SettingsInfo[];
  /**
   * the type of the action
   * @type {string}
   */
  type: 'PROJECT_SETTINGS';
}

/**
 * action to update the area path and lists in the state
 * @param {string} projectNew
 * @param {SettingsInfo[]} projects
 * @returns {IProjectSettingsAction}
 */
export function updateProjectSettingsAction(projectNew: string, projects: SettingsInfo[]): IProjectSettingsAction {
  return {
    project: projectNew,
    projectList: projects,
    type: 'PROJECT_SETTINGS',
  };
}

/**
 * Represents the currently selected area path
 * @interface ISettingsAction
 */
export interface ITeamSettingsAction {
  /**
   * the name of the currently selected team
   * @type {string}
   */
  team: string;
  /**
   * list of teams for currently selected project
   * @type {SettingsInfo[]}
   */
  teamList: SettingsInfo[];
  /**
   * the type of the action
   * @type {string}
   */
  type: 'TEAM_SETTINGS';
}

/**
 * action to update the area path and lists in the state
 * @param {string} teamNew
 * @param {SettingsInfo[]} teams
 * @returns {ITeamSettingsAction}
 */
export function updateTeamSettingsAction(teamNew: string, teams: SettingsInfo[]): ITeamSettingsAction {
  return {
    team: teamNew,
    teamList: teams,
    type: 'TEAM_SETTINGS',
  };
}

/**
 * Represents the user's information in the state
 * @interface IUserProfileAction
 */
export interface IUserProfileAction {
  /**
   * the type of the action
   * @type {string}
   */
  type: string;
  /**
   * the user's display name
   * @type {string}
   */
  displayName: string;
  /**
   * the user's email address
   * @type {string}
   */
  email: string;
  /**
   * the user's VSTS member id
   * @type {string}
   */
  memberID: string;
}

/**
 * action to update name, email, and member ID for user in the state
 * @param {string} name
 * @param {string} mail
 * @param {string} id
 * @returns {IUserProfileAction}
 */
export function updateUserProfileAction(name: string, mail: string, id: string): IUserProfileAction {
  return {
    displayName: name,
    email: mail,
    memberID: id,
    type: 'USER_PROFILE',
  };
}















