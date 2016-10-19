/// <reference path="../../../office.d.ts" />
import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { updateAccountSettingsAction, ISettingsInfo} from '../../Redux/LogInActions';
import {updatePopulatingAction } from '../../Redux/FlowActions';
import {Rest, Account} from '../../RestHelpers/rest';
require('react-select/dist/react-select.css');
let Select: any = require('react-select');

/**
 * Properties needed for the AccountDropdown component
 * @interface IAccountProps
 */
interface IAccountProps {
  /**
   * intermediate to dispatch actions to update the global store
   * @type {any}
   */
  dispatch?: any;
  /**
   * user's email address
   * @type {string}
   */
  email?: string;
  /**
   * user's VSTS member id
   * @type {string}
   */
  memberId?: string;
  /**
   * currently selected account option
   * @type {string}
   */
  account?: string;
  /**
   * list of accounts associated with user's VSTS profile
   * @type {ISettingsInfo[]}
   */
  accountList?: ISettingsInfo[];

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
function mapStateToProps(state: any): IAccountProps {
  return ({
    account: state.currentSettings.settings.account,
    accountList: state.currentSettings.lists.accountList,
    email: state.userProfile.email,
    memberId: state.userProfile.memberID,
    populationTier: state.controlState.populationTier,
  });
}

@connect(mapStateToProps)
/**
 * Smart component
 * Renders account dropdowns
 * @class {AccountDropdown} 
 */
export class AccountDropdown extends React.Component<IAccountProps, any> {

  private POPULATION_TIER: number = 3;

  public constructor() {
    super();
    this.populateAccounts = this.populateAccounts.bind(this);
  }

  /**
   * Populates accounts for user's profile before first time account drop down is rendered
   * @returns {void}
   */
  public componentWillMount(): void {
    // let defaultAccount: string = Office.context.roamingSettings.get('default_account');
    // console.log('comp will mount'+defaultAccount);
    // if (defaultAccount !== undefined) {
    //   this.props.dispatch(updateAccountSettingsAction(defaultAccount, this.props.accountList));
    // }
    console.log(this.props.populationTier);
    this.populateAccounts();
  }

  /**
   * Reaction to selection of account option from dropdown list
   * Triggers repopulation of projects for selected account
   * @param {any} option
   * @returns {void}
   */
  public onAccountSelect(option: any): void {
    console.log('AccountList: ' + JSON.stringify(option));
    let account: string;
    if (option.label) {
      account = option.label;
    } else {
      account = option;
    }
    console.log('onAccountSelect' + account);
    this.props.dispatch(updateAccountSettingsAction(account, this.props.accountList));
  }

  /**
   * Renders the react-select dropdown component
   */
  public render(): React.ReactElement<Provider> {
    let renderableName: string = this.props.account;
    if (renderableName.length > 25) {
      renderableName = renderableName.slice(0, 20) + '...';
    }
    return (
      <Select
        name='form-field-name'
        options={this.props.accountList}
        value={renderableName}
        onChange={this.onAccountSelect.bind(this) }
        autoBlur={true}
        searchable={false}
        />
    );
  }

  /**
   * Populates list of accounts for given profile from VSTS rest call
   * Updates the store for current settings and current options lists
   * @returns {void}
   */
  public populateAccounts(): void {
    this.props.dispatch(updatePopulatingAction(true, this.POPULATION_TIER));
    let accountOptions: ISettingsInfo[] = [];
    let accountNamesOnly: string[] = [];
    let selectedAccount: string = this.props.account;
    console.log('populating accounts' + this.props.email + this.props.memberId);
    Rest.getAccounts(this.props.memberId, (accountList: Account[]) => {
      accountList = accountList.sort(Account.compare);
      accountList.forEach(acc => {
        accountOptions.push({ label: acc.name, value: acc.name });
        accountNamesOnly.push(acc.name);
      });
      // console.log('AccountList: ' + JSON.stringify(accountList));
      let defaultAccount: string = Office.context.roamingSettings.get('default_account');
      if (defaultAccount !== undefined && defaultAccount !== '') {
        selectedAccount = defaultAccount;
        console.log('setting default account:' + defaultAccount);
      }
      if (selectedAccount === '' || (accountNamesOnly.indexOf(selectedAccount) === -1)) { // very first time user
        selectedAccount = accountNamesOnly[0];
        console.log('setting first account:' + selectedAccount);
      }
      console.log('popaccounts' + defaultAccount);
      this.props.dispatch(updateAccountSettingsAction(selectedAccount, accountOptions));
      this.props.dispatch(updatePopulatingAction(false, this.POPULATION_TIER));
    });
  }
}
