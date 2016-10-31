/// <reference path="../../../office.d.ts" />
import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { updateAccountSettingsAction, ISettingsInfo} from '../../Redux/LogInActions';
import { updateErrorAction, updatePopulatingAction, PopulationStage } from '../../Redux/FlowActions';
import {Rest, RestError, Account} from '../../RestHelpers/rest';
import { Dropdown, IDropdownOptions } from 'office-ui-fabric-react';

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
  populationStage?: PopulationStage;
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
    populationStage: state.controlState.populationStage,
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
    this.populateAccounts();
  }

  /**
   * Reaction to selection of account option from dropdown list
   * Triggers repopulation of projects for selected account
   * @param {any} option
   * @returns {void}
   */
  public onAccountSelect(option: any): void {
    let account: string;
    if (option.text) {
      account = option.text;
    } else {
      account = option;
    }
    this.props.dispatch(updateAccountSettingsAction(account, this.props.accountList));
  }

  /**
   * Renders the react-select dropdown component
   */
  public render(): React.ReactElement<Provider> {
    let accounts: IDropdownOptions[] = [];
    let containsAccount: boolean = false;
    this.props.accountList.forEach((option: IDropdownOptions) => {
      let isSelected: boolean = false;
      if (option.text === this.props.account) {
        isSelected = true;
        containsAccount = true;
      }
      accounts.push({
        isSelected: isSelected,
        key: option.key,
        text: option.text,
      });
    });

    return (
      <Dropdown
        label={'Account'}
        options={accounts}
        onChanged={this.onAccountSelect.bind(this)}
        disabled={this.props.populationStage < PopulationStage.accountReady}
      />);
  }

  /**
   * Populates list of accounts for given profile from VSTS rest call
   * Updates the store for current settings and current options lists
   * @returns {void}
   */
  public populateAccounts(): void {
    this.props.dispatch(updatePopulatingAction(PopulationStage.accountPopulating));
    let accountOptions: ISettingsInfo[] = [];
    let accountNamesOnly: string[] = [];
    let selectedAccount: string = this.props.account;
    console.log('populating accounts' + this.props.email + this.props.memberId);
    Rest.getAccounts(this.props.memberId, (error: RestError, accountList: Account[]) => {
      if (error) {
        this.props.dispatch(updateErrorAction(true, 'Accounts failed to populate due to ' + error.type));
        return;
      }
      accountList = accountList.sort(Account.compare);
      accountList.forEach(acc => {
        accountOptions.push({ key: acc.name, text: acc.name });
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
      try {
        this.props.dispatch(updateAccountSettingsAction(selectedAccount, accountOptions));
        this.props.dispatch(updatePopulatingAction(PopulationStage.accountReady));
      } catch (e) {
        // bug in fabricReact requires this
        this.props.dispatch(updatePopulatingAction(PopulationStage.accountReady));
        this.props.dispatch(updateAccountSettingsAction(selectedAccount, accountOptions));
      }
    });
  }
}
