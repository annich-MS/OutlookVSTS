import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { updateAccountSettingsAction, SettingsInfo } from '../../Redux/LogInActions';
import { updateNotificationAction, updatePopulatingAction, PopulationStage, NotificationType } from '../../Redux/FlowActions';
import { Rest, RestError, Account } from '../../rest';
import { RoamingSettings } from '../RoamingSettings';
import { Dropdown } from 'office-ui-fabric-react';

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
   * @type {SettingsInfo[]}
   */
  accountList?: SettingsInfo[];

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

  public constructor() {
    super();
    this.populateAccounts = this.populateAccounts.bind(this);
  }

  /**
   * Populates accounts for user's profile before first time account drop down is rendered
   * @returns {void}
   */
  public componentWillMount(): void {
    if (this.props.account === '') {
      this.populateAccounts();
    } else {
      this.runPopulate((account: string, accounts: SettingsInfo[]) => {
        if (JSON.stringify(accounts) !== JSON.stringify(this.props.accountList)) {
          RoamingSettings.GetInstance().accounts = accounts;
          RoamingSettings.GetInstance().save();
        }
        this.props.dispatch(updateAccountSettingsAction(account, accounts));
      });
    }
  }

  /**
   * Reaction to selection of account option from dropdown list
   * Triggers repopulation of projects for selected account
   * @param {any} option
   * @returns {void}
   */
  public onAccountSelect(option: any): void {
    this.props.dispatch(updateNotificationAction(NotificationType.Hide, ''));
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
    let accounts: SettingsInfo[] = [];
    let containsAccount: boolean = false;
    this.props.accountList.forEach((option: SettingsInfo) => {
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
    this.runPopulate((account: string, accounts: SettingsInfo[]) => {
      try {
        this.props.dispatch(updateAccountSettingsAction(account, accounts));
        this.props.dispatch(updatePopulatingAction(PopulationStage.accountReady));
      } catch (e) {
        // bug in fabricReact requires this
        this.props.dispatch(updatePopulatingAction(PopulationStage.accountReady));
        this.props.dispatch(updateAccountSettingsAction(account, accounts));
      }
    });
  }

  private runPopulate(callback: Function): void {
    let accountOptions: SettingsInfo[] = [];
    let accountNamesOnly: string[] = [];
    let selectedAccount: string = this.props.account;
    Rest.getAccounts(this.props.memberId, (error: RestError, accountList: Account[]) => {
      if (error) {
        this.props.dispatch(updateNotificationAction(NotificationType.Error, error.toString('populate accounts')));
        return;
      }
      accountList = accountList.sort(Account.compare);
      accountList.forEach(acc => {
        accountOptions.push({ key: acc.name, text: acc.name });
        accountNamesOnly.push(acc.name);
      });
      // console.log('AccountList: ' + JSON.stringify(accountList));
      let defaultAccount: string = RoamingSettings.GetInstance().account;
      if (defaultAccount !== undefined && defaultAccount !== '') {
        selectedAccount = defaultAccount;
        console.log('setting default account:' + defaultAccount);
      }
      if (selectedAccount === '' || (accountNamesOnly.indexOf(selectedAccount) === -1)) { // very first time user
        selectedAccount = accountNamesOnly[0];
        console.log('setting first account:' + selectedAccount);
      }
      console.log('popaccounts' + defaultAccount);
      callback(selectedAccount, accountOptions);
    });
  }
}
