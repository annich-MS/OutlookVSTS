/// <reference path="../../../office.d.ts" />
import * as React from 'react';
import { Provider, connect } from 'react-redux';
import {PageVisibility, updatePageAction} from '../../Redux/FlowActions';
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
}

/**
 * maps state in application store to properties for the component
 * @param {any} state
 */
function mapStateToProps(state: any): ISettingsProps {
  return({
      account: state.currentSettings.settings.account,
      project: state.currentSettings.settings.project,
      team: state.currentSettings.settings.team,
    });
}

@connect(mapStateToProps)

/**
 * Smart component
 * Renders area path dropdowns and save button
 * @class {Settings} 
 */
export class CancelButton extends React.Component<ISettingsProps, any> {

  /**
   * saves current selected settings to Office Roaming Settings
   * updates page state to Create Work Item page
   * @returns {void}
   */
  public Cancel(): void {
    this.props.dispatch(updatePageAction(PageVisibility.CreateItem));
  }

  /**
   * Renders the area path dropdowns and save button
   */
  public render(): React.ReactElement<Provider> {
    return (
       <div style={{float: 'right'}}>
          <Button buttonType={ButtonType.command} onClick={this.Cancel.bind(this)}>Cancel</Button>
      </div>
    );
  }
}
