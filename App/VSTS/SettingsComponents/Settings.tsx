/// <reference path="../../../office.d.ts" />
import * as React from 'react';
import { Provider, connect } from 'react-redux';
import {Error } from '../SimpleComponents/Error';
import {AccountDropdown } from './AccountDropdown';
import {ProjectDropdown } from './ProjectDropdown';
import {AreaDropdown } from './AreaDropdown';
import {SaveDefaultsButton } from './SaveDefaultsButton';
import {CancelButton } from './CancelButton';

interface ISettingsProps {
  /**
   * intermediate to dispatch actions to update the global store
   * @type {any}
   */
  dispatch?: any;
  /**
   * user's display name
   * @type {string}
   */
  name?: string;
}

/**
 * maps state in application store to properties for the component
 * @param {any} state
 */
function mapStateToProps(state: any): ISettingsProps {
  return ({
    name: state.userProfile.displayName,
    });
}

@connect(mapStateToProps)

/**
 * Smart component
 * Renders area path dropdowns and save button
 * @class {Settings} 
 */
export class Settings extends React.Component<ISettingsProps, any> {
  /**
   * Renders the area path dropdowns and save button
   */
  public render(): React.ReactElement<Provider> {
    let style_text: any = {
      color: 'rgb(0,0,0)',
      font: '15px arial, ms-segoe-ui',
    };

    let style_label: any = {
      color: 'rgb(0,0,0)',
      font: '15px arial, ms-segoe-ui',
    };

    let rowStyle: any = {
      'margin-top': '25px',
      'margin-bottom': '25px',
    };

    return (
      <div>
        <Error />
        <div>
          <p style = {style_text}> Welcome {this.props.name}!</p>
          <p/>
          <p style = {style_text}> Take a moment to configure your default settings for work item creation.</p>
        </div>
        <div style={rowStyle}>
          <label style = {style_label}> Account </label>
          <AccountDropdown />
        </div>
        <div style={rowStyle}>
          <label style = {style_label}> Project </label>
          <ProjectDropdown />
        </div>
        <div style={rowStyle}>
          <label style = {style_label}> Area </label>
          <AreaDropdown />
        </div>
        <div>
          <SaveDefaultsButton/>
          <CancelButton/>
        </div>
      </div>
    );
  }
}
