import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { Notification } from '../SimpleComponents/Notification';
import {AccountDropdown } from './AccountDropdown';
import {ProjectDropdown } from './ProjectDropdown';
import {AreaDropdown } from './AreaDropdown';
import {SaveDefaultsButton } from './SaveDefaultsButton';
import {CancelButton } from './CancelButton';
import { LogoutButton } from './LogoutButton';

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
    let style_text: string = 'ms-font-m-plus';

    return (
      <div>
        <Notification />
        <div>
          <p className={style_text}> Welcome {this.props.name}!</p>
          <p/>
          <p className={style_text}> Take a moment to configure your default settings for work item creation.</p>
        </div>
        <div>
          <AccountDropdown />
          <ProjectDropdown />
          <AreaDropdown />
        </div>
        <div>
          <SaveDefaultsButton/>
          <CancelButton/>
        </div>
        <br />
        <div>
          <LogoutButton/>
        </div>
      </div>
    );
  }
}
