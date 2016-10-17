import * as React from 'react';
import { Provider } from 'react-redux';
import { AccountDropdown } from './SettingsComponents/AccountDropdown';
import { AreaDropdown } from './SettingsComponents/AreaDropdown';
import { ProjectDropdown } from './SettingsComponents/ProjectDropdown';

/**
 * Renders the Acccount, Project, and Area components
 * @class {Classification}
 */
export class Classification extends React.Component<{}, {}> {
  /**
   * Renders the Account, Project, and Area components
   */
  public render(): React.ReactElement<Provider> {
    let titleColumnStyle: any = {
      display: 'inline-block',
      font: '16px arial, ms-segoe-ui',
      width: '21%',
      'margin-bottom': '8px',
      'margin-top': '8px',
    };
    let dataColumnStyle: any = {
      display: 'inline-block',
      position: 'absolute',
      width: '74.25%',
    };
    let rowColumnStyle: any = {
      margin: '7px',
    };
    return (
      <div>
        <div>
          <div style={rowColumnStyle}>
            <div style={titleColumnStyle}>Account </div> <div style={dataColumnStyle}>  <AccountDropdown /> </div>
          </div>
          <div style={rowColumnStyle}>
            <div style={titleColumnStyle}>Project </div> <div style={dataColumnStyle}> <ProjectDropdown /> </div>
          </div>
          <div style={rowColumnStyle}>
            <div style={titleColumnStyle}>Team </div> <div style={dataColumnStyle}>  <AreaDropdown /> </div>
          </div>
        </div>
      </div>
    );
  }
}

