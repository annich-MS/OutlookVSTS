import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { updatePageAction, PageVisibility } from '../Redux/FlowActions';
import { Button, ButtonType } from 'office-ui-fabric-react';

/**
 * Represents the Gear Properties
 * @interface IGearProps
 */
export interface IGearProps {
  /**
   * dispatch to map dispatch to props
   * @type {any}
   */
  dispatch?: any;
}

@connect()
/**
 * Renders the Gear Icon and the button underneath
 * @class { Gear }
 */
export class Gear extends React.Component<IGearProps, {}> {
  /**
   * Renders the Gear Icon and the button underneath
   */
  public render(): React.ReactElement<Provider> {
    return (
      <div style={{float: 'right'}}>
        <Button buttonType={ButtonType.icon} icon='Settings' title='Settings' onClick={this.handleGearClick}/>
      </div>
    );
  }

  /**
   * Dispatches the action to change the pageVisibility value in the store
   * @ returns {void}
   */
  private handleGearClick: () => void = () => {
    this.props.dispatch(updatePageAction(PageVisibility.Settings));
  }
}
