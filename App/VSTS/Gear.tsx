import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { updatePageAction, PageVisibility } from '../Redux/FlowActions';

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
    let style_button: any = {
      background: 'rgb(255,255,255)',
      border: 'rgb(255,255,255)',
      color: 'rgb(0,0,0)',
      float: 'right',
    };
    return (
      <div>
        <div className='16px arial, ms-segoe-ui'> Create a work item
          <button className='ms-Button' style = {style_button} id='gear' onClick = {this.handleGearClick}>
            <span className='ms-Icon ms-Icon--gear'> </span>
          </button>
        </div>
      </div>
    );
  }

  /**
   * Dispatches the action to change the pageVisibility value in the store
   * @ returns {void}
   */
  private handleGearClick: () => void = () => {
    console.log('dispatch action here to change visibility enum');
    this.props.dispatch(updatePageAction(PageVisibility.Settings));
  }
}
