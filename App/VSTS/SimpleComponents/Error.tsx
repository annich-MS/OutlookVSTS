import * as React from 'react';
import { Provider, connect } from 'react-redux';

/**
 * Properties needed for the Error component
 * @interface IErrorProps
 */
export interface IErrorProps {
  dispatch?: any;
  isVisible?: boolean;
  message?: string;
}

/**
 * maps state in application store to properties for the component
 * @param {any} state
 */
function mapStateToProps(state: any): IErrorProps {
  return (
      {
      isVisible: state.controlState.error.isVisible,
      message: state.controlState.error.message,
    });
}

@connect(mapStateToProps)

/**
 * Smart component
 * Renders error response
 * @class {Error} 
 */
export class Error extends React.Component<IErrorProps, {}> {
  /**
   * Renders the error message in parent component
   */
  public render(): React.ReactElement<Provider> {
    if (this.props.isVisible === true) {
      console.log('error');

      return (<div color='rgb(255,0,0)'>
                <span className='ms-Icon ms-Icon--infoCircle'> </span>
                <span font-family='Arial Black, Gadget, sans-serif'> {this.props.message} </span>
              </div>);
    }else {
      return (<div/>);
    }
  }
}




