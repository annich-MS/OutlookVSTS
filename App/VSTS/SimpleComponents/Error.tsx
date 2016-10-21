import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { updateErrorAction } from '../../Redux/FlowActions';

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
export class Error extends React.Component<IErrorProps, any> {

  /**
   * determines whether or not the component should re-render based on changes in state
   * @param {any} nextProps
   * @param {any} nextState
  */
  public shouldComponentUpdate(nextProps: any, nextState: any): boolean {
    return this.props.isVisible !== nextProps.isVisible;
  }


  /**
   * Renders the error message in parent component
   */
  public render(): React.ReactElement<Provider> {
    let style: any = {
       background: 'white',
       border: 'white',
       color: 'red',
       float: 'left',
     };
    if (this.props.isVisible === true) {
      return (<div>
                <button className='ms-font-s' style={style} onClick={this.onClick.bind(this)}>
                  <span className='ms-Icon ms-Icon--Error'> </span>
                  <span font-family='Arial Black, Gadget, sans-serif'> {this.props.message} </span>
                </button>
                <br />
              </div>);
    }else {
      return (<div/>);
    }
  }

  private onClick(): void {
    this.props.dispatch(updateErrorAction(false, ''));
  }
}

