import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { Rest, RestError } from '../../RestHelpers/rest';
import { AuthState, updateAuthAction, updateErrorAction } from '../../Redux/FlowActions';

/**
 * Properties needed for the LogoutButton component
 * @interface IAreaProps
 */
interface ILogoutProps {
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
}

/**
 * maps state in application store to properties for the component
 * @param {any} state
 */
function mapStateToProps(state: any): ILogoutProps {
  return ({
    email: state.userProfile.email,
  });
}

@connect(mapStateToProps)

export class LogoutButton extends React.Component<ILogoutProps, any> {

    public render(): React.ReactElement<Provider> {
        let style: any = {
            background: 'rgb(255,255,255)',
            border: 'rgb(255,255,255)',
            color: 'rgb(0,122,204)',
            font: '10px arial, ms-segoe-ui',
            'text-align': 'center',
        };

        return (
            <div style={{margin:'auto', width:'75%', 'text-align':'center'}}>
                <button style={style} onClick={this.logout.bind(this)}>
                    <span font-family='Arial Black, Gadget, sans-serif' > Disconnect From VSTS </span>
                </button>
            </div>);
    }

    private logout(): void {
        let dispatch: any = this.props.dispatch;

        Rest.removeUser((error: RestError) => {
            if (error) {
                this.props.dispatch(updateErrorAction(true, 'Failed to disconnect due to ' + error.type));
                return;
            } else {
                dispatch(updateAuthAction(AuthState.NotAuthorized));
            }
        });
    }
}
