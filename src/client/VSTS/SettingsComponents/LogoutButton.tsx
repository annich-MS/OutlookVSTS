import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { Rest, RestError } from '../../RestHelpers/rest';
import { AuthState, updateAuthAction, updateNotificationAction, NotificationType } from '../../Redux/FlowActions';
import { RoamingSettings } from '../RoamingSettings';
import { Button, ButtonType } from 'office-ui-fabric-react';

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

        return (
            <div style={{margin:'auto', width:'75%', textAlign:'center'}}>
                <Button buttonType={ButtonType.command} onClick={this.logout.bind(this)}>
                    Disconnect From VSTS
                </Button>
            </div>);
    }

    private logout(): void {
        let dispatch: any = this.props.dispatch;

        Rest.removeUser((error: RestError) => {
            if (error) {
                this.props.dispatch(updateNotificationAction(NotificationType.Error, error.toString('disconnect')));
                return;
            } else {
                RoamingSettings.GetInstance().clear();
                dispatch(updateAuthAction(AuthState.NotAuthorized));
            }
        });
    }
}
