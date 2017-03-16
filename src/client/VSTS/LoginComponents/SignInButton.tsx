import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { AuthState, updateAuthAction, updateNotificationAction, NotificationType } from '../../Redux/FlowActions';
import { updateUserProfileAction } from '../../Redux/LogInActions';
import { Rest, RestError, UserProfile } from '../../RestHelpers/rest';
import { Auth } from '../authMM';
import { RoamingSettings } from '../RoamingSettings';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

/**
 * Properties needed for the SignInButton component
 * @interface ISignInProps
 */
interface ISignInProps {
  /**
   * intermediate to dispatch actions to update the global store
   * @type {any}
   */
  dispatch?: any;
  /**
   * interval for checking the database for user token
   * @type {number}
   */
  authState?: AuthState;
}

/**
 * maps state in application store to properties for the component
 * @param {any} state
 */
function mapStateToProps(state: any): ISignInProps {
  return ({
    authState: state.controlState.authState,
  });
}

@connect(mapStateToProps)

/**
 * Dumb component
 * Renders sign in button to connect to authentication flow
 * @class {SignInButton} 
 */
export class SignInButton extends React.Component<ISignInProps, {}> {

  private authInterval: any = '';
  /**
   * On-click response for sign in button
   * Opens browser window for user to authenticate
   * @returns {void}
   */
  public authOnClick(): void {
    Rest.getUser(function (user: string): void {
      Office.context.ui.displayDialogAsync(
        `https://${document.location.host}/authenticate?user=${user}`,
        { height: 50, width: 50 },
        (result: Office.AsyncResult) => {
          let dialog: Office.DialogHandler = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (message: Office.AsyncResult) => {
            this.props.dispatch(updateAuthAction(AuthState.Request));
            this.refreshAuth();
            dialog.close();
          });
        });
    });
  }

  public constructor() {
    super();
    this.refreshAuth = this.refreshAuth.bind(this);
  }

  /**
   * Calls function that determines user authentication state and updates authState if user token is present
   * Saves user's VSTS member id to Office Roaming Settings on success
   * @return {void}
   */
  public refreshAuth(): void {
    let authKey: any = this.authInterval;
    const name: string = Office.context.mailbox.userProfile.displayName;
    const email: string = Office.context.mailbox.userProfile.emailAddress;
    let dispatch: any = this.props.dispatch;
    Auth.getAuthState(function (state: string): void {
      if (state === 'success') {
        clearInterval(authKey);
        Rest.getUserProfile((error: RestError, profile: UserProfile) => {
          if (error) {
            this.props.dispatch(updateNotificationAction(NotificationType.Error, error.toString('get user profile')));
            return;
          }
          RoamingSettings.GetInstance().id = profile.id;
          RoamingSettings.GetInstance().save();
          dispatch(updateUserProfileAction(name, email, profile.id));
          dispatch(updateAuthAction(AuthState.Authorized));
        });
      }
    });
  }

  /**
   * Renders the sign in button
   */
  public render(): React.ReactElement<Provider> {

    let style_button: any = {
      textAlign: 'center',
    };

    return (
      <div style={style_button}>
        <PrimaryButton onClick={this.authOnClick.bind(this)}> Sign in to get started </PrimaryButton>
      </div>);
  }
}

