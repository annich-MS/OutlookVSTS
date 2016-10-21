import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { AuthState, updateAuthAction, updateErrorAction } from '../../Redux/FlowActions';
import { updateUserProfileAction} from '../../Redux/LogInActions';
import {Rest, RestError, UserProfile} from '../../RestHelpers/rest';
import { Auth} from '../authMM';

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
export class SignInButton extends React.Component<ISignInProps,  {}> {

  private authInterval: any = '';
  /**
   * On-click response for sign in button
   * Opens browser window for user to authenticate
   * @returns {void}
   */
  public authOnClick(): void {
    Rest.getUser( (user: string) => {
      window.open('./authenticate?user=' + user);
      this.props.dispatch(updateAuthAction(AuthState.Request));
      let pollInterval: number = 3000;
      this.authInterval = setInterval(this.refreshAuth, pollInterval);
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
        let id: string = '';
        Rest.getUserProfile((error: RestError, profile: UserProfile) => {
          if (error) {
            this.props.dispatch(updateErrorAction(true, 'Failed to retrieve User Profile due to ' + error.type));
            return;
          }
          id = profile.id;
          Office.context.roamingSettings.set('member_ID', '' + id);
          Office.context.roamingSettings.saveAsync();
          dispatch(updateUserProfileAction(name, email, id));
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
      background: 'rgb(0,122,204)',
      border: 'rgb(255,255,255)',
      color: 'rgb(255,255,255)',
      float: 'left',
      font: '17px arial, ms-segoe-ui',
      'margin-left': '20%',
    };
    let buttonClasses: string = 'ms-Button';

    return(
      <div>
      <button className={buttonClasses} onClick = {this.authOnClick.bind(this)}> Sign in to get started </button>
      </div>);
  }
  }

