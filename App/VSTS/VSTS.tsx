/// <reference path="../../office.d.ts" />
import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { LogInPage } from './LoginComponents/LogInPage';
import { Settings} from './SettingsComponents/Settings';
import { Loading } from './SimpleComponents/Loading';
import { Connecting } from './SimpleComponents/Connecting';
import { Auth } from './authMM';
import { updateUserProfileAction} from '../Redux/LogInActions';
import { PageVisibility, AuthState, updateAuthAction, IErrorStateAction, updatePageAction } from '../Redux/FlowActions';
import { UserProfile } from '../RestHelpers/rest';
import { CreateWorkItem } from './CreateWorkItem';
import { QuickActions } from './QuickActions';
import { Rest } from '../RestHelpers/rest';

interface IRefreshCallback { (): void; }
interface IUserProfileCallback { (profile: UserProfile): void; }

/**
 * Properties needed for the main VSTS component
 * @interface IVSTSProps
 */
interface IVSTSProps {
  dispatch?: any;
  authState?: AuthState;
  pageState?: PageVisibility;
  error?: IErrorStateAction;
}

/**
 * maps state in application store to properties for the component
 * @param {any} state
 */
function mapStateToProps(state: any): IVSTSProps {
  // console.log('state:' + JSON.stringify(state));
  return ({
    authState: state.controlState.authState,
    error: state.controlState.error,
    pageState: state.controlState.pageState,
  });
}

@connect(mapStateToProps)

export class VSTS extends React.Component<IVSTSProps, any> {

  public constructor() {
    super();
    this.Initialize = this.Initialize.bind(this);
    Office.initialize = this.Initialize;
  }

  /**
   * determines whether or not the component should re-render based on changes in state
   * @param {any} nextProps
   * @param {any} nextState
   */
  public shouldComponentUpdate(nextProps: any, nextState: any): boolean {
    return (this.props.authState !== nextProps.authState) ||
      (this.props.pageState !== nextProps.pageState);
  }

  /**
   * Executed after Office.initialize is complete. 
   * Initial check for user authentication token and determines correct first page to show
   */
  public Initialize(): void {
    console.log('Initiating');
    // - TODO check for auth token
    let dispatch: any = this.props.dispatch;
    const email: string = Office.context.mailbox.userProfile.emailAddress;
    const name: string = Office.context.mailbox.userProfile.displayName;
    Auth.getAuthState(email, function (state: string): void {
      if (state === 'success') {
        let id: string = Office.context.roamingSettings.get('memberID');
        if (id) {
          dispatch(updateUserProfileAction(name, email, Office.context.roamingSettings.get('member_ID')));
          if (Office.context.roamingSettings.get('default_team') !== undefined) {
            dispatch(updatePageAction(PageVisibility.CreateItem));
          }
          dispatch(updateAuthAction(AuthState.Authorized));
        } else {
          Rest.getUserProfile(email, (profile: UserProfile) => {
            id = profile.id;
            Office.context.roamingSettings.set('member_ID', id);
            Office.context.roamingSettings.saveAsync();
            dispatch(updateUserProfileAction(name, email, id));
            dispatch(updateAuthAction(AuthState.Authorized));
          });

          if (Office.context.roamingSettings.get('default_team') !== undefined) {
            dispatch(updatePageAction(PageVisibility.CreateItem));
          }
        }
      } else {
        dispatch(updateAuthAction(AuthState.NotAuthorized));
        //this.cleanRoamingSettings();
      }
    });
  }

  public cleanRoamingSettings(): void {
      Office.context.roamingSettings.set('member_ID', undefined);
      Office.context.roamingSettings.set('default_team', undefined);
      Office.context.roamingSettings.set('default_project', undefined);
      Office.context.roamingSettings.set('default_account', undefined);
      Office.context.roamingSettings.saveAsync();
  }

  /**
   * Renders the add-in. Contains logic to determine which component/page to display
   */
  public render(): React.ReactElement<Provider> {
    let bodyStyle: any = {
      padding: '2.25%',
    };
    let body: any;
    switch (this.props.authState) {
      case AuthState.None:
        body = (<Loading />);
        break;
      case AuthState.NotAuthorized:
        body = (<LogInPage />);
        break;
      case AuthState.Request:
        body = (<Connecting/>);
        break;
      case AuthState.Authorized:
        {
          switch (this.props.pageState) {
            case PageVisibility.CreateItem:
              body = (<CreateWorkItem />);
              // body = (<QuickActions />);
              break;
            case PageVisibility.QuickActions:
              body = (<QuickActions />);
              break;
            case PageVisibility.Settings:
            default:
              body = (<Settings />);
              break;
          }
        }
        break;
      default:
       body = (<LogInPage />);
    }
    return(<div style={bodyStyle}> {body} </div>);
  }
}
