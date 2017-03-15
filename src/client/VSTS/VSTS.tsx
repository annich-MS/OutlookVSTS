import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { LogInPage } from './LoginComponents/LogInPage';
import { Settings } from './SettingsComponents/Settings';
import { Connecting } from './SimpleComponents/Connecting';
import { Saving } from './SimpleComponents/Saving';
import { Auth } from './authMM';
import {
  updateUserProfileAction, updateTeamSettingsAction, updateAccountSettingsAction, updateProjectSettingsAction, SettingsInfo
} from '../Redux/LogInActions';
import { Stage, updateAddAsAttachment, updateDescription } from '../Redux/WorkItemActions';
import {
  PageVisibility,
  AuthState,
  updateAuthAction,
  INotificationStateAction,
  updatePageAction,
  updateNotificationAction,
  NotificationType,
  updatePopulatingAction,
  PopulationStage,
} from '../Redux/FlowActions';
import { UserProfile } from '../RestHelpers/rest';
import { CreateWorkItem } from './CreateWorkItem';
import { QuickActions } from './QuickActions';
import { RoamingSettings } from './RoamingSettings';
import { Rest, RestError } from '../RestHelpers/rest';

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
  stage?: Stage;
  notification?: INotificationStateAction;
}

/**
 * maps state in application store to properties for the component
 * @param {any} state
 */
function mapStateToProps(state: any): IVSTSProps {
  return ({
    authState: state.controlState.authState,
    notification: state.controlState.notification,
    pageState: state.controlState.pageState,
    stage: state.workItem.stage,
  });
}

@connect(mapStateToProps)

export class VSTS extends React.Component<IVSTSProps, any> {

  private roamingSettings: RoamingSettings;

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
      (this.props.pageState !== nextProps.pageState) ||
      (this.props.stage !== nextProps.stage);
  }

  public iosInit(): void {
    if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
      this.props.dispatch(updateAddAsAttachment(false));
      (Office.context.mailbox.item as Office.MessageCompose).body.getAsync(Office.CoercionType.Text, (result: Office.AsyncResult) => {
        this.props.dispatch(updateDescription(result.value.trim()));
      });
    }
  }

  public authInit(): void {
    let dispatch: any = this.props.dispatch;
    let roamingSettings: RoamingSettings = this.roamingSettings;
    const email: string = Office.context.mailbox.userProfile.emailAddress;
    const name: string = Office.context.mailbox.userProfile.displayName;
    Auth.getAuthState(function (state: string): void {
      if (state === 'success') {
        if (roamingSettings.id) {
          dispatch(updateUserProfileAction(name, email, roamingSettings.id));
          dispatch(updateAuthAction(AuthState.Authorized));
          if (roamingSettings.isValid) {
            dispatch(updatePageAction(PageVisibility.CreateItem)); // todo - may cause issues here
          }
        } else {
          Rest.getUserProfile((error: RestError, profile: UserProfile) => {
            if (error) {
              dispatch(updateNotificationAction(NotificationType.Error, error.toString('retrieve user profile')));
              return;
            }
            roamingSettings.id = profile.id;
            roamingSettings.save();
            dispatch(updateUserProfileAction(name, email, profile.id));
            dispatch(updateAuthAction(AuthState.Authorized));
          });
          if (roamingSettings.isValid) {
            dispatch(updatePageAction(PageVisibility.CreateItem)); // todo - may cause issues here
          }
        }
      } else {
        dispatch(updateAuthAction(AuthState.NotAuthorized));
      }
    });
  }

  public prepopDropdowns(): void {
    this.props.dispatch(updatePopulatingAction(PopulationStage.prepopulate));
    if (this.roamingSettings.isValid) {
      console.log('prepopulating');
      this.props.dispatch(updateAccountSettingsAction(this.roamingSettings.account, this.roamingSettings.accounts));
      this.props.dispatch(updateProjectSettingsAction(this.roamingSettings.project, this.roamingSettings.projects));
      this.props.dispatch(updateTeamSettingsAction(this.roamingSettings.team, this.roamingSettings.teams));
    }
  }


  /**
   * Executed after Office.initialize is complete. 
   * Initial check for user authentication token and determines correct first page to show
   */
  public Initialize(): void {
    console.log('Initiating');
    this.roamingSettings = RoamingSettings.GetInstance();
    this.iosInit();
    this.prepopDropdowns();
    this.authInit();
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
      case AuthState.NotAuthorized:
        body = (<LogInPage />);
        break;
      case AuthState.None:
      case AuthState.Request:
        body = (<Connecting />);
        break;
      case AuthState.Authorized:
        {
          switch (this.props.pageState) {
            case PageVisibility.CreateItem:
              body = [<CreateWorkItem />];
              if (this.props.stage === Stage.Saved) {
                body.push(<Saving />);
              }
              break;
            case PageVisibility.QuickActions:
              body = (<QuickActions />);
              break;
            // case PageVisibility.Settings:
            default:
              body = (<Settings />);
              break;
          }
        }
        break;
      default:
        body = (<LogInPage />);
    }
    return (<div style={bodyStyle}> {body} </div>);
  }
}
