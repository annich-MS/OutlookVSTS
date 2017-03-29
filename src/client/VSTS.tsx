// libraries
import * as React from "react";
import { observer } from "mobx-react";

// components
import LogInPage from "./components/logInPage";
import Settings from "./components/settingsPage";
import Connecting from "./components/connectingPage";
import Saving from "./components/savingOverlay";
import CreateWorkItem from "./components/createWorkItemPage";
import QuickActions from "./components/quickActionsPage";

// utilities
import { Rest, RestError, UserProfile } from "./utils/rest";

// models
import RoamingSettings from "./models/roamingSettings";
import APTPopulationStage from "./models/aptPopulateStage";
import { AppNotificationType } from "./models/appNotification";
import Constants from "./models/constants";
import NavigationPage from "./models/navigationPage";

// stores
import APTCache from "./stores/aptCache";
import NavigationStore from "./stores/navigationStore";
import WorkItemStore from "./stores/workItemStore";


/**
 * Properties needed for the main VSTS component
 * @interface IVSTSProps
 */
interface IVSTSProps {
  aptCache: APTCache;
  navigationStore: NavigationStore;
  workItemStore: WorkItemStore;
}

@observer
export class VSTS extends React.Component<IVSTSProps, any> {

  private roamingSettings: RoamingSettings;
  private item: Office.MessageRead;

  public constructor() {
    super();
    Office.initialize = this.Initialize.bind(this);
  }

  public iosInit(): void {
    if (Office.context.mailbox.diagnostics.hostName === Constants.IOS_HOST_NAME) {
      this.props.workItemStore.toggleAttachEmail();
      this.item.body.getAsync(Office.CoercionType.Text, (result: Office.AsyncResult) => {
        this.props.workItemStore.setDescription = result.value;
      });
    }
  }

  public async authInit(): Promise<void> {
    try {

      let roamingSettings: RoamingSettings = this.roamingSettings;
      let isAuthenticated: boolean = await Rest.getIsAuthenticated();
      if (isAuthenticated) {
        // TODO: manage roamingSettings
        if (roamingSettings.isValid) {
          this.props.navigationStore.navigate(NavigationPage.CreateWorkItem);
        } else {
          try {
            let profile: UserProfile = await Rest.getUserProfile();
            roamingSettings.id = profile.id;
            roamingSettings.save();
          } catch (error) {
            let message: string;
            if (error instanceof RestError) {
              message = error.toString("retrieve user profile");
            } else {
              message = (error as Office.Error).message;
            }
            this.props.navigationStore.updateNotification({ message: message, type: AppNotificationType.Error });
          }
        }
      } else {
        this.props.navigationStore.navigate(NavigationPage.LogIn);
      }
    } catch (error) {
      console.log(`ASSERT: getIsAuthenticated rejected promise in AuthInit ${error}`);
    }
    return;
  }

  public prepopDropdowns(): void {
    if (this.roamingSettings.isValid) {
      this.props.aptCache.setPopulateStage(APTPopulationStage.PrePopulate);
      this.props.aptCache.populate(true).then(() => {
        this.props.aptCache.populate(false);
      });
    }
  }

  /**
   * Executed after Office.initialize is complete. 
   * Initial check for user authentication token and determines correct first page to show
   */
  public async Initialize(): Promise<void> {
    this.roamingSettings = await RoamingSettings.GetInstance();
    this.item = Office.context.mailbox.item as Office.MessageRead;
    this.iosInit();
    this.prepopDropdowns();
    this.workItemInit();
    this.authInit();
  }

  /**
   * Renders the add-in. Contains logic to determine which component/page to display
   */
  public render(): JSX.Element {
    let body: JSX.Element = (<div />);
    let saving: JSX.Element = (<div />);
    switch (this.props.navigationStore.currentPage) {
      case NavigationPage.Connecting:
        body = <Connecting />;
        break;
      case NavigationPage.LogIn:
        body = <LogInPage navigationStore={this.props.navigationStore} />;
        break;
      case NavigationPage.CreateWorkItem:
        body = <CreateWorkItem cache={this.props.aptCache} navigationStore={this.props.navigationStore} workItem={this.props.workItemStore} />;
        break;
      case NavigationPage.Settings:
        body = <Settings cache={this.props.aptCache} navigationStore={this.props.navigationStore} />;
        break;
      case NavigationPage.QuickActions:
        body = <QuickActions navigationStore={this.props.navigationStore} workItem={this.props.workItemStore} />;
        break;
      default:
        body = <div>Invalid navigationPage</div>;
    }
    if (this.props.navigationStore.isSaving) {
      saving = <Saving />;
    }
    return (<div> {body}{saving} </div>);
  }

  private workItemInit(): void {
    this.props.workItemStore.setTitle(this.item.normalizedSubject);
  }
}
