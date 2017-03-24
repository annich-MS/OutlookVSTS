// libraries
import * as React from "react";
import { observer } from "mobx-react";

// components
import { LogInPage } from "./LoginComponents/LogInPage";
import { Settings } from "./SettingsComponents/Settings";
import { Connecting } from "./SimpleComponents/Connecting";
import { Saving } from "./SimpleComponents/Saving";
import { CreateWorkItem } from "./CreateWorkItem";
import { QuickActions } from "./QuickActions";

// utilities
import { Rest, RestError, UserProfile } from "../rest";

// models
import { RoamingSettings } from "./RoamingSettings";
import APTPopulationStage from "../models/aptPopulateStage";
import { AppNotificationType } from "../models/appNotification";
import Constants from "../models/constants";
import NavigationPage from "../models/navigationPage";

// stores
import APTCache from "../stores/aptCache";
import NavigationStore from "../stores/navigationStore";
import WorkItemStore from "../stores/workItemStore";


interface IRefreshCallback { (): void; }
interface IUserProfileCallback { (profile: UserProfile): void; }

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

  public authInit(): void {
    let roamingSettings: RoamingSettings = this.roamingSettings;
    Rest.getIsAuthenticated()
      .then((isAuthenticated: boolean) => {
        if (isAuthenticated) {
          // TODO: manage roamingSettings
          if (roamingSettings.isValid) {
            this.props.navigationStore.navigate(NavigationPage.CreateWorkItem);
          } else {
            Rest.getUserProfile((error: RestError, profile: UserProfile) => {
              if (error) {
                this.props.navigationStore.updateNotification({ message: error.toString("retrieve user profile"), type: AppNotificationType.Error });
                return;
              }
              roamingSettings.id = profile.id;
              roamingSettings.save();
              this.props.navigationStore.navigate(NavigationPage.Settings);
            });
          }
        } else {
          this.props.navigationStore.navigate(NavigationPage.LogIn);
        }
      }).catch((error) => {
        console.log(`ASSERT: getIsAuthenticated rejected promise in AuthInit ${error}`);
      });
  }

  public prepopDropdowns(): void {
    this.props.aptCache.setPopulateStage(APTPopulationStage.PrePopulate);
    if (this.roamingSettings.isValid) {
      console.log("prepopulating");
      this.props.aptCache.setAccounts(this.roamingSettings.accounts, this.roamingSettings.account);
      this.props.aptCache.setProjects(this.roamingSettings.projects, this.roamingSettings.project);
      this.props.aptCache.setTeams(this.roamingSettings.teams, this.roamingSettings.team);
      this.props.aptCache.setPopulateStage(APTPopulationStage.PostPopulate);
    }
  }

  /**
   * Executed after Office.initialize is complete. 
   * Initial check for user authentication token and determines correct first page to show
   */
  public Initialize(): void {
    console.log("Initiating");
    this.roamingSettings = RoamingSettings.GetInstance();
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
