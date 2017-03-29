// libs
import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
// utils
import { Rest, RestError } from "../../utils/rest";
// models
import RoamingSettings from "../../models/roamingSettings";
import { AppNotificationType } from "../../models/appNotification";
import NavigationPage from "../../models/navigationPage";
// stores
import NavigationStore from "../../stores/navigationStore";

/**
 * Properties needed for the SignInButton component
 */
interface ISignInProps {
  /**
   * The navigation store, in order to redirect on successful authentication
   */
  navigationStore: NavigationStore;
}

/**
 * Renders sign in button to connect to authentication flow
 */
export class SignInButton extends React.Component<ISignInProps, {}> {

  /**
   * constructor for signInButton
   */
  public constructor() {
    super();
    this.refreshAuth = this.refreshAuth.bind(this);
  }

  /**
   * On-click response for sign in button
   * Opens popout window for user to authenticate
   */
  public async authOnClick(): Promise<void> {
    let user: string = await Rest.getUser();
    Office.context.ui.displayDialogAsync(
      `https://${document.location.host}/authenticate?user=${user}`,
      { height: 50, width: 50 },
      (result: Office.AsyncResult) => {
        let dialog: Office.DialogHandler = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (message: Office.AsyncResult) => {
          this.refreshAuth();
          dialog.close();
        });
      });
  }


  /**
   * Calls function that determines user authentication state and updates authState if user token is present
   * Saves user's VSTS member id to Office Roaming Settings on success
   */
  public async refreshAuth(): Promise<void> {
    try {
      this.props.navigationStore.navigate(NavigationPage.Connecting);
      let authenticated: boolean = await Rest.getIsAuthenticated();
      if (authenticated) {
        let userProfile = await Rest.getUserProfile();
        let roamingSettings: RoamingSettings = await RoamingSettings.GetInstance();
        roamingSettings.id = userProfile.id;
        await roamingSettings.save();
        this.props.navigationStore.navigate(NavigationPage.Settings);
      } else {
        this.props.navigationStore.updateNotification({ message: "Did not find auth info, please reauthenticate", type: AppNotificationType.Warning });
        this.props.navigationStore.navigate(NavigationPage.LogIn);
      }

    } catch (error) {
      let message: string = "";
      if (error instanceof RestError) {
        message = error.toString("get user profile");
      } else if (typeof error === "string") {
        message = error;
      } else {
        message = (error as Office.Error).message;
      }
      this.props.navigationStore.updateNotification({ message: message, type: AppNotificationType.Error });
    }
    return;
  }

  /**
   * Renders the sign in button
   */
  public render(): JSX.Element {

    let buttonStyle: any = {
      textAlign: "center",
    };

    return (
      <div style={buttonStyle}>
        <Button buttonType={ButtonType.primary} onClick={this.authOnClick.bind(this)}> Sign in to get started </Button>
      </div>);
  }
}

