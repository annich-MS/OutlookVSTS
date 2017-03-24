import * as React from "react";
import { Rest, RestError } from "../../rest";
import { RoamingSettings } from "../RoamingSettings";
import { Button, ButtonType } from "office-ui-fabric-react";
import NavigationPage from "../../models/navigationPage";
import NavigationStore from "../../stores/navigationStore";
import { AppNotificationType } from "../../models/appNotification";

interface ILogoutButtonProps {
    navigationStore: NavigationStore;
}

export class LogoutButton extends React.Component<ILogoutButtonProps, any> {

    public render(): JSX.Element {

        return (
            <div style={{ margin: "auto", textAlign: "center", width: "75%" }}>
                <Button buttonType={ButtonType.command} onClick={this.logout.bind(this)}>
                    Disconnect From VSTS
                </Button>
            </div>);
    }

    private logout(): void {

        Rest.removeUser((error: RestError) => {
            if (error) {
                this.props.navigationStore.updateNotification({message: error.toString("disconnect user"), type: AppNotificationType.Error});
                return;
            } else {
                RoamingSettings.GetInstance().clear();
                this.props.navigationStore.navigate(NavigationPage.LogIn);
            }
        });
    }
}
