import * as React from "react";
import { Rest, RestError } from "../../utils/rest";
import RoamingSettings from "../../models/roamingSettings";
import { Button, ButtonType } from "office-ui-fabric-react";
import NavigationPage from "../../models/navigationPage";
import NavigationStore from "../../stores/navigationStore";
import { AppNotificationType } from "../../models/appNotification";
import APTCache from "../../stores/aptCache";

interface ILogoutButtonProps {
    aptCache: APTCache;
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

    private async logout(): Promise<void> {
        try {
            await Rest.removeUser();
            let rs: RoamingSettings = await RoamingSettings.GetInstance();
            rs.clear();
            this.props.aptCache.clear();
            this.props.navigationStore.navigate(NavigationPage.LogIn);
        } catch (error) {
            let message: string;
            if (error instanceof RestError) {
                message = error.toString("disconnect user");
            } else {
                message = error.message;
            }
            this.props.navigationStore.updateNotification({ message: message, type: AppNotificationType.Error });
        }
    }
}
