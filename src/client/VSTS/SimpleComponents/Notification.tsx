import * as React from "react";
import { observer } from "mobx-react";
import { MessageBar, MessageBarType } from "office-ui-fabric-react";
import NavigationStore from "../../stores/navigationStore";
import { AppNotificationType } from "../../models/appNotification";

/**
 * Properties needed for the Error component
 * @interface IErrorProps
 */
export interface IErrorProps {
  navigationStore: NavigationStore;
}

/**
 * Smart component
 * Renders error response
 * @class {Notification} 
 */
@observer
export class Notification extends React.Component<IErrorProps, any> {

  /**
   * Renders the error message in parent component
   */
  public render(): JSX.Element {
    if (this.props.navigationStore.notification !== null) {
      return (<div>
        <MessageBar
          messageBarType={this.getMessageBarType(this.props.navigationStore.notification.type)}
          onDismiss={this.onClick.bind(this)}>
          {this.props.navigationStore.notification.message}
        </MessageBar>
      </div>);
    } else {
      return (<div />);
    }
  }

  private onClick(): void {
    this.props.navigationStore.clearNotification();
  }

  private getMessageBarType(type: AppNotificationType): MessageBarType {
    switch (type) {
      case AppNotificationType.Error:
        return MessageBarType.error;
      case AppNotificationType.Warning:
        return MessageBarType.warning;
      case AppNotificationType.Success:
        return MessageBarType.success;
      case AppNotificationType.Assert:
        return MessageBarType.severeWarning;
      default:
        return null;
    }
  }


}

