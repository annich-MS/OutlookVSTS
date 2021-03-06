import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { updateNotificationAction, NotificationType } from '../../Redux/FlowActions';
import { updateStage, Stage } from '../../Redux/WorkItemActions';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';

/**
 * Properties needed for the Error component
 * @interface IErrorProps
 */
export interface IErrorProps {
  dispatch?: any;
  type?: NotificationType;
  message?: string;
}

/**
 * maps state in application store to properties for the component
 * @param {any} state
 */
function mapStateToProps(state: any): IErrorProps {
  return (
      {
      message: state.controlState.notification.message,
      type: state.controlState.notification.notificationType,
    });
}

@connect(mapStateToProps)

/**
 * Smart component
 * Renders error response
 * @class {Notification} 
 */
export class Notification extends React.Component<IErrorProps, any> {

  /**
   * determines whether or not the component should re-render based on changes in state
   * @param {any} nextProps
   * @param {any} nextState
   */
  public shouldComponentUpdate(nextProps: any, nextState: any): boolean {
    return this.props.type !== nextProps.type || this.props.message !== nextProps.message;
  }

  /**
   * Renders the error message in parent component
   */
  public render(): React.ReactElement<Provider> {
    if (this.props.type !== NotificationType.Hide) {
      return (<div>
                <MessageBar
                    messageBarType={ this.getMessageBarType(this.props.type) }
                    onDismiss={this.onClick.bind(this)}>
                  {this.props.message}
                </MessageBar>
              </div>);
    } else {
      return (<div/>);
    }
  }

  private onClick(): void {
    this.props.dispatch(updateNotificationAction(NotificationType.Hide, ''));
  }

  private getMessageBarType(type: NotificationType): MessageBarType {
    switch (type) {
      case NotificationType.Error:
        return MessageBarType.error;
      case NotificationType.Warning:
        return MessageBarType.warning;
      case NotificationType.Success:
        return MessageBarType.success;
      default:
        return null;
    }
  }


}

