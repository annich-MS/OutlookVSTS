import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { Rest, WorkItemInfo } from '../RestHelpers/rest';
import { updateStage, Stage, updateSave } from '../Redux/WorkItemActions';
import { IWorkItem } from '../Redux/WorkItemReducer';
import { updatePageAction, PageVisibility } from '../Redux/FlowActions';
import { IUserProfileReducer, ISettingsAndListsReducer } from '../Redux/LogInReducer';

/**
 * Represents the Save Properties
 * @interface ISaveProps
 */
export interface ISaveProps {
  /**
   * dispatch to map dispatch to props
   * @type {any}
   */
  dispatch?: any;
  /**
   * the entire work item property
   * @type {IWorkItem}
   */
  workItem?: IWorkItem;
  /**
   * the user profile information
   * @type {IUserProfile}
   */
  userProfile?: IUserProfileReducer;
  /**
   * the current settings information
   * @type {ISettingsAndListsReducer}
   */
  currentSettings?: ISettingsAndListsReducer;
}

/**
 * Renders the Save button and makes REST api calls
 * @class { Save }
 */
function mapStateToProps(state: any): ISaveProps {
  return {
    currentSettings: state.currentSettings,
    userProfile: state.userProfile,
    workItem: state.workItem,
  };
}

@connect(mapStateToProps)
export class Save extends React.Component<ISaveProps, {}> {
  /**
   * Dispatches the action to change the Stage and make the REST call to create the work item
   * @returns {void}
   */
  public handleSave(): void {
    this.props.dispatch(updateStage(Stage.Saved));
    console.log(this.props.workItem.addAsAttachment);
    if (this.props.workItem.addAsAttachment) {
      Office.context.mailbox.getCallbackTokenAsync((result) => { this.uploadAttachment(result.value, (result) => {this.createWorkItem(result)})});
    } else {
      this.createWorkItem(null);
    }
  }

  public uploadAttachment(token, callback): void {
    let email: string = this.props.userProfile.email;
    let id: string = Office.context.mailbox.item.itemId;
    let url: string = Office.context.mailbox.ewsUrl;
    let account: string = this.props.currentSettings.settings.account;

    Rest.getMessage(email, id, url, token, (data) => {
      Rest.uploadAttachment(email, data, account, Office.context.mailbox.item.normalizedSubject, callback);
    });
 
  }
  
  public createWorkItem(attachmentUrl: string): void {
    let options: any = {
      attachment: attachmentUrl,
      type: this.props.workItem.workItemType,
      title: this.props.workItem.title,
      body: this.props.workItem.description,

    };
    let dispatch: any = this.props.dispatch;

    let user: string = this.props.userProfile.email;
    let account: string = this.props.currentSettings.settings.account;
    let project: string = this.props.currentSettings.settings.project;
    let teamName: string = this.props.currentSettings.settings.team;

    Rest.createTask(user, options, account, project, teamName, (workItemInfo: WorkItemInfo) => {
      dispatch(updateSave(workItemInfo.VSTShtmlLink, workItemInfo.id));
      // dispatch(updateStage(Stage.Saved));
      dispatch(updatePageAction(PageVisibility.QuickActions));
    });
  }


  /**
   * Renders the Save button and disables it on click
   */
  public render(): React.ReactElement<Provider> {
    /**
     * Style for the save button 
     */
    let styleEnabled: any = {
      background: 'rgb(16,130,207)',
      border: 'rgb(255,255,255)',
      color: 'rgb(255,255,255)',
      font: '15px arial, ms-segoe-ui',
      margin: '10px',
      'margin-left': '25%',
    };
    let styleDisabled: any = {
      background: 'rgb(192,192,192)',
      border: 'rgb(255,255,255)',
      color: 'rgb(255,255,255)',
      font: '15px arial, ms-segoe-ui',
      margin: '10px',
      'margin-left': '25%',
    };

    let currentStyle: any = this.props.workItem.stage === Stage.Saved ? styleDisabled : styleEnabled;
    let text: any = this.props.workItem.stage === Stage.Saved ? 'Creating...' : 'Create work item';
    return (
      <div>
        <br/>
        <button className = 'ms-Button' style= {currentStyle} disabled = {this.props.workItem.stage === Stage.Saved}
          onClick = {this.handleSave.bind(this)} > {text}
        </button>
      </div>
    );
  }
}
