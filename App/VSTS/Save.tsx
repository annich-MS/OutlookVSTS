import * as React from 'react';
import { Provider, connect } from 'react-redux';
// import { changeSave, StageEnum } from '../Reducers/ActionsET';
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
  return { workItem: state.workItem, userProfile: state.userProfile, currentSettings: state.currentSettings };
}

@connect(mapStateToProps)
export class Save extends React.Component<ISaveProps, {}> {
  /**
   * Dispatches the action to change the Stage and make the REST call to create the work item
   * @returns {void}
   */
  public handleSave(): void {
    let options: any = { account: this.props.currentSettings.settings.account, project: this.props.currentSettings.settings.project, teamName: this.props.currentSettings.settings.team };
    let token: any = Office.context.mailbox.getCallbackTokenAsync;
    let OutlookitemID: any = Office.context.mailbox.item.itemId;
    let ewsURL: any = Office.context.mailbox.ewsUrl;
    let returnInfo: WorkItemInfo;
    /*    Rest.createWorkItem('t-emtenc@microsoft.com', options, token, OutlookitemID, ewsURL,
    this.props.workItem.workItemType, this.props.workItem.title, this.props.workItem.description, (output) => console.log(output));*/
    Rest.getCurrentIteration(this.props.userProfile.email, options, this.props.workItem.workItemType,
      this.props.workItem.title, this.props.workItem.description,
      (workItemInfo: WorkItemInfo) => {
        returnInfo = workItemInfo;
        this.props.dispatch(updateSave(returnInfo.VSTShtmlLink, returnInfo.id));
        this.props.dispatch(updateStage(Stage.Saved));
        this.props.dispatch(updatePageAction(PageVisibility.QuickActions));
      });
  }

  /**
   * Renders the Save button and disables it on click
   */
  public render(): React.ReactElement<Provider> {
    /**
     * Style for the live save button 
     */
    let save: any = {
      align: 'center',
      background: '#80ccff',
      height: '35px',
      width: '250px',
    };

    /**
     * Style for the disabled save button
     */
    let disabled: any = {
      align: 'center',
      background: '#d9d9d9',
      height: '35px',
      width: '250px',
    };

    /**
     * Decides which style to use for the stage button based on the Stage
     */
    let currentStyle: any = this.props.workItem.stage === Stage.Saved ? disabled : save;
    return (
      <div>
        <br/>
        <button className = 'ms-Button' style= {currentStyle} disabled = {this.props.workItem.stage === Stage.Saved}
          onClick = {this.handleSave.bind(this) } > Create Work Item
        </button>
      </div>
    );
  }
}
