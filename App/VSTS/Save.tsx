import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { Rest, WorkItemInfo, IRestCallback } from '../RestHelpers/rest';
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

  /**
   * Represents what tier is currently being populated
   * @type {number}
   */
  populationTier?: number;
}

/**
 * Renders the Save button and makes REST api calls
 * @class { Save }
 */
function mapStateToProps(state: any): ISaveProps {
  return {
    currentSettings: state.currentSettings,
    populationTier: state.controlState.populationTier,
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
      Office.context.mailbox.getCallbackTokenAsync((tokenResult) => {
        this.uploadAttachment(tokenResult.value, (attachmentUrl) => { this.createWorkItem(attachmentUrl); });
      });
    } else {
      this.createWorkItem(null);
    }
  }

  public uploadAttachment(token: string, callback: IRestCallback): void {
    let id: string = Office.context.mailbox.item.itemId;
    let url: string = Office.context.mailbox.ewsUrl || 'https://outlook.office365.com/EWS/Exchange.asmx';
    let account: string = this.props.currentSettings.settings.account;

    Rest.getMessage(id, url, token, (data) => {
      Rest.uploadAttachment(data, account, Office.context.mailbox.item.normalizedSubject + '.eml', callback);
    });

  }

  public createWorkItem(attachmentUrl: string): void {
    let options: any = {
      attachment: attachmentUrl,
      body: this.props.workItem.description,
      title: this.props.workItem.title,
      type: this.props.workItem.workItemType,
    };
    let dispatch: any = this.props.dispatch;

    let account: string = this.props.currentSettings.settings.account;
    let project: string = this.props.currentSettings.settings.project;
    let teamName: string = this.props.currentSettings.settings.team;

    Rest.createTask(options, account, project, teamName, (workItemInfo: WorkItemInfo) => {
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

    let currentStyle: any = this.shouldBeEnabled() ? styleEnabled : styleDisabled;
    let text: any = this.isSaving ? 'Creating...' : 'Create work item';
    return (
      <div>
        <br/>
        <button
          className = 'ms-Button'
          style= {currentStyle}
          disabled = {!this.shouldBeEnabled()}
          onClick = {this.handleSave.bind(this)} > {text}
        </button>
      </div>
    );
  }

  private get isSaving(): boolean { return this.props.workItem.stage === Stage.Saved; }

  private shouldBeEnabled(): boolean {
    return !(this.isSaving || this.props.populationTier > 0);
  }
}
