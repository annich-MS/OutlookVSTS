import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { Rest, RestError, WorkItemInfo, IStringCallback} from '../RestHelpers/rest';
import { updateStage, Stage, updateSave } from '../Redux/WorkItemActions';
import { IWorkItem } from '../Redux/WorkItemReducer';
import { updateErrorAction, updatePageAction, PageVisibility, PopulationStage } from '../Redux/FlowActions';
import { IUserProfileReducer, ISettingsAndListsReducer } from '../Redux/LogInReducer';
import { Button, ButtonType } from 'office-ui-fabric-react';

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
  populationStage?: PopulationStage;
}

/**
 * Renders the Save button and makes REST api calls
 * @class { Save }
 */
function mapStateToProps(state: any): ISaveProps {
  return {
    currentSettings: state.currentSettings,
    populationStage: state.controlState.populationStage,
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
    if (this.props.workItem.addAsAttachment) {
      Office.context.mailbox.getCallbackTokenAsync((tokenResult) => {
        this.uploadAttachment(tokenResult.value, (error, attachmentUrl) => { this.createWorkItem(attachmentUrl); });
      });
    } else {
      this.createWorkItem(null);
    }
  }

  public uploadAttachment(token: string, callback: IStringCallback): void {
    let id: string = Office.context.mailbox.item.itemId;
    let url: string = Office.context.mailbox.ewsUrl || 'https://outlook.office365.com/EWS/Exchange.asmx';
    let account: string = this.props.currentSettings.settings.account;

    Rest.getMessage(id, url, token, (error, data) => {
      if (error) {
        this.props.dispatch(updateErrorAction(true, 'Failed to retrieve message file due to ' + error.type));
        return;
      }
      Rest.uploadAttachment(data, account, Office.context.mailbox.item.normalizedSubject + '.eml', (err, link) => {
        if (err) {
          this.props.dispatch(updateErrorAction(true, 'Failed to upload email attachment due to ' + err.type));
          return;
        }
        callback(null, link);
      });
    });

  }

  public createWorkItem(attachmentUrl: string): void {
    let options: any = {
      account: this.props.currentSettings.settings.account,
      attachment: attachmentUrl,
      body: this.props.workItem.description,
      project: this.props.currentSettings.settings.project,
      teamName: this.props.currentSettings.settings.team,
      title: this.props.workItem.title,
      type: this.props.workItem.workItemType,
    };
    let dispatch: any = this.props.dispatch;


    Rest.createTask(options, (error: RestError, workItemInfo: WorkItemInfo) => {
      if (error) {
        this.props.dispatch(updateErrorAction(true, 'Failed to create work item due to ' + error.type));
        return;
      }
      dispatch(updateSave(workItemInfo.VSTShtmlLink, workItemInfo.id));
      dispatch(updatePageAction(PageVisibility.QuickActions));
    });
  }


  /**
   * Renders the Save button and disables it on click
   */
  public render(): React.ReactElement<Provider> {

    let text: any = this.isSaving ? 'Creating...' : 'Create work item';
    return (
      <div style={{'text-align': 'center'}} >
        <br/>
        <Button
          buttonType={ButtonType.primary}
          disabled = {!this.shouldBeEnabled()}
          onClick={this.handleSave.bind(this)} > {text} </Button>
      </div>
    );
  }

  private get isSaving(): boolean { return this.props.workItem.stage === Stage.Saved; }

  private shouldBeEnabled(): boolean {
    return !(this.isSaving || this.props.populationStage < PopulationStage.teamReady);
  }
}
