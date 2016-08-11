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
    let options: any = {
      account: this.props.currentSettings.settings.account,
      project: this.props.currentSettings.settings.project,
      teamName: this.props.currentSettings.settings.team,
    };
    let mimeString: string = '';
    let addAsAttachment: boolean = this.props.workItem.addAsAttachment;
    let workItemType: string = this.props.workItem.workItemType;
    let title: string = this.props.workItem.title;
    let description: string = this.props.workItem.description;
    let dispatch: any = this.props.dispatch;
    let user: string = this.props.userProfile.email;
    this.props.dispatch(updateStage(Stage.Saved));
    if (this.props.workItem.addAsAttachment) {
      let request: any = '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> <soap:Header>' +
        '<RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
        '</soap:Header>' +
        '<soap:Body>' +
        '<GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> <ItemShape>' +
        '<t:BaseShape>IdOnly</t:BaseShape> <t:IncludeMimeContent>true</t:IncludeMimeContent>' +
        '</ItemShape>' +
        '<ItemIds>' +
        '<t:ItemId Id="' + Office.context.mailbox.item.itemId + '"/>' +
        '</ItemIds>' +
        '</GetItem>' +
        '</soap:Body>' +
        '</soap:Envelope>';
      Office.context.mailbox.makeEwsRequestAsync(
        request,
        function (asyncResult: any, result: any): any {
          if (asyncResult.status === 'failed') {
            console.log('EWS request failed with error: ' + asyncResult.error.code + ' - ' + asyncResult.error.message);
            if (asyncResult.error.code === 9020) {
              alert('Your file exceeds 1 MB size limit. Please modify your EWS request.');
            }
            return;
          }

          let response: any = $.parseXML(asyncResult.value);
          mimeString = $(response).find('MimeContent').text();
          Rest.getCurrentIteration(user, options, addAsAttachment, mimeString, workItemType, title,
            description, (workItemInfo: WorkItemInfo) => {
              console.log('in callback for get curr iteration');
              dispatch(updateSave(workItemInfo.VSTShtmlLink, workItemInfo.id));
              // ispatch(updateStage(Stage.Saved));
              dispatch(updatePageAction(PageVisibility.QuickActions));
            });
        });
    } else { // don't add as attachment
      Rest.getCurrentIteration(user, options, addAsAttachment, mimeString, workItemType, title,
        description, (workItemInfo: WorkItemInfo) => {
          dispatch(updateSave(workItemInfo.VSTShtmlLink, workItemInfo.id));
          // dispatch(updateStage(Stage.Saved));
          dispatch(updatePageAction(PageVisibility.QuickActions));
        });
    }
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
