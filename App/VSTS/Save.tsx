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
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
        'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
        'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> <soap:Header>' +
        '<RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types"' +
        'soap:mustUnderstand="0" />' +
        '</soap:Header>' +
        '<soap:Body>' +
        '<GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> <ItemShape>' +
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
        function (asyncResult, result) {
          if (asyncResult.status === 'failed') {
            return;
          }
          let response: any = $.parseXML(asyncResult.value);
          let value: string = $(response).find('MimeContent').text();
          mimeString = value;
          Rest.getCurrentIteration(user, options, addAsAttachment, mimeString, workItemType, title,
                                   description, (workItemInfo: WorkItemInfo) => {
                                   console.log('in callback for get curr iteration');
                                   dispatch(updateSave(workItemInfo.VSTShtmlLink, workItemInfo.id));
                                   dispatch(updateStage(Stage.Saved));
                                   dispatch(updatePageAction(PageVisibility.QuickActions));
            });
        });
    }
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
