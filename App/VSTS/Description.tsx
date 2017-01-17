import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { updateAddAsAttachment, updateDescription, Stage } from '../Redux/WorkItemActions';
import { Checkbox, TextField } from 'office-ui-fabric-react';

/**
 * Represents the Description Properties
 * @interface IDescriptionProps
 */
export interface IDescriptionProps {
  /**
   * dispatch to map dispatch to props
   * @type {any}
   */
  dispatch?: any;
  /**
   * the text in the description
   * @type {string}
   */
  description?: string;
  /**
   * whether to attach the email on the work item
   * @type {boolean}
   */
  addAsAttachment?: boolean;
  /**
   * indicates if form has been changed or not
   * @type {Stage}
   */
  stage?: Stage;
}

/**
 * Renders the Description heading, Add Email as Attachment checkbox, and description textbox
 * @class { Description }
 */
function mapStateToProps(state: any): IDescriptionProps {
  return { addAsAttachment: state.workItem.addAsAttachment, description: state.workItem.description, stage: state.workItem.stage };
}

@connect(mapStateToProps)
export class Description extends React.Component<IDescriptionProps, {}> {

  private CHECKBOX_LABEL: string = 'Add e-mail as attachment';

  /**
   * Dispatches the action to change the description value in the store
   * @ returns {void}
   * @param {any} event
   */
  public handleChangeDescription(event: any): void {
    this.props.dispatch(updateDescription(event));
  }

  /**
   * Dispatches the action to update the addAsAttachment and description values in the store
   * @ returns {void}
   */
  public handleChangeAddAsAttachment(event: any, isChecked: boolean): void {
    if (isChecked === true) {
      this.props.dispatch(updateDescription('For more details, please refer to the attached mail thread. ' + this.props.description));
    } else {
      this.props.dispatch(updateDescription(
        this.props.description.replace('For more details, please refer to the attached mail thread. ', '')));
    }
    this.props.dispatch(updateAddAsAttachment(isChecked));
  }

  /**
   * Renders the Description heading, the Add Email as Attachment checkbox, and the Description textbox
   */
  public render(): React.ReactElement<Provider> {

    let checkbox: any = null;

    if (Office.context.mailbox.diagnostics.hostName !== 'OutlookIOS') {
      checkbox = <Checkbox label={this.CHECKBOX_LABEL} onChange={this.handleChangeAddAsAttachment.bind(this) } defaultChecked={true} />;
    }
    return (
      <div>
        {checkbox}
        <TextField
          id='description'
          label='Description'
          value={this.props.description}
          onChanged={this.handleChangeDescription.bind(this) }
          multiline={true} />
      </div>
    );
  }
}

