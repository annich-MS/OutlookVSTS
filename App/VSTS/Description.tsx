import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { updateAddAsAttachment, updateDescription, Stage } from '../Redux/WorkItemActions';

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
  /**
   * Dispatches the action to change the description value in the store
   * @ returns {void}
   * @param {any} event
   */
  public handleChangeDescription(event: any): void {
    this.props.dispatch(updateDescription(event.target.value));
  }

/**
 * Dispatches the action to update the addAsAttachment and description values in the store
 * @ returns {void}
 */
public handleChangeAddAsAttachment (event: any): void {
  // console.log(JSON.stringify(event.target.checked));
  if (event.target.checked === true) {
    // console.log(this.props.addAsAttachment);
    this.props.dispatch(updateDescription('For more details, please refer to the attached mail thread. ' + this.props.description));
    this.props.dispatch(updateAddAsAttachment(true));
  } else {
    // console.log('false');
    this.props.dispatch(updateDescription(
    this.props.description.replace('For more details, please refer to the attached mail thread. ', '')));
    this.props.dispatch(updateAddAsAttachment(false));
  }
}
  /**
   * Renders the Description heading, the Add Email as Attachment checkbox, and the Description textbox
   */
  public render(): React.ReactElement<Provider> {
    let descriptionStyle: any = {
      height: '150px',
      overflow: 'auto',
      padding: '10px',
      resize: 'none',
      width: '98%',
      'padding-top': '5px',
    };
    let checkboxStyle: any = {
      height: '15px',
      margin: '5px',
      width: '15px',
    };
    return (
      <div>
        <div className='ms-font-1x ms-fontWeight-semibold ms-fontColor-black'> DESCRIPTION </div>
        <label className='15px arial, ms-segoe-ui' >
          <input type='checkbox' style={checkboxStyle} id='cbox' onClick={this.handleChangeAddAsAttachment.bind(this)} defaultChecked/>
          Add e-mail as attachment
        </label>
        <br/>
        <textarea className='15px arial, ms-segoe-ui' style={descriptionStyle} id='description'
          value={this.props.description} onChange={this.handleChangeDescription.bind(this) }>
        </textarea>
        <br/><br/>
      </div>
    );
  }
}

