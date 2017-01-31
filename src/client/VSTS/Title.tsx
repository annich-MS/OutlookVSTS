import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { updateTitle, updateStage, Stage } from '../Redux/WorkItemActions';
import { TextField } from 'office-ui-fabric-react';

/**
 * Represents the Title Properties
 * @interface ITitleProps
 */
export interface ITitleProps {
  /**
   * dispatch to map dispatch to props
   * @type {any}
   */
  dispatch?: any;
  /**
   * title of the work item
   * @type {string}
   */
  title?: string;
  /**
   * Flag to signal the stage the user is on: New if no edits are make, Draft if edits were made, Saved if the user created the work item
   * @type {Stage}
   */
  stage?: Stage;
}

/**
 * Renders the Title heading and Title textbox
 * @class { Title }
 */
function mapStateToProps (state: any): ITitleProps  {
  return {
    stage: state.workItem.stage,
    title: state.workItem.title,
  };
}

@connect (mapStateToProps)
export class Title extends React.Component<ITitleProps, {}> {
  /**
   * Dipatches the action to change the value of title in the store 
   * @returns {void}
   * @param {any} event
   */
  public handleChangeTitle(event: any): void {
    this.props.dispatch(updateTitle (event));
  }
  /**
   * Rendersthe Title heading and the Title textbox
   */
  public render(): React.ReactElement<Provider> {

    /**
     * Gets the normalizedSubject from Office and depending on the Stage, dispatches an action to update the value of title in store
     */
    let normalizedSubject: string = Office.context.mailbox.item.normalizedSubject;
    let currentTitle: string = this.props.title;
    if (currentTitle === '' && this.props.stage === Stage.New) {
        this.props.dispatch(updateTitle (normalizedSubject));
        this.props.dispatch(updateStage (Stage.Draft));
    }
    return (
      <div>
        <TextField
          label='Title'
          onChanged={this.handleChangeTitle.bind(this)}
          value={this.props.title} />
      </div>
    );
  }
}

