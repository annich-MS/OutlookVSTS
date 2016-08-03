import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { changeTitle, changeStage, Stage } from '../Reducers/ActionsET';

 /**
  * Represents the Title Properties
  * @interface ITitleProp
  */
export interface ITitleProp {
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
 * Maps elements of the state to properties
 * @returns {ITitleProp}
 * @param {any} state
 */
function mapStateToProps (state: any): ITitleProp  {
  console.log('mapStateToProps' + JSON.stringify(state));
  return {stage: state.workItem.stage, title: state.workItem.title} ;
   }

@connect (mapStateToProps)
/**
 * Renders the Title heading and Title textbox
 * @class { Title }
 */
export class Title extends React.Component<ITitleProp, {}> {
/**
 * Dipatches the action to change the value of title in the store 
 * @returns {void}
 * @param {any} event
 */
public handlechangeTitle(event: any): void {
  this.props.dispatch(changeTitle (event.target.value));
}
/**
 * Rendersthe Title heading and the Title textbox
 */
public render(): React.ReactElement<Provider> {


  let title: any = {
    align: 'left',
    height: '22px',
    width: '250px',
  };
/**
 * Gets the normalizedSubject from Office and depending on the Stage, dispatches an action to update the value of title in store
 */
  let normalizedSubject: string = Office.context.mailbox.item.normalizedSubject;
  let currentTitle: string = this.props.title;
  if (currentTitle === '' && this.props.stage === Stage.New) {
      this.props.dispatch(changeTitle (normalizedSubject));
      this.props.dispatch(changeStage (Stage.Draft));
  }

  return (<div>
     <br/>
     <div className='ms-font-1x ms-fontWeight-semibold ms-fontColor-black'> TITLE </div>
     <input type='text' style={title} className='ms-font-1x ms-fontWeight-light ms-fontColor-black'
        id= 'titleval' value={this.props.title} onChange={this.handlechangeTitle.bind(this)}  />
     <br/> <br/>
     </div>);

  }
}
