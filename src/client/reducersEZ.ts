import { Reducer, combineReducers } from 'redux';
import { InitialState, IWorkItem, IChangeFollowAction } from './statesEZ';

/**
 * Reducer that describes the change to the followState member of the work item
 * @returns { IWorkItem }
 */
function changeFollowReducer(state: IWorkItem = InitialState, action: IChangeFollowAction): IWorkItem {
  switch (action.type) {
    case 'ChangeFollowState':
      return Object.assign({}, state, {followState: action.followState});
    default:
      return state;
  }
}

/**
 * The reducer for the QuickActions component
 * @const 
 */
export const quickActionsReducer: Reducer = combineReducers({ workItemState: changeFollowReducer });
