/// <reference path="../../../typings/tsd.d.ts" />
import {AuthState, PageVisibility, PopulationStage, NotificationType} from './FlowActions';

/**
 * Represents the error data in the store
 * @interface IControlStateReducer
 */
export interface IControlStateReducer {
  populationStage: PopulationStage;
  authState: AuthState;
  pageState: PageVisibility;
  notification: INotificationStateReducer;
}


/**
 * Represents the error data in the store
 * @interface INotificationStateReducer
 */
export interface INotificationStateReducer {
  notificationType: NotificationType;
  message: string;
}

/**
 * Represents the initial state for the control flow of the application
 * @type {IControlStateReducer}
 */
const initialControlState: IControlStateReducer = {
  authState : AuthState.None,
  notification: {
    message: '',
    notificationType: NotificationType.Hide,
  },
  pageState : PageVisibility.Settings,
  populationStage: PopulationStage.accountPopulating,
};

/**
 * reducer to update the control state in the store
 * @param {IControlStateReducer} state
 * @param {any} action
 * @returns {IControlStateReducer}
 */
export function updateControlStateReducer(state: IControlStateReducer = initialControlState, action: any): IControlStateReducer {
  switch (action.type) {
    case 'NotificationState':
      return Object.assign({}, state, { notification: {message: action.message , notificationType: action.notificationType}});
    case 'AUTH_STATE':
      return Object.assign({}, state, { authState: action.authState});
    case 'PAGE_STATE':
       return Object.assign({}, state, { pageState: action.pageState});
    case 'DropdownState':
      return Object.assign({}, state, { populationStage: action.populationStage});
    default:
      return state;
  }
}
