import * as React from 'react';
import { Provider } from 'react-redux';

/**
 * Dumb component
 * Renders the static add-in description text
 * @class {AddInDescription} 
 */
export class AddInDescription extends React.Component<{}, {}> {

  /**
   * Renders the add-in description text
   */
  public render(): React.ReactElement<Provider> {
    let titleClasses: string = 'ms-font-l ms-fontWeight-semibold ms-fontColor-themePrimary';
    let bodyClasses: string = 'ms-font-l';

    return(<div>
      <div>
        <p className={titleClasses} > Create work items</p>
        <p className={bodyClasses}>Turn an email thread into a work item directly from Outlook!</p>
      </div>
      <div>
        <p className={titleClasses}> Communicate with your team </p>
        <p className={bodyClasses}> Once the work item is created,
        use the reply-all feature to close the thread with a link and details to the work item. </p>
      </div>
      </div>
    );
  }
  }


