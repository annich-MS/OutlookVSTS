import * as React from "react";

/**
 * Props for ItemHyperlink Component
 */
interface IItemHyperlinkProps {
  /**
   * workItemHyperlink
   */
  workItemHyperlink: string;
}

/**
 * Renders a ReactHTML div 
 */
export class ItemHyperlink extends React.Component<IItemHyperlinkProps, {}> {
  /**
   * Renders the ItemHyperlink Component and reads IItemHyperlinkProps
   */
  public render(): JSX.Element {
    return(
        <div dangerouslySetInnerHTML={{__html: this.props.workItemHyperlink}} />
    );
  }
 }
