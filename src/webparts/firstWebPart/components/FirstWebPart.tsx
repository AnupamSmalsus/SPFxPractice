import * as React from 'react';
import type { IFirstWebPartProps } from './IFirstWebPartProps';
import DocumentSearch from './DocumentSearch';

export default class FirstWebPart extends React.Component<IFirstWebPartProps, {}> {
  public render(): React.ReactElement<IFirstWebPartProps> {
    const {
      searchService
    } = this.props;

    return (
      <DocumentSearch searchService={searchService} />
    );
  }
}
