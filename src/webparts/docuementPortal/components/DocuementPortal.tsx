import * as React from 'react';
import { IDocuementPortalProps } from './IDocuementPortalProps';
import DocuementPortalComponent from './DocumentPortalComponent';

export default class DocuementPortal extends React.Component<IDocuementPortalProps, {}> {
  public render(): React.ReactElement<IDocuementPortalProps> {

    return (
      <DocuementPortalComponent allProps={this.props}/>
    );
  }
}
