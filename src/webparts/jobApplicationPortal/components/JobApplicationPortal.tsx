import * as React from 'react';
import { IJobApplicationPortalProps } from './IJobApplicationPortalProps';
import JobApplicationPortalComponent from './JobApplicationPortalComponent';

export default class JobApplicationPortal extends React.Component<IJobApplicationPortalProps, {}> {
  public render(): React.ReactElement<IJobApplicationPortalProps> {
    

    return (
      <>
      <JobApplicationPortalComponent props={this.props} />
      </>
    );
  }
}
