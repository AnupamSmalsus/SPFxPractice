import * as React from 'react';
import type { IUserRegistrationFormProps } from './IUserRegistrationFormProps';
import UserRegistration from './UserRegistration';

export default class UserRegistrationForm extends React.Component<IUserRegistrationFormProps, {}> {
  public render(): React.ReactElement<IUserRegistrationFormProps> {

    return (
      <>
      <UserRegistration allProps={this.props}/>
      </>
    );
  }
}
