import * as React from 'react';
import { IReduxUserFormProps } from './IReduxUserFormProps';
import { Provider } from 'react-redux';
import StudentForm from './StudentForm';
import { store } from '../state/store';

export default class ReduxUserForm extends React.Component<IReduxUserFormProps, {}> {
  public render(): React.ReactElement<IReduxUserFormProps> { 
    return (
      <React.StrictMode>
        <Provider store={store}>
          <StudentForm props={this.props}/>
        </Provider>
      </React.StrictMode>
    );
  }
}
