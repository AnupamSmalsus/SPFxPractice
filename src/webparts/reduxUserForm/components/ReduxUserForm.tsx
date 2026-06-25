import * as React from 'react';
import { IReduxUserFormProps } from './IReduxUserFormProps';
import { Provider } from 'react-redux';
import { store } from '../state/store';
import StudentsTable from './AllStudentsTable';

export default class ReduxUserForm extends React.Component<IReduxUserFormProps, {}> {
  public render(): React.ReactElement<IReduxUserFormProps> { 
    return (
      <React.StrictMode>
        <Provider store={store}>
          <StudentsTable props={this.props}/>
        </Provider>
      </React.StrictMode>
    );
  }
}
