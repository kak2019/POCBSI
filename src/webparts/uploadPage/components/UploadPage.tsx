import * as React from 'react';

import { IUploadPageProps } from './IUploadPageProps';
import App from "./App";
import AppContext from "../../../common/AppContext";
import { Provider } from "react-redux";
import store from "../../../common/store";

export default class UploadPage extends React.Component<IUploadPageProps, {}> {
  public render(): React.ReactElement<IUploadPageProps> {
    const { context} = this.props;


    return (
      <AppContext.Provider value={{ context }}>
      <Provider store={store}>
        <App />
      </Provider>
    </AppContext.Provider>
    );
  }
}
