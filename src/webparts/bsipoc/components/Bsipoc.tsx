import * as React from 'react';
import { IBsipocProps } from './IBsipocProps';
import AppContext from "../../../common/AppContext";
import { Provider } from "react-redux";
import store from "../../../common/store";
import App from './App';
export default class Bsipoc extends React.Component<IBsipocProps, {}> {
  public render(): React.ReactElement<IBsipocProps> {
    const { context } = this.props;

    return (
      <AppContext.Provider value={{ context }}>
      <Provider store={store}>
        <App />
      </Provider>
    </AppContext.Provider>
    );
  }
}
