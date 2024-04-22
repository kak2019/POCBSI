import { createContext } from "react";

export interface AppContextProps {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    context: any;
}

const AppContext = createContext<AppContextProps>(undefined);

export default AppContext;