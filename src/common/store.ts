import { configureStore } from "@reduxjs/toolkit";
// import { casesReducer } from "./features/cases/casesSlice";
// import { packagingsReducer } from "./features/packagings/packagingSlice";
const store = configureStore({
  reducer: {
    // cases: casesReducer,
    // packagings: packagingsReducer,
  },
});

export type AppDispatch = typeof store.dispatch;
export type RootState = ReturnType<typeof store.getState>;
export default store;
