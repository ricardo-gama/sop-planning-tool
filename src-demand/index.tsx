/// <reference types="office-js" />
import * as React from "react";
import * as ReactDOM from "react-dom";
import App from "./App";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    ReactDOM.render(
      <React.StrictMode>
        <App />
      </React.StrictMode>,
      document.getElementById("root")
    );
  }
});