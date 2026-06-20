/// <reference types="node" />
import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, module, require, HTMLElement */

Office.onReady(() => {
  const rootElement = document.getElementById("container");
  if (rootElement) {
    const root = createRoot(rootElement);
    root.render(
      <FluentProvider theme={webLightTheme}>
        <App />
      </FluentProvider>
    );
  }
});

// if ((module as any).hot) {
//   (module as any).hot.accept("./App", () => {
//     const NextApp = require("./App").default;
//     const rootElement = document.getElementById("container");
//     if (rootElement) {
//       const root = createRoot(rootElement);
//       root.render(
//         <FluentProvider theme={webLightTheme}>
//           <NextApp />
//         </FluentProvider>
//       );
//     }
//   });
// }