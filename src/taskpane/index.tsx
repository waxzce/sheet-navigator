import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { App } from "./App";
import "./taskpane.css";

Office.onReady(() => {
  const container = document.getElementById("root");
  if (!container) {
    return;
  }
  createRoot(container).render(
    <React.StrictMode>
      <FluentProvider theme={webLightTheme} style={{ height: "100%" }}>
        <App />
      </FluentProvider>
    </React.StrictMode>
  );
});
