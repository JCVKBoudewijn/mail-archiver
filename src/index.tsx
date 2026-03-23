import React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { Taskpane } from "./components/Taskpane";

Office.onReady(() => {
  const container = document.getElementById("root");
  if (!container) throw new Error("Root element not found");

  const root = createRoot(container);
  root.render(
    <FluentProvider theme={webLightTheme}>
      <Taskpane />
    </FluentProvider>
  );
});
