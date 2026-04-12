import React from "react";
import ReactDOM from "react-dom/client";
import { App } from "./App";
import { initializeMicrosoftAuth } from "./auth";
import "./styles.css";

async function bootstrap() {
  await initializeMicrosoftAuth();

  ReactDOM.createRoot(document.getElementById("root")!).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
}

bootstrap().catch((error) => {
  console.error("Failed to initialize Microsoft authentication", error);
});
