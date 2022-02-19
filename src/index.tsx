import { render } from "react-dom";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import App from "./App";

const instance = new PublicClientApplication({
  auth: {
    clientId: "07d03982-8b30-48cb-ae23-fcce0c041cca",
    authority://如果是多租户应用，则为 /common
      "https://login.microsoftonline.com/common"
  },
  cache: {
    cacheLocation: "localStorage"//还可以设置为localStorage
  }
});

const rootElement = document.getElementById("root");

Office.onReady(() => {
  render(
    <MsalProvider instance={instance}>
      <App />
    </MsalProvider>,
    rootElement
  );
})
