import React, { useState } from 'react';
import { useMsal } from '@azure/msal-react';
import { AccountInfo } from '@azure/msal-browser';
import { BrowserRouter, Route } from "react-router-dom";

const App = () => {
  return <BrowserRouter>
    <Route path='/' element={Home}></Route>
    <Route path='/login' element={Login}></Route>
  </BrowserRouter>
}

const Login = () => {
  return <h1>Login
    <button onClick={() => {
      Office.context.ui.messageParent("test");
    }}>提交</button>

  </h1>
}

const Home = () => {
  const { instance } = useMsal();
  const [account, setAccount] = useState<AccountInfo | null>();

  return (
    <div>
      <button onClick={() => {
        instance.loginPopup().then(x => setAccount(x.account));
      }}>登陆</button>

      <button onClick={async () => {
        await Excel.run(async (context) => {
          context.workbook.getActiveCell().values = [["1"]];
          await context.sync();
        })
      }}>设置单元格</button>


      <button onClick={() => {
        Office.context.ui.displayDialogAsync("https://nice-moss-06bb30900.1.azurestaticapps.net/login", {}, (result) => {
          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (message) => {
            alert(message);
          });
        })
      }}>弹出对话框</button>


      {account && <pre>{JSON.stringify(account)}</pre>}
    </div>
  );
}

export default App;
