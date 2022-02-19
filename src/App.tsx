import React, { useState } from 'react';
import { useMsal } from '@azure/msal-react';
import { AccountInfo } from '@azure/msal-browser';

function App() {
  const { instance } = useMsal();
  const [account, setAccount] = useState<AccountInfo | null>();

  return (
    <div>
      <button onClick={() => {
        instance.loginPopup().then(x => setAccount(x.account));
      }}>登陆</button>


      {account && <pre>{JSON.stringify(account)}</pre>}
    </div>
  );
}

export default App;
