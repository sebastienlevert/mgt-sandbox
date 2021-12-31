import { useState, useEffect } from 'react';
import './App.css';
import { Login, Person, ViewType } from '@microsoft/mgt-react';
import { Providers, ProviderState } from '@microsoft/mgt-element';

function useIsSignedIn(): [boolean] {
  const [isSignedIn, setIsSignedIn] = useState(false);

  useEffect(() => {
    const updateState = () => {
      const provider = Providers.globalProvider;
      setIsSignedIn(provider && provider.state === ProviderState.SignedIn);
    };

    Providers.onProviderUpdated(updateState);
    updateState();

    return () => {
      Providers.removeProviderUpdatedListener(updateState);
    }
  }, []);

  return [isSignedIn];
}

function App() {
  const [isSignedIn] = useIsSignedIn();

  return (
    <div className="App">
      <header>
        <Login />
      </header>
      <div>
        {isSignedIn && (
          <Person personQuery="me" view={ViewType.threelines}></Person>
        )}        
      </div>
    </div>
  );
}

export default App;
