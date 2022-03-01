import React, { useState, useEffect } from 'react';
import './App.css';
import { Agenda, FileList, Get, Login, MgtTemplateProps, Person, PersonCardInteraction, PersonViewType } from '@microsoft/mgt-react';
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { Pivot, PivotItem } from '@fluentui/react';

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
      {isSignedIn &&
        <Pivot aria-label="Basic Pivot Example">
          <PivotItem headerText="Agenda">
            <Agenda />
          </PivotItem>
          <PivotItem headerText="Files">
            <FileList></FileList> 
          </PivotItem>
          <PivotItem headerText="Emails">
            <div className="mgt-get-email">
              <Get resource='/me/mailFolders/Inbox/messages' maxPages={3}>
                <EmailsComponent template="value"></EmailsComponent>
                <LoadingComponent template='loading'></LoadingComponent>
                <ErrorComponent template='error'></ErrorComponent>
              </Get>
            </div>
          </PivotItem>
        </Pivot>
      }
    </div>
  );
}

function EmailsComponent(props: MgtTemplateProps) {
  const email = props.dataContext;
  return (
    <div>
      <div className="email">
				<div className="header">
					<div>
						<Person personQuery={email.sender.emailAddress.address} view={PersonViewType.oneline} personCardInteraction={PersonCardInteraction.hover} />
					</div>
				</div>
				<div className="title">
					<a href={email.webLink} target="_blank" rel="noreferrer">
						<h3>{email.subject}</h3>
					</a>
					<span className="date">
						{new Date(email.receivedDateTime).toLocaleDateString()}
					</span>
				</div>
        {email.bodyPreview ?? <div className="preview">{email.bodyPreview}</div>}
        {!email.bodyPreview ?? <div className="preview empty-body">...</div>}
			</div>

    </div>    
  )
}

function LoadingComponent(props: MgtTemplateProps) {
  return (<div>Loading...</div>)
}

function ErrorComponent(props: MgtTemplateProps) {
  return (<pre>{props.dataContext}</pre>)
}

export default App;

