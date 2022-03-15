import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PersonCardInteraction, PersonViewType, ViewType } from '@microsoft/mgt-spfx';
import { Agenda, Get, MgtTemplateProps, Person, FileList } from '@microsoft/mgt-react/dist/es6/spfx';
import { Pivot, PivotItem } from '@fluentui/react';


export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  public render(): React.ReactElement<IHelloWorldProps> {
    const {
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.helloWorld} ${hasTeamsContext ? styles.teams : ''} ${isDarkTheme ? 'mgt-dark' : 'mgt-light'}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <Person personQuery="me" view={ViewType.twolines}></Person>

          <Pivot aria-label="Basic Pivot Example">
            <PivotItem headerText="Agenda">
              <Agenda />
            </PivotItem>
            <PivotItem headerText="Files">
              <FileList></FileList> 
            </PivotItem>
            <PivotItem headerText="Emails">
              <div className={styles.mgtGetEmail}>
                <Get resource='/me/mailFolders/Inbox/messages' maxPages={3}>
                  <EmailsComponent template="value"></EmailsComponent>
                  <LoadingComponent template='loading'></LoadingComponent>
                  <ErrorComponent template='error'></ErrorComponent>
                </Get>
              </div>
            </PivotItem>
          </Pivot>
        </div>
      </section>
    );
  }  
}

function EmailsComponent(props: MgtTemplateProps) {
  const email = props.dataContext;
  return (
    <div>
      <div className={styles.email}>
				<div className={styles.header}>
					<div>
						<Person personQuery={email.sender.emailAddress.address} view={PersonViewType.oneline} personCardInteraction={PersonCardInteraction.hover} />
					</div>
				</div>
				<div className={styles.title}>
					<a href={email.webLink} target="_blank" rel="noreferrer">
						<h3>{email.subject}</h3>
					</a>
					<span className={styles.date}>
						{new Date(email.receivedDateTime).toLocaleDateString()}
					</span>
				</div>
        {email.bodyPreview ?? <div className={styles.preview}>{email.bodyPreview}</div>}
        {!email.bodyPreview ?? <div className={`${styles.preview} ${styles.emptyBody}`}>...</div>}
			</div>

    </div>    
  );
}

function LoadingComponent(props: MgtTemplateProps) {
  return (<div>Loading...</div>);
}

function ErrorComponent(props: MgtTemplateProps) {
  return (<pre>{props.dataContext}</pre>);
}
