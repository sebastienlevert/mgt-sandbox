import { Component, OnInit } from '@angular/core';
import { environment } from '../environments/environment';
import { Providers, Msal2Provider, TemplateHelper, ProviderState } from '@microsoft/mgt';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {
  
  public isLoggedIn(){
    return Providers.globalProvider.state === ProviderState.SignedIn;
  }
  
  ngOnInit() {
    Providers.globalProvider = new Msal2Provider({
      clientId: environment.clientId
    });
    TemplateHelper.setBindingSyntax("[[", "]]");
  }
}
