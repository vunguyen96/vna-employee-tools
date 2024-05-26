import { Injectable } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { AuthenticationProvider, AuthenticationProviderOptions, Client, ClientOptions } from '@microsoft/microsoft-graph-client';
import { firstValueFrom } from 'rxjs';



@Injectable()
export class GraphService {
  private client: Client;

  constructor(
    private msalService: MsalService,
  ) { }

  public getGraphClient() {
    if (!this.client) {
      const accounts = this.msalService.instance.getAllAccounts();
      let account;

      if (accounts.length > 0) {
        account = accounts[0];

        const accessTokenRequest = {
          scopes: ['openid', 'profile', 'User.Read', 'email', 'Mail.Read', 'Mail.ReadWrite', 'Mail.ReadBasic'],
          account: account,
        };
  
        const authProvider: AuthenticationProvider = {
          getAccessToken: (authenticationProviderOptions?: AuthenticationProviderOptions | undefined) => {
            return firstValueFrom(this.msalService.acquireTokenSilent(accessTokenRequest)).then((accessTokenResponse) => accessTokenResponse.accessToken);
          },
        };
        let clientOptions: ClientOptions = {
          // MyFirstMiddleware is the first middleware in my custom middleware chain
          authProvider: authProvider,
        };
        this.client = Client.initWithMiddleware(clientOptions);
      }
    }
    return this.client;
  }
}
