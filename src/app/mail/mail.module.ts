import { NgModule, importProvidersFrom } from '@angular/core';
import { CommonModule } from '@angular/common';
import { MailComponent } from './mail.component';
import { RouterModule, Routes, provideRouter } from '@angular/router';
import { MSAL_GUARD_CONFIG, MSAL_INSTANCE, MSAL_INTERCEPTOR_CONFIG, MsalBroadcastService, MsalGuard, MsalGuardConfiguration, MsalInterceptor, MsalInterceptorConfiguration, MsalRedirectComponent, MsalService } from '@azure/msal-angular';
import { BrowserCacheLocation, IPublicClientApplication, InteractionType, LogLevel, PublicClientApplication } from '@azure/msal-browser';
import { HTTP_INTERCEPTORS, provideHttpClient, withFetch, withInterceptorsFromDi } from '@angular/common/http';
import { MatToolbarModule } from '@angular/material/toolbar';
import { MatButtonModule } from '@angular/material/button';
import { provideNoopAnimations } from '@angular/platform-browser/animations';
import { GraphService } from './service/graph.service';
// import { MailSearchComponent } from './component/mail-search/mail-search.component';
import { MatInputModule } from '@angular/material/input';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatIconModule } from '@angular/material/icon';
import { MsLoginComponent } from './component/ms-login/ms-login.component';
import { MailboxComponent } from './component/mailbox/mailbox.component';
import { DemoFlexyModule } from '../demo-flexy-module';
import { NgApexchartsModule } from 'ng-apexcharts';

const routes: Routes = [
  { 
    path: '', component: MailComponent, children: [
      // { path: 'sender', component: SenderComponent, canActivate: [MsalGuard] },
      { path: 'mailbox', component: MailboxComponent, canActivate: [MsalGuard] },
      { path: 'ms-login', component: MsLoginComponent },
    ]
  },
];

@NgModule({
  imports: [
    CommonModule,
    RouterModule.forChild(routes),
    MatToolbarModule,
    MatInputModule,
    FormsModule,
    ReactiveFormsModule,
    MatFormFieldModule,
    MatIconModule,
    DemoFlexyModule,
    NgApexchartsModule,
  ],
  declarations: [
    MailComponent,
    // SenderComponent,
    MsLoginComponent,
    MailboxComponent,
    
  ],
  providers: [
    provideRouter(routes), 
    importProvidersFrom(MatButtonModule, MatToolbarModule),
    provideNoopAnimations(),
    provideHttpClient(withInterceptorsFromDi(), withFetch()),
    {
        provide: HTTP_INTERCEPTORS,
        useClass: MsalInterceptor,
        multi: true
    },
    {
        provide: MSAL_INSTANCE,
        useFactory: MSALInstanceFactory
    },
    {
        provide: MSAL_GUARD_CONFIG,
        useFactory: MSALGuardConfigFactory
    },
    {
        provide: MSAL_INTERCEPTOR_CONFIG,
        useFactory: MSALInterceptorConfigFactory
    },
    MsalService,
    MsalGuard,
    MsalBroadcastService,
    GraphService,
  ],
  bootstrap: [MailComponent, MsalRedirectComponent]
})
export class MailModule { }

export function MSALInstanceFactory(): IPublicClientApplication {
  return new PublicClientApplication({
    auth: {
      clientId: "047bcd51-4b13-487e-9899-fbddd7690f1d",
      authority: "https://login.microsoftonline.com/e8c68ddc-2298-43d0-8b43-48f719b469d1/oauth2/v2.0/authorize",
      redirectUri: "http://localhost:4200/mail/ms-login",
      postLogoutRedirectUri: '/logout'
    },
    cache: {
      cacheLocation: BrowserCacheLocation.LocalStorage,
      storeAuthStateInCookie: isIE,
    },
    system: {
      allowNativeBroker: false, // Disables WAM Broker
      loggerOptions: {
        loggerCallback: (logLevel, message, piiEnabled) => {
          // console.log('MSAL Logging: ', message, logLevel, piiEnabled);
        },
        // logLevel: LogLevel.Info,
        piiLoggingEnabled: false
      }
    }
  });
}

const isIE =
  window.navigator.userAgent.indexOf("MSIE ") > -1 ||
  window.navigator.userAgent.indexOf("Trident/") > -1;

export function MSALInterceptorConfigFactory(): MsalInterceptorConfiguration {
  const protectedResourceMap = new Map([ 
    ['https://graph.microsoft.com/v1.0/me', ['user.read']],
    ['https://graph.microsoft.com/v1.0/me/messages', ['openid', 'profile', 'User.Read', 'email', 'Mail.Read', 'Mail.ReadWrite', 'Mail.ReadBasic']]
  ]);

  return {
    interactionType: InteractionType.Redirect,
    protectedResourceMap
  };
}

export function MSALGuardConfigFactory(): MsalGuardConfiguration {
  return { 
    interactionType: InteractionType.Redirect,
    authRequest: {
      scopes: ['openid', 'https://graph.microsoft.com/mail.send', 'user.read', 'profile', 'email', 'Mail.Read', 'Mail.ReadWrite', 'Mail.ReadBasic']
    },
    loginFailedRoute: '/login-failed'
  };
}