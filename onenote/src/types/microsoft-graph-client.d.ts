declare module '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials' {
  import { AuthenticationProvider } from '@microsoft/microsoft-graph-client';
  import { TokenCredential } from '@azure/identity';

  export class TokenCredentialAuthenticationProvider implements AuthenticationProvider {
    constructor(credential: TokenCredential, options?: { scopes: string[] });
    getAccessToken(): Promise<string>;
  }
} 