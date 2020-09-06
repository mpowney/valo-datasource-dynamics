import { BaseDataSourceProvider, IDataSourceData } from "@valo/extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import { HttpClient, SPHttpClient } from '@microsoft/sp-http';
import * as Msal from "msal";

const AAD_CONNECT_STORAGE_ENTITY = "ValoAadClientId";
const AAD_LOGIN_URL = "https://login.microsoftonline.com";
const LOADFRAME_TIMEOUT = 6000;

export class Dynamics365DataSource extends BaseDataSourceProvider<IDataSourceData> {

    private clientId: string = '';
    private msalInstance: Msal.UserAgentApplication | undefined = undefined;
    private msalConfig: Msal.Configuration;
    private msalLoginRequest: Msal.AuthenticationParameters;
    private msalAuthResponse: Msal.AuthResponse | undefined;

    public async getData(): Promise<IDataSourceData> {

        const apiUrl = `${this.ctx.pageContext.web.absoluteUrl}/_api/web/GetStorageEntity('${AAD_CONNECT_STORAGE_ENTITY}')`;
        const storageEntity = await this.ctx.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1).then(data => data.json());

        if (storageEntity && storageEntity.Value) {
            this.clientId = storageEntity.Value;
        } else {
            throw `Storage entity ${AAD_CONNECT_STORAGE_ENTITY} was not found`;
        }

        await this.initAuth();

        if (!this.msalAuthResponse) {
            console.log(`Not authenticated`);
            return { items: [] };
        }

        const data = await this.ctx.httpClient.get(this.properties.apiUrl, HttpClient.configurations.v1, {
            headers: {
                "Authorization": `Bearer ${this.msalAuthResponse.accessToken}`,
                "content-type": "application/json",
                "accept": "application/json"
            }
        });
        if (data && data.ok) {
            return { items: [await data.json()] };
        }
  
        return { items: [] };

    }

    public getConfigProperties(): IPropertyPaneGroup[] {

        return [
            {
              groupName: "Dynamics 365",
              groupFields: [
                PropertyPaneTextField('apiUrl', {
                  value: this.properties.apiUrl,
                  label: "API URL"
                })
              ],
              isCollapsed: false
            }
          ];
      
    }

    public async initAuth() {

        let tenantId: string = this.ctx.pageContext.aadInfo.tenantId;
        let tenantUrl: string = this.ctx.pageContext.site.absoluteUrl.replace(this.ctx.pageContext.site.serverRelativeUrl, "");

        const aadScope = (this.properties.apiUrl && this.properties.apiUrl.indexOf("https://") > -1) ? this.properties.apiUrl.substring(0, this.properties.apiUrl.indexOf("/", 8)) : '';
        if (aadScope === '') return;
        
        console.log(`DynamicsDataSource scope = ${aadScope}/user_impersonation`);
        this.msalLoginRequest = { 
            scopes: [`${aadScope}/user_impersonation`],
            loginHint: this.ctx.pageContext.user.loginName,
            
        };

        this.msalConfig = {
            auth: {
                clientId: this.clientId,
                authority: `${AAD_LOGIN_URL}/${tenantId ? tenantId : "common"}`,
                redirectUri: `${tenantUrl}/_layouts/images/blank.gif`,
            },
            system: {
                loadFrameTimeout: LOADFRAME_TIMEOUT
            }
        };

        if (!this.msalInstance) {
            if ((window as any).apiMsalInstance) {
                this.msalInstance = (window as any).apiMsalInstance;
            }
            else {
                this.msalInstance = new Msal.UserAgentApplication(this.msalConfig);
                (window as any).apiMsalInstance = this.msalInstance;
            }
        }

        if (this.msalInstance) {

            this.msalInstance.handleRedirectCallback((error: any, response: any) => {
                // handle redirect response or error
                if (error) {
                    console.log(`Error: ${error.errorMessage}`);
                } else if (response) {
                    console.log(`Response from MSAL: ${response.account}`);
                }
            });
        
        }

        console.log(`Calling acquireTokenSilent()`);

        try {
            this.msalAuthResponse = await Dynamics365DataSource.acquireTokenSilent(this.msalLoginRequest);
        } catch (err) {
            console.log(err);
        }

    }

    private static async acquireTokenSilent(msalLoginRequest: Msal.AuthenticationParameters): Promise<Msal.AuthResponse | undefined> {
        if (Dynamics365DataSource.msalInstance) {
            return Dynamics365DataSource.msalInstance().acquireTokenSilent(msalLoginRequest);
        }
        return undefined;
    }

    private static msalInstance(): Msal.UserAgentApplication {
        return (window as any).apiMsalInstance;
    }

    private static renewAccessToken(msalLoginRequest: Msal.AuthenticationParameters) {
        (window as any).apiMsalTokenTimeout = null;
        Dynamics365DataSource.ensureRefresh(msalLoginRequest);
    }

    private static async ensureRefresh(msalLoginRequest: Msal.AuthenticationParameters) {

        if (Dynamics365DataSource.msalInstance().getAccount()) {
            const { DateTime } = require("luxon");
            const expiresOn = (await Dynamics365DataSource.acquireTokenSilent(msalLoginRequest) as Msal.AuthResponse).expiresOn;
            if (expiresOn) {
                const diffDuration = DateTime.fromJSDate(expiresOn).diff(DateTime.local(), "seconds");
                console.log(`render() accessTokenExpires diff ${JSON.stringify(diffDuration.seconds)}`);
                window.clearTimeout((window as any).apiMsalTokenTimeout);
                (window as any).apiMsalTokenTimeout = window.setTimeout(Dynamics365DataSource.renewAccessToken, diffDuration.seconds * 1000);
    
            }
        }

    }




}
