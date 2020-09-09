import { BaseDataSourceProvider, IDataSourceData } from "@valo/extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import { HttpClient, SPHttpClient } from '@microsoft/sp-http';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import * as Msal from "msal";

const AAD_CONNECT_STORAGE_ENTITY = "ValoAadClientId";
const AAD_LOGIN_URL = "https://login.microsoftonline.com";
const LOADFRAME_TIMEOUT = 6000;

interface Map<T> {
    [K: string]: T;
}

export class ChainedDataSource extends BaseDataSourceProvider<IDataSourceData> {

    private clientId: string = '';
    private msalInstance: Msal.UserAgentApplication | undefined = undefined;
    private msalConfig: Map<Msal.Configuration>;
    private msalLoginRequest: Map<Map<Msal.AuthenticationParameters>>;
    private msalAuthResponse: Map<Map<Msal.AuthResponse | undefined>>;
    private propertyFieldCollectionData;
    private customCollectionFieldType;
    private tenantId: string;
    private tenantUrl: string;

    public async getData(): Promise<IDataSourceData> {

        const apiUrl = `${this.ctx.pageContext.web.absoluteUrl}/_api/web/GetStorageEntity('${AAD_CONNECT_STORAGE_ENTITY}')`;
        const storageEntity = await this.ctx.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1).then(storageData => storageData.json());

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
            return await data.json();
        }
  
        return { items: [] };

    }

    public getConfigProperties(): IPropertyPaneGroup[] {

        this.propertyFieldCollectionData = PropertyFieldCollectionData;
        this.customCollectionFieldType = CustomCollectionFieldType;

        const parametersControl = this.propertyFieldCollectionData("templateParams", {
            key: "templateParams",
            label: `API URLs`,
            panelHeader: "API chain",
            manageBtnLabel: "Configure API chain",
            value: this.properties.apiUrl,
            enableSorting: true,
            fields: [
              {
                id: "apiUrl",
                title: "API URL",
                type: this.customCollectionFieldType.string,
                required: true,
                disableEdit: (this.properties.intTemplate && !this.properties.external)
              },
              {
                id: "method",
                title: "Method",
                type: this.customCollectionFieldType.dropdown,
                required: true,
                options: [ { key: 'GET', text: 'GET' }, { key: 'POST', text: 'POST' }, { key: 'OPTIONS', text: 'OPTIONS' }, { key: 'PATCH', text: 'PATCH' }, { key: 'PUT', text: 'PUT' }, { key: 'DELETE', text: 'DELETE' } ]
              },
              {
                id: "clientId",
                title: "Client Id",
                type: this.customCollectionFieldType.string,
              },
              {
                id: "resource",
                title: "Resource",
                type: this.customCollectionFieldType.string,
              }
            ]
        });
        return [
            {
              groupName: "Chained",
              groupFields: [
                parametersControl
              ],
              isCollapsed: false
            }
          ];
      
    }

    private ensureMsalConfig(clientId: string, tenantScopedAuth: boolean) {

        this.msalConfig = this.msalConfig || {};
        this.msalConfig[clientId] = this.msalConfig[clientId] || {
            auth: {
                clientId: this.clientId,
                authority: `${AAD_LOGIN_URL}/${tenantScopedAuth ? this.tenantId : "common"}`,
                redirectUri: `${this.tenantUrl}/_layouts/images/blank.gif`,
            },
            system: {
                loadFrameTimeout: LOADFRAME_TIMEOUT
            }
        };

    }

    private ensureMsalLoginRequest(clientId: string, scope: string, loginName: string) {

        this.msalLoginRequest = this.msalLoginRequest || {};
        this.msalLoginRequest[clientId] = this.msalLoginRequest[clientId] || {};
        this.msalLoginRequest[clientId][scope] = this.msalLoginRequest[clientId][scope] || {
            scopes: [scope],
            loginHint: this.ctx.pageContext.user.loginName,
        };

    }
    
    public async initAuth() {


        for (let x = 0; x < this.properties.apiUrl.length; x++ ) {

            const chainItem: any = this.properties.apiUrl[x];
            const defaultAadScope = (chainItem.apiUrl && chainItem.apiUrl.indexOf("https://") > -1) ? `${chainItem.apiUrl.substring(0, chainItem.apiUrl.indexOf("/", 8))}/user_impersionation` : '';
            this.ensureMsalConfig(chainItem.clientId || this.clientId, true);
            this.ensureMsalLoginRequest(chainItem.clientId || this.clientId, chainItem.resource || defaultAadScope, this.ctx.pageContext.user.loginName);

        }


        console.log(`DynamicsDataSource scope = ${aadScope}/user_impersonation`);
        this.msalLoginRequest = { 
            scopes: [`${aadScope}/user_impersonation`],
            loginHint: this.ctx.pageContext.user.loginName,
            
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
            this.msalAuthResponse = await ChainedDataSource.acquireTokenSilent(this.msalLoginRequest);
        } catch (err) {
            console.log(err);
        }

    }

    private static async acquireTokenSilent(msalLoginRequest: Msal.AuthenticationParameters): Promise<Msal.AuthResponse | undefined> {
        if (ChainedDataSource.msalInstance) {
            return ChainedDataSource.msalInstance().acquireTokenSilent(msalLoginRequest);
        }
        return undefined;
    }

    private static msalInstance(): Msal.UserAgentApplication {
        return (window as any).apiMsalInstance;
    }

    private static renewAccessToken(msalLoginRequest: Msal.AuthenticationParameters) {
        (window as any).apiMsalTokenTimeout = null;
        ChainedDataSource.ensureRefresh(msalLoginRequest);
    }

    private static async ensureRefresh(msalLoginRequest: Msal.AuthenticationParameters) {

        if (ChainedDataSource.msalInstance().getAccount()) {
            const { DateTime } = require("luxon");
            const expiresOn = (await ChainedDataSource.acquireTokenSilent(msalLoginRequest) as Msal.AuthResponse).expiresOn;
            if (expiresOn) {
                const diffDuration = DateTime.fromJSDate(expiresOn).diff(DateTime.local(), "seconds");
                console.log(`render() accessTokenExpires diff ${JSON.stringify(diffDuration.seconds)}`);
                window.clearTimeout((window as any).apiMsalTokenTimeout);
                (window as any).apiMsalTokenTimeout = window.setTimeout(ChainedDataSource.renewAccessToken, diffDuration.seconds * 1000);
    
            }
        }

    }




}
