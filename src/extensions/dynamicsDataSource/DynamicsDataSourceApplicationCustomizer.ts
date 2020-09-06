import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPHttpClient } from "@microsoft/sp-http";
import { DataSourceService } from '@valo/extensibility';
import * as Msal from "msal";

import * as strings from 'DynamicsDataSourceApplicationCustomizerStrings';
import { Dynamics365DataSource } from './dataSource/DynamicsDataSource';

const LOG_SOURCE: string = 'DynamicsDataSourceApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IDynamicsDataSourceApplicationCustomizerProperties {
    // This is an example; replace with your own property
    testMessage: string;
}

const AAD_STORED_TOKEN = "Valo:DynamicsAccessToken";
const AAD_CONNECT_STORAGE_ENTITY = "ValoAadClientId";
const DYNAMICS_RESOURCE_STORAGE_ENTITY = "ValoAadDynamicsStorageEntity";
const AAD_LOGIN_URL = "https://login.microsoftonline.com";
const AAD_AUTHORIZE_URL = "oauth2/authorize";
const AAD_DYNAMICS_SCOPE = "https://admin.services.crm.dynamics.com/user_impersonation";
const LOADFRAME_TIMEOUT = 60000;
const POPUP_TIMEOUT = null; // Turn off timeout for popup -> needed for when user login expired

/** A Custom Action which can be run during execution of a Client Side Application */
export default class DynamicsDataSourceApplicationCustomizer extends BaseApplicationCustomizer<IDynamicsDataSourceApplicationCustomizerProperties> {

    private dataSourceService: DataSourceService = null;
    private clientId: string = '';
    private msalInstance: Msal.UserAgentApplication | undefined = undefined;
    private msalConfig: Msal.Configuration;
    private msalLoginRequest: Msal.AuthenticationParameters;

    @override
    public async onInit(): Promise<void> {

        console.log(`Initialized DynamicsDataSourceApplicationCustomizer`);

        this.dataSourceService = DataSourceService.getInstance();

        this.dataSourceService.registerDataSource({
            dataSourcePrototype: Dynamics365DataSource.prototype,
            id: "DynamicsDataSource",
            name: "Dynamics 365"
        });

        return Promise.resolve();
    }




}
