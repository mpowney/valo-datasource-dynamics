import { override } from '@microsoft/decorators';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { DataSourceService } from '@valo/extensibility';

import { Dynamics365DataSource } from './dataSource/DynamicsDataSource';

const LOG_SOURCE: string = 'DynamicsDataSourceApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IDynamicsDataSourceApplicationCustomizerProperties {
}


/** A Custom Action which can be run during execution of a Client Side Application */
export default class DynamicsDataSourceApplicationCustomizer extends BaseApplicationCustomizer<IDynamicsDataSourceApplicationCustomizerProperties> {

    private dataSourceService: DataSourceService = null;

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
