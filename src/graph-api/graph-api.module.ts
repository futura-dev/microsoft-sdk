import {DynamicModule, Module} from '@nestjs/common';
import {GraphApiService} from './graph-api.service';

export type GraphModuleOptions = {
    tenantId: string;
    clientId: string;
    clientSecret: string;
    grantType: 'authorization_code' | 'client_credentials';
    scopes: string;
}

@Module({})
export class GraphApiModule {
    static forFeature(token: string, options: GraphModuleOptions): DynamicModule {

        const provider = {
            provide: token,
            useFactory: async () => {
                return new GraphApiService({...options});
            }
        }
        return {
            module: GraphApiModule,
            imports: [],
            providers: [provider],
            exports: [provider]
        }
    }
}
