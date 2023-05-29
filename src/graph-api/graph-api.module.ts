import { DynamicModule, Module, Provider } from "@nestjs/common";
import { GraphApiService } from "./graph-api.service";

export type GraphModuleOptions = {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  scopes: string;
};

@Module({})
export class GraphApiModule {
  static forFeature(token: string, options: GraphModuleOptions): DynamicModule {
    const provider: Provider = {
      provide: token,
      useFactory: async () => {
        return new GraphApiService({ ...options });
      }
    };
    return {
      module: GraphApiModule,
      imports: [],
      providers: [provider],
      exports: [provider]
    };
  }
}
