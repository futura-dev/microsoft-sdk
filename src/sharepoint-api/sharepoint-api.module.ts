import { DynamicModule, Module, Provider } from "@nestjs/common";
import { SharepointApiService } from "./sharepoint-api.service";

export type SharepointModuleOptions = {
  tenantId: string;
  clientId: string;
  thumbprint: string;
  privateKey: string;
  scopes: string;
};

@Module({})
export class SharepointApiModule {
  static forFeature(
    token: string,
    options: SharepointModuleOptions
  ): DynamicModule {
    const provider: Provider = {
      provide: token,
      useFactory: async () => {
        return new SharepointApiService({ ...options });
      }
    };

    return {
      module: SharepointApiModule,
      imports: [],
      providers: [provider],
      exports: [provider]
    };
  }
}
