import { Injectable } from "@nestjs/common";
import * as msal from "@azure/msal-node";
import { SharepointModuleOptions } from "./sharepoint-api.module";
import { Client } from "@microsoft/microsoft-graph-client";

@Injectable()
export class SharepointApiService {
  private readonly tenant_id: string;
  private readonly client_id: string;
  private readonly thumbprint: string;
  private readonly private_key: string;
  private readonly scopes: string[];
  private readonly msal_client: msal.ConfidentialClientApplication;
  private readonly sharepoint_client: Client;

  constructor(readonly options: SharepointModuleOptions) {
    this.tenant_id = options.tenantId;
    this.client_id = options.clientId;
    this.scopes = options.scopes.split(" ") || [""];
    this.thumbprint = options.thumbprint;
    this.private_key = options.privateKey;

    // init msal client
    this.msal_client = new msal.ConfidentialClientApplication({
      auth: {
        authority: `https://login.microsoftonline.com/${this.tenant_id}`,
        clientId: `${this.client_id}`,
        clientCertificate: {
          thumbprint: this.thumbprint,
          privateKey: this.private_key
        }
      }
    });
    // init graph client
    this.sharepoint_client = Client.init({
      authProvider: async resolve => {
        this.msal_client
          .acquireTokenByClientCredential({
            scopes: this.scopes
          })
          .then(token => resolve(null, token.accessToken))
          .catch(error => resolve(error, null));
      },
      customHosts: new Set(["futuraitsrl.sharepoint.com"])
    });
  }

  /**
   *
   * @param listId
   */
  getListSubscriptions = async (listId: string) => {
    return this.sharepoint_client
      .api(
        `https://futuraitsrl.sharepoint.com/sites/hr/_api/web/lists('${listId}')/subscriptions`
      )
      .get();
  };

  /**
   *
   * @param listId
   * @param notificationUrl
   * @param expirationTimestamp
   */
  createListSubscription = async (
    listId: string,
    notificationUrl: string,
    expirationTimestamp: number
  ) => {
    // read key from files
    return this.sharepoint_client
      .api(
        `https://futuraitsrl.sharepoint.com/sites/hr/_api/web/lists('${listId}')/subscriptions`
      )
      .post({
        resource:
          "https://futuraitsrl.sharepoint.com/sites/hr/Lists/Recruiting%20Board/AllItems.aspx",
        notificationUrl: `${notificationUrl}`,
        expirationDateTime: new Date(expirationTimestamp).toISOString()
      });
  };

  /**
   *
   * @param listId
   * @param subscriptionId
   */
  deleteListSubscription = async (listId: string, subscriptionId: string) => {
    return this.sharepoint_client
      .api(
        `https://futuraitsrl.sharepoint.com/sites/hr/_api/web/lists('${listId}')/subscriptions('${subscriptionId}')`
      )
      .delete();
  };

  /**
   *
   * @param options
   */
  getListLogs = async (options: { listId: string; from: Date }) => {
    const startTick = options.from.getTime() * 10000 + 621355968000000000;
    const start = `1;3;${options.listId};${startTick};-1`;
    return this.sharepoint_client
      .api(
        `https://futuraitsrl.sharepoint.com/sites/hr/_api/web/Lists(guid'${options.listId}')/getChanges`
      )
      .post({
        query: {
          Add: true,
          Alert: true,
          DeleteObject: true,
          Update: true,
          Item: true,
          ChangeTokenStart: {
            StringValue: start
          }
        }
      });
  };
}
