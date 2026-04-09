import { Injectable } from "@nestjs/common";
import * as msal from "@azure/msal-node";
import { SharepointModuleOptions } from "./sharepoint-api.module";
import { Client } from "@microsoft/microsoft-graph-client";
import {
  DriveItem,
  DriveItemPermissionsResponse,
  DriveRecipient,
  InviteDriveItemPermissionsResponse,
} from "./types";

@Injectable()
export class SharepointApiService {
  private readonly tenant_id: string;
  private readonly client_id: string;
  private readonly client_secret: string;
  private readonly scopes: string[];
  private readonly msal_client: msal.ConfidentialClientApplication;
  private readonly sharepoint_client: Client;

  constructor(readonly options: SharepointModuleOptions) {
    this.tenant_id = options.tenantId;
    this.client_id = options.clientId;
    this.client_secret = options.clientSecret;
    this.scopes = options.scopes.split(" ") || [""];

    this.msal_client = new msal.ConfidentialClientApplication({
      auth: {
        authority: `https://login.microsoftonline.com/${this.tenant_id}`,
        clientId: `${this.client_id}`,
        clientSecret: `${this.client_secret}`,
      },
    });

    this.sharepoint_client = Client.init({
      authProvider: async (resolve) => {
        this.msal_client
          .acquireTokenByClientCredential({
            scopes: this.scopes,
          })
          .then((token) => {
            if (!token) throw new Error();
            resolve(null, token.accessToken);
          })
          .catch((error) => resolve(error, null));
      },
      customHosts: new Set(["futuraitsrl.sharepoint.com"]),
    });
  }

  private encodePath = (path: string): string => {
    return path
      .split("/")
      .filter(Boolean)
      .map((segment) => encodeURIComponent(segment))
      .join("/");
  };

  getListSubscriptions = async (listId: string) => {
    return this.sharepoint_client
      .api(
        `https://futuraitsrl.sharepoint.com/sites/hr/_api/web/lists('${listId}')/subscriptions`,
      )
      .get();
  };

  createListSubscription = async (
    listId: string,
    notificationUrl: string,
    expirationTimestamp: number,
  ) => {
    return this.sharepoint_client
      .api(
        `https://futuraitsrl.sharepoint.com/sites/hr/_api/web/lists('${listId}')/subscriptions`,
      )
      .post({
        resource:
          "https://futuraitsrl.sharepoint.com/sites/hr/Lists/Recruiting%20Board/AllItems.aspx",
        notificationUrl: `${notificationUrl}`,
        expirationDateTime: new Date(expirationTimestamp).toISOString(),
      });
  };

  deleteListSubscription = async (listId: string, subscriptionId: string) => {
    return this.sharepoint_client
      .api(
        `https://futuraitsrl.sharepoint.com/sites/hr/_api/web/lists('${listId}')/subscriptions('${subscriptionId}')`,
      )
      .delete();
  };

  getListLogs = async (options: { listId: string; from: Date }) => {
    const startTick = options.from.getTime() * 10000 + 621355968000000000;
    const start = `1;3;${options.listId};${startTick};-1`;
    return this.sharepoint_client
      .api(
        `https://futuraitsrl.sharepoint.com/sites/hr/_api/web/Lists(guid'${options.listId}')/getChanges`,
      )
      .post({
        query: {
          Add: true,
          Alert: true,
          DeleteObject: true,
          Update: true,
          Item: true,
          ChangeTokenStart: {
            StringValue: start,
          },
        },
      });
  };

  uploadDriveItemIntoSite = async (input: {
    siteId: string;
    fileName: string;
    content: Buffer;
  }): Promise<DriveItem> => {
    return await this.sharepoint_client
      .api(`/sites/${input.siteId}/drive/root:/${input.fileName}:/content`)
      .headers({ "Content-Type": "application/octet-stream" })
      .put(input.content);
  };

  getDriveItemBySitePath = async (input: {
    siteId: string;
    itemPath: string;
  }): Promise<DriveItem> => {
    return this.sharepoint_client
      .api(
        `/sites/${input.siteId}/drive/root:/${this.encodePath(input.itemPath)}`,
      )
      .get();
  };

  createFolderIntoSite = async (input: {
    siteId: string;
    parentPath: string;
    folderName: string;
    conflictBehavior?: "fail" | "rename" | "replace";
  }): Promise<DriveItem> => {
    return this.sharepoint_client
      .api(
        `/sites/${input.siteId}/drive/root:/${this.encodePath(input.parentPath)}:/children`,
      )
      .post({
        name: input.folderName,
        folder: {},
        "@microsoft.graph.conflictBehavior":
          input.conflictBehavior ?? "replace",
      });
  };

  ensureFolderPathIntoSite = async (input: {
    siteId: string;
    basePath: string;
    folderPath: string;
  }): Promise<void> => {
    const segments = input.folderPath.split("/").filter(Boolean);
    let currentPath = input.basePath;

    for (const segment of segments) {
      const nextPath = `${currentPath}/${segment}`;

      try {
        await this.getDriveItemBySitePath({
          siteId: input.siteId,
          itemPath: nextPath,
        });
      } catch (error) {
        const statusCode = (error as { statusCode?: number })?.statusCode;
        const code = (error as { code?: string })?.code;

        if (statusCode !== 404 && code !== "itemNotFound") {
          throw error;
        }

        await this.createFolderIntoSite({
          siteId: input.siteId,
          parentPath: currentPath,
          folderName: segment,
          conflictBehavior: "replace",
        });
      }

      currentPath = nextPath;
    }
  };

  getDriveItemPermissions = async (input: {
    siteId: string;
    itemId: string;
  }): Promise<DriveItemPermissionsResponse> => {
    return this.sharepoint_client
      .api(`/sites/${input.siteId}/drive/items/${input.itemId}/permissions`)
      .get();
  };

  inviteDriveItemPermissions = async (input: {
    siteId: string;
    itemId: string;
    recipients: DriveRecipient[];
    roles: string[];
    requireSignIn?: boolean;
    sendInvitation?: boolean;
  }): Promise<InviteDriveItemPermissionsResponse> => {
    return this.sharepoint_client
      .api(`/sites/${input.siteId}/drive/items/${input.itemId}/invite`)
      .post({
        recipients: input.recipients,
        roles: input.roles,
        requireSignIn: input.requireSignIn ?? true,
        sendInvitation: input.sendInvitation ?? false,
      });
  };

  deleteDriveItemPermission = async (input: {
    siteId: string;
    itemId: string;
    permissionId: string;
  }) => {
    return this.sharepoint_client
      .api(
        `/sites/${input.siteId}/drive/items/${input.itemId}/permissions/${input.permissionId}`,
      )
      .delete();
  };
}
