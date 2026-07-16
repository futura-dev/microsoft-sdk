import { Injectable } from "@nestjs/common";
import { GraphModuleOptions } from "./graph-api.module";
import { Client } from "@microsoft/microsoft-graph-client";
import * as msal from "@azure/msal-node";
import {
  ListUsersItemDTO,
  ListUsersPageDTO,
  UserExpandKeys,
  UserResponseDTO,
} from "./dto/response/user.response.dto";

// Fields requested from the directory listing. Derived from the keys of
// ListUsersItemDTO via `satisfies`, so the query and the return type share one
// source of truth: a typo or a field removed from the DTO is a compile error
// here, never a silent mismatch between what we ask for and what we type.
const LIST_USERS_SELECT_FIELDS = [
  "id",
  "displayName",
  "givenName",
  "surname",
  "mail",
  "userPrincipalName",
  "jobTitle",
  "accountEnabled",
  "userType",
] as const satisfies readonly (keyof ListUsersItemDTO)[];

const LIST_USERS_SELECT = LIST_USERS_SELECT_FIELDS.join(",");

// Graph caps $top at 999 for /users; clamp so a caller can't trigger a 400.
const LIST_USERS_MIN_PAGE_SIZE = 1;
const LIST_USERS_MAX_PAGE_SIZE = 999;
const LIST_USERS_DEFAULT_PAGE_SIZE = 100;

@Injectable()
export class GraphApiService {
  private readonly tenant_id: string;
  private readonly client_id: string;
  private readonly client_secret: string;
  private readonly scopes: string[];
  private readonly msal_client: msal.ConfidentialClientApplication;
  private readonly graph_client: Client;

  constructor(readonly options: GraphModuleOptions) {
    this.tenant_id = options.tenantId;
    this.client_id = options.clientId;
    this.client_secret = options.clientSecret;
    this.scopes = options.scopes?.split(" ") || [
      "https://graph.microsoft.com/.default",
    ];

    // init msal client
    this.msal_client = new msal.ConfidentialClientApplication({
      auth: {
        authority: `https://login.microsoftonline.com/${this.tenant_id}`,
        clientId: `${this.client_id}`,
        clientSecret: `${this.client_secret}`,
      },
    });
    // init graph client
    this.graph_client = Client.init({
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
    });
  }

  /**
   * @param identifier - user email or id
   */
  getUser = async <Expand extends UserExpandKeys = never>(options: {
    identifier: string;
    expand?: Expand;
  }): Promise<UserResponseDTO<Expand>> => {
    return this.graph_client
      .api(
        `https://graph.microsoft.com/v1.0/users('${options.identifier}')?$expand=${options.expand}`,
      )
      .get();
  };

  /**
   * Lists every user in the directory, following pagination to completion.
   *
   * Uses the app's own credentials (client-credentials flow), so it requires
   * the application permission `User.Read.All` granted + admin-consented on the
   * app registration — otherwise Graph returns 403.
   *
   * Returns the full set: intended for a one-off backoffice sync, not for
   * high-frequency calls. `pageSize` only controls how many round trips it
   * takes, not the result (all pages are aggregated).
   */
  listUsers = async (options?: {
    /**
     * Users fetched per Graph round trip. Does NOT change the result — every
     * page is aggregated and the full directory is returned regardless. Only
     * affects how many requests it takes. Clamped to Graph's 1–999 range;
     * defaults to 100.
     */
    pageSize?: number;
  }): Promise<ListUsersItemDTO[]> => {
    const pageSize = Math.min(
      Math.max(options?.pageSize ?? LIST_USERS_DEFAULT_PAGE_SIZE, LIST_USERS_MIN_PAGE_SIZE),
      LIST_USERS_MAX_PAGE_SIZE,
    );

    const users: ListUsersItemDTO[] = [];

    let page: ListUsersPageDTO = await this.graph_client
      .api(
        `https://graph.microsoft.com/v1.0/users?$select=${LIST_USERS_SELECT}&$top=${pageSize}`,
      )
      .get();
    users.push(...(page.value ?? []));

    // `@odata.nextLink` is an absolute URL that already carries $select/$top and
    // the skip token — pass it through unchanged until the directory runs out.
    while (page["@odata.nextLink"]) {
      page = await this.graph_client.api(page["@odata.nextLink"]).get();
      users.push(...(page.value ?? []));
    }

    return users;
  };

  /**
   * @param identifier - user email or id
   */
  getUserProfilePhoto = async (options: {
    identifier: string;
  }): Promise<Blob> => {
    const profilePhoto = await this.graph_client
      .api(
        `https://graph.microsoft.com/v1.0/users('${options.identifier}')/photo/$value`,
      )
      .get();

    return profilePhoto;
  };

  /**
   *
   * @param options
   */
  getListItem = async (options: {
    siteId: string;
    listId: string;
    itemId: string;
  }) => {
    return this.graph_client
      .api(
        `https://graph.microsoft.com/v1.0/sites/${options.siteId}/lists/${options.listId}/items/${options.itemId}?expand=fields`,
      )
      .get();
  };

  /**
   *
   * @param options
   */
  getListColumns = async <T>(options: { siteId: string; listId: string }) => {
    return this.graph_client
      .api(
        `https://graph.microsoft.com/v1.0/sites/${options.siteId}/lists/${options.listId}/columns`,
      )
      .get();
  };

  /**
   *
   * @param options
   */
  createListItemFile = async (options: {
    siteId: string;
    driveId: string;
    itemId: string;
    fileName: string;
    file: Buffer;
  }) => {
    return this.graph_client
      .api(
        `https://graph.microsoft.com/v1.0/sites/${options.siteId}/drives/${options.driveId}/items/${options.itemId}:/${options.fileName}:/content`,
      )
      .put(options.file);
  };

  /**
   *
   * @param options
   */
  createListItem = async (options: {
    siteId: string;
    listId: string;
    body: { fields: Record<string, any> };
  }) => {
    return this.graph_client
      .api(
        `https://graph.microsoft.com/v1.0/sites/${options.siteId}/lists/${options.listId}/items`,
      )
      .post(options.body);
  };

  getSites = async () => {
    return this.graph_client
      .api(`https://graph.microsoft.com/v1.0/sites`)
      .get();
  };
  getSite = async (input: { siteId: string }) => {
    return this.graph_client
      .api(`https://graph.microsoft.com/v1.0/sites/${input.siteId}`)
      .get();
  };

  listSiteDrives = async (input: { siteId: string }) => {
    return this.graph_client
      .api(`https://graph.microsoft.com/v1.0/sites/${input.siteId}/drives`)
      .get();
  };
  listSiteDriveChildrens = async (input: {
    siteId: string;
    driveId: string;
  }) => {
    return this.graph_client
      .api(
        `https://graph.microsoft.com/v1.0/sites/${input.siteId}/drives/${input.driveId}/items`,
      )
      .get();
  };
}
