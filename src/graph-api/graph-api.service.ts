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
import { Site, SiteSearchResponse } from "./types";
import { DriveItem } from "../sharepoint-api/types";
import { fetchGraphPage, fetchGraphNextPage, GraphPage } from "./pagination";

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
   * Lists users in the directory (GET /users).
   *
   * Uses the app's own credentials (client-credentials flow), so it requires
   * the application permission `User.Read.All` granted + admin-consented on the
   * app registration — otherwise Graph returns 403.
   *
   * Paginated like every other collection method on this client: `unroll`
   * defaults to false (first page only, plus `@odata.nextLink` if there's
   * more) — pass `unroll: true` to follow every page and get the full
   * directory aggregated in one call (intended for a one-off backoffice sync,
   * not for high-frequency calls).
   */
  listUsers = async (options?: {
    /**
     * Users fetched per Graph round trip. Does NOT change the result set,
     * only how many requests it takes to get there. Clamped to Graph's
     * 1–999 range; defaults to 100.
     */
    pageSize?: number;
    unroll?: boolean;
  }): Promise<ListUsersPageDTO> => {
    const pageSize = Math.min(
      Math.max(
        options?.pageSize ?? LIST_USERS_DEFAULT_PAGE_SIZE,
        LIST_USERS_MIN_PAGE_SIZE,
      ),
      LIST_USERS_MAX_PAGE_SIZE,
    );

    return fetchGraphPage<ListUsersItemDTO>(
      this.graph_client,
      `https://graph.microsoft.com/v1.0/users?$select=${LIST_USERS_SELECT}&$top=${pageSize}`,
      options?.unroll,
    );
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

  /**
   * Searches SharePoint sites the app registration can see, via Graph's
   * `/sites?search=` — Graph rejects `/sites` with no `search` term, so an
   * empty/omitted query defaults to `*` (Graph's documented "match every
   * site" wildcard) instead of failing.
   *
   * Paginated like every other collection method on this client — see
   * `unroll` on `listUsers` above.
   */
  searchSites = async (input?: {
    query?: string;
    unroll?: boolean;
  }): Promise<SiteSearchResponse> => {
    const query = input?.query?.trim() || "*";
    return fetchGraphPage<Site>(
      this.graph_client,
      `https://graph.microsoft.com/v1.0/sites?search=${encodeURIComponent(query)}`,
      input?.unroll,
    );
  };

  /**
   * Lists every SharePoint site the app registration can see, via Graph's
   * `GET /sites/getAllSites`. Unlike `searchSites` (`/sites?search=`, served
   * from SharePoint's search index — eventually consistent, so a freshly
   * created or newly-permissioned site can be missing for a while), this reads
   * the sites directly, so new sites show up immediately. Prefer it for a
   * "pick a site" picker where staleness is confusing. Requires the same
   * `Sites.Read.All` application permission.
   *
   * Paginated like every other collection method on this client — see `unroll`
   * on `listUsers` above.
   */
  getAllSites = async (input?: {
    unroll?: boolean;
  }): Promise<SiteSearchResponse> => {
    return fetchGraphPage<Site>(
      this.graph_client,
      `https://graph.microsoft.com/v1.0/sites/getAllSites`,
      input?.unroll,
    );
  };

  /**
   * Fetches the next page of any paginated collection call made without
   * `unroll: true` (sites, users, ...) — pass the `@odata.nextLink` from the
   * previous response verbatim (it's an absolute URL that already carries
   * the original query and skip token). `T` is the item type of that
   * collection, e.g. `getNextPage<Site>(...)`.
   */
  getNextPage = async <T>(input: {
    nextLink: string;
  }): Promise<GraphPage<T>> => {
    return fetchGraphNextPage<T>(this.graph_client, input.nextLink);
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

  // Path segments are individually percent-encoded (not the whole path, so
  // literal "/" keeps separating segments) — same approach as
  // SharepointApiService.encodePath, duplicated here since it's a stateless
  // one-liner and the two services don't share a base class.
  private encodeDrivePath(path: string): string {
    return path
      .split("/")
      .filter(Boolean)
      .map((segment) => encodeURIComponent(segment))
      .join("/");
  }

  /**
   * Lists the children of a folder in a site's default document library,
   * addressed by path instead of item id — lets a caller drill down a
   * SharePoint folder tree (e.g. to pick a `docsBasePath`) knowing only
   * the site id and the path built up so far. Omit `path` (or pass "") for
   * the drive root.
   */
  getSiteDriveChildrenByPath = async (input: {
    siteId: string;
    path?: string;
  }): Promise<{ value: DriveItem[] }> => {
    const trimmedPath = input.path?.split("/").filter(Boolean).join("/");
    const url = trimmedPath
      ? `https://graph.microsoft.com/v1.0/sites/${input.siteId}/drive/root:/${this.encodeDrivePath(trimmedPath)}:/children`
      : `https://graph.microsoft.com/v1.0/sites/${input.siteId}/drive/root/children`;

    return this.graph_client.api(url).get();
  };
}
