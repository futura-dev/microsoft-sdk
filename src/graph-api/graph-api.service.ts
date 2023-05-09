import {Injectable} from '@nestjs/common';
import {GraphModuleOptions} from "./graph-api.module";
import {Client} from '@microsoft/microsoft-graph-client'
import * as msal from "@azure/msal-node";

@Injectable()
export class GraphApiService {

    private readonly tenant_id: string;
    private readonly client_id: string;
    private readonly client_secret: string;
    private readonly scopes: string;
    private readonly msal_client: msal.ConfidentialClientApplication;
    private readonly graph_client: Client;

    constructor(readonly options: GraphModuleOptions) {
        this.tenant_id = options.tenantId;
        this.client_id = options.clientId;
        this.client_secret = options.clientSecret;
        this.scopes = options.scopes;
        // init msal client
        this.msal_client = new msal.ConfidentialClientApplication({
            auth: {
                knownAuthorities: [
                    `https://login.microsoftonline.com/${this.tenant_id}`,
                ],
                clientId: `${this.client_id}`,
                clientSecret: `${this.client_secret}`,
            },
        });
        // init graph client
        this.graph_client = Client.init({
            authProvider: async (resolve) => {
                this.msal_client.acquireTokenByClientCredential({
                    authority: `https://login.microsoftonline.com/${this.tenant_id}`,
                    scopes: this.scopes?.split(" ") || ["https://graph.microsoft.com/.default"],
                })
                    .then((token) => resolve(null, token.accessToken))
                    .catch(error => resolve(error, null))
            }
        })
    }

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
            .api(`https://graph.microsoft.com/beta/sites/${options.siteId}/lists/${options.listId}/items/${options.itemId}?expand=fields`)
            .get()
    };

    /**
     *
     * @param options
     */
    getListColumns = async <T>(options: {
        siteId: string
        listId: string
    }) => {
        return this.graph_client
            .api(`https://graph.microsoft.com/beta/sites/${options.siteId}/lists/${options.listId}/columns`)
            .get();
    }

    /**
     *
     * @param options
     */
    createListItemFile = async (options: {
        siteId: string,
        driveId: string,
        itemId: string,
        fileName: string
        file: Buffer
    }) => {
        return this.graph_client
            .api(`https://graph.microsoft.com/beta/sites/${options.siteId}/drives/${options.driveId}/items/${options.itemId}:/${options.fileName}:/content`)
            .put(options.file);
    }

    /**
     *
     * @param options
     */
    createListItem = async (options: {
        siteId: string,
        listId: string,
        body: { fields: Record<string, any> }
    }) => {
        return this.graph_client
            .api(`https://graph.microsoft.com/beta/sites/${options.siteId}/lists/${options.listId}/items`)
            .post(options.body)
    }


}
