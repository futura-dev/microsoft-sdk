import { GraphApiService } from "./graph-api/graph-api.service";
import { GraphApiModule } from "./graph-api/graph-api.module";
import { SharepointApiService } from "./sharepoint-api/sharepoint-api.service";
import { SharepointApiModule } from "./sharepoint-api/sharepoint-api.module";
import {
  DriveItem,
  DriveItemPermission,
  DriveItemPermissionsResponse,
  DriveRecipient,
  InviteDriveItemPermissionsResponse,
  ListItem,
  ListLog,
  ListWebhook,
  HRListItem,
} from "./sharepoint-api/types";

export {
  GraphApiService,
  GraphApiModule,
  SharepointApiService,
  SharepointApiModule,
};

export type {
  DriveItem,
  DriveItemPermission,
  DriveItemPermissionsResponse,
  DriveRecipient,
  InviteDriveItemPermissionsResponse,
  ListWebhook,
  ListLog,
  ListItem,
  HRListItem,
};
