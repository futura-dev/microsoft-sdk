type MicrosoftGraphIdentitySet = {
  id: string;
  email: string;
  displayName: string;
};

export type ListWebhook = {
  value: {
    subscriptionId: string;
    clientState: null;
    expirationDateTime: string;
    resource: string;
    tenantId: string;
    siteUrl: string;
    webId: string;
  }[];
};

export type ListLog = {
  "odata.type": string;
  "odata.id": string;
  "odata.editLink": string;
  ChangeToken: {
    StringValue: `${number};${number};${string};${string};${string}`;
  };
  ChangeType: number;
  SiteId: string;
  Time: string;
  Editor: string;
  EditorEmailHint: null | unknown;
  ItemId: number;
  ListId: string;
  ServerRelativeUrl: string;
  SharedByUser: null | unknown;
  SharedWithUsers: null | unknown;
  UniqueId: string;
  WebId: string;
};

export type ListItem<T> = {
  "@odata.context": string;
  "@odata.etag": string;
  createdDateTime: string;
  eTag: string;
  id: string;
  name: string;
  lastModifiedDateTime: string;
  description: string;
  webUrl: string;
  createdBy: {
    user: MicrosoftGraphIdentitySet;
  };
  lastModifiedBy: {
    user: MicrosoftGraphIdentitySet;
  };
  parentReference: {
    id: string;
    siteId: string;
  };
  contentType: {
    id: string;
    name: string;
  };
  "fields@odata.context": string;
  fields: T;
};
export type HRListItem = ListItem<{
  "@odata.etag": string;
  Title: string;
  Attachments: boolean;
  LinkTitle: string;
  Position: string;
  Progress: string;
  RecruiterLookupId: string;
  Email: string;
  Telefono: string;
  ApplicationDate: string;
  PhoneScreenDate: string;
  PhoneScreenerLookupId: string;
  InterviewDate: string;
  Interviewers: {
    LookupId: number;
    LookupValue: string;
    Email: string;
  };
  Notes: string;
  Profilo_x0020_LinkedIn: string;
  Curriculum_x0020_URL: string;
  id: string;
  ContentType: string;
  Modified: string;
  Created: string;
  AuthorLookupId: string;
  EditorLookupId: string;
  _UIVersionString: string;
  Edit: string;
  LinkTitleNoMenu: string;
  ItemChildCount: string;
  FolderChildCount: string;
  _ComplianceFlags: string;
  _ComplianceTag: string;
  _ComplianceTagWrittenTime: string;

  _ComplianceTagUserId: string;
}>;

export type ColumnValue = {
  columnGroup: string;
  description: string;
  displayName: string;
  enforceUniqueValues: boolean;
  hidden: boolean;
  id: string;
  indexed: boolean;
  name: string;
  readOnly: boolean;
  required: boolean;
  geolocation?: Record<string, never>;
  boolean?: Record<string, never>;
  calculated?: {
    format: string;
    formula: string;
    outputType: "boolean" | "currency" | "dateTime" | "number" | "text";
  };
  choice?: {
    allowTextEntry: boolean;
    choices: string[];
    displayAs: "checkBoxes" | "dropDownMenu" | "radioButtons";
  };
  currency?: { locale: string };
  dateTime?: { displayAs: "default" | "friendly" | "standard"; format: string };
  lookup?: {
    allowMultipleValues: boolean;
    allowUnlimitedLength: boolean;
    columnName: string;
    listId: string;
    primaryLookupColumnId: string;
  };
  number?: {
    decimalPlaces: string;
    displayAs: "number" | "percentage";
    maximum: number;
    minimum: number;
  };
  personOrGroup?: {
    allowMultipleSelection: boolean;
    displayAs: string;
    chooseFromType: "peopleAndGroups" | "peopleOnly";
  };
  text?: {
    allowMultipleLines: boolean;
    appendChangesToExistingText: boolean;
    linesForEditing: number;
    maxLength: number;
    textType: "plain" | "richText";
  };
};
export type GetColumnsResponse = {
  "@odata.context": string;
  value: ColumnValue[];
};
