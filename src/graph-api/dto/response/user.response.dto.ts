export interface BaseUserResponseDTO {
  id: string;
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users/$entity";
  businessPhones: string[];
  displayName: string;
  givenName: string;
  jobTitle: string;
  mail: string;
  mobilePhone: string;
  officeLocation: string | null;
  preferredLanguage: string;
  surname: string;
  userPrincipalName: string;
}

export type UserExpandKeys = keyof UserResponseExpands;
export type UserResponseExpands = {
  manager?: {
    id: string;
    displayName: string;
    givenName: string;
    jobTitle: string;
    mail: string;
  };
};

export type UserResponseDTO<ExpandKeys extends UserExpandKeys = never> =
  ExpandKeys extends UserExpandKeys
    ? BaseUserResponseDTO & Pick<UserResponseExpands, ExpandKeys>
    : BaseUserResponseDTO;

// Azure AD userType. Documented values are "Member" and "Guest"; the
// `(string & {})` keeps those two as autocomplete suggestions while still
// accepting any string, so an unforeseen value never breaks type-checking at
// the call site (a plain `| string` would erase the suggestions entirely).
export type MicrosoftUserType = "Member" | "Guest" | (string & {});

// Single user as returned by the directory listing endpoint (GET /users).
// Narrower than BaseUserResponseDTO: exactly the fields listUsers $selects,
// nothing more. The `| null` fields are the ones the directory does not
// guarantee per account (e.g. a mailbox-less account has no `mail`, a service
// account no given/surname); `id` / `userPrincipalName` / `accountEnabled` are
// always present, so they are non-null.
//
// The field set here is the single source of truth: listUsers derives its
// `$select` from `keyof ListUsersItemDTO`, so this interface and the query can
// never drift apart.
export interface ListUsersItemDTO {
  id: string;
  displayName: string | null;
  givenName: string | null;
  surname: string | null;
  mail: string | null;
  userPrincipalName: string;
  jobTitle: string | null;
  accountEnabled: boolean;
  userType: MicrosoftUserType | null;
}

// Raw shape of a single page from Graph's GET /users. `@odata.nextLink`, when
// present, is the absolute URL of the next page — followed until absent to
// return the full directory. Internal to listUsers; not part of its return.
export interface ListUsersPageDTO {
  value: ListUsersItemDTO[];
  "@odata.nextLink"?: string;
}
