export interface BaseUserResponseDTO {
    "id": string,
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users/$entity",
    "businessPhones": Array<string>,
    "displayName": string,
    "givenName": string,
    "jobTitle": string,
    "mail": string,
    "mobilePhone": string,
    "officeLocation": string | null,
    "preferredLanguage": string,
    "surname": string,
    "userPrincipalName": string
}

export type UserExpandKeys = keyof UserResponseExpands;
export type UserResponseExpands = {
    manager?: {
        id: string,
        displayName: string;
        givenName: string;
        jobTitle: string;
        mail: string;
    }
}

export type UserResponseDTO<ExpandKeys extends UserExpandKeys = never> =
            ExpandKeys extends UserExpandKeys
                ? BaseUserResponseDTO & Pick<UserResponseExpands, ExpandKeys>
                : BaseUserResponseDTO
