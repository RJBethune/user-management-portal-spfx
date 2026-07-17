export type MembershipAction = 'Add Member' | 'Remove Member' | 'Add Owner' | 'Remove Owner';

export interface IUser {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
  /** True when this member is itself a group (security/O365/SharePoint group) nested in a SharePoint site group. */
  isGroup?: boolean;
}

export interface IOfficeGroup {
  /** SharePoint list item id of the row in the Group list. */
  id: number;
  title: string;
  /** O365 group GUID (PROD) or SharePoint site-group id (DEV mock). */
  groupId: string;
  description?: string;
  mail?: string;
  mailNickname?: string;
  visibility?: string;
  createdDateTime?: string;
  isTeamsConnected?: boolean;
  siteUrl?: string;
  siteTitle?: string;
  authorizationItemId?: number;
  /** Optional Group-list column used to flag unmanageable types (dynamic, distribution, etc.). */
  groupType?: string;
}

export interface IRequestSummary {
  id: number;
  action: string;
  groupName: string;
  groupId: string;
  memberDisplayName: string;
  status: string;
  resultMessage?: string;
  requestedOn?: string;
  requesterName?: string;
  /** SharePoint Modified — when the request row was last edited (e.g. the flow's write-back). */
  modified?: string;
}

export interface IMembershipRequest {
  id: number;
  status: string;
  resultMessage?: string;
  authorizationChecked?: boolean;
  authorizationResult?: string;
}

export interface IMembershipChangeInput {
  action: MembershipAction;
  group: IOfficeGroup;
  member: IUser;
  justification?: string;
}

export interface ICurrentUser {
  Id: number;
  Title?: string;
  Email?: string;
  UserPrincipalName?: string;
}

/** A SharePoint site (and permission level) a group is used on — curated, from the Group Site Permissions list. */
export interface ISitePermission {
  siteName: string;
  siteUrl?: string;
  permission: string;
}
