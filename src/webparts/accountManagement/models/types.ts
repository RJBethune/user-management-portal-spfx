export type MembershipAction = 'Add Member' | 'Remove Member';

export interface IUser {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
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
}

export interface ICurrentUser {
  Id: number;
  Title?: string;
  Email?: string;
  UserPrincipalName?: string;
}
