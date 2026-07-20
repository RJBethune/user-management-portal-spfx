/**
 * Expected SharePoint schema for the four lists this web part reads.
 *
 * Single source of truth for BOTH the runtime health check and the on-canvas setup
 * panel, so the guidance shown to an admin can never drift from what the app queries.
 * Mirrors docs/SharePoint-Lists-Setup.md.
 */

export type ListKey = 'groups' | 'requests' | 'admins' | 'sitePerms';

export interface IExpectedColumn {
  /** Internal name the app queries by (must match exactly — SharePoint freezes it at creation). */
  name: string;
  /** Acceptable SharePoint TypeAsString values; the first is the documented one. */
  types: string[];
  /** Friendly type shown in the setup table. */
  label: string;
  /** Optional columns never raise an issue when absent (the app retries without them). */
  optional?: boolean;
}

export interface IExpectedList {
  key: ListKey;
  /** Default list title (the property pane can override it per page). */
  defaultTitle: string;
  /** Property-pane field that sets this list's title. */
  paneField: string;
  /** One line on what stops working without it. */
  purpose: string;
  /** Optional lists degrade silently — the dependent feature just switches off. */
  optional?: boolean;
  columns: IExpectedColumn[];
}

export interface IColumnIssue {
  column: string;
  /** Friendly expected type, for the message. */
  expected: string;
  /** Actual TypeAsString when the column exists but is the wrong type. */
  actual?: string;
  kind: 'missing' | 'type';
  /** A field that looks like the expected one under a different internal name (the space trap). */
  lookalike?: string;
}

export interface IListHealth {
  key: ListKey;
  /** The configured title actually probed. */
  title: string;
  optional: boolean;
  exists: boolean;
  /** Set when the schema could not be read at all (permissions, network). */
  error?: string;
  issues: IColumnIssue[];
}

const TEXT: string[] = ['Text'];
const NOTE: string[] = ['Note'];
const CHOICE: string[] = ['Choice'];
const DATE: string[] = ['DateTime'];
const YESNO: string[] = ['Boolean'];
// Hyperlink columns report as URL; the docs allow plain text as well.
const LINK: string[] = ['URL', 'Text'];
const PERSON: string[] = ['User'];
const LOOKUP: string[] = ['Lookup', 'LookupMulti'];
const LOOKUP_MULTI: string[] = ['LookupMulti', 'Lookup'];

export const EXPECTED_LISTS: IExpectedList[] = [
  {
    key: 'groups',
    defaultTitle: 'Managed Groups',
    paneField: 'Group (offices) list title',
    purpose: 'The groups this tool manages. Without it there is nothing to show.',
    columns: [
      { name: 'GroupId', types: TEXT, label: 'Single line of text' },
      { name: 'Description', types: NOTE, label: 'Multiple lines of text' },
      { name: 'Mail', types: TEXT, label: 'Single line of text' },
      { name: 'MailNickname', types: TEXT, label: 'Single line of text' },
      { name: 'Visibility', types: TEXT, label: 'Single line of text' },
      { name: 'CreatedDateTime', types: DATE, label: 'Date and Time' },
      { name: 'IsTeamsConnected', types: YESNO, label: 'Yes/No' },
      { name: 'SiteUrl', types: LINK, label: 'Hyperlink (or text)' },
      { name: 'SiteTitle', types: TEXT, label: 'Single line of text' },
      { name: 'GroupType', types: TEXT, label: 'Single line of text', optional: true }
    ]
  },
  {
    key: 'admins',
    defaultTitle: 'Group Management Authorized Admins',
    paneField: 'Authorized Admins list title',
    purpose: 'Who may manage which groups. Without it nothing is manageable.',
    columns: [
      { name: 'User', types: PERSON, label: 'Person or Group (single)' },
      { name: 'OfficeGroupRecord', types: LOOKUP_MULTI, label: 'Lookup → Managed Groups (allow multiple)' }
    ]
  },
  {
    key: 'requests',
    defaultTitle: 'Group Membership Requests',
    paneField: 'Request list title',
    purpose: 'Queue + audit log for Microsoft 365 group changes. Without it those changes cannot be filed.',
    columns: [
      { name: 'Action', types: CHOICE, label: 'Choice (Add/Remove Member, Add/Remove Owner)' },
      { name: 'GroupId', types: TEXT, label: 'Single line of text' },
      { name: 'GroupName', types: TEXT, label: 'Single line of text' },
      { name: 'MemberId', types: TEXT, label: 'Single line of text' },
      { name: 'MemberDisplayName', types: TEXT, label: 'Single line of text' },
      { name: 'Status', types: CHOICE, label: 'Choice (Pending, Completed, Failed)' },
      { name: 'ResultMessage', types: NOTE, label: 'Multiple lines of text' },
      { name: 'RequestedOn', types: DATE, label: 'Date and Time' },
      { name: 'TargetUserPrincipalName', types: TEXT, label: 'Single line of text' },
      { name: 'TargetUserEmail', types: TEXT, label: 'Single line of text' },
      { name: 'CorrelationId', types: TEXT, label: 'Single line of text' },
      { name: 'OfficeGroupRecord', types: LOOKUP, label: 'Lookup → Managed Groups' },
      { name: 'Justification', types: NOTE, label: 'Multiple lines of text' },
      { name: 'MemberEntraId', types: PERSON, label: 'Person or Group' },
      { name: 'RequestedBy', types: PERSON, label: 'Person or Group', optional: true },
      { name: 'AuthorizationChecked', types: YESNO, label: 'Yes/No' },
      { name: 'AuthorizationResult', types: NOTE, label: 'Multiple lines of text' }
    ]
  },
  {
    key: 'sitePerms',
    defaultTitle: 'Group Site Permissions',
    paneField: 'Site permissions list title',
    optional: true,
    purpose: 'Feeds the "Used On" tab. Optional — without it that tab simply does not appear.',
    columns: [
      { name: 'GroupId', types: TEXT, label: 'Single line of text' },
      { name: 'SiteName', types: TEXT, label: 'Single line of text' },
      { name: 'SiteUrl', types: LINK, label: 'Hyperlink (or text)' },
      { name: 'Permission', types: TEXT, label: 'Single line of text' }
    ]
  }
];

/**
 * The classic trap: a column first created with a space in its name is frozen with an
 * escaped internal name (e.g. "Group Id" -> `Group_x0020_Id`), so the app's query for
 * `GroupId` finds nothing. Given the actual fields, return the look-alike that explains it.
 */
export function findLookalike(expected: string, actualInternalNames: string[]): string | undefined {
  const normalize = (s: string): string => s.replace(/_x0020_/g, '').replace(/[^a-z0-9]/gi, '').toLowerCase();
  const target: string = normalize(expected);
  return actualInternalNames.filter((n: string) => n !== expected).filter((n: string) => normalize(n) === target)[0];
}
