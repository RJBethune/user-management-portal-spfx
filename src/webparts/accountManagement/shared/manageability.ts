import { IOfficeGroup } from '../models/types';

export interface IManageability {
  manageable: boolean;
  /** Short reason shown on the card when not manageable. */
  reason?: string;
}

const GUID_RE: RegExp = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Decide whether a group row can actually be managed by this tool, so the card can
 * disable Add/Remove up front instead of failing after the admin has done work.
 * Driven by GroupId validity + the optional GroupType column on the Group list.
 */
export function getManageability(group: IOfficeGroup): IManageability {
  const id: string = (group.groupId || '').trim();
  if (!id) {
    return { manageable: false, reason: 'This office has no Group ID configured. Contact an administrator.' };
  }
  const isInteger: boolean = /^\d+$/.test(id);
  const isGuid: boolean = GUID_RE.test(id);
  if (!isInteger && !isGuid) {
    return {
      manageable: false,
      reason: 'This office’s Group ID is malformed (must be an Entra group GUID or a SharePoint group id). Contact an administrator.'
    };
  }

  const type: string = (group.groupType || '').trim().toLowerCase();
  const unmanageable: { [key: string]: string } = {
    dynamic: 'Membership of this group is rule-based (dynamic) and cannot be edited here.',
    'mailenabledsecurity': 'This is a mail-enabled security group — manage it in the Exchange admin center.',
    distribution: 'This is a distribution list — manage it in the Exchange admin center.',
    roleassignable: 'This is a role-assignable group and cannot be edited here.',
    onprem: 'This group is synced from on-premises and must be managed on-premises.',
    onpremises: 'This group is synced from on-premises and must be managed on-premises.'
  };
  const key: string = type.replace(/[\s_-]/g, '');
  if (unmanageable[key]) {
    return { manageable: false, reason: unmanageable[key] };
  }
  return { manageable: true };
}
