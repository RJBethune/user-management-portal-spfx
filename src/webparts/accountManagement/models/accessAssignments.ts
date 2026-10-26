export const DEFAULT_ACCESS_ADMIN_GROUP = '365AccessManagement-Admins';

export interface IManagedGroupReference {
  id: number;
  title: string;
  groupId: string;
  isTeamsConnected: boolean;
}
export interface IAccessAssignment {
  id: number;
  userId?: number;
  userName: string;
  userEmail: string;
  groups: IManagedGroupReference[];
  missingGroupIds: number[];
}
export interface IAccessOverview {
  assignments: IAccessAssignment[];
  groups: IManagedGroupReference[];
  unassigned: IManagedGroupReference[];
  duplicateUserIds: number[];
  listUrl: string;
  canAdd: boolean;
}
export interface IAssignmentEditor {
  id?: number;
  etag?: string;
  userId?: number;
  userName: string;
  userEmail: string;
  groupIds: number[];
  groups: IManagedGroupReference[];
  multiple: boolean;
  required: boolean;
}

/** Both historic single lookups and current multi-value REST representations remain supported. */
export function lookupIds(value: unknown): number[] {
  if (!value) return [];
  const record = value as { results?: unknown[]; Id?: number };
  const entries = Array.isArray(value) ? value : Array.isArray(record.results) ? record.results : [value];
  return Array.from(new Set(entries.map((entry: { Id?: number }) => entry && entry.Id)
    .filter((id): id is number => typeof id === 'number' && Number.isInteger(id) && id > 0)));
}

export function summarizeAssignments(assignments: IAccessAssignment[], groups: IManagedGroupReference[]): {
  unassigned: IManagedGroupReference[]; duplicateUserIds: number[];
} {
  const assigned = new Set<number>();
  const counts = new Map<number, number>();
  assignments.forEach((assignment) => {
    if (!assignment.userId) return;
    assignment.groups.forEach((group) => assigned.add(group.id));
    if (assignment.userId) counts.set(assignment.userId, (counts.get(assignment.userId) || 0) + 1);
  });
  return {
    unassigned: groups.filter((group) => !assigned.has(group.id)),
    duplicateUserIds: Array.from(counts.keys()).filter((id) => (counts.get(id) || 0) > 1)
  };
}
