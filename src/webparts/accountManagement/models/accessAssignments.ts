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

export interface IGroupAssignmentCoverage {
  group: IManagedGroupReference;
  assignments: IAccessAssignment[];
  adminCount: number;
  duplicateGrants: boolean;
}

/** Counts people separately from records, and never treats a missing user as an admin. */
export function groupAssignmentCoverage(assignments: IAccessAssignment[], groups: IManagedGroupReference[]): IGroupAssignmentCoverage[] {
  const byGroup = new Map<number, IAccessAssignment[]>();
  assignments.forEach(row => new Set(row.groups.map(group => group.id)).forEach(id => {
    const records = byGroup.get(id) || [];
    records.push(row); byGroup.set(id, records);
  }));
  return groups.map(group => {
    const records = byGroup.get(group.id) || [];
    const valid = records.filter(row => !!row.userId);
    const adminCount = new Set(valid.map(row => row.userId)).size;
    return { group, assignments: records, adminCount, duplicateGrants: valid.length > adminCount };
  });
}

export interface IAssignmentChange { id: number; title: string; otherRecordIds: number[] }

/** Describe only this draft's changes; other saved records remain independent grants. */
export function assignmentChanges(model: IAssignmentEditor, selected: number[], saved: IAccessAssignment[]): {
  added: IAssignmentChange[]; removed: IAssignmentChange[];
} {
  const describe = (id: number): IAssignmentChange => ({
    id,
    title: model.groups.find(group => group.id === id)?.title || `Unavailable group (record ${id})`,
    otherRecordIds: model.userId ? saved.filter(row => row.id !== model.id && row.userId === model.userId && row.groups.some(group => group.id === id)).map(row => row.id) : []
  });
  return {
    added: Array.from(new Set(selected)).filter(id => model.groupIds.indexOf(id) === -1).map(describe),
    removed: Array.from(new Set(model.groupIds)).filter(id => selected.indexOf(id) === -1).map(describe)
  };
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
