export const SPHttpClient = { configurations: { v1: {} } };
export const Guid = { newGuid: () => ({ toString: () => 'preview-correlation' }) };
export class SPPermission {
  static manageLists = 2048;
  static editListItems = 4;
  constructor(private mask: { Low: number; High: number }) {}
  hasPermission(value: number): boolean { return (this.mask.Low & value) === value; }
}
