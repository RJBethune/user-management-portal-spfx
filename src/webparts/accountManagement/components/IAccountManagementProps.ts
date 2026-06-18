import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IListConfig } from '../services/AccountManagementService';

export interface IAccountManagementProps {
  description: string;
  userDisplayName: string;
  context: WebPartContext;
  buildVersion: string;
  listConfig: IListConfig;
  /** Comma-separated office names or Group IDs to show on this page; blank = all authorized. */
  visibleOffices: string;
  pollTimeoutMs: number;
}
