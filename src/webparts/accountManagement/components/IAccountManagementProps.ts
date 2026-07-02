import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
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
  /** Optional per-page intro/help text shown under the heading (blank = default line). */
  introText: string;
  /** Optional "need help?" line; helpUrl makes helpText a link. */
  helpText: string;
  helpUrl: string;
  /** Start every group card collapsed (no auto-expand of the first office). */
  startCollapsed: boolean;
  /** Require a justification note before Add/Remove (captured into the request item). */
  requireJustification: boolean;
  /** Load and show real M365 group photos (lazily, on expand); off = always use the initials tile. */
  showGroupPhotos: boolean;
  /** SharePoint section theme, mapped onto Fluent v9 tokens (supports dark/inverted sections). */
  sectionTheme?: IReadonlyTheme;
}
