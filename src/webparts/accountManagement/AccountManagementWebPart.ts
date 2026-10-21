import { buildInfo } from '../../delivery/buildInfo';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base';
import { createSpfxUiHost, SpfxUiHost, SpfxUiHostProvider } from '@spfx-kit/ui-profile';
import { MessageBar, MessageBarBody } from '../../ui/Controls';
import { initialTokens } from '../../ui/theme';
import '../../ui/styles';

import * as strings from 'AccountManagementWebPartStrings';
import { AccountManagementShell } from './components/AccountManagementShell';
import { IListConfig, DEFAULT_LIST_CONFIG } from './services/AccountManagementService';
import { toMessage } from './shared/errors';
import { setVerbose } from './shared/log';

export interface IAccountManagementWebPartProps {
  description: string;
  introText: string;
  helpText: string;
  helpUrl: string;
  requestListTitle: string;
  groupListTitle: string;
  authorizedAdminsListTitle: string;
  sitePermissionsListTitle: string;
  listSiteUrl: string;
  visibleOffices: string;
  pollTimeoutSeconds: number;
  startCollapsed: boolean;
  requireJustification: boolean;
  showGroupPhotos: boolean;
  verboseLogging: boolean;
}

interface IErrorBoundaryProps {
  children?: React.ReactNode;
}

interface IErrorBoundaryState {
  errorMessage?: string;
}

class ErrorBoundary extends React.Component<IErrorBoundaryProps, IErrorBoundaryState> {
  public constructor(props: IErrorBoundaryProps) {
    super(props);
    this.state = {};
  }

  public static getDerivedStateFromError(error: unknown): IErrorBoundaryState {
    return { errorMessage: toMessage(error, 'An unknown error occurred while loading 365 Account Management.') };
  }

  public componentDidCatch(error: unknown): void {
    console.error('365 Account Management web part error', error);
  }

  public render(): React.ReactNode {
    if (this.state.errorMessage) {
      return React.createElement(
        MessageBar,
        { intent: 'error' },
        React.createElement(MessageBarBody, {}, this.state.errorMessage)
      );
    }
    return this.props.children;
  }
}

export default class AccountManagementWebPart extends BaseClientSideWebPart<IAccountManagementWebPartProps> {
  public static readonly buildVersion: string = buildInfo.version;

  private _theme: IReadonlyTheme | undefined;
  private _themeProvider: ThemeProvider | undefined;
  private _uiHost: SpfxUiHost | undefined;
  private _windowErrorHandler: ((e: ErrorEvent) => void) | undefined;
  private _unhandledRejectionHandler: ((e: PromiseRejectionEvent) => void) | undefined;

  protected onInit(): Promise<void> {
    const themeProvider: ThemeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._themeProvider = themeProvider;
    this._theme = themeProvider.tryGetTheme();
    themeProvider.themeChangedEvent.add(this, this._onThemeChanged);
    this._uiHost = createSpfxUiHost({
      mountPoint: this.domElement,
      portalParent: this.domElement.ownerDocument.body,
      targetDocument: this.domElement.ownerDocument,
      instanceId: this.instanceId,
      style: 'luma',
      theme: initialTokens
    });
    this._uiHost.portalHost.style.position = 'relative';
    this._uiHost.portalHost.style.zIndex = '1000';
    if (this.properties.showGroupPhotos === undefined) {
      this.properties.showGroupPhotos = true; // default ON; preserves prior behavior for existing instances
    }
    this._registerDiagnostics();
    console.info(`365 Account Management web part ${AccountManagementWebPart.buildVersion} loaded`, {
      componentId: this.context.manifest.id,
      siteUrl: this.context.pageContext.web.absoluteUrl
    });
    return Promise.resolve();
  }

  public render(): void {
    try {
      if (!this._uiHost) return;
      setVerbose(!!this.properties.verboseLogging);
      const listConfig: IListConfig = {
        requestListTitle: this.properties.requestListTitle || DEFAULT_LIST_CONFIG.requestListTitle,
        groupListTitle: this.properties.groupListTitle || DEFAULT_LIST_CONFIG.groupListTitle,
        authorizedAdminsListTitle:
          this.properties.authorizedAdminsListTitle || DEFAULT_LIST_CONFIG.authorizedAdminsListTitle,
        sitePermissionsListTitle:
          this.properties.sitePermissionsListTitle || DEFAULT_LIST_CONFIG.sitePermissionsListTitle,
        listSiteUrl: this.properties.listSiteUrl || ''
      };
      const seconds: number =
        this.properties.pollTimeoutSeconds && this.properties.pollTimeoutSeconds > 0
          ? this.properties.pollTimeoutSeconds
          : 120;
      const element: React.ReactElement = React.createElement(
        SpfxUiHostProvider,
        { host: this._uiHost },
        React.createElement(ErrorBoundary, {}, React.createElement(AccountManagementShell, {
          host: this._uiHost,
          description: this.properties.description,
          userDisplayName: this.context.pageContext.user.displayName,
          context: this.context,
          buildVersion: AccountManagementWebPart.buildVersion,
          listConfig: listConfig,
          visibleOffices: this.properties.visibleOffices || '',
          pollTimeoutMs: seconds * 1000,
          introText: this.properties.introText || '',
          helpText: this.properties.helpText || '',
          helpUrl: this.properties.helpUrl || '',
          startCollapsed: !!this.properties.startCollapsed,
          sectionTheme: this._theme,
          requireJustification: !!this.properties.requireJustification,
          showGroupPhotos: this.properties.showGroupPhotos !== false,
          // Schema warnings are shown only while the page is being edited — the audience who can fix them.
          isEditMode: this.displayMode === DisplayMode.Edit
        }))
      );
      ReactDom.render(element, this._uiHost.appRoot);
    } catch (e) {
      this._renderFallbackError(e);
    }
  }

  protected onDispose(): void {
    this._unregisterDiagnostics();
    if (this._themeProvider) this._themeProvider.themeChangedEvent.remove(this, this._onThemeChanged);
    if (this._uiHost) {
      ReactDom.unmountComponentAtNode(this._uiHost.appRoot);
      this._uiHost.dispose();
      this._uiHost = undefined;
    }
  }

  private _onThemeChanged(args: ThemeChangedEventArgs): void {
    this._theme = args.theme;
    this.render();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: 'Display',
              groupFields: [
                PropertyPaneTextField('description', { label: 'Title' }),
                PropertyPaneTextField('introText', {
                  label: 'Intro / help text',
                  description: 'Shown under the heading. Blank = the default line. Tailor it per team/page.',
                  multiline: true,
                  rows: 3
                }),
                PropertyPaneTextField('helpText', {
                  label: 'Support line text',
                  description: 'e.g. "Questions? Contact the IT Service Desk." Blank = hidden.'
                }),
                PropertyPaneTextField('helpUrl', {
                  label: 'Support link URL (optional)',
                  description: 'If set, the support line becomes a link (e.g. a mailto: or ticket URL).'
                })
              ]
            },
            {
              groupName: 'Offices on this page',
              groupFields: [
                PropertyPaneTextField('visibleOffices', {
                  label: 'Offices to show',
                  description:
                    'Comma-separated office names or Group IDs. Blank = every office you are authorized to manage. Authorization is still enforced by the Authorized Admins list — listing an office you are not authorized for has no effect.',
                  multiline: true,
                  rows: 3
                })
              ]
            },
            {
              groupName: 'Data source (lists)',
              groupFields: [
                PropertyPaneLabel('listSchemaHelp', {
                  text:
                    'While this page is in edit mode, the web part itself reports any missing list or mistyped column — check the canvas for a setup panel or schema banner.'
                }),
                PropertyPaneTextField('listSiteUrl', {
                  label: 'List site URL (optional)',
                  description:
                    'Absolute URL of the site holding the four lists (e.g. the site-collection root). Blank = the current site. Set this to run the web part on any page while reading one central set of lists. Must be in this tenant.'
                }),
                PropertyPaneTextField('groupListTitle', { label: 'Group (offices) list title' }),
                PropertyPaneTextField('authorizedAdminsListTitle', { label: 'Authorized Admins list title' }),
                PropertyPaneTextField('requestListTitle', {
                  label: 'Request list title',
                  description:
                    'O365 (PROD) requests are written here. Point different pages at different request lists to route them to different Power Automate flows.'
                }),
                PropertyPaneTextField('sitePermissionsListTitle', {
                  label: 'Site permissions list title',
                  description:
                    'Optional. Curated list of which SharePoint sites (and permission level) each group is used on.'
                })
              ]
            },
            {
              groupName: 'Behavior',
              groupFields: [
                PropertyPaneSlider('pollTimeoutSeconds', {
                  label: 'Status wait timeout (seconds, O365 flow)',
                  min: 30,
                  max: 600,
                  step: 10,
                  value: 120
                }),
                PropertyPaneToggle('startCollapsed', {
                  label: 'Start office cards collapsed',
                  onText: 'Collapsed',
                  offText: 'First office expanded'
                }),
                PropertyPaneToggle('showGroupPhotos', {
                  label: 'Show Microsoft 365 group photos',
                  onText: 'On',
                  offText: 'Off (use initials tile)'
                }),
                PropertyPaneToggle('requireJustification', {
                  label: 'Require a reason for each change',
                  onText: 'Required',
                  offText: 'Optional'
                }),
                PropertyPaneToggle('verboseLogging', {
                  label: 'Verbose diagnostic logging (browser console)',
                  onText: 'On',
                  offText: 'Off'
                })
              ]
            },
            { groupName: 'About', groupFields: [PropertyPaneLabel('deliveryAttribution', { text: `Account Management ${buildInfo.version} · ${buildInfo.channel.toUpperCase()} — M/EX - Executive Office` })] }
          ]
        }
      ]
    };
  }

  private _renderFallbackError(e: unknown): void {
    const target = this._uiHost ? this._uiHost.appRoot : this.domElement;
    ReactDom.unmountComponentAtNode(target);
    const div: HTMLDivElement = target.ownerDocument.createElement('div');
    div.setAttribute('role', 'alert');
    div.style.color = 'var(--spfx-ui-color-destructive, inherit)';
    div.style.padding = '16px';
    div.textContent = toMessage(e, '365 Account Management failed to load.');
    while (target.firstChild) {
      target.removeChild(target.firstChild);
    }
    target.appendChild(div);
  }

  private _registerDiagnostics(): void {
    this._windowErrorHandler = (e: ErrorEvent): void => {
      console.error('365 Account Management diagnostic window error', {
        message: toMessage(e.error || e.message, 'Window error'),
        filename: e.filename,
        lineno: e.lineno,
        colno: e.colno,
        error: e.error
      });
    };
    this._unhandledRejectionHandler = (e: PromiseRejectionEvent): void => {
      // Some third-party controls (notably the @pnp LivePersona profile card) reject promises
      // with no reason while prefetching profile data. These are benign and non-actionable, so
      // swallow them entirely (preventDefault keeps the browser from logging them too) and only
      // surface rejections that actually carry a reason.
      if (e.reason === undefined || e.reason === null) {
        e.preventDefault();
        return;
      }
      console.error('365 Account Management diagnostic unhandled rejection', {
        message: toMessage(e.reason, 'Unhandled promise rejection'),
        reason: e.reason
      });
    };
    window.addEventListener('error', this._windowErrorHandler);
    window.addEventListener('unhandledrejection', this._unhandledRejectionHandler);
  }

  private _unregisterDiagnostics(): void {
    if (this._windowErrorHandler) {
      window.removeEventListener('error', this._windowErrorHandler);
      this._windowErrorHandler = undefined;
    }
    if (this._unhandledRejectionHandler) {
      window.removeEventListener('unhandledrejection', this._unhandledRejectionHandler);
      this._unhandledRejectionHandler = undefined;
    }
  }
}
