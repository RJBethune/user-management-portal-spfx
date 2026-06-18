import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MessageBar, MessageBarType } from '@fluentui/react';

import * as strings from 'AccountManagementWebPartStrings';
import AccountManagement from './components/AccountManagement';
import { IListConfig, DEFAULT_LIST_CONFIG } from './services/AccountManagementService';
import { toMessage } from './shared/errors';
import { setVerbose } from './shared/log';

export interface IAccountManagementWebPartProps {
  description: string;
  requestListTitle: string;
  groupListTitle: string;
  authorizedAdminsListTitle: string;
  visibleOffices: string;
  pollTimeoutSeconds: number;
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
      return React.createElement(MessageBar, { messageBarType: MessageBarType.error }, this.state.errorMessage);
    }
    return this.props.children;
  }
}

export default class AccountManagementWebPart extends BaseClientSideWebPart<IAccountManagementWebPartProps> {
  public static readonly buildVersion: string = '1.0.7';

  private _windowErrorHandler: ((e: ErrorEvent) => void) | undefined;
  private _unhandledRejectionHandler: ((e: PromiseRejectionEvent) => void) | undefined;

  protected onInit(): Promise<void> {
    this._registerDiagnostics();
    console.info(`365 Account Management web part ${AccountManagementWebPart.buildVersion} loaded`, {
      componentId: this.context.manifest.id,
      siteUrl: this.context.pageContext.web.absoluteUrl
    });
    return Promise.resolve();
  }

  public render(): void {
    try {
      setVerbose(!!this.properties.verboseLogging);
      const listConfig: IListConfig = {
        requestListTitle: this.properties.requestListTitle || DEFAULT_LIST_CONFIG.requestListTitle,
        groupListTitle: this.properties.groupListTitle || DEFAULT_LIST_CONFIG.groupListTitle,
        authorizedAdminsListTitle:
          this.properties.authorizedAdminsListTitle || DEFAULT_LIST_CONFIG.authorizedAdminsListTitle
      };
      const seconds: number =
        this.properties.pollTimeoutSeconds && this.properties.pollTimeoutSeconds > 0
          ? this.properties.pollTimeoutSeconds
          : 120;
      const element: React.ReactElement = React.createElement(
        ErrorBoundary,
        {},
        React.createElement(AccountManagement, {
          description: this.properties.description,
          userDisplayName: this.context.pageContext.user.displayName,
          context: this.context,
          buildVersion: AccountManagementWebPart.buildVersion,
          listConfig: listConfig,
          visibleOffices: this.properties.visibleOffices || '',
          pollTimeoutMs: seconds * 1000
        })
      );
      ReactDom.render(element, this.domElement);
    } catch (e) {
      this._renderFallbackError(e);
    }
  }

  protected onDispose(): void {
    this._unregisterDiagnostics();
    ReactDom.unmountComponentAtNode(this.domElement);
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
              groupFields: [PropertyPaneTextField('description', { label: 'Title' })]
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
                PropertyPaneTextField('groupListTitle', { label: 'Group (offices) list title' }),
                PropertyPaneTextField('authorizedAdminsListTitle', { label: 'Authorized Admins list title' }),
                PropertyPaneTextField('requestListTitle', {
                  label: 'Request list title',
                  description:
                    'O365 (PROD) requests are written here. Point different pages at different request lists to route them to different Power Automate flows.'
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
                PropertyPaneToggle('verboseLogging', {
                  label: 'Verbose diagnostic logging (browser console)',
                  onText: 'On',
                  offText: 'Off'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private _renderFallbackError(e: unknown): void {
    const div: HTMLDivElement = document.createElement('div');
    div.style.color = '#a4262c';
    div.style.padding = '16px';
    div.textContent = toMessage(e, '365 Account Management failed to load.');
    while (this.domElement.firstChild) {
      this.domElement.removeChild(this.domElement.firstChild);
    }
    this.domElement.appendChild(div);
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
