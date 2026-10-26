import * as React from 'react';
import { SpfxUiHost, useSpfxUiId } from '@spfx-kit/ui-profile';
import { AppHeader, AppHeaderExpandButton } from '@spfx-kit/ui-patterns/app-header';
import { DarkTheme20Regular, PeopleTeam20Regular, MoreHorizontal20Regular } from '@fluentui/react-icons';
import { Button } from '@spfx-kit/ui-profile/button';
import { DropdownMenu, DropdownMenuTrigger, DropdownMenuContent, DropdownMenuItem, DropdownMenuRadioGroup, DropdownMenuRadioItem } from '@spfx-kit/ui-profile/dropdown-menu';
import { applyAppTheme, readThemeChoice, ThemeChoice } from '../../../ui/theme';
import { AccessAdministrationService } from '../services/AccessAdministrationService';
import { IAccountManagementProps } from './IAccountManagementProps';
import { AccessAssignments } from './AccessAssignments';
import { Notice } from '@spfx-kit/ui-patterns/notice';
import { useExpandedWorkspace } from '../../../ui/useExpandedWorkspace';
import { RecentRequests } from './RecentRequests';
import AccountManagement from './AccountManagement';
import styles from './AccountManagement.module.scss';

export function AccountManagementShell(props: IAccountManagementProps & { host: SpfxUiHost }): JSX.Element {
  const preferenceKey = `account-management:theme:${props.context.pageContext.site.id}:${props.context.pageContext.user.loginName}`;
  const [choice, setChoice] = React.useState<ThemeChoice>(() => readThemeChoice(preferenceKey));
  const [canAdminister, setCanAdminister] = React.useState(false);
  const [view, setView] = React.useState<'membership' | 'requests' | 'access'>('membership');
  const [requestsOpened, setRequestsOpened] = React.useState(false);
  const [accessOpened, setAccessOpened] = React.useState(false);
  const [themeMenuOpen, setThemeMenuOpen] = React.useState(false);
  const configKey = JSON.stringify(props.listConfig);
  const service = React.useMemo(() => new AccessAdministrationService(props.context, props.listConfig), [props.context, configKey]);
  const menuId = useSpfxUiId('theme-menu');
  const workspaceMenuId = useSpfxUiId('workspace-menu');
  const workspace = useExpandedWorkspace(props.host, styles.expandedFrame);
  React.useEffect(() => {
    let current = true;
    setCanAdminister(false);
    setView('membership');
    setAccessOpened(false);
    service.isAvailable().then((value) => { if (current) setCanAdminister(value); }).catch(() => undefined);
    return () => { current = false; };
  }, [service]);
  React.useLayoutEffect(() => { applyAppTheme(props.host, choice, props.sectionTheme); }, [props.host, choice, props.sectionTheme]);
  const selectTheme = (value: string): void => {
    const next: ThemeChoice = value === 'light' || value === 'dark' ? value : 'sharepoint';
    setChoice(next);
    setThemeMenuOpen(false);
    try { localStorage.setItem(preferenceKey, next); } catch { /* The theme still works when storage is unavailable. */ }
  };
  return <div className={styles.shell}>
    <AppHeader className={workspace.expanded ? styles.expandedHeader : undefined} role="region" aria-label="Account Management header" brand={{ name: props.description || '365 Account Management', mark: <PeopleTeam20Regular aria-hidden="true" /> }} user={{ name: props.userDisplayName }} actions={<>
      <AppHeaderExpandButton expanded={workspace.expanded} onExpandedChange={workspace.toggle} aria-disabled={workspace.pending || undefined} data-workspace-toggle="true" />
      {workspace.expanded && <DropdownMenu><DropdownMenuTrigger render={<Button variant="ghost" size="icon" aria-label="Workspace actions" title="Workspace actions" />}><MoreHorizontal20Regular aria-hidden="true" /></DropdownMenuTrigger>
        <DropdownMenuContent id={workspaceMenuId} align="end"><DropdownMenuItem disabled={workspace.pending || workspace.native} onClick={workspace.enterBrowser}>Enter browser full screen</DropdownMenuItem></DropdownMenuContent>
      </DropdownMenu>}
      <DropdownMenu open={themeMenuOpen} onOpenChange={setThemeMenuOpen}><DropdownMenuTrigger render={<Button variant="ghost" size="icon" aria-label="Choose theme" title="Choose theme" />}><DarkTheme20Regular aria-hidden="true" /></DropdownMenuTrigger>
        <DropdownMenuContent id={menuId} align="end" style={{ width: 'max-content', whiteSpace: 'nowrap' }} finalFocus={() => {
          const document = props.host.targetDocument;
          const active = document.activeElement;
          // An exiting menu must not steal focus from a newly opened editor or confirmation.
          return !active || active === document.body || !!document.getElementById(menuId)?.contains(active);
        }}>
          <DropdownMenuRadioGroup value={choice} onValueChange={selectTheme}>
            <DropdownMenuRadioItem value="sharepoint">SharePoint</DropdownMenuRadioItem>
            <DropdownMenuRadioItem value="light">Light</DropdownMenuRadioItem>
            <DropdownMenuRadioItem value="dark">Dark</DropdownMenuRadioItem>
          </DropdownMenuRadioGroup>
        </DropdownMenuContent>
      </DropdownMenu>
    </>} />
    {workspace.error && <Notice tone="warning" onDismiss={workspace.clearError}>{workspace.error}</Notice>}
    <span className={styles.visuallyHidden} role="status">{workspace.expanded ? 'Workspace expanded. Press Escape to exit.' : 'Workspace in page.'}</span>
    <nav className={styles.viewNavigation} aria-label="Account management views">
      <Button variant={view === 'membership' ? 'secondary' : 'ghost'} aria-current={view === 'membership' ? 'page' : undefined} onClick={() => setView('membership')}>Membership</Button>
      <Button variant={view === 'requests' ? 'secondary' : 'ghost'} aria-current={view === 'requests' ? 'page' : undefined} onClick={() => { setRequestsOpened(true); setView('requests'); }}>Recent requests</Button>
      {canAdminister && <Button variant={view === 'access' ? 'secondary' : 'ghost'} aria-current={view === 'access' ? 'page' : undefined} onClick={() => { setAccessOpened(true); setView('access'); }}>Access assignments</Button>}
    </nav>
    <div hidden={view !== 'membership'}><AccountManagement {...props} /></div>
    {requestsOpened && <div hidden={view !== 'requests'}><RecentRequests {...props} active={view === 'requests'} /></div>}
    {accessOpened && canAdminister && <div hidden={view !== 'access'}><AccessAssignments key={configKey} service={service} /></div>}
  </div>;
}
