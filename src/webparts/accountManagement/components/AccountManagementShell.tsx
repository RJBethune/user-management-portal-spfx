import * as React from 'react';
import { SpfxUiHost, useSpfxUiId } from '@spfx-kit/ui-profile';
import { AppHeader } from '@spfx-kit/ui-patterns/app-header';
import { Button } from '@spfx-kit/ui-profile/button';
import { DropdownMenu, DropdownMenuTrigger, DropdownMenuContent, DropdownMenuRadioGroup, DropdownMenuRadioItem } from '@spfx-kit/ui-profile/dropdown-menu';
import { applyAppTheme, readThemeChoice, ThemeChoice } from '../../../ui/theme';
import { AccessAdministrationService } from '../services/AccessAdministrationService';
import { IAccountManagementProps } from './IAccountManagementProps';
import { AccessAssignments } from './AccessAssignments';
import AccountManagement from './AccountManagement';
import styles from './AccountManagement.module.scss';

export function AccountManagementShell(props: IAccountManagementProps & { host: SpfxUiHost }): JSX.Element {
  const preferenceKey = `account-management:theme:${props.context.pageContext.site.id}:${props.context.pageContext.user.loginName}`;
  const [choice, setChoice] = React.useState<ThemeChoice>(() => readThemeChoice(preferenceKey));
  const [canAdminister, setCanAdminister] = React.useState(false);
  const [view, setView] = React.useState<'membership' | 'access'>('membership');
  const [themeMenuOpen, setThemeMenuOpen] = React.useState(false);
  const configKey = JSON.stringify(props.listConfig);
  const service = React.useMemo(() => new AccessAdministrationService(props.context, props.listConfig), [props.context, configKey]);
  const menuId = useSpfxUiId('theme-menu');
  React.useEffect(() => {
    let current = true;
    setCanAdminister(false);
    setView('membership');
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
    <AppHeader brand={{ name: props.description || '365 Account Management' }} user={{ name: props.userDisplayName }} actions={
      <DropdownMenu open={themeMenuOpen} onOpenChange={setThemeMenuOpen}><DropdownMenuTrigger render={<Button variant="ghost" />} aria-label="Choose theme">Theme</DropdownMenuTrigger>
        <DropdownMenuContent id={menuId}>
          <DropdownMenuRadioGroup value={choice} onValueChange={selectTheme}>
            <DropdownMenuRadioItem value="sharepoint">SharePoint</DropdownMenuRadioItem>
            <DropdownMenuRadioItem value="light">Light</DropdownMenuRadioItem>
            <DropdownMenuRadioItem value="dark">Dark</DropdownMenuRadioItem>
          </DropdownMenuRadioGroup>
        </DropdownMenuContent>
      </DropdownMenu>
    } />
    {canAdminister && <nav className={styles.viewNavigation} aria-label="Account management views">
      <Button variant={view === 'membership' ? 'secondary' : 'ghost'} aria-pressed={view === 'membership'} onClick={() => setView('membership')}>Membership</Button>
      <Button variant={view === 'access' ? 'secondary' : 'ghost'} aria-pressed={view === 'access'} onClick={() => setView('access')}>Access assignments</Button>
    </nav>}
    <div hidden={view !== 'membership'}><AccountManagement {...props} /></div>
    {view === 'access' && canAdminister && <AccessAssignments service={service} />}
  </div>;
}
