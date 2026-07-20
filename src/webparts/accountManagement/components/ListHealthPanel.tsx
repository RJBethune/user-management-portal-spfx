import * as React from 'react';
import { MessageBar, MessageBarBody, MessageBarTitle } from '@fluentui/react-components';

import { EXPECTED_LISTS, IColumnIssue, IExpectedColumn, IExpectedList, IListHealth } from '../models/listSchema';
import styles from './AccountManagement.module.scss';

export interface IListHealthPanelProps {
  health: IListHealth[];
  /** True while the page is being edited — the only audience shown schema plumbing. */
  isEditMode: boolean;
}

const specFor = (key: string): IExpectedList | undefined =>
  EXPECTED_LISTS.filter((s: IExpectedList) => s.key === key)[0];

const issueText = (i: IColumnIssue): string => {
  if (i.kind === 'type') {
    return `is a "${i.actual}" column but must be ${i.expected}. Delete it and re-create it with the correct type.`;
  }
  if (i.lookalike) {
    return `is missing — but this list has "${i.lookalike}", which is the same name created with a space in it. SharePoint freezes the internal name at creation, so re-create the column as ${i.column} (no spaces).`;
  }
  return `is missing. Add it as ${i.expected}.`;
};

/**
 * Self-diagnosing setup guidance.
 *  - A missing REQUIRED list replaces the old "not found" error with a real setup panel.
 *  - Column problems surface only while the page is in edit mode; ordinary viewers never
 *    see plumbing warnings.
 */
export const ListHealthPanel: React.FunctionComponent<IListHealthPanelProps> = (props: IListHealthPanelProps) => {
  const missingRequired: IListHealth[] = props.health.filter((h: IListHealth) => !h.exists && !h.optional && !h.error);
  const withIssues: IListHealth[] = props.health.filter((h: IListHealth) => h.exists && h.issues.length > 0);
  const missingOptional: IListHealth[] = props.health.filter((h: IListHealth) => !h.exists && h.optional && !h.error);

  if (missingRequired.length > 0) {
    return (
      <div className={styles.setupPanel}>
        <MessageBar intent="error" layout="multiline">
          <MessageBarBody>
            <MessageBarTitle>
              Setup needed — {missingRequired.length === 1 ? 'a required list was not found' : 'required lists were not found'}
            </MessageBarTitle>
            This web part reads its data from SharePoint lists. Create the list(s) below with the exact
            title shown, then add each column using its <strong>exact internal name</strong>.
          </MessageBarBody>
        </MessageBar>

        {missingRequired.map((h: IListHealth) => {
          const spec: IExpectedList | undefined = specFor(h.key);
          return (
            <div className={styles.setupList} key={h.key}>
              <h3>
                List: <code>{h.title}</code>
              </h3>
              {spec && <p className={styles.emptyText}>{spec.purpose}</p>}
              {spec && (
                <table className={styles.setupTable}>
                  <thead>
                    <tr>
                      <th>Column (internal name)</th>
                      <th>Type</th>
                      <th>Required?</th>
                    </tr>
                  </thead>
                  <tbody>
                    {spec.columns.map((c: IExpectedColumn) => (
                      <tr key={c.name}>
                        <td>
                          <code>{c.name}</code>
                        </td>
                        <td>{c.label}</td>
                        <td>{c.optional ? 'Optional' : 'Required'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          );
        })}

        <ol className={styles.setupSteps}>
          <li>
            On the site that holds the lists, choose <strong>New &rarr; List</strong> and name it exactly as shown
            above.
          </li>
          <li>
            For each column choose <strong>Add column</strong>, pick the type, and type the name{' '}
            <strong>with no spaces</strong>. SharePoint freezes a column&rsquo;s internal name at creation and turns
            a space into <code>_x0020_</code>, which the app will not find.
          </li>
          <li>Rename the display label afterwards if you want a friendlier caption — that does not change the internal name.</li>
          <li>
            If the lists live on a different site, set <strong>List site URL</strong> in this web part&rsquo;s property
            pane, under <em>Data source (lists)</em>.
          </li>
        </ol>
      </div>
    );
  }

  // Plumbing warnings are for editors only.
  if (!props.isEditMode || (withIssues.length === 0 && missingOptional.length === 0)) {
    return null;
  }

  return (
    <div className={styles.setupPanel}>
      <MessageBar intent="warning" layout="multiline">
        <MessageBarBody>
          <MessageBarTitle>List schema needs attention — shown only while you are editing this page</MessageBarTitle>
          {withIssues.map((h: IListHealth) => (
            <div className={styles.schemaIssueGroup} key={h.key}>
              <strong>{h.title}</strong>
              <ul>
                {h.issues.map((i: IColumnIssue) => (
                  <li key={i.column}>
                    <code>{i.column}</code> {issueText(i)}
                  </li>
                ))}
              </ul>
            </div>
          ))}
          {missingOptional.map((h: IListHealth) => {
            const spec: IExpectedList | undefined = specFor(h.key);
            return (
              <div className={styles.schemaIssueGroup} key={h.key}>
                <strong>{h.title}</strong> was not found (optional).{' '}
                {spec ? spec.purpose : 'The dependent feature is switched off.'}
              </div>
            );
          })}
        </MessageBarBody>
      </MessageBar>
    </div>
  );
};

export default ListHealthPanel;
