import * as React from 'react';
import styles from './Ups.module.scss';
import type { IUpsProps } from './IUpsProps';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { getTheme } from '@fluentui/react/lib/Styling';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import {
  ISiteUser,
  IUserProperty,
  IUserProfileService,
  UserProfileService,
  UserProfileServiceMock
} from '../../../services';

// Longest a column may grow before its text is truncated with an ellipsis.
const MAX_PROPERTY_WIDTH = 300;
const MAX_VALUE_WIDTH = 800;

const measureCanvas = document.createElement('canvas');

/** Rendered pixel width of the widest string, in the given CSS font. */
function measureMaxWidth(texts: string[], font: string): number {
  const ctx = measureCanvas.getContext('2d');
  if (!ctx) {
    return 0;
  }
  ctx.font = font;
  return texts.reduce((max, text) => Math.max(max, ctx.measureText(text).width), 0);
}

/** Columns sized to their content, so the list (and the web part) is only as wide as needed. */
function buildColumns(items: IUserProperty[]): IColumn[] {
  const { fonts } = getTheme();
  const cellFont = `${fonts.small.fontSize} ${fonts.small.fontFamily}`;
  const headerFont = `600 ${fonts.medium.fontSize} ${fonts.medium.fontFamily}`;

  const fit = (header: string, values: string[], max: number): number =>
    Math.min(max, Math.ceil(Math.max(
      measureMaxWidth([header], headerFont),
      measureMaxWidth(values, cellFont)
    )) + 4);

  const propertyWidth = fit('Property', items.map(i => i.property ?? ''), MAX_PROPERTY_WIDTH);
  const valueWidth = fit('Value', items.map(i => String(i.value ?? '')), MAX_VALUE_WIDTH);

  return [
    { key: 'property', name: 'Property', fieldName: 'property', minWidth: propertyWidth, maxWidth: propertyWidth },
    { key: 'value', name: 'Value', fieldName: 'value', minWidth: valueWidth, maxWidth: valueWidth }
  ];
}

export interface IUpsState {
  users: ISiteUser[];
  selectedLoginName: string;
  userProperties: IUserProperty[];
  loading: boolean;
}

export default class Ups extends React.Component<IUpsProps, IUpsState> {
  private _ups: IUserProfileService;
  private _currentUser: ISiteUser;
  private _requestedLoginName = '';

  constructor(props: IUpsProps) {
    super(props);

    // Use the live profile service in SharePoint, the mock everywhere else (local workbench).
    this._ups = Environment.type === EnvironmentType.SharePoint
      ? new UserProfileService(props.context)
      : new UserProfileServiceMock();

    const { user } = props.context.pageContext;
    this._currentUser = {
      loginName: `i:0#.f|membership|${user.loginName.toLowerCase()}`,
      name: user.displayName,
      email: user.email
    };

    this.state = {
      users: [this._currentUser],
      selectedLoginName: this._currentUser.loginName,
      userProperties: [],
      loading: true
    };
  }

  public async componentDidMount(): Promise<void> {
    const [users] = await Promise.all([
      this._ups.GetSiteUsers(),
      this._loadProperties(this._currentUser.loginName)
    ]);

    // Make sure the current user is always selectable, even if the site list misses them.
    const current = this._currentUser.loginName;
    if (!users.some(u => u.loginName.toLowerCase() === current)) {
      users.push(this._currentUser);
      users.sort((a, b) => a.name.localeCompare(b.name));
    }
    this.setState({ users });
  }

  public render(): React.ReactElement<IUpsProps> {
    const { users, selectedLoginName, userProperties, loading } = this.state;
    const selected = users.find(u => u.loginName.toLowerCase() === selectedLoginName) ?? this._currentUser;
    const options: IDropdownOption[] = users.map(u => ({ key: u.loginName.toLowerCase(), text: u.name, data: u }));

    return (
      <div className={styles.ups}>
        <div className={styles.container}>
          <span className={styles.title}>User Profile Explorer</span>
          <Dropdown
            label="User"
            className={styles.userPicker}
            options={options}
            selectedKey={selectedLoginName}
            onChange={this._onUserChanged}
          />
          <div className={styles.subTitle}>Properties for {selected.name} ({selected.email})</div>
          {loading
            ? <p>Loading profile properties&hellip;</p>
            : <DetailsList
                // Remount per user and render every row, so the list sizes to exactly the
                // selected user's properties instead of keeping the previous user's height.
                key={selectedLoginName}
                onShouldVirtualize={() => false}
                // Drop the header's built-in 16px top padding so the list sits right under the subtitle.
                styles={{ root: { marginTop: 0 }, headerWrapper: { selectors: { '.ms-DetailsHeader': { paddingTop: 0 } } } }}
                layoutMode={DetailsListLayoutMode.fixedColumns}
                selectionMode={SelectionMode.none}
                items={userProperties}
                columns={buildColumns(userProperties)}
              />
          }
        </div>
      </div>
    );
  }

  private _onUserChanged = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      this._loadProperties(option.key as string).catch(console.error);
    }
  };

  private async _loadProperties(loginName: string): Promise<void> {
    this._requestedLoginName = loginName;
    this.setState({ selectedLoginName: loginName, loading: true });

    // The current user goes through myProperties; everyone else through getPropertiesFor.
    const userProperties = loginName === this._currentUser.loginName
      ? await this._ups.GetUserProfileProperties()
      : await this._ups.GetUserProfileProperties(loginName);

    // Ignore a slow response for a user who is no longer selected.
    if (loginName === this._requestedLoginName) {
      this.setState({ userProperties, loading: false });
    }
  }
}
