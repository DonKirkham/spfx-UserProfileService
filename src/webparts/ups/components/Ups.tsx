import * as React from 'react';
import styles from './Ups.module.scss';
import type { IUpsProps } from './IUpsProps';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { DetailsList, DetailsListLayoutMode, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import {
  ISiteUser,
  IUserProperty,
  IUserProfileService,
  UserProfileService,
  UserProfileServiceMock
} from '../../../services';

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
          <span className={styles.title}>User Profile Service Property Demo</span>
          <Dropdown
            label="User"
            className={styles.userPicker}
            options={options}
            selectedKey={selectedLoginName}
            onChange={this._onUserChanged}
          />
          <p className={styles.subTitle}>Properties for {selected.name} ({selected.email})</p>
          {loading
            ? <p>Loading profile properties&hellip;</p>
            : <DetailsList
                // Remount per user and render every row, so the list sizes to exactly the
                // selected user's properties instead of keeping the previous user's height.
                key={selectedLoginName}
                onShouldVirtualize={() => false}
                layoutMode={DetailsListLayoutMode.fixedColumns}
                selectionMode={SelectionMode.none}
                items={userProperties}
                columns={[
                  { key: 'property', name: 'Property', fieldName: 'property', minWidth: 20, maxWidth: 200 },
                  { key: 'value', name: 'Value', fieldName: 'value', minWidth: 200, maxWidth: 1000 }
                ]}
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
