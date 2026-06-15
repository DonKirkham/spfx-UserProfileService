import * as React from 'react';
import styles from './Ups.module.scss';
import type { IUpsProps } from './IUpsProps';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { DetailsList, DetailsListLayoutMode, SelectionMode } from '@fluentui/react/lib/DetailsList';
import {
  IUserProperty,
  IUserProfileService,
  UserProfileService,
  UserProfileServiceMock
} from '../../../services';

export interface IUpsState {
  userProperties: IUserProperty[];
  loading: boolean;
}

export default class Ups extends React.Component<IUpsProps, IUpsState> {
  private _ups: IUserProfileService;

  constructor(props: IUpsProps) {
    super(props);

    // Use the live profile service in SharePoint, the mock everywhere else (local workbench).
    this._ups = Environment.type === EnvironmentType.SharePoint
      ? new UserProfileService(props.context)
      : new UserProfileServiceMock();

    this.state = {
      userProperties: [],
      loading: true
    };
  }

  public async componentDidMount(): Promise<void> {
    const userProperties = await this._ups.GetUserProfileProperties();
    this.setState({ userProperties, loading: false });
  }

  public render(): React.ReactElement<IUpsProps> {
    const { userProperties, loading } = this.state;

    return (
      <div className={styles.ups}>
        <div className={styles.container}>
          <span className={styles.title}>User Profile Service Demo</span>
          <p className={styles.subTitle}>My Properties</p>
          {loading
            ? <p>Loading profile properties&hellip;</p>
            : <DetailsList
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
}
