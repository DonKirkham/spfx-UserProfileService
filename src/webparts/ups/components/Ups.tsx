import * as React from 'react';
import styles from './Ups.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { List } from "office-ui-fabric-react/lib/List";
import { DetailsList, DetailsListLayoutMode, SelectionMode } from "office-ui-fabric-react/lib/DetailsList";

import { IUserProperty, UserProfileService, UserProfileServiceMock } from "../../../services";

export interface IUpsProps {
}

export interface IUpsState {
  userProperties: IUserProperty[];
  webpartState: any;
}
enum WebpartState {
  loading,
  showMyProperties,
  showUpdateProperties,
  showOtherProperties
}

export default class Ups extends React.Component<IUpsProps, IUpsState> {
  private _ups: any;
  
  constructor (props: IUpsProps) {
    super(props);
    switch (Environment.type) {
      case EnvironmentType.SharePoint: {
        this._ups = new UserProfileService();
        break;
      }
      default: {
        this._ups = new UserProfileServiceMock();
      }
    }
		this.state = {
      userProperties: [],
      webpartState: WebpartState.loading
		};

  } 

  public async componentDidMount() {
    const userProperties = await this._ups.GetUserProfileProperties();
    this.setState({
      userProperties: userProperties,
      webpartState: WebpartState.showMyProperties
    });
  }

  public render(): React.ReactElement<IUpsProps> {
    const items = this.state.userProperties;
    return (
      <div className={ styles.ups }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>User Profile Service Demo</span>
              <div>
              {/* { this.state.webpartState != WebpartState.loading {return <div></div> : */}

              <p className={ styles.subTitle }>My Properties</p>
              <p className={ styles.description }>
                <button className={ styles.button }>
                  <span className={ styles.label }>Show My Properties</span>
                </button>
                <button className={ styles.button }>
                  <span className={ styles.label }>Update A Custom Properties</span>
                </button>
              </p>
              <div style={ {width: '100%', display: 'block'} } >
              <DetailsList  
                layoutMode={DetailsListLayoutMode.fixedColumns}
                selectionMode={SelectionMode.none}
                items={ items }
                //setKey="property"
                columns={[
                  { key: "property", name: "Property", fieldName: "property", minWidth: 20, maxWidth: 200 },
                  { key: "value", name: "Value", fieldName: "value", minWidth: 200, maxWidth:1000 }
                ]}
              />
              </div>
              <p className={ styles.subTitle }>Other User Properties</p>
              <p className={ styles.description }>
                <button className={ styles.button }>
                  <span className={ styles.label }>Show User Properties</span>
                </button>
              </p>
            </div>
            </div>
          </div>
        </div>
      </div>
    );
    // if (this.state.userProperties) {
    //     this._renderList(this.state.userProperties);
    // }
  }

  // private _renderList(items: IUserProperty[]): string {
  //   let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;">';
  //   html += `<th>Property</th><th>Value</th>`;
  //   items.forEach((item: IUserProperty) => {
  //     html += `
  //         <tr>
  //         <td>${item.property}</td>
  //         <td>${item.value}</td>
  //         </tr>
  //         `;
  //   });
  //   html += `</table>`;
  //   return html;
  // }
}
