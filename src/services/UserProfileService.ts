import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/profiles";

export interface IUserProperty {
  property: string;
  value: string;
}

export interface IUserProfileService {
  GetUserProfileProperties(forceRefresh?: boolean): Promise<IUserProperty[]>;
}

export class UserProfileService implements IUserProfileService {
  private _sp: SPFI;
  private _profile: IUserProperty[] = [];

  constructor(context: WebPartContext) {
    // PnPjs v4: no global `sp`; build an SPFI instance bound to the web part context.
    this._sp = spfi().using(SPFx(context));
  }

  public async GetUserProfileProperties(forceRefresh?: boolean): Promise<IUserProperty[]> {
    if (this._profile.length === 0 || forceRefresh) {
      this._profile = [];
      // v4: myProperties is invokable directly (no `.get()`).
      const profile = await this._sp.profiles.myProperties();

      const AADProperties = [
        "AccountName",
        "DirectReports",
        "DisplayName",
        "Email",
        "ExtendedManagers",
        "ExtendedReports",
        "IsFollowed",
        "LatestPost",
        "odata.metadata",
        "odata.type",
        "Peers",
        "PersonalSiteHostUrl",
        "PersonalUrl",
        "PictureUrl",
        "Title",
        "UserUrl"
      ];
      AADProperties.forEach(property => {
        this._profile.push({ property: `AAD-${property}`, value: profile[property] as string });
      });

      profile.UserProfileProperties.forEach((prop: { Key: string; Value: string }) => {
        this._profile.push({ property: prop.Key, value: prop.Value });
      });
    }
    return this._profile.sort((a, b) => (a.property > b.property ? 1 : -1));
  }
}

export class UserProfileServiceMock implements IUserProfileService {
  private _profile: IUserProperty[] | undefined = undefined;

  public async GetUserProfileProperties(forceRefresh?: boolean): Promise<IUserProperty[]> {
    if (this._profile === undefined || forceRefresh) {
      this._profile = this.MockData;
    }
    return this._profile;
  }

  private MockData: IUserProperty[] = [
    { property: "Property1", value: "Value1" },
    { property: "Property2", value: "Value2" },
    { property: "Property3", value: "Value3" }
  ];
}
