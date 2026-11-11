import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/profiles";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";

export interface IUserProperty {
  property: string;
  value: string;
}

export interface ISiteUser {
  /** Claims login name, e.g. `i:0#.f|membership|user@contoso.com` */
  loginName: string;
  /** Display name, or "First Last" when the user has no display name */
  name: string;
  email: string;
}

export interface IUserProfileService {
  /** Profile properties for `loginName` (claims format), or the current user when omitted. */
  GetUserProfileProperties(loginName?: string, forceRefresh?: boolean): Promise<IUserProperty[]>;
  /** Real (non-system) users known to the current site, sorted by name. */
  GetSiteUsers(): Promise<ISiteUser[]>;
}

const CURRENT_USER = "";

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

const byProperty = (a: IUserProperty, b: IUserProperty): number => (a.property > b.property ? 1 : -1);

export class UserProfileService implements IUserProfileService {
  private _sp: SPFI;
  private _profiles = new Map<string, IUserProperty[]>();

  constructor(context: WebPartContext) {
    // PnPjs v4: no global `sp`; build an SPFI instance bound to the web part context.
    this._sp = spfi().using(SPFx(context));
  }

  public async GetUserProfileProperties(loginName: string = CURRENT_USER, forceRefresh?: boolean): Promise<IUserProperty[]> {
    const cached = this._profiles.get(loginName);
    if (cached && !forceRefresh) {
      return cached;
    }

    // v4: both are invokable directly (no `.get()`).
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const profile: any = loginName === CURRENT_USER
      ? await this._sp.profiles.myProperties()
      : await this._sp.profiles.getPropertiesFor(loginName);

    const properties: IUserProperty[] = AADProperties.map(property => ({
      property: `AAD-${property}`,
      value: profile[property] as string
    }));

    // Users without a profile come back without UserProfileProperties.
    (profile.UserProfileProperties ?? []).forEach((prop: { Key: string; Value: string }) => {
      properties.push({ property: prop.Key, value: prop.Value });
    });

    properties.sort(byProperty);
    this._profiles.set(loginName, properties);
    return properties;
  }

  public async GetSiteUsers(): Promise<ISiteUser[]> {
    // The User Information List carries FirstName/LastName, which siteUsers does not.
    const items: { Name: string; Title: string; FirstName: string; LastName: string; EMail: string }[] =
      await this._sp.web.siteUserInfoList.items
        .select("Name", "Title", "FirstName", "LastName", "EMail")
        .filter("startswith(Name, 'i:0#.f|membership|')")
        .top(5000)();

    return items
      .map(item => ({
        loginName: item.Name,
        name: item.Title?.trim() || [item.FirstName, item.LastName].filter(Boolean).join(" ") || item.EMail || item.Name,
        email: item.EMail ?? ""
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }
}

export class UserProfileServiceMock implements IUserProfileService {
  public async GetUserProfileProperties(loginName: string = CURRENT_USER): Promise<IUserProperty[]> {
    const user = loginName || "current user";
    return [
      { property: "AccountName", value: user },
      { property: "Property1", value: `Value1 for ${user}` },
      { property: "Property2", value: `Value2 for ${user}` },
      { property: "Property3", value: `Value3 for ${user}` }
    ];
  }

  public async GetSiteUsers(): Promise<ISiteUser[]> {
    return [
      { loginName: "i:0#.f|membership|adele@contoso.com", name: "Adele Vance", email: "adele@contoso.com" },
      { loginName: "i:0#.f|membership|alex@contoso.com", name: "Alex Wilber", email: "alex@contoso.com" },
      { loginName: "i:0#.f|membership|megan@contoso.com", name: "Megan Bowen", email: "megan@contoso.com" }
    ];
  }
}
