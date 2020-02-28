import { sp } from "@pnp/sp";
import "@pnp/sp/profiles";
import { stringIsNullOrEmpty } from "@pnp/common";

export interface IUserProperty {
    property: string;
    value: string; 
}
export class UserProfileService {

    private _profile: IUserProperty[] = [];

    public async GetUserProfileProperties (forceRefresh?: boolean)  {
        if (this._profile.length == 0 || forceRefresh) {
            const profile = await sp.profiles.myProperties.get();
            // profile.forEach((prop) => {
            //     //if (typeOf(prop.Value) == "String") {
            //         this._profile.push( { property: prop.Key, value: prop.Value } );
            //     //}
            // });
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
                "UserUrl"];
            AADProperties.forEach(property => {
                this._profile.push( { property: `AAD-${property}`, value: <string>profile[property] } );
            });

            profile.UserProfileProperties.forEach((prop) => {
                this._profile.push( { property: prop.Key, value: prop.Value } );
            });
        }
        return this._profile.sort((a,b) => { return a.property > b.property ? 1 : -1; });
    }
}

export class UserProfileServiceMock {
    private _profile: IUserProperty[] = null;

    public async GetUserProfileProperties (forceRefresh?: boolean)  {
        if (this._profile == null || forceRefresh) {
            this._profile = this.MockData;
        }
        return this._profile;
    }

    private MockData: IUserProperty[] = [
        { property: "Property1", value: "Value1" },
        { property: "Property2", value: "Value2" },
        { property: "Property3", value: "Value3" },
    ];
}
