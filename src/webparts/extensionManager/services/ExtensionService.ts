/**
 * ExtensionService
 */
import {
    SPHttpClient,
    SPHttpClientResponse
} from "@microsoft/sp-http";
import {
    IWebPartContext
} from "@microsoft/sp-webpart-base";
import { IExtensionService } from "./IExtensionService";
import { IUserCustomAction } from "./IUserCustomAction";
import { IUserCustomActionCollection } from "./IUserCustomActionCollection";
import { sp, Web } from '@pnp/sp';

export class ExtensionService implements IExtensionService {
    constructor(private context: IWebPartContext) {
        //
    }

    public async getExtensions(): Promise<IUserCustomAction[]> {
        const webUrl: string = this.context.pageContext.web.absoluteUrl;
        const spWeb = new Web(webUrl);
        return spWeb.userCustomActions.get();
    }

    public async getExtensionById(id) : Promise<IUserCustomAction>{
        const webUrl: string = this.context.pageContext.web.absoluteUrl;
        const spWeb = new Web(webUrl);
        return spWeb.userCustomActions.getById(id).get();
    }

    public async editExtension(id, customActionProps): Promise<any> {
        const webUrl: string = this.context.pageContext.web.absoluteUrl;
        const spWeb = new Web(webUrl);
        return spWeb.userCustomActions.getById(id).update(customActionProps);
    }

    public async deleteExtension(id):Promise<any>{
        const webUrl: string = this.context.pageContext.web.absoluteUrl;
        const spWeb = new Web(webUrl);
        return spWeb.userCustomActions.getById(id).delete();
    }

    public async getExtensionsByUrl(url: string): Promise<IUserCustomAction[]> {
        const apiUrl: string = `${url}/_api/web/UserCustomActions`;

        try {
            // get tasks
            return await this.context.spHttpClient.get(
                apiUrl,
                SPHttpClient.configurations.v1)
                .then((data: SPHttpClientResponse) => data.json())
                .then((data: IUserCustomActionCollection) => {
                    return data.value;
                });
        } catch (error) {
            console.error("Error loading extensions: ", error);
        }
    }
}