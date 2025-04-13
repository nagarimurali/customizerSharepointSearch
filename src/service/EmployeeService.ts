/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import { SPFI } from "@pnp/sp";
import { getSP } from "../pnpjs-config";
import "@pnp/sp/webs";
import "@pnp/sp/sites";
import "@pnp/sp/lists";
import "@pnp/sp/items";

class EmployeeService {

    private sp: SPFI;
    private context: FormCustomizerContext

    public async init(context: FormCustomizerContext) {
        if (!!context) {
            this.sp = getSP(context);
            this.context = context;
        }
        return this.sp;
    }
    public async getBaseLineContentTypeId() {
        if ('list' in this.context) {
            const cType = await this.sp.web.lists.getByTitle(this.context.list.title)();
            console.log(cType);


        }
    }
    public async createListItem<T extends Record<string, unknown>>(listTitle: string, object: T) {
        const i = await this.sp.web.lists.getByTitle(listTitle).items.add(object);
        return i;
    }
    public async updateListItem(listTitle: string, object: any, itemId: number) {
        const i = await this.sp.web.lists.getByTitle(listTitle).items.getById(itemId).update(object);
        return i;
    }

    // public static async updateListItem(listName: string, itemId: number, data: any): Promise<any> {
    //     try {
    //         const sp = getSP(); // Ensure SP context is initialized
    //         const result = await sp.web.lists.getByTitle(listName).items.getById(itemId).update(data);
    //         return result;
    //     } catch (error) {
    //         console.error("Error updating list item:", error);
    //         throw error;
    //     }
    // }
}
export default new EmployeeService();