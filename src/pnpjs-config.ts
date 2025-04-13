/* eslint-disable @typescript-eslint/no-non-null-assertion */
/* eslint-disable eqeqeq */
/* eslint-disable no-var */
/* eslint-disable @rushstack/no-new-null */
import { WebPartContext } from "@microsoft/sp-webpart-base";

// import pnp and pnp logging system
import { spfi, SPFI, SPFx } from "@pnp/sp";
// import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";

let _sp: SPFI | null = null;

export const getSP = (context: WebPartContext | FormCustomizerContext): SPFI => {
    if (!_sp) {
        _sp = spfi().using(SPFx(context));
    }
    return _sp;
};