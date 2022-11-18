import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/attachments";
import "@pnp/sp/files";

export default class Service {

    public mysitecontext: any;

    public constructor(siteUrl: string, Sitecontext: any) {
        this.mysitecontext = Sitecontext;

        sp.setup({
            sp: {
                baseUrl: siteUrl

            },
        });

    }
    
    public async getItemByID(ItemID:any): Promise<any> {
        try {
            const selectedList = 'BulkApprovalDetatilsList';
            const Item: any[] = await sp.web.lists.getByTitle(selectedList).items
                .select("*,Title,ApproverName,ApproverID,WF_ReviewDueDate,FieldValuesAsText/WF_ReviewDueDate").expand("FieldValuesAsText")
                .filter("ID eq '" + ItemID + "'")
                .get();
            return Item[0];
        } catch (error) {
            console.log(error);
        }
    }
    public async getUserByLogin(LoginName: string): Promise<any> {

        try {

            const user = await sp.web.siteUsers.getByLoginName(LoginName).get();

            return user;

        } catch (error) {

            console.log(error);

        }

    }

    public async getItemByIDs(ItemID: any): Promise<any> {

        try {



    const selectedList = 'BulkApprovalDetatilsList';

    const Item: any[] = await sp.web.lists.getByTitle(selectedList).items.select("*,ApproverName").filter("ID eq '" + ItemID + "'").get();

            return Item[0];

        } catch (error) {

            console.log(error);

        }

    }

    public async EmailTemplate(): Promise<any> {

        try {


let EMailbody="REVIEW INITIATED"
    const selectedList = 'EmailFormat';

    const Item: any[] = await sp.web.lists.getByTitle(selectedList).items.select("*,Title,Body").filter("Title eq '" + EMailbody + "'").get();

            return Item[0];

        } catch (error) {

            console.log(error);

        }

    }
 
    public async FinanceandHRMSEmailTemplate(): Promise<any> {

        try {


let EMailbody="HRMSANDFINANCE"
    const selectedList = 'EmailFormat';

    const Item: any[] = await sp.web.lists.getByTitle(selectedList).items.select("*,Title,Body,Mynewcc").filter("Title eq '" + EMailbody + "'").get();

            return Item[0];

        } catch (error) {

            console.log(error);

        }

    }
    public async getNewApproverid(NewApproverName:any): Promise<any> {

        try {



    const selectedList = 'PeopleSoft';

    const Item: any[] = await sp.web.lists.getByTitle(selectedList).items.select("*,EMPID").filter("EmailID eq '" + NewApproverName + "'").get();

            return Item[0];

        } catch (error) {

            console.log(error);

        }

    }
 


    

    public async getCurrentUser(): Promise<any> {

        try {

            return await sp.web.currentUser.get().then(result => {

                return result;

            });

        } catch (error) {

            console.log(error);

        }

      }



















}
