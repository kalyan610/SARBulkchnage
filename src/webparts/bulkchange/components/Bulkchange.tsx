import * as React from 'react';
import styles from './Bulkchange.module.scss';
import { IBulkchangeProps } from './IBulkchangeProps';
import Service from './Service';
import Moment from 'react-moment';

import { escape } from '@microsoft/sp-lodash-subset';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Checkbox, PrimaryButton } from 'office-ui-fabric-react';
import { ChoiceGroup, IChoiceGroupOption, textAreaProperties, Stack, IStackTokens, StackItem, IStackStyles, TextField, CheckboxVisibility, BaseButton } from 'office-ui-fabric-react';
import { sp } from "@pnp/sp";
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { stringIsNullOrEmpty, isArray, objectDefinedNotNull } from '@pnp/common/util';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Items } from 'sp-pnp-js';
import { Image } from '@microsoft/office-ui-fabric-react-bundle';
import * as moment from 'moment';
let RedirectUrl = '';
let RecordId='';
let TitleIds='';
let ListApplicationName='';
let newbody='';
let mynewcc='';
let newdate='';
let   Imagesnew:any;
export interface ISecurityBeyState {
  isLoading: boolean;
  AppliName:string;
  Appid:string;
  NewAppid:any;
  ItemInfo: any;
  ItemId:number;
  NewApproverName:string;
  TitleId:any;
  ApplicationName:any;
  EMPName:any;
  EMPids:any;
  loginNames:any;
  ToMail:any;
  SendBody:string;
  Mycc:string;
  newbody:string,
  newApplication:string,
  Quater:string,
  Year:string,
  ListApplicationName:string,
  WFReviewDuedate:string,
  EmpName:string;
  EmpidName:string;
  
 
}


export default class Bulkchange extends React.Component<IBulkchangeProps, ISecurityBeyState> {

  public _service: Service;
  protected ppl: PeoplePicker;
  public constructor(props: IBulkchangeProps) {
    super(props);
    this.state = {
      isLoading: true,
      AppliName:"",
      Appid:"",
      NewAppid:"",
      ItemInfo:"",
      ItemId:null,
      NewApproverName:"",
      TitleId:"",
      ListApplicationName:"",
      EMPName:"",
      EMPids:"",
      loginNames:"",
      ToMail:[],
      SendBody:"",
      Mycc:"",
      newbody:"",
      newApplication:"",
      Quater:"",
      Year:"",
      ApplicationName:"",
      WFReviewDuedate:"",
      EmpName:"",
      EmpidName:"",
    };
    this._service = new Service(this.props.url, this.props.context);

    RedirectUrl = this.props.url;
  
    let myitemId = this.getParam('SID');

    RecordId=myitemId;

    this.getHRandAdminGroupUserorNot();
  }

  public async getHRandAdminGroupUserorNot() {

    let itemId = this.getParam('SID');

    if(itemId!="")

    {

    this.getuserrecords();

    }

  }

  public async getuserrecords()

  {

    let myitemId = this.getParam('SID');

    RecordId=myitemId;

    let ItemInfo = await this._service.getItemByIDs(RecordId);
     this.setState({AppliName:ItemInfo.ApproverName})
    this.setState({Appid:ItemInfo.ApproverID})
    this.setState({TitleId:ItemInfo.Title})
    this.setState({Quater:ItemInfo.MyQuarter})
    this.setState({Year:ItemInfo.MyYear})
    this.setState({ApplicationName:ItemInfo.NameofList})
    this.setState({WFReviewDuedate:ItemInfo.WF_ReviewDueDate})
    this.setState({EmpName:ItemInfo.EmpName})
    this.setState({EmpidName:ItemInfo.EmpId})

    if(ItemInfo.NameofList=="Sharepoint")
    {
      this.setState({ListApplicationName:"SAR_SharePoint"})

    }

    else if(ItemInfo.NameofList=="Egencia")
    {
      this.setState({ListApplicationName:"SAR_Egencia"})
    }
    else if(ItemInfo.NameofList=="Greenhouse")
    {
      this.setState({ListApplicationName:"SAR_Greenhouse"})
    }
    else if(ItemInfo.NameofList=="PeopleSoftHRMS")
    {
      this.setState({ListApplicationName:"SAR_PeopleSoftHRM"})
    }
    else if(ItemInfo.NameofList=="PeopleSoftFinance")
    {
      this.setState({ListApplicationName:"SAR_PeopleSoftFinance"})
    }
    else if(ItemInfo.NameofList=="Cornerstone")
    {
      this.setState({ListApplicationName:"SAR_Cornerstone"})
    }
    //TitleIds+=this.state.TitleId
    TitleIds=this.state.TitleId.split(",")  
  }

 
  public getParam(name: string) {
    name = name.replace(/[\[]/, "\\\[").replace(/[\]]/, "\\\]");
    var regexS = "[\\?&]" + name + "=([^&#]*)";
    var regex = new RegExp(regexS);
    var results = regex.exec(window.location.href);
    if (results == null)
      return "";
    else
      return results[1];
  }

  public handleChangeApplicationName = (event) => {

    console.log(event.target.value);
  
    this.setState({
  
      AppliName:event.target.value,
     
  
    });
  
  };
  public handleChangeApproverID = (event) => {
    console.log(event.target.value);
  
    this.setState({
  
      Appid:event.target.value,
     
  
    });
  
  };

  public handleChangeNewwApproverID = (event) => {

    console.log(event.target.value);
  
    this.setState({
  
      NewAppid:event.target.value,
     
  
    });
  };
    public handleChangeEmpName  = (event) => {

      console.log(event.target.value);
    
      this.setState({
    
        EmpName:event.target.value,
       
    
      });
  
  };

  public handleChangeEMPid  = (event) => {

    console.log(event.target.value);
  
    this.setState({
  
      EmpidName:event.target.value,
     
  
    });

};
  private _getPeoplePickerItems7 = async (items: any[]): Promise<void> => {
    console.log('Items:', items);
    if (items.length > 0) {

      let userInfo = this._service.getUserByLogin(items[0].loginName).then((info) => {
        this.setState({ NewApproverName: info.Email});
        this.setState({ loginNames: info.LoginName});
        console.log(info);
      });
    }
    else {
      this.setState({ NewApproverName: null });
    }
  }
   
  
  private async SendEMail(e:any)
  {

    if(this.state.NewApproverName==null || this.state.NewApproverName=="")
  
    {
      alert('Please enter New Approver Name');
    }
else
{
    if(this.state.ListApplicationName=='SAR_SharePoint')
    {
    //let SARappName=ApplicationName;
    let myemailid=this.state.NewApproverName
    let ItemInfo = await this._service.getNewApproverid(myemailid);
    //let email = this.state.ItemInfo.Email
    //let ItemInfo = await this._service.getNewApproverid(myemailid);
   //alert(email)
    
    //NewAppid:ItemInfo['EMPID'],
    this.setState({NewAppid:ItemInfo.EMPID});
    this.setState({EMPids:ItemInfo.EMPID});
    this.setState({EMPName:ItemInfo.EmplName})
    for(let count=0;count<TitleIds.length;count++)
    {

let listid =TitleIds[count];
await sp.web.lists.getByTitle("SAR_SharePoint").items
.getById(+listid) .update({
  WF_ApproverEmail:myemailid,
  ApproverID:this.state.EMPids,
  ApproverName:this.state.EMPName,
  ApproverStatus:"Pending"
  

});

    }
  }
    else if(this.state.ListApplicationName=='SAR_Egencia')
    {
      let myemailid=this.state.NewApproverName
    let ItemInfo = await this._service.getNewApproverid(myemailid);
    //let email = this.state.ItemInfo.Email
    //let ItemInfo = await this._service.getNewApproverid(myemailid);
   //alert(email)
    
    //NewAppid:ItemInfo['EMPID'],
    this.setState({NewAppid:ItemInfo.EMPID});
    this.setState({EMPids:ItemInfo.EMPID});
    this.setState({EMPName:ItemInfo.EmplName})
    for(let count=0;count<TitleIds.length;count++)
    {

let listid =TitleIds[count];
await sp.web.lists.getByTitle("SAR_Egencia").items
.getById(+listid) .update({
  WF_ApproverEmail:myemailid,
  ApproverID:this.state.EMPids,
  ApproverName:this.state.EMPName,
  ApproverStatus:"Pending"
  

});
    }
    
    }
    else if(this.state.ListApplicationName=='SAR_PeopleSoftFinance')
    {
      let myemailid=this.state.NewApproverName
    let ItemInfo = await this._service.getNewApproverid(myemailid);
    //let email = this.state.ItemInfo.Email
    //let ItemInfo = await this._service.getNewApproverid(myemailid);
   //alert(email)
    
    //NewAppid:ItemInfo['EMPID'],
    this.setState({NewAppid:ItemInfo.EMPID});
    this.setState({EMPids:ItemInfo.EMPID});
    this.setState({EMPName:ItemInfo.EmplName})
    for(let count=0;count<TitleIds.length;count++)
    {

let listid =TitleIds[count];
await sp.web.lists.getByTitle("SAR_PeopleSoftFinance").items
.getById(+listid) .update({
  WF_ApproverEmail:myemailid,
  ApproverID:this.state.EMPids,
  ApproverName:this.state.EMPName,
  ApproverStatus:"Pending"
  

});
    }
    
    }

    else if(this.state.ListApplicationName=='SAR_Cornerstone')
    {
      let myemailid=this.state.NewApproverName
    let ItemInfo = await this._service.getNewApproverid(myemailid);
    //let email = this.state.ItemInfo.Email
    //let ItemInfo = await this._service.getNewApproverid(myemailid);
   //alert(email)
    
    //NewAppid:ItemInfo['EMPID'],
    this.setState({NewAppid:ItemInfo.EMPID});
    this.setState({EMPids:ItemInfo.EMPID});
    this.setState({EMPName:ItemInfo.EmplName})
    for(let count=0;count<TitleIds.length;count++)
    {

let listid =TitleIds[count];
await sp.web.lists.getByTitle("SAR_Cornerstone").items
.getById(+listid) .update({
  WF_ApproverEmail:myemailid,
  ApproverID:this.state.EMPids,
  ApproverName:this.state.EMPName,
  ApproverStatus:"Pending"
  

});
    }
    
    }
    else if(this.state.ListApplicationName=='SAR_PeopleSoftHRM')
    {
      let myemailid=this.state.NewApproverName
    let ItemInfo = await this._service.getNewApproverid(myemailid);
    //let email = this.state.ItemInfo.Email
    //let ItemInfo = await this._service.getNewApproverid(myemailid);
   //alert(email)
    
    //NewAppid:ItemInfo['EMPID'],
    this.setState({NewAppid:ItemInfo.EMPID});
    this.setState({EMPids:ItemInfo.EMPID});
    this.setState({EMPName:ItemInfo.EmplName})
    for(let count=0;count<TitleIds.length;count++)
    {

let listid =TitleIds[count];
await sp.web.lists.getByTitle("SAR_PeopleSoftHRM").items
.getById(+listid) .update({
  WF_ApproverEmail:myemailid,
  ApproverID:this.state.EMPids,
  ApproverName:this.state.EMPName,
  ApproverStatus:"Pending"
  

});
    }
    
    }
    else if(this.state.ListApplicationName=='SAR_Greenhouse')
    {
      let myemailid=this.state.NewApproverName
    let ItemInfo = await this._service.getNewApproverid(myemailid);
    //let email = this.state.ItemInfo.Email
    //let ItemInfo = await this._service.getNewApproverid(myemailid);
   //alert(email)
    
    //NewAppid:ItemInfo['EMPID'],
    this.setState({NewAppid:ItemInfo.EMPID});
    this.setState({EMPids:ItemInfo.EMPID});
    this.setState({EMPName:ItemInfo.EmplName})
    for(let count=0;count<TitleIds.length;count++)
    {

let listid =TitleIds[count];
await sp.web.lists.getByTitle("SAR_Greenhouse").items
.getById(+listid) .update({
  WF_ApproverEmail:myemailid,
  ApproverID:this.state.EMPids,
  ApproverName:this.state.EMPName,
  ApproverStatus:"Pending"
  

});
    }
    
    }
  }
    this.addusertogroup(this.state.loginNames)
  }

  private async addusertogroup(myemailid) {

    //let newApplication=this.state.ApplicationName

    if(this.state.ListApplicationName=='SAR_SharePoint')
    {
      let testdate = new Date(this.state.WFReviewDuedate);
      testdate.setDate(testdate.getDate()-1);


      newdate=moment(testdate).format("MMM-DD-yyyy");
    let ItemInfos = await this._service.EmailTemplate();
    //"https://capcoinc.sharepoint.com/sites/SARModren_Dev/SitePages/SharePoint.aspx?SID="+this.state.EMPids+"&Status=null"
    this.setState({SendBody:ItemInfos.Body})
    this.setState({Mycc:ItemInfos.CC})
    mynewcc=this.state.Mycc
    newbody=this.state.SendBody
  Imagesnew = "<img src='https://capcoinc.sharepoint.com/sites/DepartmentsDevlopment/IT/GITA/SAR/PublishingImages/SAR.jpg'/>"
    //let myname='Gangareddy';
let newtest=newbody.replace("Image",Imagesnew).replace('NAME',this.state.EMPName).replace("QUARTER",this.state.Quater).replace("Year",this.state.Year).replace("APPLICATION",this.state.ApplicationName).replace("DATE",newdate).replace("URL","https://capcoinc.sharepoint.com/sites/SecurityAccessReview/SitePages/SharePoint.aspx?SID="+this.state.EMPids+"&Status=null");
    return await sp.web.siteGroups.getByName("Sharepoint_Approvers").users.add(myemailid).then(function(results){
      console.log('done')
      sp.utility.sendEmail({

        To:[myemailid],
        CC:[mynewcc],
        Subject:'Action Required: Security Access Review',
        Body:newtest,
      })
      alert("New Approver assigned and Email Sent Successfully")
      window.location.href=RedirectUrl;
    })
    }

    else if(this.state.ListApplicationName=='SAR_Egencia')

    {
      let testdate = new Date(this.state.WFReviewDuedate);
      testdate.setDate(testdate.getDate()-1);


      newdate=moment(testdate).format("MMM-DD-yyyy");
    
      //newdate=moment(this.state.WFReviewDuedate).format("MMM-DD-yyyy");
      Imagesnew = "<img src='https://capcoinc.sharepoint.com/sites/DepartmentsDevlopment/IT/GITA/SAR/PublishingImages/SAR.jpg'/>"
      let ItemInfos = await this._service.EmailTemplate();
    let Egencialink="https://capcoinc.sharepoint.com/sites/SARModren_Dev/SitePages/SharePoint.aspx?SID="+this.state.EMPids+"&Status=null"
    this.setState({SendBody:ItemInfos.Body})
    this.setState({Mycc:ItemInfos.CC})
    mynewcc=this.state.Mycc
    newbody=this.state.SendBody
    this.setState({})
    let myname='S,Gangareddy';
let newtest=newbody.replace("Image",Imagesnew).replace('NAME',this.state.EMPName).replace("QUARTER",this.state.Quater).replace("Year",this.state.Year).replace("APPLICATION",this.state.ApplicationName).replace("DATE",newdate).replace("URL","https://capcoinc.sharepoint.com/sites/SecurityAccessReview/SitePages/Egencia.aspx?SID="+this.state.EMPids+"&Status=null");
    return await sp.web.siteGroups.getByName("Egencia_Approvers").users.add(myemailid).then(function(results){
      console.log('done')

      sp.utility.sendEmail({

        To:[myemailid],
        CC:[mynewcc],
        Subject:'Action Required: Security Access Review',
        Body:newtest,
      })
      alert("New Approver assigned and Email Sent Successfully")
      window.location.href=RedirectUrl;

    })
    
    }

    else if(this.state.ListApplicationName=='SAR_PeopleSoftFinance')

    {
      
      let testdate = new Date(this.state.WFReviewDuedate);
      testdate.setDate(testdate.getDate()-1);


      newdate=moment(testdate).format("MMM-DD-yyyy");
      Imagesnew = "<img src='https://capcoinc.sharepoint.com/sites/DepartmentsDevlopment/IT/GITA/SAR/PublishingImages/SAR.jpg'/>"
      //newdate=moment(this.state.WFReviewDuedate).format("MMM-DD-yyyy");
      let ItemInfos = await this._service.FinanceandHRMSEmailTemplate();
    let Egencialink="https://capcoinc.sharepoint.com/sites/SARModren_Dev/SitePages/SharePoint.aspx?SID="+this.state.EMPids+"&Status=null"
    this.setState({SendBody:ItemInfos.Body})
    this.setState({Mycc:ItemInfos.Mynewcc})
    mynewcc=this.state.Mycc
    newbody=this.state.SendBody
    let myname='S,Gangareddy';
let newtest=newbody.replace("Image",Imagesnew).replace('NAME',this.state.EMPName).replace("QUARTER",this.state.Quater).replace("Year",this.state.Year).replace("APPLICATION",this.state.ApplicationName).replace("DATE",newdate).replace("URL","https://capcoinc.sharepoint.com/sites/SecurityAccessReview/SitePages/PeoplesoftFinance.aspx?SID="+this.state.EMPids+"&Status=null");
    return await sp.web.siteGroups.getByName("PeopleSoftFinance_Approvers").users.add(myemailid).then(function(results){
      console.log('done')

      sp.utility.sendEmail({

        To:[myemailid],
        CC:[mynewcc],
        Subject:'Action Required: Security Access Review',
        Body:newtest,
      })
      alert("New Approver assigned and Email Sent Successfully")
      window.location.href=RedirectUrl;
    })
    }

    else if(this.state.ListApplicationName=='SAR_PeopleSoftHRM')

    {
      
      let testdate = new Date(this.state.WFReviewDuedate);
      testdate.setDate(testdate.getDate()-1);


      newdate=moment(testdate).format("MMM-DD-yyyy");

      Imagesnew = "<img src='https://capcoinc.sharepoint.com/sites/DepartmentsDevlopment/IT/GITA/SAR/PublishingImages/SAR.jpg'/>"
     // newdate=moment(this.state.WFReviewDuedate).format("MMM-DD-yyyy");
      let ItemInfos = await this._service.FinanceandHRMSEmailTemplate();
    let Egencialink="https://capcoinc.sharepoint.com/sites/SARModren_Dev/SitePages/SharePoint.aspx?SID="+this.state.EMPids+"&Status=null"
    this.setState({SendBody:ItemInfos.Body})
    this.setState({Mycc:ItemInfos.CC})
    mynewcc=this.state.Mycc
    newbody=this.state.SendBody
    let myname='S,Gangareddy';
let newtest=newbody.replace("Image",Imagesnew).replace('NAME',this.state.EMPName).replace("QUARTER",this.state.Quater).replace("Year",this.state.Year).replace("APPLICATION",this.state.ApplicationName).replace("DATE",newdate).replace("URL","https://capcoinc.sharepoint.com/sites/SecurityAccessReview/SitePages/PeopleSoft%20HRMS.aspx?SID="+this.state.EMPids+"&Status=null");
    return await sp.web.siteGroups.getByName("PeopleSoftHRMS_Approvers").users.add(myemailid).then(function(results){
      console.log('done')

      sp.utility.sendEmail({

        To:[myemailid],
        CC:[mynewcc],
        Subject:'Action Required: Security Access Review',
        Body:newtest,
      })
      alert("New Approver assigned and Email Sent Successfully")
      window.location.href=RedirectUrl;
    })
    }
    else if(this.state.ListApplicationName=='SAR_Cornerstone')
{
  let testdate = new Date(this.state.WFReviewDuedate);
  testdate.setDate(testdate.getDate()-1);


  newdate=moment(testdate).format("MMM-DD-yyyy");
  Imagesnew = "<img src='https://capcoinc.sharepoint.com/sites/DepartmentsDevlopment/IT/GITA/SAR/PublishingImages/SAR.jpg'/>"
    //newdate=moment(this.state.WFReviewDuedate).format("MMM-DD-yyyy");
      let ItemInfos = await this._service.EmailTemplate();
    let Egencialink="https://capcoinc.sharepoint.com/sites/SecurityAccessReview/SitePages/SharePoint.aspx?SID="+this.state.EMPids+"&Status=null"
    this.setState({SendBody:ItemInfos.Body})
    this.setState({Mycc:ItemInfos.CC})
    mynewcc=this.state.Mycc
    newbody=this.state.SendBody
    let myname='S,Gangareddy';
let newtest=newbody.replace("Image",Imagesnew).replace('NAME',this.state.EMPName).replace("QUARTER",this.state.Quater).replace("Year",this.state.Year).replace("APPLICATION",this.state.ApplicationName).replace("DATE",newdate).replace("URL","https://capcoinc.sharepoint.com/sites/SecurityAccessReview/SitePages/CornerStone.aspx?SID="+this.state.EMPids+"&Status=null");
    return await sp.web.siteGroups.getByName("Cornerstone_Approvers").users.add(myemailid).then(function(results){
      console.log('done')

      sp.utility.sendEmail({

        To:[myemailid],
        CC:[mynewcc],
        Subject:'Action Required: Security Access Review',
        Body:newtest,
      })
      alert("New Approver assigned and Email Sent Successfully")
      window.location.href=RedirectUrl;
    })
    }
    else if(this.state.ListApplicationName=='SAR_Greenhouse')

    {

      let testdate = new Date(this.state.WFReviewDuedate);
      testdate.setDate(testdate.getDate()-1);


      newdate=moment(testdate).format("MMM-DD-yyyy");
      Imagesnew = "<img src='https://capcoinc.sharepoint.com/sites/DepartmentsDevlopment/IT/GITA/SAR/PublishingImages/SAR.jpg'/>"
     // newdate=moment(this.state.WFReviewDuedate).format("MMM-DD-yyyy");
      
      
      let ItemInfos = await this._service.EmailTemplate();
    let Egencialink="https://capcoinc.sharepoint.com/sites/SARModren_Dev/SitePages/SharePoint.aspx?SID="+this.state.EMPids+"&Status=null"
    this.setState({SendBody:ItemInfos.Body})
    this.setState({Mycc:ItemInfos.CC})
    mynewcc=this.state.Mycc
    newbody=this.state.SendBody
    let myname='S,Gangareddy';
let newtest=newbody.replace("Image",Imagesnew).replace('NAME',this.state.EMPName).replace("QUARTER",this.state.Quater).replace("Year",this.state.Year).replace("APPLICATION",this.state.ApplicationName).replace("DATE",newdate).replace("URL","https://capcoinc.sharepoint.com/sites/SecurityAccessReview/SitePages/Greenhouse.aspx?SID="+this.state.EMPids+"&Status=null");
    return await sp.web.siteGroups.getByName("GreenHouse_Approvers").users.add(myemailid).then(function(results){
      console.log('done')

      sp.utility.sendEmail({

        To:[myemailid],
        CC:[mynewcc],
        Subject:'Action Required: Security Access Review',
        Body:newtest,
      })
      alert("New Approver assigned and Email Sent Successfully")
      window.location.href=RedirectUrl;
    })
    }

  }
  public render(): React.ReactElement<IBulkchangeProps> {
    return (
      <div className={ styles.bulkchange }>
        <div>
          <b>Employee Name</b><br></br>
          </div><br></br>
          <div className={styles.commonsize}>
          <TextField name="Applitxt" readOnly={true}  disabled={true} value={this.state.EmpName == null ? 'N/A' : this.state.EmpName} onChange={this.handleChangeEmpName}/>

        </div><br></br>
        <div>
          <b>Employee ID </b><br></br>
          </div><br></br>
          <div className={styles.commonsize}>
          <TextField name="Applitxt" readOnly={true}  disabled={true} value={this.state.EmpidName == null ? 'N/A' : this.state.EmpidName} onChange={this.handleChangeEMPid}/>

        </div><br></br>
        <div>
          <b>Approver Name</b><br></br>
          </div><br></br>
          <div className={styles.commonsize}>
          <TextField name="Applitxt" readOnly={true}  disabled={true} value={this.state.AppliName == null ? 'N/A' : this.state.AppliName} onChange={this.handleChangeApplicationName}/>

        </div><br></br>
        <div>
          <b>Approver ID</b><br></br>
          </div><br></br>
          <div className={styles.commonsize}>
          <TextField name="Appidtxt" readOnly={true} disabled={true}  value={this.state.Appid == null ? 'N/A' : this.state.Appid} onChange={this.handleChangeApproverID}/>

        </div><br></br>

        <div>
          <b>New Approver Name</b><br></br>
          </div><br></br>
          <div className={styles.commonsize}>
          <PeoplePicker
                      context={this.props.context as any}
                      //titleText="User Name"
                      personSelectionLimit={1}
                      showtooltip={true}
                      required={true}
                      onChange={this._getPeoplePickerItems7}
                      ensureUser={true}
                      showHiddenInUI={false}
                      principalTypes={[PrincipalType.User]}
                      defaultSelectedUsers={this.state.NewApproverName ? [this.state.NewApproverName] : []}
                      ref={c => (this.ppl = c)}
                      resolveDelay={1000} />
          

        </div><br></br>

      

        <div>
        <PrimaryButton onClick={e => this.SendEMail(e)}>Send E-mail</PrimaryButton>
        </div>
        
      </div>
    );
  }
}
