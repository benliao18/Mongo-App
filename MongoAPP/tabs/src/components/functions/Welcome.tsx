import { useContext, useEffect, useRef, useState } from "react";
import { Button, CloseIcon, datepickerBehavior, Dialog, Divider, Flex, Header, Input, Loader } from "@fluentui/react-northstar";
import "./Welcome.css";
import EventList from "./EventList"

import * as xlsx from "xlsx";
import { tsXLXS } from 'ts-xlsx-export';

import { IEventSellerMapping } from "../models/EventSellerMappingModel";
import { addEvent, addLog, addMapping, CheckUserIsInDirectory, CreateInvitation, getADMappingList, queryLogs } from "../services/apiservices";
import { useGraph } from "@microsoft/teamsfx-react";
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";
import { Providers, ProviderState } from "@microsoft/mgt-element";
import { TeamsFxContext } from "../Context";
import { IADMappingModel } from "../models/ADMappingModel";
import { TeamsFx, createMicrosoftGraphClient } from "@microsoft/teamsfx";
import { ISellerBookingsModel } from "../models/SellerBookingsModel";
import { IEventListProps } from "./IEventListProps";
import { Link } from "react-router-dom";
import { ILogs } from "../models/TableModel";

export function Welcome(props: { showFunction?: boolean; environment?: string }) {
  const [admin, setAdmin] = useState(false)
  const [open, setOpen] = useState(false)
	const [selectedFile, setSelectedFile] = useState();
  const [userMail, setUserMail] = useState('')
  const [isImport, setIsImport] = useState(false)

  const childRef = useRef<IEventListProps>(null);

  const scope = ["Bookings.Read.All", "BookingsAppointment.ReadWrite.All", "Bookings.ReadWrite.All", "Bookings.Manage.All", "User.Read"]

  function filePathset(e: any, data: any) {
    // alert(`filePathset.`)
    e.stopPropagation();
    e.preventDefault();
    const file = e.target.files[0];
    console.log(file);
    setSelectedFile(file);

    console.log(file);
  }

  function readUploadFile() {
    setIsImport(true)
    if (selectedFile) {
        const reader = new FileReader();
        reader.onload = async (e) => {
          const data = e.target ? e.target.result : null ;
          if(data !== null) {
            const workbook = xlsx.read(data, { type: "array" });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const json: any[] = xlsx.utils.sheet_to_json(worksheet);
            console.log(json);

            //先宣告Graph
            let result: boolean = false
            try {
              var idx = 0;
              
              for (var i = 0 ; i < json.length;i++){
                var obj : IEventSellerMapping =json[i];
                var rtn = await addEvent(obj);  //先將資料寫入資料庫

                if (rtn != "重複新增") {
                  console.log('第( '+i.toString() + ' )筆開始', new Date());
                  await AddBookings(obj);  //建立Bookings
                  console.log('第( '+i.toString() + ' )筆結束', new Date());
                } else {
                  alert(`${obj.EventName}活動及O365Account(${obj.O365Account})已經存在`);
                }

                if (i == json.length - 1) {
                  console.log('匯入成功');
                  alert('匯入成功');
                }
              }
            } catch (error) {
              console.error(`Import Event List fail. ${error}`)
            }
          } else {
            console.log(`file data is null.`);
          }
        };
        reader.readAsArrayBuffer(selectedFile);
    }else{
      alert(`請選擇檔案!!`)
    }
    setIsImport(false)
  }

  async function AddBookings(obj: IEventSellerMapping) {
    addLog('Import Events', userMail)

    const owners = obj.EventOwner.split(`;`)
    const index = owners.indexOf(data!.mail, 0); //找出登入者，並移除，其餘人後續再加入
    if (index > -1) {
      owners.splice(index, 1);
    }

    const eventEntity = {
      "displayName": obj.EventName + '_' + obj.CompanyName_CN,
      "email": data!.mail,
      "defaultCurrencyIso":"TWD",
      "businessType":"Other"
    }
    console.log(`eventEntity=${JSON.stringify(eventEntity)}`)
    const teamsfx = new TeamsFx();
    const provider = new TeamsFxProvider(teamsfx, scope);

    Providers.globalProvider = provider;
    // await teamsfx.login(scope);
    Providers.globalProvider.setState(ProviderState.SignedIn);

    const graph = createMicrosoftGraphClient(teamsfx, scope)
    
    // 先新增
    const bookings = await graph.api("/solutions/bookingBusinesses").post(eventEntity)
    if(bookings) {
      // Publis Bookings
      try {
        await graph.api(`/solutions/bookingBusinesses/${bookings.id}/publish`).post({})
      } catch (error) {
        alert(`publish error, ${error}`)
      }
      
      // 將其他Owner加入到staff中，並設定為管理者
      for (let it = 0; it < owners.length; it++) {
        const bookingStaffMemberOwner = {
          '@odata.type': '#microsoft.graph.bookingStaffMember',
          displayName: owners[it].substring(0, 1),
          emailAddress: owners[it],
          'role@odata.type': '#microsoft.graph.bookingStaffRole',
          // role: 'externalGuest',
          role: 'administrator',
          timeZone: 'Asia/Taipei',
          useBusinessHours: true,
          'workingHours@odata.type': '#Collection(microsoft.graph.bookingWorkHours)',
          workingHours: [
            {
              '@odata.type': '#microsoft.graph.bookingWorkHours',
              'day@odata.type': '#microsoft.graph.dayOfWeek',
              day: 'monday',
              'timeSlots@odata.type': '#Collection(microsoft.graph.bookingWorkTimeSlot)',
              timeSlots: [
                {
                  '@odata.type': '#microsoft.graph.bookingWorkTimeSlot',
                  end: '17:00:00.0000000',
                  start: '08:00:00.0000000'
                }
              ]
            },
            {
              '@odata.type': '#microsoft.graph.bookingWorkHours',
              'day@odata.type': '#microsoft.graph.dayOfWeek',
              day: 'tuesday',
              'timeSlots@odata.type': '#Collection(microsoft.graph.bookingWorkTimeSlot)',
              timeSlots: [
                {
                  '@odata.type': '#microsoft.graph.bookingWorkTimeSlot',
                  end: '17:00:00.0000000',
                  start: '08:00:00.0000000'
                }
              ]
            },
            {
              '@odata.type': '#microsoft.graph.bookingWorkHours',
              'day@odata.type': '#microsoft.graph.dayOfWeek',
              day: 'wednesday',
              'timeSlots@odata.type': '#Collection(microsoft.graph.bookingWorkTimeSlot)',
              timeSlots: [
                {
                  '@odata.type': '#microsoft.graph.bookingWorkTimeSlot',
                  end: '17:00:00.0000000',
                  start: '08:00:00.0000000'
                }
              ]
            },
            {
              '@odata.type': '#microsoft.graph.bookingWorkHours',
              'day@odata.type': '#microsoft.graph.dayOfWeek',
              day: 'thursday',
              'timeSlots@odata.type': '#Collection(microsoft.graph.bookingWorkTimeSlot)',
              timeSlots: [
                {
                  '@odata.type': '#microsoft.graph.bookingWorkTimeSlot',
                  end: '17:00:00.0000000',
                  start: '08:00:00.0000000'
                }
              ]
            },
            {
              '@odata.type': '#microsoft.graph.bookingWorkHours',
              'day@odata.type': '#microsoft.graph.dayOfWeek',
              day: 'friday',
              'timeSlots@odata.type': '#Collection(microsoft.graph.bookingWorkTimeSlot)',
              timeSlots: [
                {
                  '@odata.type': '#microsoft.graph.bookingWorkTimeSlot',
                  end: '17:00:00.0000000',
                  start: '08:00:00.0000000'
                }
              ]
            }
          ]
        };

        try {
          const staffAdmin = await graph.api(
            `/bookingBusinesses/${bookings.id}/staffMembers`)
            .version('beta')
            .post(bookingStaffMemberOwner)
          console.log(`add EventOwner to staffmember, role is ${staffAdmin.role}`)
        } catch (error) {
          console.log('將其他Owner加入到staff中，並設定為管理者', error);
          alert(error);
        }
      }

      owners.forEach(async element => {
        const bookingStaffMemberOwner = {
          '@odata.type':'#microsoft.graph.bookingStaffMember',
          displayName: element.substring(0, 1),
          emailAddress: element,
          'role@odata.type':'#microsoft.graph.bookingStaffRole',
          // role: 'externalGuest',
          role: 'administrator',
          timeZone: 'Asia/Taipei',
          useBusinessHours: true,
          'workingHours@odata.type':'#Collection(microsoft.graph.bookingWorkHours)',
          workingHours: [
              {
                  '@odata.type':'#microsoft.graph.bookingWorkHours',
                  'day@odata.type':'#microsoft.graph.dayOfWeek',
                  day: 'monday',
                  'timeSlots@odata.type':'#Collection(microsoft.graph.bookingWorkTimeSlot)',
                  timeSlots: [
                      {
                          '@odata.type':'#microsoft.graph.bookingWorkTimeSlot',
                          end: '17:00:00.0000000',
                          start: '08:00:00.0000000'
                      }
                  ]
              },
              {
                  '@odata.type':'#microsoft.graph.bookingWorkHours',
                  'day@odata.type':'#microsoft.graph.dayOfWeek',
                  day: 'tuesday',
                  'timeSlots@odata.type':'#Collection(microsoft.graph.bookingWorkTimeSlot)',
                  timeSlots: [
                      {
                          '@odata.type':'#microsoft.graph.bookingWorkTimeSlot',
                          end: '17:00:00.0000000',
                          start: '08:00:00.0000000'
                      }
                  ]
              },
              {
                  '@odata.type':'#microsoft.graph.bookingWorkHours',
                  'day@odata.type':'#microsoft.graph.dayOfWeek',
                  day: 'wednesday',
                  'timeSlots@odata.type':'#Collection(microsoft.graph.bookingWorkTimeSlot)',
                  timeSlots: [
                      {
                          '@odata.type':'#microsoft.graph.bookingWorkTimeSlot',
                          end: '17:00:00.0000000',
                          start: '08:00:00.0000000'
                      }
                  ]
              },
              {
                  '@odata.type':'#microsoft.graph.bookingWorkHours',
                  'day@odata.type':'#microsoft.graph.dayOfWeek',
                  day: 'thursday',
                  'timeSlots@odata.type':'#Collection(microsoft.graph.bookingWorkTimeSlot)',
                  timeSlots: [
                      {
                          '@odata.type':'#microsoft.graph.bookingWorkTimeSlot',
                          end: '17:00:00.0000000',
                          start: '08:00:00.0000000'
                      }
                  ]
              },
              {
                  '@odata.type':'#microsoft.graph.bookingWorkHours',
                  'day@odata.type':'#microsoft.graph.dayOfWeek',
                  day: 'friday',
                  'timeSlots@odata.type':'#Collection(microsoft.graph.bookingWorkTimeSlot)',
                  timeSlots: [
                      {
                          '@odata.type':'#microsoft.graph.bookingWorkTimeSlot',
                          end: '17:00:00.0000000',
                          start: '08:00:00.0000000'
                      }
                  ]
              }
          ]
        };

        try {
          // const staffAdmin = await graph.api(
          //   `/bookingBusinesses/${bookings.id}/staffMembers`)
          //   .version('beta')
          //   .post(bookingStaffMemberOwner)
          // console.log(`add EventOwner to staffmember, role is ${staffAdmin.role}`)
        } catch (error) {
          console.log('將其他Owner加入到staff中，並設定為管理者',error);
          alert(error);
        }
      });

      // 將廠商加入到staff，先檢查帳號是否已存在tenant中，若不存在則先加入到AAD
      try {
        const arrayValue: any[] = await CheckUserIsInDirectory(obj.O365Account)
        if(arrayValue.length === 0) {
          CreateInvitation(obj.O365Account)
          console.log(`Create Invitation to ${obj.O365Account}`)
        }
      } catch (error) {
        alert(`Check user is in directory fail. ${error}`)
      }

      // 完成帳號處理就將Seller帳號加入到建立的Bookings中的staffMember中
      const eventid = bookings.id
      const bookingStaffMember = {
          '@odata.type':'#microsoft.graph.bookingStaffMember',
          displayName: obj.SellerName_CN,
          emailAddress: obj.O365Account,
          'role@odata.type':'#microsoft.graph.bookingStaffRole',
          // role: 'externalGuest',
          role: 'administrator',
          timeZone: 'Asia/Taipei',
          useBusinessHours: true,
          'workingHours@odata.type':'#Collection(microsoft.graph.bookingWorkHours)',
          workingHours: [
              {
                  '@odata.type':'#microsoft.graph.bookingWorkHours',
                  'day@odata.type':'#microsoft.graph.dayOfWeek',
                  day: 'monday',
                  'timeSlots@odata.type':'#Collection(microsoft.graph.bookingWorkTimeSlot)',
                  timeSlots: [
                      {
                          '@odata.type':'#microsoft.graph.bookingWorkTimeSlot',
                          end: '17:00:00.0000000',
                          start: '08:00:00.0000000'
                      }
                  ]
              },
              {
                  '@odata.type':'#microsoft.graph.bookingWorkHours',
                  'day@odata.type':'#microsoft.graph.dayOfWeek',
                  day: 'tuesday',
                  'timeSlots@odata.type':'#Collection(microsoft.graph.bookingWorkTimeSlot)',
                  timeSlots: [
                      {
                          '@odata.type':'#microsoft.graph.bookingWorkTimeSlot',
                          end: '17:00:00.0000000',
                          start: '08:00:00.0000000'
                      }
                  ]
              },
              {
                  '@odata.type':'#microsoft.graph.bookingWorkHours',
                  'day@odata.type':'#microsoft.graph.dayOfWeek',
                  day: 'wednesday',
                  'timeSlots@odata.type':'#Collection(microsoft.graph.bookingWorkTimeSlot)',
                  timeSlots: [
                      {
                          '@odata.type':'#microsoft.graph.bookingWorkTimeSlot',
                          end: '17:00:00.0000000',
                          start: '08:00:00.0000000'
                      }
                  ]
              },
              {
                  '@odata.type':'#microsoft.graph.bookingWorkHours',
                  'day@odata.type':'#microsoft.graph.dayOfWeek',
                  day: 'thursday',
                  'timeSlots@odata.type':'#Collection(microsoft.graph.bookingWorkTimeSlot)',
                  timeSlots: [
                      {
                          '@odata.type':'#microsoft.graph.bookingWorkTimeSlot',
                          end: '17:00:00.0000000',
                          start: '08:00:00.0000000'
                      }
                  ]
              },
              {
                  '@odata.type':'#microsoft.graph.bookingWorkHours',
                  'day@odata.type':'#microsoft.graph.dayOfWeek',
                  day: 'friday',
                  'timeSlots@odata.type':'#Collection(microsoft.graph.bookingWorkTimeSlot)',
                  timeSlots: [
                      {
                          '@odata.type':'#microsoft.graph.bookingWorkTimeSlot',
                          end: '17:00:00.0000000',
                          start: '08:00:00.0000000'
                      }
                  ]
              }
          ]
      };
      var staffId = '';
      try {
       var staffMember = await graph.api(
          `/solutions/bookingBusinesses/${eventid}/staffMembers`)
          .post(JSON.stringify(bookingStaffMember));
console.log('staffMember',staffMember);
          if(staffMember){
            staffId = staffMember.id;
          }
      } catch (error) {
        alert(`Add seller to staff fail. ${error}`)
      }

      try {

        console.log('start to get service ');
        console.log(`/solutions/bookingBusinesses/${eventid}/services`);

        var serviceList = await graph.api(
          `/solutions/bookingBusinesses/${eventid}/services`)
          .get();

        console.log('serviceList count = ' + serviceList.value.length);

        if (serviceList.value.length > 0) {
          var serviceId = serviceList.value[0].id;
          const updateBody = {
            "displayName": "買家視訊會議",
            "defaultDuration": "PT0H30M",
            "defaultPrice": 0,
            "defaultPriceType": "notSet",
            "preBuffer": "PT0S",
            "postBuffer": "PT0S",
            "isLocationOnline": true,
            "defaultLocation": null,
            "staffMemberIds":[staffId]
          }

          console.log('updateBody',updateBody);

          await graph.api(
            `/solutions/bookingBusinesses/${eventid}/services/${serviceId}`)
            .patch(JSON.stringify(updateBody));

        }
        const addServiceBody = {
          "@odata.type": "#microsoft.graph.bookingService",
          "defaultDuration": "PT0H30M",
          "defaultLocation": null,
          "defaultPrice": 0,
          "defaultPriceType@odata.type": "#microsoft.graph.bookingPriceType",
          "defaultPriceType": "notSet",
          "defaultReminders@odata.type": "#Collection(microsoft.graph.bookingReminder)",
          "defaultReminders": [
            {
              "@odata.type": "#microsoft.graph.bookingReminder",
              "message": "這是提醒您服務不久後就可以使用的快速提醒.",
              "offset": "P1D",
              "recipients@odata.type": "#microsoft.graph.bookingReminderRecipients",
              "recipients": "customer"
            }
          ],
          "description": "買家視訊會議",
          "displayName": "買家視訊會議",
          "isLocationOnline": true,
          "smsNotificationsEnabled": true,
          "isHiddenFromCustomers": false,
          "notes": "",
          "postBuffer": "PT0S",
          "preBuffer": "PT0S",
          "schedulingPolicy": null,
          "staffMemberIds@odata.type": "#Collection(String)",
          "staffMemberIds": []
        }

      } catch(error){
        alert(`add service fail. ${error}`);
      }

      // Bookings處理完畢後將資料寫回SellerBookingsInfo
      const sellerBookings: ISellerBookingsModel = {
        BookingsID: eventid,
        CompanyName_CN: obj.CompanyName_CN,
        CompanyName_EN: obj.CompanyName_EN,
        SellerEmail: obj.O365Account,
        EventName: obj.EventName
      }
      try {
        const result = await addMapping(sellerBookings)

        if(result) {
          
          setOpen(false);
          setIsImport(false);

          if(childRef.current?.getEventData) {
            childRef.current?.getEventData()
            
          }
        }
      } catch (error) {
        alert(`Add Mapping fail. ${error}`)
      }
      
    }
    console.log(`bookings=${bookings}`)
    
  }

  const { teamsfx } = useContext(TeamsFxContext);
  const { loading, error, data, reload } = useGraph(
    async (graph, teamsfx, scope) => {
      // Call graph api directly to get user profile information
      var profile : any;

      // Initialize Graph Toolkit TeamsFx provider
      const provider = new TeamsFxProvider(teamsfx, scope);
      Providers.globalProvider = provider;
      Providers.globalProvider.setState(ProviderState.SignedIn);

      let mail = "";
      try {
        profile = await graph.api("/me").get();
        const profilemail = await graph.api("/me/mail/$value").get();
        mail = profilemail;
      } catch (err) {
        console.log('useGraph出現錯誤 - ', err);
        const teamsfxg = new TeamsFx();
        await teamsfxg.login(scope);
        profile = await graph.api("/me").get();
        const profilemail = await graph.api("/me/mail/$value").get();
        mail = profilemail;
        // Could not fetch photo from user's profile, return empty string as placeholder.
      }
      const adminUsers = process.env.adminUsers || "admin@yourdomain.onmicrosoft.com"
      console.log(`admin user : ${adminUsers}`);
      console.log(`login user : ${mail.toLowerCase()}`);
      const values = adminUsers.split(`;`)
      if(values.find(obj => obj.toLowerCase() === mail.toLowerCase()))
        setAdmin(true)

      setUserMail(mail)
      return { profile, mail };
    },
    { scope: ["Bookings.Read.All", "BookingsAppointment.ReadWrite.All", "Bookings.ReadWrite.All", "Bookings.Manage.All", "User.Read"], teamsfx: teamsfx }
  );

  async function exportADAccountMapping() {
    addLog('Export AD Account Mapping', userMail)
    try {
      const csvData: IADMappingModel[] = await getADMappingList()
      if(csvData) {
        tsXLXS().exportAsExcelFile(csvData).saveAsExcelFile(`aad user account mapping.xlsx`);
      }else{
        alert(`No Data can export!!`)
      }
    } catch (error) {
      alert(error)
    }
  }

  async function exportLogs() {
    try {
      const csvData: ILogs[] = await queryLogs()
      if(csvData) {
        tsXLXS().exportAsExcelFile(csvData).saveAsExcelFile(`SystemLogs.xlsx`);
      }else{
        alert(`No Data can export!!`)
      }
    } catch (error) {
      alert(error)
    }
    addLog('Export Logs', userMail)
  }

  useEffect(() => {
    
  }, []);

  return loading ? (
    <Loader />
  ) : ( 
    <div className="mainContent">
      <div>
        <Flex space="between">
          <Header as="h3" content="TAITRA Eevents Management" />
          <Flex gap="gap.small" className="adminFunc">
            <Button content="Import Events" primary onClick={() => setOpen(true)} />
            {
              admin ?             
                <><Button content="Export Logs" primary onClick={exportLogs} /><Button content="Export Company/MS AAD Mapping" primary onClick={exportADAccountMapping} /></>
              : <></>
            }
          </Flex>
        </Flex>
      </div>
      <Divider fitted size={2} color="Brand" />
      <div>
        <EventList userMail={ userMail } ref={childRef}  />
        {/* {loading ? <Loader /> : <EventList userMail={ userMail } /> } */}
      </div>
      <div>
        <Dialog
          header="Import File"
          open={open}
          headerAction={{ icon: <CloseIcon />, title: 'Close', onClick: () => setOpen(false) }}
          content={
            <>
              <Flex gap="gap.small">
                <Input 
                  type="file"
                  accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
                  onChange={filePathset}
                />
                <Button 
                  loading={ isImport }
                  content={ isImport ? `Importing` : `Import` }
                  disabled={ isImport }
                  onClick={readUploadFile} />
              </Flex>
              <Flex>
                <div className="spaceTop">
                  下載匯入的範例檔案：<a href="#" onClick={() => { window.open(`https://fileurl.com.tw`) }}>範例檔案</a>
                </div>
              </Flex>
            </>
          }
        />
      </div>
    </div>
  );
}

