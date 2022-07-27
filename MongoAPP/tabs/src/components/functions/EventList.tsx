import { Button, Flex, Header, Segment, Table, gridNestedBehavior, Popup, ArrowLeftIcon, Loader } from "@fluentui/react-northstar";
import { forwardRef, useEffect, useImperativeHandle, useState } from "react";
import './EventList.css';
import { IEventListProps } from "./IEventListProps";
import { IEventSellerMapping } from "../models/EventSellerMappingModel";
import { addLog, deleteEvent, queryEventCompanyList, queryEventList } from "../services/apiservices";
import { IEventSurvey } from "../models/EventSurvey";
import { ITableModel } from "../models/TableModel";
import { EventSurvey } from "./EventSurvey";
import { createMicrosoftGraphClient, TeamsFx } from "@microsoft/teamsfx";
import { Providers, ProviderState } from "@microsoft/mgt-element";
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";
import { ISellerBookingsLinkModel } from "../models/SellerBookingLinkModel";
import { tsXLXS } from 'ts-xlsx-export';
import { ISellerBuyerModel, ISellerModel } from "../models/SellerBuyerModel";
import * as xlsx from "xlsx";
import FileSaver from 'file-saver';

const EventList = forwardRef((props: IEventListProps, ref) => {
  const [viewMode, setViewMode] = useState("event");
  const [eventList, setEventList] = useState<IEventSurvey[]>();
  const [eventRow, setEventRow] = useState<ITableModel[]>([]);
  const [companyList, setCompanyList] = useState<IEventSellerMapping[]>();
  const [companyRow, setCompanyRow] = useState<ITableModel[]>([]);

  const scope = ["Bookings.Read.All", "BookingsAppointment.ReadWrite.All", "Bookings.ReadWrite.All", "Bookings.Manage.All", "User.Read"]

  useImperativeHandle(
    ref,
    () => ({
         getEventData() {
          GetEventListData()
        }
    }),
  )

  async function GetEventListData() {
    const userMail = props?.userMail
    const result: any[] = await queryEventList(userMail)
    const eventRowPlain: ITableModel[] = []
    for(let i = 0; i < result.length; i++) {
      const element: IEventSurvey = result[i]
      const EventName = element.EventName
      const item = [
        { content: EventName, styles: { maxWidth: '27% !important' } },
        { ...eventMoreActionCell(element.EventName, element.SurveyURL), styles: { maxWidth: '73% !important' } }
      ]
      const value: ITableModel = {
        key: element.EventName,
        items: item,
        className : (i==0 ? 'rowStyle' : (i==result.length-1 ? 'rowLastStyle' : 'rowStyle'))
      }
      eventRowPlain.push(value);
    }

    setEventRow(eventRowPlain);
    setEventList(result);
  }

  async function DeleteEvent(id: string) {
    if(!window.confirm('確認刪除Event嗎?')){
      return;
    }

    console.log(`delete event eventName: ${id}`)
      addLog('Delete Event', props.userMail)
  
      const message = await deleteEvent(id)
      alert(message)
      GetEventListData()
  }

  async function exportBookingsLink(eventName: string) {
    console.log(`exportBookingsLink: ${eventName}`)
    addLog('Export Bookings Url', props.userMail)

    let jsonData: ISellerBookingsLinkModel[] = []

    // 先從資料庫取得所有活動廠商的Bookings ID
    const sellerData = await queryEventCompanyList(eventName)
    if(sellerData) {
      //從Bookings ID取得publicUrl回來
      const teamsfx = new TeamsFx();
      const provider = new TeamsFxProvider(teamsfx, scope);
  
      Providers.globalProvider = provider;
      // await teamsfx.login(scope);
      Providers.globalProvider.setState(ProviderState.SignedIn);
  
      const graph = createMicrosoftGraphClient(teamsfx, scope)
      for(let i = 0; i < sellerData.length; i++) {
        const bookings = await graph.api(`/solutions/bookingBusinesses/${sellerData[i].BookingsID}?$select=displayName,publicUrl`).get()
        if(bookings) {
          const nameValue = bookings.displayName.split(`_`)
          const newObj: ISellerBookingsLinkModel = {
            EventName: nameValue[0],
            CompanyName: nameValue[1],
            BookingsURL: bookings.publicUrl
          }
          jsonData.push(newObj)
        }
      }
    }
    return jsonData;
  }

  async function handleExportBookingLink (eventName: string) {
    const data: any[] = await exportBookingsLink(eventName)
    tsXLXS().exportAsExcelFile(data).saveAsExcelFile(`Event Bookings Link.xlsx`);
  }

  async function exportBuyerData(eventName: string) {
    // 先從資料庫取得所有活動廠商的Bookings ID
    console.log(`exportBuyerData: ${eventName}`)

    let jsonData: ISellerBuyerModel[] = []

    const sellerData = await queryEventCompanyList(eventName)
    if(sellerData) {
      //從Bookings ID取得buyer回來
      const teamsfx = new TeamsFx();
      const provider = new TeamsFxProvider(teamsfx, scope);
  
      Providers.globalProvider = provider;
      // await teamsfx.login(scope);
      Providers.globalProvider.setState(ProviderState.SignedIn);
  
      const graph = createMicrosoftGraphClient(teamsfx, scope)
      for(let i = 0; i < sellerData.length; i++) {
        const bookings = await graph.api(`/bookingBusinesses/${sellerData[i].BookingsID}/customers`)
                                    .version(`beta`)
                                    .get()
        if(bookings) {
          const values: any[] = bookings.value
          for(let j = 0; j < values.length; j++) {
            const newObj: ISellerBuyerModel = {
              SellerName: sellerData[i].CompanyName_CN,
              displayName: values[j].displayName,
              emailAddress: values[j].emailAddress,
              addresses: values[j].addresses,
              phones: values[j].phones
            }
            jsonData.push(newObj)
          }
        }
      }
    }

    return jsonData
  }

  async function exportSellerData(eventName: string) {
    // 先從資料庫取得所有活動廠商的Bookings ID
    console.log(`exportBookingsLink: ${eventName}`)
    let jsonData: ISellerModel[] = []

    const sellerData = await queryEventCompanyList(eventName)
    if(sellerData) {

      for(let i = 0; i < sellerData.length; i++) {
        const newObj: ISellerModel = {
          CompanyName_CN: sellerData[i].CompanyName_CN,
          CompanyName_EN: sellerData[i].CompanyName_EN,
          SellerName_CN: sellerData[i].SellerName_CN,
          SellerName_EN: sellerData[i].SellerName_EN,
          SellerEmail: sellerData[i].SellerEmail
        }
        jsonData.push(newObj)
        
      }
    }

    return jsonData
  }

  async function handleExportBuyerData(eventName: string) {
    addLog('Export Buyers/Sellers Data', props.userMail)

    const dataBuyers: any[] = await exportBuyerData(eventName)
    const dataSellers: any[] = await exportSellerData(eventName)
    const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
    const fileExtension = '.xlsx';
    
    if(dataBuyers && dataSellers) {
      const wsBuyers = xlsx.utils.json_to_sheet(dataBuyers);
      const wsSellers = xlsx.utils.json_to_sheet(dataSellers);
      
      const wb = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(wb, wsBuyers, "Buyers");
      xlsx.utils.book_append_sheet(wb, wsSellers, "Sellers");

      const excelBuffer = xlsx.write(wb, { bookType: 'xlsx', type: 'array' });
      const data = new Blob([excelBuffer], {type: fileType});
      FileSaver.saveAs(data, `Seller and Buyers List_${eventName}` + fileExtension);
    }
  }

  const eventMoreActionCell = (id:string, url: string) => {
    return (
      <Flex gap="gap.small" vAlign="center">
        <Button size="small" content="參展廠商列表" primary  onClick={() => showCompanyList(id)} className="btnAction" />
        <Popup
          autoSize={true}
          trigger={<Button content="問卷" primary className="btnAction" />}
          content={ <EventSurvey eventName={id} url={url} getEventData={GetEventListData} /> }
        />
        <Button size="small" content="匯出供應商/買家資料" primary className="btnAction" onClick={() => handleExportBuyerData(id)} />
        <Button size="small" content="匯出廠商預約連結" primary className="btnAction" onClick={() => handleExportBookingLink(id)} />
        <Button size="small" content="刪除活動" primary className="btnAction" onClick={() => DeleteEvent(id)} />
      </Flex>
    )
  };

  const eventHeader = {
      key: 'header',
      items: [
        { content: 'EventName', key: 'eventName', styles: { maxWidth: '27% !important' } },
        { content: 'Action', key: 'action', styles: { maxWidth: '73% !important' } },
      ],
      className: 'tableHead'
  };

  async function showCompanyList(eventName: string) {
    const result = await queryEventCompanyList(eventName)
    console.log('EventCompanyList : '+ JSON.stringify(result));
    const companyRowsPlain: any[] = []
    for(let i = 0; i < result.length; i++) {
      const element:IEventSellerMapping = result[i]
      const item = [
        { content: element.CompanyName_CN },
        { content: element.CompanyName_EN },
        { content: element.SellerName_CN,className:"col-company-cn" },
        { content: element.SellerName_EN },
        { content: element.SellerEmail ,className:"col-Mail" },
        { ...companyMoreActionCell(element.BookingsID) }
      ]

      const value: ITableModel = {
        key: element.No.toString(),
        items: item,
        className:'rowStyle'
      }
      companyRowsPlain.push(value);
    }

    setCompanyRow(companyRowsPlain);
    setCompanyList(result);
    setViewMode("company");
  }

  const companyMoreActionCell = (id:any) => {
      return (
        <Flex gap="gap.small" vAlign="center">
          <Button size="small" content="開啟Bookings" primary className="btnAction" onClick={() => OpenBookings(id)} />
        </Flex>
      )
  };

  function OpenBookings(id: any) {
    window.open(`https://outlook.office365.com/bookings?calendar=${id}&action=goto`);
  }

  const companyHeader = {
      key: 'header',
      items: [
        { content: '廠商公司名稱(中)' },
        { content: '廠商公司名稱(英)' },
        { content: '聯絡人姓名(中)' ,className:'col-company-cn'},
        { content: '聯絡人姓名(英)' },
        { content: 'Email' },
        { content: 'Action' },
      ],
      className: 'tableHead'
  };

  useEffect(() => {
    GetEventListData();
  }, []);

  const load = () => {
    return <Loader />
  }
  return viewMode === "event" ? (
    <div>
      <Header as="h3" content="Event List" className="listTitle" /><Segment content="Main" color="Brand">
        <Table
          variables={{ cellContentOverflow: 'none' }}
          header={eventHeader}
          rows={eventRow}
          aria-label="Events List"
          accessibility={gridNestedBehavior}
          styles={{ width: '100%' }}
        />
      </Segment>

    </div>
  ) : (
      <div>
          <Flex gap="gap.small">
              <Button icon={<ArrowLeftIcon/>} iconOnly title="Back" className="backButton" onClick={() => { setViewMode("event") }} />
              <Header as="h3" content="Company List" />
          </Flex>
          <Segment content="Main" color="Brand">
              <Table
                  variables={{ cellContentOverflow: 'none' }}
                  header={companyHeader}
                  rows={companyRow}
                  aria-label="Company List"
                  accessibility={gridNestedBehavior}
                  onLoad={load}
              />
          </Segment>
      </div>
  );
});

export default EventList;