import { Button, Flex, Form, FormButton, FormInput, Layout, Segment, Text, Image, Box, List, gridCellMultipleFocusableBehavior, gridNestedBehavior, Table, Status, ErrorIcon, AcceptIcon, Header, Input, SearchIcon, Dropdown, Dialog, Loader, LinkIcon } from "@fluentui/react-northstar"
import { useContext, useState } from "react";
import { TeamsFxContext } from "../Context";
import "./LoginForm.css"
import { SellerEventModel } from "../../common/sellerEventsModel";
import { useGraph } from "@microsoft/teamsfx-react";
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";
import { Providers, ProviderState } from "@microsoft/mgt-element";
import { createMicrosoftGraphClient, TeamsFx } from "@microsoft/teamsfx";
import { bookingsCustomerModel, bookingsAppointmentModel } from "../../common/bookingsCustomerModel";
import { SellerO365AccountModel } from "../../common/SellerO365AccountModel";

export function LoginForm() {
    const [loginStatus, setLoginStatus] = useState(false);
    const [selectedIndex, setSelectedIndex] = useState(-1);
    const [account, setAccount] = useState('');
    const [password, setPassword] = useState('');
    const [actionm, setActionm] = useState<string>('');
    const [events, setEvents] = useState<SellerEventModel[]>([]);
    const [customers, setCustomers] = useState<bookingsCustomerModel[]>([]);
    const [inputItems, setInputItems] = useState<string[]>([]);
    const [searchtext, setSearchtext] = useState<string>('');
    const [tablerows, setTablerows] = useState<Array<any>>([]);
    const [message, setMessage] = useState<string>('');
    const [open, setOpen] = useState<boolean>(false);
    const [bookingsUrl,setBookingsUrl] = useState<string>('');
    const cfg = {
        webapiUrl: process.env.TAITRAAPI || 'https://yourapi.azurewebsites.net/',
        bookingsGraphUrl: process.env.BookingsGraphAPI || 'https://graph.microsoft.com/beta/bookingBusinesses/'
    };
    var _customers: Array<bookingsCustomerModel> = [];
    var _appointments: Array<bookingsAppointmentModel> = [];
    const scope = ["User.Read", "Bookings.Read.All", "Bookings.ReadWrite.All", "BookingsAppointment.ReadWrite.All"];

    const { teamsfx } = useContext(TeamsFxContext);
    const graph = useGraph(
        async (graph, teamsfx, scope) => {
            // Call graph api directly to get user profile information
            var profile: any;

            // Initialize Graph Toolkit TeamsFx provider
            const provider = new TeamsFxProvider(teamsfx, scope);
            Providers.globalProvider = provider;
            Providers.globalProvider.setState(ProviderState.SignedIn);

            let mail = "";
            try {
                profile = await graph.api("/me").get();
            } catch (err) {
                console.log('useGraph出現錯誤 - ', err);
                const teamsfxg = new TeamsFx();
                await teamsfxg.login(scope);
                profile = await graph.api("/me").get();
            }

            return { profile };
        },
        { scope: scope, teamsfx: teamsfx }
    );

    async function login() {
        console.log(cfg.webapiUrl);
        addLog("seller login", graph?.data?.profile.mail);
        var requestOptions = {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ account: account, password: password })
        };
       
        console.log(cfg.webapiUrl + 'login/loginCheck');
        await fetch(cfg.webapiUrl + 'login/loginCheck', requestOptions).then(res => res.json()).then(
            async result => {
                console.log(result);

                if (result.data.code == '200') {
                    var SellerEmail = graph?.data?.profile.mail;//result.data.memberAccount;
                    var reqOption = {
                        method: 'GET',
                        headers: { 'Content-Type': 'application/json' }
                    };
                    await fetch(cfg.webapiUrl + 'sellerEvents' + '?SellerEmail=' + SellerEmail, reqOption).then(res => res.json()).then(
                        async result => {
                            console.log(result);
                            var i = 0;
                            result.data.forEach((element: SellerEventModel) => {
                                setEvents(el => [...el, element]);
                                if (i == 0) {
                                    setActionm(element.EventName);
                                }
                                i++;
                            });

                            showDialog('請點選"開啟Bookings"按鈕，進行預約時段設定。');
                            //todo : call api write log
                            var model = new SellerO365AccountModel();
                            var cmpC = '';
                            var cmpE = '';
                            if (result.data.length > 0) {
                                model.CompanyName_CN = result.data[0].CompanyName_CN;
                                model.CompanyName_EN = result.data[0].CompanyName_EN;
                            }

                            model.SellerEmail = account;
                            // model.AADAccount = profile.id;
                            model.AADAccount = graph?.data?.profile.id;
                            var reqOption = {
                                method: 'POST',
                                headers: { 'Content-Type': 'application/json' },
                                body: JSON.stringify(model)
                            };
                            await fetch(cfg.webapiUrl + 'login/SellerO365Account', reqOption).then(res => res.json()).then(
                                r => {
                                    console.log(r);

                                }
                            );

                            getEvent(result.data);
                            setLoginStatus(true);
                        }
                    );
                } else {
                    showDialog('登入失敗，請重新輸入帳號密碼，謝謝');
                }
            }
        ).catch(async err => {
            console.log(err);
            showDialog('系統發生錯誤，請聯絡管理人員，謝謝！');
        });
    }

    function getEvent(eventList: Array<any>) {
        let array: Array<string> = [];
        if (inputItems.length == 0) {
            eventList.forEach(element => {
                setInputItems(el => [...el, element.EventName]);
            });

            array = eventList.map(x => x.EventName);
            if (array.length > 0) {
                SelectEventChange(null, array[0], eventList);
            }
        } else {
            array = inputItems;
        }
        console.log('getEvent-events', array);
    }

    async function SelectEventChange(e: any, evNm: string, eventList: Array<any>) {
        //todo : 取得活動名稱
        let eventName = '';
        if (e == null)
            eventName = evNm;
        else
            eventName = e.value;

        addLog(`select event : ${eventName}`, graph?.data?.profile.mail);
        setCustomers([]);
        _customers = [];

        //todo : 取得BookingsID
        let event = events.filter(x => { return x.EventName == eventName });
        if (eventList != null && eventList.length > 0)
            event = eventList.filter(x => { return x.EventName == eventName });

        if (event != null && event.length > 0) {
            let bookingsID = event[0].BookingsID;
            console.log('https://outlook.office365.com/bookings?calendar='+bookingsID+'&action=goto')
            setBookingsUrl(`https://outlook.office365.com/bookings?calendar=${bookingsID}&action=goto`);
            const teamsfx1 = new TeamsFx();
            const token = await teamsfx1.getCredential().getToken(["User.Read"]);
            console.log(token);
            var t = token?.token;

            const start = new Date();
            start.setDate(start.getDate() - 8);
            const end = new Date();
            end.setMonth(end.getMonth() + 1);

            var url = cfg.bookingsGraphUrl + bookingsID + '/calendarView?start=' + start.toISOString() + '&end=' + end.toISOString();
            var reqOption = {
                method: 'GET',
                headers: { 'Authorization': 'Bearer ' + t }
            };
            console.log(url);
            try {
                await fetch(url, reqOption).then(res => res.json()).then(
                    result => {
                        result.value.forEach((el: any) => {
                            _appointments.push({
                                id: el.id,
                                startDateTime: el.start.dateTime,
                                endDateTime: el.end.dateTime,
                                joinWebUrl: el.joinWebUrl,
                                customerId: el.customerId,
                                customerName: el.customerName,
                                customerEmailAddress: el.customerEmailAddress
                            });
                        });

                    }
                );

            } catch (err) {
                console.log(err);
            }

            try {
                url = cfg.bookingsGraphUrl + bookingsID + '/customers';
                await fetch(url, reqOption).then(res => res.json()).then(
                    result => {
                        if (result.value != null && result.value.length > 0) {
                            result.value.forEach((element: any) => {
                                _customers.push({
                                    id: element.id,
                                    displayName: element.displayName,
                                    emailAddress: element.emailAddress,
                                    endDateTime: '',
                                    joinWebUrl: ''
                                });
                            });

                        }
                    }
                );

            } catch (err) { console.log(err); }

            console.log('_appointments',_appointments);
            if (_customers.length > 0) {
                _customers.forEach(element => {
                    var m1 = _appointments.filter(obj => {
                        return obj.customerEmailAddress == element.emailAddress;

                    });

                    if (m1 != null && m1.length > 0) {

                        var dtNow = new Date();
                        //判斷是否有進行中的會議
                        var a1 = m1.filter(function (obj) {
                            var sd = obj.startDateTime.split('T')[0];
                            var st = obj.startDateTime.split('T')[1].substring(0, 8);
                            var sdate = new Date(sd + ' ' + st);

                            var ed = obj.endDateTime.split('T')[0];
                            var et = obj.endDateTime.split('T')[1].substring(0, 8);
                            var edate = new Date(ed + ' ' + et);

                            return dtNow > sdate && dtNow < edate;
                        });

                        if (a1.length > 0) {
                            var sd = a1[0].startDateTime.split('T')[0];
                            var st = a1[0].startDateTime.split('T')[1].substring(0, 8);
                            element.endDateTime = sd + ' ' + st;
                            element.joinWebUrl = a1[0].joinWebUrl;
                            console.log('a1 ${element.displayName} - ' + element.endDateTime);
                        } else {
                            //判斷是否有尚未開始的會議
                            var a2 = m1.filter(function (obj) {
                                var sd = obj.startDateTime.split('T')[0];
                                var st = obj.startDateTime.split('T')[1].substring(0, 8);
                                var sdate = new Date(sd + ' ' + st);

                                return dtNow < sdate;
                            });

                            if (a2.length > 0) {
                                var sd = a2[0].startDateTime.split('T')[0];
                                var st = a2[0].startDateTime.split('T')[1].substring(0, 8);
                                var sdate = new Date(sd + ' ' + st);
                                element.joinWebUrl = a2[0].joinWebUrl;
                                
                                if (a2.length > 1) {
                                    let maxObj = m1.reduce((min, obj) => (min.startDateTime < obj.startDateTime) ? min : obj);

                                    sd = maxObj.startDateTime.split('T')[0];
                                    st = maxObj.startDateTime.split('T')[1].substring(0, 8);
                                    element.joinWebUrl = maxObj.joinWebUrl;
                                    console.log('a2 langth > 1 - ' + element.displayName);
                                }

                                element.endDateTime = sd + ' ' + st;
                                console.log('a2 ' + element.displayName + ' - ' + element.endDateTime);
                            } else {
                                //判斷是否有已經結束的會議
                                var a3 = m1.filter(function (obj) {
                                    var sd = obj.startDateTime.split('T')[0];
                                    var st = obj.startDateTime.split('T')[1].substring(0, 8);
                                    var sdate = new Date(sd + ' ' + st);

                                    return dtNow > sdate;
                                });

                                if (a3.length > 0) {
                                    var sd = a3[0].startDateTime.split('T')[0];
                                    var st = a3[0].startDateTime.split('T')[1].substring(0, 8);
                                    var sdate = new Date(sd + ' ' + st);

                                    if (a3.length > 1) {
                                        let maxObj = m1.reduce((max, obj) => (max.startDateTime > obj.startDateTime) ? max : obj);

                                        sd = maxObj.startDateTime.split('T')[0];
                                        st = maxObj.startDateTime.split('T')[1].substring(0, 8);
                                        console.log('a3 langth > 1 - ' + element.displayName);
                                    }

                                    element.endDateTime = sd + ' ' + st;
                                    console.log('a3 ' + element.displayName + ' - ' + element.endDateTime);
                                }

                            }
                        }
                    }
                    setCustomers(el => [...el, element]);
                });
            }


            console.log('SelectEventChange-customers', _customers);
            getBuyer(_customers);
        }
    }

    const header = {
        key: 'header',
        items: [
            { content: 'Name', key: 'name' },
            { content: 'Email', key: 'email' },
            { content: 'Meeting Booking', key: 'meeting' },
            { content: 'Action', key: 'Action' }
        ],
        className : 'tableHead'
    };



    function getBuyer(custList: bookingsCustomerModel[]) {
        console.log('getBuyer-searchtext', searchtext);
        setTablerows([]);

        let rowsPlain: Array<any> = [];
        var allCusts : Array<bookingsCustomerModel> = [];
        console.log('getBuyer-custList',custList);
        if (custList != null && custList.length > 0) {
            allCusts = custList;
        }

        let custs = [];
        if (searchtext != '') {
            custs = allCusts.filter(obj => {
                return obj.displayName.indexOf(searchtext) >= 0 || obj.emailAddress.indexOf(searchtext) >= 0;
            });
        } else {
            custs = allCusts;
        }
        console.log('getBuyer-custs',custs);
        console.log('getBuyer-rowsPlain',rowsPlain);
        
        var idx = 1;
        let contents = [];
        custs.forEach(obj => {
            contents = [];
            contents.push({ content: obj.displayName, key: idx.toString() + '-1', id: 'name-' + idx.toString() });
            contents.push({ content: obj.emailAddress, key: idx.toString() + '-2', id: 'email-' + idx.toString() });
            contents.push({ content: obj.endDateTime, key: idx.toString() + '-3', id: 'end-' + idx.toString() });
            contents.push({ key: idx.toString() + '-4', ...moreActionCell(obj) });
            rowsPlain.push(
                {
                    key: idx,
                    items: contents,
                    className: (idx == 1 ? 'rowStyle' : (idx == custs.length ? 'rowLastStyle' : 'rowStyle'))
                }
            )
            idx++;
        });
        console.log('getBuyer-rowsPlain',rowsPlain);
        setTablerows(rowsPlain);
        console.log(rowsPlain);
    }

    const moreActionCell = (obj: bookingsCustomerModel) => {
        var item: any = {};
        let dtNow = new Date();

        if (obj.joinWebUrl != '' && obj.joinWebUrl != undefined) {
            item = {
                content: (
                    <Flex gap="gap.small" vAlign="center">
                        <a className="tableBtn" href={obj.joinWebUrl} target="_blank">Join Meeting</a>
                    </Flex>
                ),
                accessibility: gridCellMultipleFocusableBehavior,
            };
        }

        return item;
    };

    function showDialog(msg: string) {
        setOpen(true);
        setMessage(msg);

    }

    async function addLog(funcName: string, userMail: string) {
        var requestOptions = {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ FunctionName: funcName, Executer: userMail })
        };
        try {
            var url = cfg.webapiUrl + "logs";
            const resp = await fetch(url, requestOptions).then(res => res.json());
            console.log(resp);
        } catch (err) {
            console.log(err);
        }
    }

    return (
        <>
            {loginStatus === false ?
                <div className="page-logingForm-padding">
                    <Form
                        onSubmit={(e) => {
                            login();
                        }}
                        className="logingPage"
                    >
                        <div className="Header">
                            <Text size="large" weight="bold" as="h1">
                                會員登入
                            </Text>
                        </div>
                        <FormInput
                            label="帳號"
                            name="account"
                            id="account"
                            inline
                            required
                            showSuccessIndicator={false}
                            value={account}
                            onChange={event => setAccount((event.target as HTMLInputElement).value)}
                            fluid={true} />
                        <FormInput
                            label="密碼"
                            name="passWord"
                            id="password"
                            inline
                            required
                            type="password"
                            onChange={event => setPassword((event.target as HTMLInputElement).value)}
                            showSuccessIndicator={false}
                            fluid={true} />
                        <Flex gap="gap.smaller" hAlign="center">
                            <FormButton content="登入" primary={true} />
                            <a href="" target="_blank" className="forgotPwd">忘記密碼</a>
                        </Flex>
                        <Dialog
                            cancelButton="關閉"
                            header="提醒您"
                            content={message}
                            open={open}
                            onOpen={() => setOpen(true)}
                            onCancel={() => setOpen(false)} />
                    </Form>
                </div> : <div className="page-content-padding"><div></div><Flex column>
                    <Flex gap="gap.smaller">
                        <Header as="h3" content="請選擇活動" />

                        <Dropdown
                            items={inputItems}
                            placeholder="Start typing a name"
                            noResultsMessage="We couldn't find any matches."
                            a11ySelectedItemsMessage="Press Delete or Backspace to remove"
                            defaultValue={actionm}
                            className="EventSelect"
                            onChange={(element, event) => SelectEventChange(event, "", [])}
                        />
                        <a href={bookingsUrl} className="tableBtn bookingLink" target="_blank">開啟Bookings</a>
                        <Text important content="(開啟Bookings可以修改可預約時間)" className="Bookings" />
                    </Flex>
                    <Segment content="Main" color="Brand">
                        <div>
                            <Header as="h3" content="買家列表" />
                        </div>
                        <div>
                            <Flex gap="gap.medium">
                                <Flex.Item grow>
                                    <Flex>
                                        <Flex.Item grow>
                                            <Input fluid placeholder="輸入買家姓名或Email進行搜尋..." onChange={event => setSearchtext((event.target as HTMLInputElement).value)} />
                                        </Flex.Item>
                                    </Flex>
                                </Flex.Item>

                                <Button className="SearchButton" icon={<SearchIcon />} onClick={event => getBuyer(customers)} iconOnly title="Search" />
                            </Flex>
                        </div>
                        <div className="QueryUserSegment">
                            <Box
                                styles={({ theme: { siteVariables } }) => ({
                                    backgroundColor: siteVariables.colorScheme.default.background4,
                                })}
                            >
                                <Table
                                    variables={{ cellContentOverflow: 'none' }}
                                    header={header}
                                    rows={tablerows}
                                    aria-label="Nested navigation"
                                    accessibility={gridNestedBehavior}
                                />
                            </Box>
                        </div>
                    </Segment>
                </Flex>
                    <Dialog
                        cancelButton="關閉"
                        header="提醒您"
                        content={message}
                        open={open}
                        onOpen={() => setOpen(true)}
                        onCancel={() => setOpen(false)} /></div>}
        </>
    )
}