import { config } from 'dotenv';
import * as path from 'path';
import axios from "axios";
import { IEventSellerMapping } from "../models/EventSellerMappingModel";
import { IEventSurvey } from "../models/EventSurvey";
import { Client, ClientOptions } from '@microsoft/microsoft-graph-client';
import { ClientCredentialAuthenticationProvider } from './AuthenticationProvider';
import { ISellerBookingsModel } from '../models/SellerBookingsModel';
import { ILogs } from '../models/TableModel';
import { TeamsFx, createMicrosoftGraphClient } from "@microsoft/teamsfx";
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";
import { Providers, ProviderState } from "@microsoft/mgt-element";

// const ENV_FILE = path.join(__dirname, '.env');
// config({ path: ENV_FILE });

const url = process.env.APIURL || "https://yourapi.azurewebsites.net"
const scope = ["Bookings.Read.All", "BookingsAppointment.ReadWrite.All", "Bookings.ReadWrite.All", "Bookings.Manage.All", "User.Read"]

console.log(`apiservices.ts url=${String(process.env.APIURL)}`)
console.log(`apiservices.ts url=${process.env.APIURL}`)

export async function addEvent(data: IEventSellerMapping): Promise<string> {
    let result = ""
    try {

        const apiUrl = `${url}/events`
        console.log(`url=${apiUrl}`)
        //判斷有沒有重複的
        console.log('addEvent-url',`${apiUrl}?eventName=${data.EventName}`);
        var rp = await axios.get(`${apiUrl}?eventName=${data.EventName}`);
      
        if (rp.data.data != null && rp.data.data.length > 0) {
            var model: Array<IEventSellerMapping> = rp.data.data;
            var t = model.filter(item => {
                return item.SellerEmail.toLowerCase() == data.O365Account.toLowerCase();
            });

            if (t.length > 0)
                return "重複新增";
        }

        const response = await axios.post(apiUrl, data)
        console.log(response);

        if (response.status === 200) {
            result = response.data.message
        }
    } catch (error) {
        console.error(error);
    }

    return result
}

export async function deleteEvent(id: string): Promise<string> {
    let result = ""
    try {
        const apiUrl = `${url}/events`
        console.log(`url=${apiUrl}`)

        const teamsfx = new TeamsFx();
        const provider = new TeamsFxProvider(teamsfx, scope);
    
        Providers.globalProvider = provider;
        Providers.globalProvider.setState(ProviderState.SignedIn);
        const graph = createMicrosoftGraphClient(teamsfx, scope)
        const rsp1 = await axios.get(`${apiUrl}?eventName=${id}`);

        console.log(`deleteEvent-rsp1 : ${JSON.stringify(rsp1.data.data)}`);
        if (rsp1.data.data != null && rsp1.data.data.length > 0) {

            var m: Array<ISellerBookingsModel> = rsp1.data.data;
            for (let it = 0; it < m.length; it++) {
                try {
                    console.log(`deleteEvent-delete bookings id : ${m[it].BookingsID}`);
                    await graph.api(`/solutions/bookingBusinesses/${m[it].BookingsID}`).delete();
                } catch (err) {
                    console.log(err);
                }
            }
        }

        const response = await axios.delete(apiUrl, { data: { eventName: id } });
        console.log(response);

        if (response.status === 200) {
            result = response.data.message
        }
    } catch (error) {
        console.error(error);
    }

    return result
}

export async function addSurvey(data: IEventSurvey): Promise<string> {
    console.log(`add survey.`)
    let result = ""
    try {
        const apiUrl = `${url}/eventsurvey`
        console.log(`url=${apiUrl}`)
        const response = await axios.post(apiUrl, data)
        console.log(response);

        result = response.data.message

    } catch (error) {
        console.error(error);
    }

    return result
}

export async function updateSurvey(data: IEventSurvey): Promise<string> {
    console.log(`update survey`)
    let result = ""
    try {
        const apiUrl = `${url}/eventsurvey`
        console.log(`url=${apiUrl}`)
        const response = await axios.put(apiUrl, data)
        console.log(response);

        result = response.data.message
    } catch (error) {
        console.error(error);
    }
    return result
}

export async function queryEventList(userMail: string): Promise<any[]> {
    let result: any[] = []
    try {
        const apiUrl = `${url}/eventsurvey?userMail=${userMail}`

        console.log(`url=${apiUrl}`)
        const response = await axios.get(apiUrl)
        console.log(`queryEventList data=${response.data}`)

        result = response.data.data
    } catch (error) {
        console.error(error);
    }

    return result
}

export async function queryEventCompanyList(eventName: string): Promise<any[]> {
    let result: any[] = []
    try {
        const apiUrl = `${url}/events?eventName=${eventName}`

        console.log(`url=${apiUrl}`)
        const response = await axios.get(apiUrl)
        console.log(response.data)
        result = response.data.data
    } catch (error) {
        console.error(error);
    }
    return result
}

export async function CheckUserIsInDirectory(userMail: string): Promise<any[]> {
    let result: any[] = []
    try {
        const client = getClient();
        const request = await client.api("/users")
            .filter(`startswith(mail,'${userMail}')`)
            .get()
            .catch((error) => {
                console.log(error);
            });
        console.log(request.value);
        result = request.value;
    } catch (error) {
        console.error(`CheckUserIsInDirectory error, ${error}`)
    }
    return result
}

export async function CreateInvitation(userMail: string): Promise<string> {
    let result: string = "UnCompleted"
    try {
        const client = getClient()
        const invitation = {
            invitedUserEmailAddress: userMail,
            inviteRedirectUrl: 'https://myapp.' + process.env.domainName || "yourdomain.onmicrosoft.com"
        };
        const request = await client.api("/invitations")
            .post(invitation)
            .catch((error) => {
                console.log(error);
            });
        console.log(request.status);
        result = request.status;
    } catch (error) {
        console.error(`CreateInvitation error, ${error}`)
    }
    return result
}

export async function SendMail(eventName: string, recipientName: string, toRecipients: string, surveyUrl: string): Promise<string> {
    let result: string = "UnCompleted"
    const mailFrom = process.env.emailUser || "admin@yourdomain.onmicrosoft.com"
    const mailBody = `<html><head></head><body><div><p><span>Hi ${recipientName},</span></p><p><span>針對本次活動還請您協助點選下列連結進行活動調查，謝謝！</span></p><p><span><a href='${surveyUrl}'>${eventName}活動問卷</a></span></p><p><b><span style='color: #767171'>此信件透過Tasker出任務系統發送，請勿直接回覆</span></b></p></div></body></html>`
    try {
        const client = getClient()
        const message = {
            message: {
                subject: `[TAITRA]${eventName}問卷調查`,
                body: {
                    contentType: 'html',
                    content: mailBody
                },
                toRecipients: [
                    {
                        emailAddress: {
                            address: toRecipients
                        }
                    }
                ]
            },
            saveToSentItems: 'false'
        };
        const request = await client.api(`/users/${mailFrom}/sendMail`)
            .headers({
                'Content-Type': 'application/json'
            })
            .post(message)
            .catch((error) => {
                console.log(error);
            });
        console.log(request.status);
        result = request.status;
    } catch (error) {
        console.error(`CreateInvitation error, ${error}`)
    }
    return result
}

export async function addMapping(data: ISellerBookingsModel): Promise<string> {
    let result = ""
    try {
        const apiUrl = `${url}/events`
        console.log(`url=${apiUrl}`)
        const response = await axios.put(apiUrl, data)
        console.log(response);

        if (response.status === 200) {
            result = response.data.message
        }
    } catch (error) {
        console.error(error);
    }

    return result
}

export async function getADMappingList(): Promise<any[]> {
    let result: any[] = []
    try {
        const apiUrl = `${url}/admapping`
        console.log(`url=${apiUrl}`)
        const response = await axios.get(apiUrl)
        console.log(response);

        if (response.status === 200) {
            result = response.data.data
        }
    } catch (error) {
        console.error(error);
    }

    return result
}

// ************************ Common Functions ************************
function getClient(): Client {
    const clientOptions: ClientOptions = {
        defaultVersion: "v1.0",
        debugLogging: false,
        authProvider: new ClientCredentialAuthenticationProvider()
    }

    const client = Client.initWithMiddleware(clientOptions);
    return client
}

export async function addLog(functionName: string, userMail: string): Promise<string> {
    console.log(`add Log.`)
    const data: ILogs = {
        FunctionName: functionName,
        Executer: userMail
    }
    let result = ""
    try {
        const apiUrl = `${url}/logs`
        console.log(`url=${apiUrl}`)
        const response = await axios.post(apiUrl, data)
        console.log(response);

        result = response.data.message

    } catch (error) {
        console.error(error);
    }

    return result
}

export async function queryLogs(): Promise<any[]> {
    let result: any[] = []
    try {
        const apiUrl = `${url}/logs`

        console.log(`url=${apiUrl}`)
        const response = await axios.get(apiUrl)
        console.log(response.data)
        result = response.data.data
    } catch (error) {
        console.error(error);
    }
    return result
}
// ************************ Common Functions ************************