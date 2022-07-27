import { Button, Flex, Header, Input } from "@fluentui/react-northstar";
import { Providers, ProviderState } from "@microsoft/mgt-element";
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";
import { TeamsFx, createMicrosoftGraphClient } from "@microsoft/teamsfx";
import { useState } from "react";
import { IEventSurvey } from "../models/EventSurvey";
import { updateSurvey, addSurvey, queryEventCompanyList, SendMail } from "../services/apiservices";
import { IEventSurveyProps } from "./IEventSurvey";

export function EventSurvey(props: IEventSurveyProps) {
    const url = props?.url
    const id = props?.eventName
    const scope = ["Bookings.Read.All", "BookingsAppointment.ReadWrite.All", "Bookings.ReadWrite.All", "Bookings.Manage.All", "User.Read"]

    const [surveyUrl, setSurveyUrl] = useState('')

    function handleSurveyFieldChange(value: string) {
        setSurveyUrl(value)
        // alert(`surveyUrl=${surveyUrl}`)
    }

    async function updateEventSurvey(id: string) {
        if (!surveyUrl)
            return alert(`活動問卷網址不得為空`)

        console.log(`update event survey: ${surveyUrl}`)
        const obj: IEventSurvey = {
            EventName: id,
            SurveyURL: surveyUrl,
        }
        const message = await updateSurvey(obj)
        alert(message)
        props.getEventData()
    }

    async function addEventSurvey(id: string) {
        if (!surveyUrl)
            return alert(`活動問卷網址不得為空`)

        console.log(`add event survey: ${surveyUrl}`)
        const obj: IEventSurvey = {
            EventName: id,
            SurveyURL: surveyUrl,
        }
        const message = await addSurvey(obj)
        alert(message)
        props.getEventData()
    }

    async function sendSurveytoBuyers(eventName: string) {
        // 先從資料庫取得所有活動廠商的Bookings ID
        console.log(`sendSurveytoBuyers: ${eventName}`)
        if (url) {
            const sellerData = await queryEventCompanyList(eventName)
            let count = 0;

            if (sellerData) {
                //從Bookings ID取得buyer回來
                const teamsfx = new TeamsFx();
                const provider = new TeamsFxProvider(teamsfx, scope);

                Providers.globalProvider = provider;
                Providers.globalProvider.setState(ProviderState.SignedIn);

                const graph = createMicrosoftGraphClient(teamsfx, scope)
                for (let i = 0; i < sellerData.length; i++) {
                    const bookings = await graph.api(`/bookingBusinesses/${sellerData[i].BookingsID}/customers`)
                        .version(`beta`)
                        .get()
                    if (bookings) {
                        const values: any[] = bookings.value
                        for (let j = 0; j < values.length; j++) {
                            SendMail(eventName, values[j].displayName, values[j].emailAddress, url)
                            count++
                        }
                    }
                }
                alert(`已傳送${count}封信成功`)
            }
        } else {
            alert(`必須先填入問卷網址!!`)
        }
    }

    async function sendSurveytoSeller(eventName: string) {
        // 先從資料庫取得所有活動廠商的Bookings ID
        console.log(`sendSurveytoSeller: ${eventName}`)

        if (url) {
            const sellerData = await queryEventCompanyList(eventName)
            if (sellerData) {

                for (let i = 0; i < sellerData.length; i++) {
                    console.log('ContactEmail : ' + sellerData[i].ContactEmail);
                    await SendMail(eventName, sellerData[i].CompanyName_CN, sellerData[i].ContactEmail, url)
                }
            }
            alert(`已傳送${sellerData.length}封信成功`)
        }

    }

    return (
        <>
            <Flex gap="gap.small" vAlign="center">
                <Header as="h4">活動問卷網址</Header>
            </Flex>
            <Flex gap="gap.small" vAlign="center">
                <Input placeholder="請輸入問卷網址" fluid defaultValue={url ? url : ""} onChange={(e) => { handleSurveyFieldChange((e.target as HTMLInputElement).value); }} />
            </Flex><br /><Flex gap="gap.small" vAlign="center">
                <Button size="small" content="儲存" className="btnAction" onClick={url ? () => updateEventSurvey(id) : () => addEventSurvey(id)} />
                <Button size="small" content="發送給買家" className="btnAction" onClick={() => sendSurveytoBuyers(id)} />
                <Button size="small" content="發送給廠商" className="btnAction" onClick={() => sendSurveytoSeller(id)} />
            </Flex>
        </>
    )
}

