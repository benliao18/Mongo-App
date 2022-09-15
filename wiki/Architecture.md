## Solution Overview

Refer the following image for architecture.

![Overview](images/Architecture-1.png)

The **Mongo-App** has the following main components:
* **Web API**: Create or Get Teams Online Meeting.
* **Azure Bot**: Call to recording bot.
* **Reaording Bot**: The Azure Cloud Services (extended support) for teams policy-based recording.(https://docs.microsoft.com/en-us/microsoftteams/teams-recording-policy)
* **Downloader App**: The Downloader App will download recording files to local from SharePoint.
* **Microsoft Graph API**: The app leverages Microsoft graph api's to [Create Online  Meeting](https://docs.microsoft.com/en-us/graph/api/application-post-onlinemeetings?view=graph-rest-1.0&tabs=csharp) , [Get Online Meeting](https://docs.microsoft.com/en-us/graph/api/onlinemeeting-get?view=graph-rest-1.0&tabs=http) , [Get User](https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http)

---

## The API used by the APP
API Name | Method | Endpoint
--- | --- | --- 
Create bookingBusiness | POST | [https://graph.microsoft.com/v1.0/solutions/bookingBusinesses](https://graph.microsoft.com/v1.0/solutions/bookingBusinesses)
Delete bookingBusiness | DELETE | [https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/{bookingsID}](https://graph.microsoft.com/v1.0/solutions/bookingBusinesses)
Publish bookingBusiness | POST | [https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/{bookingsID}/publish](https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/{bookingsID}/publish)
Create bookingStaffMember | POST | [https://graph.microsoft.com/beta/bookingBusinesses/{bookingsID}/staffMembers](https://graph.microsoft.com/beta/bookingBusinesses/{bookingsID}/staffMembers)
Create bookingService | POST | [https://graph.microsoft.com/beta/bookingBusinesses/{bookingsID}/calendarView](https://graph.microsoft.com/beta/bookingBusinesses/{bookingsID}/calendarView)
Delete bookingService | POST | [https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/{bookingsID}/publish](https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/{bookingsID}/publish)
List customers | POST | [https://graph.microsoft.com/beta/bookingBusinesses/{bookingsID}/staffMembers](https://graph.microsoft.com/beta/bookingBusinesses/{bookingsID}/staffMembers)
List business calendarView | POST | [https://graph.microsoft.com/beta/bookingBusinesses/{bookingsID}/calendarView](https://graph.microsoft.com/beta/bookingBusinesses/{bookingsID}/calendarView)