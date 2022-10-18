# Mongo-App

| [Documentation](https://github.com/benliao18/Mongo-App/blob/main/wiki) | [Architecture](https://github.com/benliao18/Mongo-App/blob/main/wiki/Architecture.md) | [Deployment guide](https://github.com/benliao18/Mongo-App/blob/main/wiki/DeploymentGuide.md)
| ---- | ---- | ---- |


# Introduction

The Mogo APP provides buyers with meetings to book meetings through the Bookings booking page, where buyers can display buyers' appointment information and meeting links.

(i)	Seller Search for Find Buyer and booking schedule for online meeting. 

(ii)	Export meeting data for analysis as csv file format.

(iii)	Export sellers and buyers information as csv file format.

(iv)	Mapping Seller account to Office 365 Account.

(v)	Integrate seller information.

# Getting Started

### Prerequisites
  * Install the prerequisites:
    * [Visual Studio Code](https://code.visualstudio.com/)
    * [Teams toolkit](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension)

### Language and Formwork
  * Teamstoolkit
  * React + TypeScript
  
### The API used by the APP
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

### Architecture
![Architecture](https://github.com/benliao18/Mongo-App/blob/main/wiki/images/Architecture.png)

### Add the following Permissions to MongoAPP

  * Application Permissions
    * Directory.Read.All
    * Directory.ReadWrite.All
    * Mail.Send
  * Delegated permissions
    * Bookings.Read.All
    * BookingsAppointment.ReadWrite.All
    * Bookings.ReadWrite.All
    * Bookings.Manage.All
	
### Add the following Permissions to MongoAPP(Seller)

  * Delegated permissions
    * Bookings.Read.All
    * BookingsAppointment.ReadWrite.All
    * Bookings.ReadWrite.All

# Legal Notice

This compliance recording bot sample is provided under the [MIT License](https://github.com/shawnlien/CustomComplianceRecordingBot/blob/master/LICENSE) terms.  In addition to these terms, by using this compliance recording bot sample you agree to the following:

- You will license the use of your app to users or organization. 

- This compliance recording bot sample is not intended to substitute your own regulatory due diligence or make you or your app compliant with respect to any applicable regulations, including but not limited to privacy, healthcare, employment, or financial regulations.

- You are responsible for complying with all applicable privacy and security regulations including those related to use, collection and handling of any personal data by your app. This includes complying with all internal privacy and security policies of your organization if your app is developed to be sideloaded internally within your organization. Where applicable, you may be responsible for data related incidents or data subject requests for data collected through your app.

- Any trademarks or registered trademarks of Systex in the United States and/or other countries and logos included in this repository are the property of Systex, and the license for this project does not grant you rights to use any Systex names, logos or trademarks outside of this repository. Systex’s copyright and privacy statement can be found [here](https://tw.systex.com/privacy/).

- If the compliance recording bot sample enables access to any Microsoft Internet-based services (e.g., Office365), use of those services will be subject to the separately-provided terms of use. In such cases, Microsoft may collect telemetry data related to compliance recording bot sample usage and operation. Use and handling of telemetry data will be performed in accordance with such terms of use.

- Use of this compliance recording bot sample does not guarantee acceptance of your app to the Teams app store. To make this app available in the Teams app store, you will have to comply with the [submission and validation process](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/appsource/publish), and all associated requirements such as including your own privacy statement and terms of use for your app.
