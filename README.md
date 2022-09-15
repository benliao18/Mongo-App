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

### Add the following Permissions to TaitraAPP:

  * Application Permissions
    * Directory.Read.All
    * Directory.ReadWrite.All
    * Mail.Send
  * Delegated permissions
    * Bookings.Read.All
    * BookingsAppointment.ReadWrite.All
    * Bookings.ReadWrite.All
	* Bookings.Manage.All
	
### Add the following Permissions to SellerAPP:

  * Delegated permissions
    * Bookings.Read.All
    * BookingsAppointment.ReadWrite.All
    * Bookings.ReadWrite.All

