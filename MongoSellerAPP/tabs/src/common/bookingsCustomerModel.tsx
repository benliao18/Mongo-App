export class bookingsCustomerModel {
    id: string;
    displayName: string;
    emailAddress: string;
    endDateTime: string;
    joinWebUrl:string;

    public constructor() {
        this.id='';
        this.displayName='';
        this.emailAddress='';
        this.endDateTime='';
        this.joinWebUrl=''
    }
}

export class bookingsAppointmentModel {
    id: string;
    startDateTime: string;
    endDateTime: string;
    joinWebUrl:string;
    customerId:string;
    customerName:string;
    customerEmailAddress:string;

    public constructor() {
        this.id='';
        this.startDateTime='';
        this.endDateTime='';
        this.joinWebUrl='';
        this.customerId='';
        this.customerName='';
        this.customerEmailAddress='';
    }
}