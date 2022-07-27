export class SellerEventModel {
    EventName: string;
    CompanyName_CN: string;
    CompanyName_EN: string;
    SellerName_CN: string;
    SellerName_EN: string;
    SellerEmail: string;
    EventOwner: string;
    BookingsID:string;

    public constructor() {
        this.EventName='';
        this.CompanyName_CN='';
        this.CompanyName_EN='';
        this.SellerName_CN='';
        this.SellerName_EN='';
        this.SellerEmail='';
        this.EventOwner='';
        this.BookingsID=''
    }
}