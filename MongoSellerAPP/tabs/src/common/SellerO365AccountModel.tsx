export class SellerO365AccountModel {
    SellerEmail: string;
    TenantID: string;
    AADAccount: string;
    CompanyName_CN: string;
    CompanyName_EN: string;
   

    public constructor() {
        this.SellerEmail='';
        this.TenantID='';
        this.AADAccount='';
        this.CompanyName_CN='';
        this.CompanyName_EN='';
    }
}