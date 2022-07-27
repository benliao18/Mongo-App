export interface IBookings {
    id: string;
    displayName: string;
    businessType?: string;
    phone?: string;
    email: string;
    webSiteUrl?: string;
    defaultCurrencyIso?: string;
    isPublished?: string;
    publicUrl?: string;
    address?: {
        street?: string;
        city?: string;
        state?: string;
        countryOrRegion?: string;
        postalCode?: string;
    };
    businessHours?: [];
    schedulingPolicy?: {
        timeSlotInterval?: string;
        minimumLeadTime?: string;
        maximumAdvance?: string;
        sendConfirmationsToOwner?: string;
        allowStaffSelection?: string;
    }
}