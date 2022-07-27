import { MutableRefObject } from "react";

export interface IEventListProps {
    userMail: string;
    getEventData?:() => void;
}