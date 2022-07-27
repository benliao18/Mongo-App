export interface ITableModel {
    key: string;
    items: any[];
    className:string;
}

export interface ILogs {
    id?: string;
    FunctionName: string;
    Executer: string;
    ExecuteTime?: string;
}