import { getADMappingList } from "../services/apiservices";
import * as xlsx from "xlsx";
import FileSaver from 'file-saver';

export interface IExportExcelProps {
    data: any[];
    fileName: string;
}

export function ExportExcel(props: IExportExcelProps){
    async function exportADAccountMapping() {
        const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
        const fileExtension = '.xlsx';
        const csvData = props.data
        if(csvData) {
          const ws = xlsx.utils.json_to_sheet(csvData);
          const wb = { Sheets: { 'data': ws }, SheetNames: ['data'] };
          const excelBuffer = xlsx.write(wb, { bookType: 'xlsx', type: 'array' });
          const data = new Blob([excelBuffer], {type: fileType});
          FileSaver.saveAs(data, props.fileName + fileExtension);
        }
    }

    return(exportADAccountMapping)
}