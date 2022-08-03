*** Settings ***
Documentation   Excel Operation Keyword File
Library         RPA.Tables
Library         RPA.Excel.Files
Library         RPA.Excel.Application
Library         String
Library         RPA.PDF
Library         RPA.Desktop
  
*** Keywords ***
Read Excel File WorkSheet As Table
    [Arguments]      ${EXCEL_FILE}    ${WORK_SHEET_NAME}
    RPA.Excel.Files.Open Workbook    ${EXCEL_FILE}
    ${worksheet} =    Read Worksheet   name=${WORK_SHEET_NAME}   header=${TRUE}
    RPA.Excel.Files.Close Workbook
    ${policies} =       Create table     ${worksheet}
    [Return]         ${policies}
    
    
Create New WorkBook From Map
    [Arguments]      ${EXCEL_FILE_NAME} 
    RPA.Excel.Files.Create Workbook  ${EXCEL_FILE_NAME}   sheet_name=MyOrders 
    FOR    ${index}    IN RANGE    15
        ${ret} =	Generate Random String
        &{row} =      Create Dictionary
        ...           Order ID   ${index}
        ...           Amount   ${index * 25}
        ...           Invoice No    ${ret}
        Append Rows to Worksheet  ${row}    header=${TRUE}    
    END
    Save Workbook
    
    
Iterate WorkSheets From Workbook
    [Arguments]      ${EXCEL_FILE}
    RPA.Excel.Files.Open Workbook    ${EXCEL_FILE}
    @{work_sheets} =     List Worksheets
    FOR  ${work_sheet}  IN   @{work_sheets}
        ${wsheet_data} =  Read Worksheet   ${work_sheet}  
        ${rows} =         Get Length  ${wsheet_data}
        ${emp_row} =      Find Empty Row  ${work_sheet}
        Log     Worksheet '${work_sheet}' contain ${rows} rows with first empty row at ${emp_row}
    END
    Close Workbook
      
Export Workbook as PDF
    [Arguments]    ${EXCEL_FILE}
    RPA.Excel.Files.Open Workbook    ${EXCEL_FILE}
    Export as PDF    ${OUTPUT_DIR}${/}myfile.pdf    ${EXCEL_FILE}


Open Files
    #[Arguments]    ${EXCEL_FILE}
    #Open File    ${EXCEL_FILE}
    Open File    ./DataSets/sampledatainsurance.xlsx   
    [Arguments]    ${EXCEL_FILE}
    RPA.Excel.Files.Open Workbook    ${EXCEL_FILE}
    RPA.Excel.Application.Open Application     TRUE
    Open File    ${OUTPUT_DIR}${/}myfile.pdf