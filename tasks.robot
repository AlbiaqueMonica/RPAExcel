*** Settings ***
Documentation   Excel File Related Keyword Examples
Resource        KeywordLibrary/excel.robot


*** Tasks ***
Read Policies As Table
    ${policydata} =  Read Excel File WorkSheet As Table      ./DataSets/sampledatainsurance.xlsx   PolicyData
    
    @{table_dim} =  Get Table Dimensions  ${policydata}  
    ${row_value} =  Get Table Row       ${policydata}   ${0}   False    
    
    FOR    ${i}    IN RANGE    ${table_dim}[0]
        ${row_data} =  Get Table Row       ${policydata}   ${i}   False
        Log  ${row_data}
    END

*** Tasks ***
Create New WorkSheet Tasks
     Create New WorkBook From Map   my_new_wbook.xlsx       
     ${policydata} =  Read Excel File WorkSheet As Table      my_new_wbook.xlsx    MyOrders

*** Tasks ***
Iterate WorkSheets Example
    Iterate WorkSheets From Workbook   ./DataSets/sampledatainsurance.xlsx
 
    
*** Tasks ***
Export To PDF Example
    Export Workbook as PDF  my_new_wbook.xlsx
Open Files Example
    Open Files    
    