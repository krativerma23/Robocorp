*** Settings ***
Documentation       QAD Task.

Library            RPA.Excel.Files
Library            RPA.Browser.Selenium    auto_close=${False}
Library            RPA.JSON
Library            RPA.Desktop
Library            RPA.FileSystem
Library            DateTime
Library            RPA.Outlook.Application
Library            RPA.Tables
Library            Collections
Library            String
Library            RPA.Robocorp.Vault
Library            RPA.Outlook.Application
Task Setup         RPA.Outlook.Application.Open Application
Suite Teardown     Quit Application
Library            Process



*** Variables ***

${recipient}        krati.verma@yash.com  
# ${Read_data}        ${TRUE}

*** Tasks ***
QAD task
      
    TRY
        Read config file 
        ${InputFile_exists}=    Does file exist    ${config_data}[Input File path]
        IF  ${InputFile_exists} 
            ${file}=    Find Files  ${config_data}[Input File path]
            TRY
                Read data from input excel

            EXCEPT    
                    Log To Console    Read input file error
            END
        ELSE
                  Send Message       recipients=${RECIPIENT}
    ...          subject=Input File Error
    ...          body=Hello User, Input File not found    
        END    
            
    EXCEPT 
             Send Message       recipients=${RECIPIENT}
    ...      subject=Config File Error
    ...      body=Hello User, Config file not found 
    END

    
*** Keywords ***


Read config file 
    ${ConfigfileStartDate}=    Get Current Date
    ${config_exists}=    Does file exist    D:/Robocorp/QAD/configfile.json
    IF     ${config_exists}
            ${auth}=    Load JSON from file    D:/Robocorp/QAD/configfile.json
            Set Global Variable    ${config_data}    ${auth}
            ${keys}=    Get Dictionary Keys    ${config_data}
            Set Global Variable    ${configfile_keys}    ${keys}
            # ${keys}=    Get Dictionary Keys    ${config_data}
            Set Global Variable    ${config_data}    ${auth}
            Set Global Variable    ${configfile_keys}    ${keys}
            Log To Console    ${configfile_keys}
           ${len1}=    Get Length    ${auth}
        
        IF    ${len1}==0
            Open Workbook         D:/Robocorp/QAD/Business Exception.xlsx
            Set Active Worksheet    Sheet1
            Set Cell Value    3    A    Config File Error
            Set Cell Value    3    B    Config File is empty 
            Set Cell Value    3    C    ${ConfigfileStartDate}
            ${ConfigFileEndDate}=    Get Current Date
            Set Cell Value    3    D    ${ConfigFileEndDate}
            Save Workbook
            # Fail
    ELSE
        ${auth}=    Load JSON from file    D:/Robocorp/QAD/configfile.json
        Set Global Variable    ${config_data}    ${auth}
        ${keys}=    Get Dictionary Keys    ${config_data}
        Set Global Variable    ${configfile_keys}    ${keys}
        # Log To Console    Config file is not empty
    END
    ELSE
           Open Workbook    D:/Robocorp/QAD/Business Exception.xlsx
            Set Active Worksheet    Sheet1
            ${range}=    Clear Cell Range    A2:D5
            Log To Console    ${range}
            Set Cell Value    2    A    Config File Error
            Set Cell Value    2    B    Config File does not exists
            Set Cell Value    2    C    ${ConfigfileStartDate}
            ${ConfigFileEndDate}=    Get Current Date
            Set Cell Value    2    D    ${ConfigFileEndDate}
            Save Workbook
    END
    
    
    
Open the website
    TRY
        ${LaunchChromeStartTime}=    Get Current Date
        Open Available Browser
        ...     ${config_data}[URL]    browser_selection=Chrome
        ...    maximized=True    
        # Open Browser    ${config_data}[URL]    Chrome
    EXCEPT
            Open Workbook    ${config_data}[Technical Exception]
            Set Active Worksheet    Sheet1
            Clear Cell Range    A2:D6
            Set Cell Value    2    A    Launch Chrome
            Set Cell Value    2    B    ${LaunchChromeStartTime}
            ${ChromeEndTime}=    Get Current Date
            Set Cell Value    2    C    ${ChromeEndTime} 
            Set Cell Value    2   D    Failed
            Set Cell Value    2   E    Unable to open QAD website
            Save Workbook
            Log To Console    Unable to open QAD website
            Fatal Error
            
    END

Log into the website
    
    TRY
    ${LoginStartTime}=    Get Current Date
    ${secrets}=     Get Secret    ${config_data}[QAD_Credentials]
    Input Text    username    ${secrets}[Username]
    Input Password    password    ${secrets}[Password]

    EXCEPT
        Open Workbook    ${config_data}[Technical Exception]
        Set Active Worksheet    Sheet1
        Set Cell Value    3    A    Login Error
        Set Cell Value    3    B    ${LoginStartTime}
        ${LoginEndTime}=    Get Current Date
        Set Cell Value    3    C    ${LoginEndTime} 
        Set Cell Value    3   D    Failed
        Set Cell Value    3   E    Username or password is invalid or not found
        Save Workbook
        Log To Console    Username or password not found
        Fatal Error
    END
    

    TRY
        Submit Form
    EXCEPT
        Log To Console    Sign in button not found
    END
    TRY
        Wait Until Page Contains Element    id:rightMenu
        
    EXCEPT
        Log To Console    Page doesn't load properly
    END
    
Click on the developer option
    Sleep    3
    Click Element    xpath://*[@id="menuSelectorDropdown"]/span/span[1]
    Sleep    2
    Click Element    xpath://*[@id="menuSelector_listbox"]/li[10]
    Sleep    2
    Click Element    xpath://*[@id="kMenuWebShell"]/li[2]/span
    Sleep    1
    Click Element    xpath://*[@id="kMenuWebShell_mn_active"]/div/ul/li[1]
    Sleep    3
    Click Element    xpath://*[@id="ToolBtnNew"]
    

Get data from input excel
            Open Workbook    ${config_data}[Input File path]
            ${data}=    Read Worksheet As Table    header=True
            Close Workbook
            FOR    ${data1}    IN    @{data}
                Fill the form available on app    ${data1}
            END
            
    
Fill the form available on app
    [Arguments]    ${data1}
    ${FormField1Start}=    Get Current Date
    ${App1}=    Set Variable    ${data1}[App]
    Log To Console    ${App1}
    
    IF    "${data1}[App]"== "None"
            Open Workbook    ${config_data}[Technical Exception]
            Set Active Worksheet    Sheet1
            Set Cell Value    4    A    First Field
            Set Cell Value    4    B    ${FormField1Start}
            ${FormField1End}=    Get Current Date
            Set Cell Value    4    C    ${FormField1End} 
            Set Cell Value    4   D    Failed
            Set Cell Value    4   E    Data not found for App Field
            Save Workbook
            Log To Console    App Field not found
        
    ELSE
        Sleep    3
        Click Element    xpath://*[@id="ToolBtnNew"]
        Sleep    3
        Input Text    xpath://*[@id="AppNameAutoField"]    ${data1}[App]
    END
    Sleep    2
        ${DescriptionStartTime}=    Get Current Date
        IF    "${data1}[Description]"== "None"
            Open Workbook    ${config_data}[Technical Exception]
            Set Active Worksheet    Sheet1
            Set Cell Value    5    A    Description Field
            Set Cell Value    5    B    ${DescriptionStartTime}
            ${DescriptionEndTime}=    Get Current Date
            Set Cell Value    5    C    ${DescriptionEndTime} 
            Set Cell Value    5   D    Failed
            Set Cell Value    5   E    Data not found for Description Field
            Save Workbook
            Log To Console    Description Field not found
            
        ELSE
           
            Sleep    3
            Input Text    xpath://*[@id="DescriptionAutoField"]    ${data1}[Description]  
            
        
    END
    Sleep    2

    TRY
        ${ReleasedStartTime}=    Get Current Date
        IF    "${data1}[Released]"== "None"
            Open Workbook    ${config_data}[Technical Exception]
            Set Active Worksheet    Sheet1
            Set Cell Value    6    A    Released Field
            Set Cell Value    6    B    ${ReleasedStartTime}
            ${ReleasedEndTime}=    Get Current Date
            Set Cell Value    6    C    ${ReleasedEndTime} 
            Set Cell Value    6   D    Failed
            Set Cell Value    6   E    Data not found for Released Field
            Save Workbook
            Log To Console    Released Field not found
            
        ELSE
            Sleep    3
            Select Checkbox    xpath://*[@id="IsReleasedAutoField"]
        END
        
    EXCEPT    
        Log To Console    Released checkbox not found
        
    END
    Sleep    2

    TRY
        Click Button   id:ToolBtnSave
    EXCEPT    
        Log To Console    Save button not found
        
    END
    Sleep    2
    

    TRY
        Click Button    id:qModalDialogConfirm  
    EXCEPT    
        Log To Console    Confirm button not found
        
    END
    Sleep    3


    # Development
    Click Element     xpath://*[@id="kMenuWebShell"]/li[2]/span

    Sleep    3
    # App 
    Click Element    xpath://*[@id="kMenuWebShell_mn_active"]/div/ul/li[1]
    Sleep    3
    # Leave Button
    TRY
        Sleep    3
        Click Element When Visible    id:qModalDialogConfirm
    EXCEPT    
        
        Click Element    xpath://*[@id="ToolBtnNew"]
        Sleep    3
    END
    # Fatal Error
    
Read data from input excel
    ${InputFileDate}=    Get Current Date
    ${InputFile_exists}=    Does file exist    ${config_data}[Input File path]
    IF    ${InputFile_exists}
        Open Workbook    ${config_data}[Input File path]
        ${data}=    Read Worksheet As Table    
        ${type}=    Evaluate    type($data)
        Close Workbook
        ${InputFile_data}=    Get Length    ${data}
        Log To Console    ${InputFile_data}
            IF    ${InputFile_data}==0
                Open Workbook    ${config_data}[Business Exception]
                Set Active Worksheet    Sheet1
                Set Cell Value    5    A    Input File Error
                Set Cell Value    5    B    Input File is empty
                Set Cell Value    5    C    ${InputFileDate}
                ${InputFileEndDate}=    Get Current Date
                Set Cell Value    5    D    ${InputFileEndDate}
                Save Workbook
            ELSE
                    ${header1}=    RPA.Tables.Get Table Cell    ${data}    0    A
                    ${header2}=    RPA.Tables.Get Table Cell    ${data}    0    B
                    ${header3}=    RPA.Tables.Get Table Cell    ${data}    0    C
                    ${header4}=    RPA.Tables.Get Table Cell    ${data}    0    D
                    ${header5}=    RPA.Tables.Get Table Cell    ${data}    0    E
                    ${header6}=    RPA.Tables.Get Table Cell    ${data}    0    F
                    IF  '"${header1}" == "${configfile_keys}[0]" AND "${header2}" == "${configfile_keys}[2]" AND "${header3}"== "${configfile_keys}[1]" AND "${header4}"== "${configfile_keys}[4]" AND "${header5}"== "${configfile_keys}[8]" AND "${header6}"== "${configfile_keys}[7]"'
                        Open the website
                        Log into the website
                        Click on the developer option
                        Get data from input excel
                        Log To Console    Config file and input file headers are same
                    ELSE
                        Open Workbook    ${config_data}[Business Exception]
                        Set Active Worksheet    Sheet1
                        Set Cell Value    6    A    Input File Error
                        Set Cell Value    6    B    Input File is invalid
                        Set Cell Value    6    C    ${InputFileDate}
                        ${InputFileEndDate}=    Get Current Date
                        Set Cell Value    6    D    ${InputFileEndDate}
                        Save Workbook
                    END
                        
            END
        
    ELSE
        Open Workbook    ${config_data}[Business Exception]
        Set Active Worksheet    Sheet1
        Set Cell Value    4    A    Input File Error
        Set Cell Value    4    B    Input File does not exists
        Set Cell Value    4    C    ${InputFileDate}
        ${InputFileEndDate}=    Get Current Date
        Set Cell Value    4    D    ${InputFileEndDate}
        Save Workbook
           
    END

Send mail to user
    Send Message       recipients=${RECIPIENT}
    ...                subject=Business and Technical Exception file
    ...                body=Hello User, Please find the attached business and technical exception file
    ...                attachments=${config_data}[Business Exception],${config_data}[Technical Exception]