*** Settings ***
Documentation       Template robot main suite.

Library    RPA.Excel.Files
Library    RPA.Browser.Selenium
Library    RPA.Dialogs
Library    OperatingSystem
Library    RPA.Tables
Library    RPA.RobotLogListener
Library    RPA.PDF
Library    RPA.Archive

*** Tasks ***
Minimal task
    Open Browser
    Close the annoying modal
    @{read_data}=    Fill the Excel
    Fill the Excel
    Create a zip folder of all pdf 
    # FOR    ${row}    IN    @{read_data}
    #     #   Close the annoying modal
    #       Fill the form on website    ${row}
    #     #   Preview the robot
    #     #   Submit the order
    #     #   ${Store}    Store the receipt as pdf file    $row
    #     #   Take screenshot of image    $row
    #     #   Embed screenshot of image to pdf file    $row
    #     #   Click on order another robot
    # END
   

*** Keywords ***

Open Browser
    Open Available Browser    https://robotsparebinindustries.com/#/robot-order

Fill the form on website
   
    [Arguments]    ${row}
    Select From List By Value    head    ${row}[Head]
    Sleep    3
    Select Radio Button    body    ${row}[Body]
    Sleep    3
    Input Text    css:input[class='form-control']    ${row}[Legs]
    Sleep    3
    Input Text    address    ${row}[Address]
    Preview the robot
    Submit the order
    Sleep    3
    Wait Until Element Is Visible    id:order-completion
    ${Store}=    Get Element Attribute    id:order-completion    outerHTML
    
    Html To Pdf    ${Store}    pdf${/}${row}[Order number].pdf    overwrite=True
    Screenshot    //*[@id="robot-preview-image"]    ${OUTPUT_DIR}${/}images${/}${row}[Order number].png        
    ${files}=    Create List
     ...     pdf${/}${row}[Order number].pdf
     ...     images${/}${row}[Order number].png:align=center

             Add Files To PDF    ${files}    newdoc${row}[Order number].pdf    append=True
     Click Button    order-another
     Close the annoying modal
Close the annoying modal
     Click Element    xpath://*[@id="root"]/div/div[2]/div/div/div/div/div/button[1]


Fill the Excel
    ${read_data}=    Read table from CSV    orders.csv    header=True
     FOR    ${row}    IN    @{read_data}
         Fill the form on website    ${row}
        
    END   

Preview the robot
      Click Button    preview    

Submit the order
    Click Button    order


Store the receipt as pdf file    
    [Arguments]    ${row}
     Wait Until Element Is Visible    id:order-completion
     ${Store}=    Get Element Attribute    id:order-completion    outerHTML
     ${SubmitButtonExist}=    Get Element Attribute    id:order-completion    outerHTML
     Html To Pdf    ${Store}    pdf${/}${row}[Order number].pdf    overwrite=True

Take screenshot of image
    [Arguments]    ${row}
    Screenshot    //*[@id="robot-preview-image"]    ${OUTPUT_DIR}${/}images${/}${row}[Order number].png        

Embed screenshot of image to pdf file
    [Arguments]    ${row}
    ${files}=    Create List
    ...     pdf${/}${row}[Order number].pdf
    ...     images${/}${row}[Order number].png:align=center

            Add Files To PDF    ${files}    newdoc${row}[Order number].pdf    append=True

Click on order another robot
    Sleep    3
    Click Button    order-another

Create a zip folder of all pdf
    Archive Folder With Zip     ${OUTPUT_DIR}    ReceiptPDF.Zip    true    include=*.pdf
    