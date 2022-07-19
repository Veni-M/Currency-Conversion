*** Settings ***
Documentation       Template robot main suite.

# library name for excel
# Library   RPA.Excel.Application
Library    RPA.Excel.Files
Library    RPA.Browser.Selenium     auto_close=${False}
Library    RPA.Email.ImapSmtp    smtp_server=smtp.gmail.com    smtp_port=587


# declaring variables
*** Variables ***
${excelFileName}   InputItems.xlsx
${rowIndex}      2
${counter}   1
${USERNAME}     sankaravenie5@gmail.com
${PASSWORD}     oymadhiqcrozavsc
${RECIPIENT}    sankaravenie5@gmail.com

*** Keywords ***
read excel
   Open Workbook    ${excelFileName}
   # reading excel and assigning in the varaible rows
   ${rows}=    Read Worksheet   Sheet1   True
   # iterating through each row 
   Set Worksheet Value    1    5    Converted Amount
   Save Workbook
   FOR  ${eachValue}  IN   @{rows}
       Log   ${eachValue}[From]
      #  giving amount
      Click Element     id=amount
      Press Keys        id=amount    CTRL+a+BACKSPACE
      Input Text        id=amount    ${eachValue}[Amout]    clear=True
      Click Element    id=midmarketFromCurrency
      # giving Currency From
      Input Text    id=midmarketFromCurrency   ${eachValue}[From]    
      Press Keys    id=midmarketFromCurrency   ENTER
      Click Element    id=midmarketToCurrency
      # giving TO currency
      Input Text    id=midmarketToCurrency    ${eachValue}[To]
      # selecting the currency
      Click Element    id=midmarketToCurrency
      # pressing enter
      Press Keys    id=midmarketToCurrency  ENTER
      # clicking convert button
      Run Keyword And Ignore Error  Click Button When Visible    //button[@type="submit"]   
      Run Keyword And Ignore Error  Click Element When Visible  id=yie-close-button-985b8a0a-1505-56d6-9b3a-5ee95bd0e8b5   
      ${finalOutputValue}  Get Text    //p[@class="result__BigRate-sc-1bsijpp-1 iGrAod"]
      Log   ${finalOutputValue}
      Open Workbook    ${excelFileName}
      Set Worksheet Value    ${rowIndex}    5    ${finalOutputValue}
      Save Workbook
      ${rowIndex}=       Evaluate    ${rowIndex} + ${counter}
      
   END
   

   

*** Tasks ***
opening browser
  
   #  open available browser - install necessary driven and opens the browser
   # open browser - manual installation required to open the browser
    Open Available Browser  https://www.xe.com/  maximized=True  alias=FirstBrowser


reading a excel
    read excel

Send test email
    Authorize    account=${USERNAME}    password=${PASSWORD}
    Send Message    sender=${USERNAME}
    ...    recipients=${RECIPIENT}
    ...    subject=Converted Amount
    ...    body=Hi, I have attached the converted amount excel. Please find it below
    ...    attachments=InputItems.xlsx




    
    
    
    

    
    
    
    