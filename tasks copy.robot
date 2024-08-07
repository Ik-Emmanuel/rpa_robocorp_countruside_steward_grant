*** Settings ***
Documentation       The robot for keila to go to gov.uk website to extract information from website
...                 also modify excel sheets.
...

Library             RPA.Excel.Files
Library             Collections
Library             RPA.Browser.Selenium
Library             RPA.Tables


*** Variables ***
${INPUT_EXCEL}                          input_workbook.xlsx
${OUTPUT_EXCEL}                         ukgov_output_workbook.xlsx
${SHEET_NAME}                           Sheet1
${RECOMMENDED_MANAGEMENT_SELECTOR}      ${EMPTY}    # recommended-management
${prohibited_texts}                     No Item
${CODE_VALUE}                           AB14


*** Tasks ***
Countryside Stewardship grant finder
    [Documentation]    Read the contents of an excel sheet and take those values to run web UI automation
    Open Workbook    ${INPUT_EXCEL}
    ${column_values}=    Read Column Values    ${SHEET_NAME}    A
    FOR    ${code}    IN    @{column_values}
        ${how_much_ispaid_text}
        ...    ${benefit_environment_text}
        ...    ${prohibited_texts}=
        ...    Get Website Contents With Codes
        ...    ${code}

        Write To Excel
        ...    ${code}
        ...    ${how_much_ispaid_text}
        ...    ${benefit_environment_text}
        ...    ${prohibited_texts}
    END


*** Keywords ***
Read Column Values
    [Arguments]    ${sheet_name}    ${column_name}
    ${data}=    Read Worksheet As Table    ${sheet_name}
    ${values}=    Get Column Values From Table    ${data}    ${column_name}
    RETURN    ${values}

Get Column Values From Table
    [Arguments]    ${table}    ${column_name}
    ${values}=    Create List
    FOR    ${row}    IN    @{table}
        ${value}=    Evaluate    ${row}.get('${column_name}')
        Append To List    ${values}    ${value}
    END
    RETURN    ${values}

Get Website Contents With Codes
    [Arguments]    ${code_to_lookup}
    Open Available Browser    https://www.gov.uk/countryside-stewardship-grants?keywords=${code_to_lookup}

    ${result_item}=    Is Element Visible    css:.finder-results .gem-c-document-list__item a
    IF    ${result_item}
        Click Element If Visible    css:.finder-results .gem-c-document-list__item a
        # Using the adjacent sibling combinator
        ${how_much_ispaid_text}=    Get Text
        ...    css=#how-much-will-be-paid + *
        # Using the adjacent sibling combinator
        ${benefit_environment_text}=    Get Text
        ...    css=#how-this-option-will-benefit-the-environment + *

        ${prohibited_exists_count}=    Is Element Visible    css:#prohibited-activities
        IF    ${prohibited_exists_count}
            ${prohibited_texts}=    Get Recommended Management Text
        ELSE
            Set Variable    ${prohibited_texts}    No item
        END
        Close Browser
        RETURN    ${how_much_ispaid_text}    ${benefit_environment_text}    ${prohibited_texts}
    ELSE
        RETURN    Not Available    Not Available    Not Available
    END

Get Recommended Management Text
    ${prohibited_texts}=    Run Keyword And Ignore Error
    ...    Execute JavaScript
    ...    return Array.from(document.querySelector("#recommended-management").previousElementSibling.previousElementSibling.querySelectorAll('li')).map(li => li.textContent).join('\\n');
    RETURN    ${prohibited_texts}

# Get Recommended Management Text
#    ${prohibited_texts}=    Run JavaScript Safely
#    RETURN    ${prohibited_texts}

# Run JavaScript Safely
#    ${prohibited_texts}=    Evaluate    None
#    Run Keyword And Ignore Error
#    ...    Execute JavaScript
#    ...    return Array.from(document.querySelector("#recommended-management").previousElementSibling.previousElementSibling.querySelectorAll('li')).map(li => li.textContent).join('\\n');
#    IF    '${prohibited_texts}' == 'None'
#    Set Variable    ${prohibited_texts}    null
#    END
#    RETURN    ${prohibited_texts}

Write To Excel
    [Arguments]    ${Code_value}    ${price_to_pay}    ${benefits_to_environ}    ${prohibited_text}
    # Create Workbook    ${OUTPUT_EXCEL}
    Open Workbook    ${OUTPUT_EXCEL}

    # @{table_column}=    Create List    Code    Price_To_Pay    Benefits    Prohibited_Actions
    # @{table_data}=    Create List    ${code_value}    ${price_to_pay}    ${benefits_to_environ}    ${prohibited_text}
    # &{table}=    Create Dictionary    column=${table_column}    data=${table_data}
    # Create Table    ${_table_}

    &{table_row1}=    Create Dictionary
    ...    Code=${Code_value}
    ...    Price_To_Pay=${price_to_pay}
    ...    Benefits=${benefits_to_environ}
    ...    Prohibited_Actions=${prohibited_text}
    &{table_row2}=    Create Dictionary
    ...    Code=${Code_value}
    ...    Price_To_Pay=${price_to_pay}
    ...    Benefits=${benefits_to_environ}
    ...    Prohibited_Actions=${prohibited_text}

    @{table_data}=    Create List
    ...    ${table_row1}
    ...    ${table_row2}
    Create Table    ${table_data}
    Append Rows To Worksheet    ${table_data}    header=False
    Save Workbook
    Close Workbook
