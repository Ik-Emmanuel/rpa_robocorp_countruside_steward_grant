*** Settings ***
Documentation       The robot for keila to go to gov.uk website to extract information from website
...                 also modify excel sheets.
...

Library             RPA.Excel.Files
Library             Collections
Library             RPA.Browser.Selenium
Library             RPA.Tables
Library             OperatingSystem


*** Variables ***
${INPUT_EXCEL}                          Scheme_codes.xlsx
${OUTPUT_EXCEL}                         ukgov_output_workbook.xlsx
${SHEET_NAME}                           Sheet1
${RECOMMENDED_MANAGEMENT_SELECTOR}      ${EMPTY}    # recommended-management
${prohibited_texts}                     No Item
${CODE_VALUE}                           AB14

@{BENEFIT_IDS}
...                                     how-this-item-will-benefit-the-environment
...                                     how-this-option-will-benefit-the-environment
...                                     how-this-supplement-will-benefit-the-environment


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

        # ${benefit_environment_text}=    Get Text
        # ...    css=#how-this-option-will-benefit-the-environment + *
        ${benefit_environment_text}=    Get All Benefit Paragraph Texts

        ${prohibited_exists_count}=    Is Element Visible    css:#prohibited-activities
        IF    ${prohibited_exists_count}
            ${result}=    Run Keyword And Ignore Error    Get Recommended Management Text
            ${status}=    Set Variable If    '${result[0]}' == 'PASS'    PASS    FAIL
            ${prohibited_texts}=    Set Variable If    '${status}' == 'PASS'    ${result[1]}    None
            Log    ${prohibited_texts}
        ELSE
            Set Variable    ${prohibited_texts}    No item
        END
        Close Browser
        RETURN    ${how_much_ispaid_text}    ${benefit_environment_text}    ${prohibited_texts}
    ELSE
        Close Browser
        RETURN    Not Available    Not Available    Not Available
    END

Get All Benefit Paragraph Texts
    [Documentation]    Extract and process paragraph texts for benefits excluding certain sections.

    # Get all paragraph texts following the element with ID 'how-this-option-will-benefit-the-environment'
    ${p_texts}=    Execute JavaScript
    ...    return Array.from(document.querySelectorAll('#how-this-option-will-benefit-the-environment ~ p'))
    ...    .map(el => el.textContent.trim())
    ...    .filter((text, index, array) => document.querySelectorAll('#how-this-option-will-benefit-the-environment ~ p')[index].nextElementSibling
    ...    && document.querySelectorAll('#how-this-option-will-benefit-the-environment ~ p')[index].nextElementSibling.className !== 'call-to-action');

    # Get the text of elements that follow the call-to-action elements

    ${Aims_section}=    Is Element Visible    css:.call-to-action
    IF    ${Aims_section}
        ${aim_texts}=    Get Text
        ...    css=.call-to-action ~ *

        # Find the index of the aim_texts in p_texts if it exists
        ${index}=    Evaluate    next((i for i, text in enumerate(${p_texts}) if text == """${aim_texts}"""), None)

        # Slice the list of p_texts up to the found index or return the full list if index is not found
        ${p_texts_filtered}=    Evaluate
        ...    list(filter(lambda x: x != "", ${p_texts}))[:${index}] if ${index} != None else list(filter(lambda x: x != "", ${p_texts}))

        # Get and return the full text content for benefit-the-environment
        # ${text}=    Get Text
        # ...    css=#how-this-option-will-benefit-the-environment ~ *
        RETURN    ${p_texts_filtered}.''
    ELSE
        # # Get and return the full text content for benefit-the-environment
        # ${benefit_item}=    Is Element Visible    css:#how-this-item-will-benefit-the-environment
        # IF    ${benefit_item}
        #    ${text}=    Get Text
        #    ...    css=#how-this-item-will-benefit-the-environment + *
        #    RETURN    ${text}
        # ELSE
        #    ${text}=    Get Text
        #    ...    css=#how-this-option-will-benefit-the-environment + *
        #    RETURN    ${text}

        #    # next how-this-supplement-will-benefit-the-environment
        # END

        # Loop through possible IDs and return the first found text
        FOR    ${id}    IN    @{BENEFIT_IDS}
            ${benefit_item}=    Is Element Visible    css=#${id}
            IF    ${benefit_item}
                ${text}=    Get Text    css=#${id} + *
                RETURN    ${text}
            END
        END
    END

Get Recommended Management Text
    ${prohibited_texts}=    Execute JavaScript
    ...    return Array.from(document.querySelector("#recommended-management").previousElementSibling.previousElementSibling.querySelectorAll('li')).map(li => li.textContent).join('\\n');

    RETURN    ${prohibited_texts}

Write To Excel
    [Arguments]    ${Code_value}    ${price_to_pay}    ${benefits_to_environ}    ${prohibited_text}
    # Create Workbook    ${OUTPUT_EXCEL}
    Open Workbook    ${OUTPUT_EXCEL}

    &{table_row1}=    Create Dictionary
    ...    Code=${Code_value}
    ...    Price_To_Pay=${price_to_pay}
    ...    Benefits=${benefits_to_environ}
    ...    Prohibited_Actions=${prohibited_text}

    @{table_data}=    Create List
    ...    ${table_row1}

    Create Table    ${table_data}
    Append Rows To Worksheet    ${table_data}    header=False
    Save Workbook
    Close Workbook
