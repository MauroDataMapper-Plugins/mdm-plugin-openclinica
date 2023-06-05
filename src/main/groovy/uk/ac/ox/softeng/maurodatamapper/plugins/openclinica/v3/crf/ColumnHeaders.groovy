/*
 * Copyright 2020-2023 University of Oxford and NHS England
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 * SPDX-License-Identifier: Apache-2.0
 */
package uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf

class ColumnHeaders {

    /* Map field names to spreadsheet column names */
    static final Map<String, String> CRF_SHEET_COLUMNS = [crfName           : 'CRF_NAME',
                                                          version           : 'VERSION',
                                                          versionDescription: 'VERSION_DESCRIPTION',
                                                          revisionNotes     : 'REVISION_NOTES']

    static final Map<String, String> SECTIONS_SHEET_COLUMNS = [sectionLabel : 'SECTION_LABEL',
                                                               sectionTitle : 'SECTION_TITLE',
                                                               subtitle     : 'SUBTITLE',
                                                               instructions : 'INSTRUCTIONS',
                                                               pageNumber   : 'PAGE_NUMBER',
                                                               parentSection: 'PARENT_SECTION']

    static final Map<String, String> GROUPS_SHEET_COLUMNS = [groupLabel        : 'GROUP_LABEL',
                                                             groupLayout       : 'GROUP_LAYOUT',
                                                             groupHeader       : 'GROUP_HEADER',
                                                             groupRepeatNumber : 'GROUP_REPEAT_NUMBER',
                                                             groupRepeatMax    : 'GROUP_REPEAT_MAX',
                                                             groupDisplayStatus: 'GROUP_DISPLAY_STATUS']

    static final Map<String, String> ITEMS_SHEET_COLUMNS = [itemName                    : 'ITEM_NAME',
                                                            descriptionLabel            : 'DESCRIPTION_LABEL',
                                                            leftItemText                : 'LEFT_ITEM_TEXT',
                                                            units                       : 'UNITS',
                                                            rightItemText               : 'RIGHT_ITEM_TEXT',
                                                            sectionLabel                : 'SECTION_LABEL',
                                                            groupLabel                  : 'GROUP_LABEL',
                                                            header                      : 'HEADER',
                                                            subheader                   : 'SUBHEADER',
                                                            parentItem                  : 'PARENT_ITEM',
                                                            columnNumber                : 'COLUMN_NUMBER',
                                                            pageNumber                  : 'PAGE_NUMBER',
                                                            questionNumber              : 'QUESTION_NUMBER',
                                                            responseType                : 'RESPONSE_TYPE',
                                                            responseLabel               : 'RESPONSE_LABEL',
                                                            responseOptionsText         : 'RESPONSE_OPTIONS_TEXT',
                                                            responseValuesOrCalculations: 'RESPONSE_VALUES_OR_CALCULATIONS',
                                                            responseLayout              : 'RESPONSE_LAYOUT',
                                                            defaultValue                : 'DEFAULT_VALUE',
                                                            dataType                    : 'DATA_TYPE',
                                                            widthDecimal                : 'WIDTH_DECIMAL',
                                                            validation                  : 'VALIDATION',
                                                            validationErrorMessage      : 'VALIDATION_ERROR_MESSAGE',
                                                            phi                         : 'PHI',
                                                            required                    : 'REQUIRED',
                                                            itemDisplayStatus           : 'ITEM_DISPLAY_STATUS',
                                                            simpleConditionalDisplay    : 'SIMPLE_CONDITIONAL_DISPLAY']

    /* Map field names to spreadsheet column regex patterns */
    static final Map<String, String> CRF_SHEET_COLUMN_PATTERNS = CRF_SHEET_COLUMNS
    static final Map<String, String> SECTIONS_SHEET_COLUMN_PATTERNS = SECTIONS_SHEET_COLUMNS
    static final Map<String, String> GROUPS_SHEET_COLUMN_PATTERNS = GROUPS_SHEET_COLUMNS.clone().tap {
        groupRepeatNumber = 'GROUP_REPEAT_NUM(BER)?'
    }
    static final Map<String, String> ITEMS_SHEET_COLUMN_PATTERNS = ITEMS_SHEET_COLUMNS
}