/*
 * Copyright 2020-2023 University of Oxford and Health and Social Care Information Centre, also known as NHS Digital
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

package uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.workbook

import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiBadRequestException
import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiException
import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiInternalException

import org.apache.poi.EncryptedDocumentException
import org.apache.poi.openxml4j.exceptions.InvalidFormatException
import org.apache.poi.openxml4j.util.ZipSecureFile
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory

class WorkbookHandler {
    /* Adapted from mdm-plugin-excel SimpleExcelDataModelImporterProviderService. */

    static DataFormatter dataFormatter = new DataFormatter()

    String getCellValueAsString(Cell cell) {
        cell ? dataFormatter.formatCellValue(cell).replaceAll(/’/, '\'').replaceAll(/—/, '-').trim() : ''
    }

    List<Map<String, String>> getSheetValues(Map<String,String> intendedColumnMappings, Sheet sheet) {
        List<Map<String, String>> returnValues = []
        Map<String, Integer> expectedSheetColumns = [:]
        Map<String, Integer> otherSheetColumns = [:]
        Row row = sheet.getRow(0)
        int col = 0
        while (row.getCell(col)) {
            String headerText = getCellValueAsString(row.getCell(col))
            boolean found = false
            intendedColumnMappings.each {columnName, regex ->
                if (headerText.replace('*', '').toLowerCase().trim() ==~regex.toLowerCase().trim()) {
                    expectedSheetColumns[columnName] = col
                    found = true
                }
            }
            if (!found) {
                otherSheetColumns[headerText] = col
            }
            col++
        }
        if (expectedSheetColumns.size() != intendedColumnMappings.size()) {
            throw new ApiBadRequestException('EIS03',
                                             "Missing header: ${sheet.getSheetName()} sheet should include the following headers: ${intendedColumnMappings.values()}")
        }

        Iterator<Row> rowIterator = sheet.rowIterator()
        // burn the header row
        rowIterator.next()

        while (rowIterator.hasNext()) {
            row = rowIterator.next()
            Map<String, String> rowValues = [:]
            intendedColumnMappings.each {columnName,regex ->
                String value = getCellValueAsString(row.getCell(expectedSheetColumns[columnName]))
                rowValues[columnName] = value
            }
            otherSheetColumns.keySet().each {columnName ->
                String value = getCellValueAsString(row.getCell(otherSheetColumns[columnName]))
                rowValues[columnName] = value
            }
            returnValues.add(rowValues)
        }
        return returnValues
    }

    Workbook loadWorkbookFromInputStream(String filename, InputStream inputStream) throws ApiException {
        if (!inputStream) throw new ApiInternalException('EFS01', "No inputstream for ${filename}")
        try {
            ZipSecureFile.setMinInflateRatio(0);
            return WorkbookFactory.create(inputStream)
        } catch (EncryptedDocumentException ignored) {
            throw new ApiInternalException('EFS02', "Excel file ${filename} could not be read as it is encrypted")
        } catch (InvalidFormatException ignored) {
            throw new ApiInternalException('EFS03', "Excel file ${filename} could not be read as it is not a valid format")
        } catch (IOException ex) {
            throw new ApiInternalException('EFS04', "Excel file ${filename} could not be read", ex)
        }
    }
}
