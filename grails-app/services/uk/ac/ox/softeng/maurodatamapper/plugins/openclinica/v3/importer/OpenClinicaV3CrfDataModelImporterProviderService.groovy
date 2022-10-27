/*
 * Copyright 2020-2022 University of Oxford and Health and Social Care Information Centre, also known as NHS Digital
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

package uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.importer

import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiInternalException
import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiNotYetImplementedException
import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiUnauthorizedException
import uk.ac.ox.softeng.maurodatamapper.core.facet.Metadata
import uk.ac.ox.softeng.maurodatamapper.core.model.CatalogueItem
import uk.ac.ox.softeng.maurodatamapper.core.model.facet.MetadataAware
import uk.ac.ox.softeng.maurodatamapper.core.provider.importer.parameter.FileParameter
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModelService
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModelType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClass
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.importer.DataModelImporterProviderService
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.importer.parameter.DataModelFileImporterProviderServiceParameters
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datamodel.provider.importer.SimpleExcelDataModelImporterProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.OpenClinicaV3CrfProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.section.OpenClinicaV3CrfSectionProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.profile.domain.ProfileField
import uk.ac.ox.softeng.maurodatamapper.profile.object.JsonProfile
import uk.ac.ox.softeng.maurodatamapper.profile.object.Profile
import uk.ac.ox.softeng.maurodatamapper.profile.provider.ProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.security.User

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

class OpenClinicaV3CrfDataModelImporterProviderService extends DataModelImporterProviderService<DataModelFileImporterProviderServiceParameters> {

    static final String CONTENT_TYPE = 'application/vnd.ms-excel'
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

    DataModelService dataModelService
    SimpleExcelDataModelImporterProviderService simpleExcelDataModelImporterProviderService
    OpenClinicaV3CrfProfileProviderService openClinicaV3CrfProfileProviderService
    OpenClinicaV3CrfSectionProfileProviderService openClinicaV3CrfSectionProfileProviderService

    @Override
    String getDisplayName() {
        'OpenClinica 3.x CRF (XLS) Importer'
    }

    @Override
    String getVersion() {
        getClass().getPackage().getSpecificationVersion() ?: 'SNAPSHOT'
    }

    @Override
    Boolean canImportMultipleDomains() {
        false
    }

    @Override
    String getNamespace() {
        'uk.ac.ox.softeng.maurodatamapper.plugins.excel.openclinica'
    }

    @Override
    Boolean handlesContentType(String contentType) {
        contentType.equalsIgnoreCase(CONTENT_TYPE)
    }

    @Override
    DataModel importModel(User currentUser, DataModelFileImporterProviderServiceParameters importerParameters) {
        if (!currentUser) throw new ApiUnauthorizedException('OC301', 'User must be logged in to import model')
        log.info('Importing {} as {}', importerParameters.importFile.fileName, currentUser.emailAddress)
        FileParameter importFile = importerParameters.importFile

        DataModel dataModel
        simpleExcelDataModelImporterProviderService.loadWorkbookFromInputStream(importFile.fileName, importFile.inputStream).withCloseable {Workbook workbook ->
            // Import CRF as DataModel
            Sheet crfSheet = workbook.getSheet('CRF')

            if (!crfSheet) {
                throw new ApiInternalException('OC302', 'The CRF file must include a sheet "CRF"')
            }

            List<Map<String, String>> crfSheetValues = simpleExcelDataModelImporterProviderService.getSheetValues(CRF_SHEET_COLUMNS, crfSheet)
            Map<String, String> crfValues = crfSheetValues.findAll {it.crfName && it.version}.last()

            dataModel = new DataModel(label: crfValues.crfName, type: DataModelType.DATA_ASSET)
            addMetadataFromColumnValues(dataModel, openClinicaV3CrfProfileProviderService, CRF_SHEET_COLUMNS, crfValues)

            // Import Sections as tree of Data Classes
            Sheet sectionsSheet = workbook.getSheet('Sections')

            if (!sectionsSheet) {
                throw new ApiInternalException('OC303', 'The CRF file must include a sheet "Sections"')
            }

            List<Map<String, String>> sectionsSheetValues = simpleExcelDataModelImporterProviderService.getSheetValues(SECTIONS_SHEET_COLUMNS, sectionsSheet)

            List<DataClass> sections = []
            sectionsSheetValues.each {Map<String, String> sectionValues ->
                sections << new DataClass(label: sectionValues.sectionLabel, description: sectionValues.sectionTitle).tap {DataClass dataClass ->
                    addMetadataFromColumnValues(dataClass, openClinicaV3CrfSectionProfileProviderService, SECTIONS_SHEET_COLUMNS, sectionValues)
                }
            }
            sectionsSheetValues.eachWithIndex {Map<String, String> sectionValues, Integer i ->
                if (sectionValues.parentSection) {
                    DataClass parentClass = sections.find {it.label == sectionValues.parentSection}
                    if (parentClass) parentClass.addToDataClasses(sections[i])
                    else dataModel.addToDataClasses(sections[i])
                } else {
                    dataModel.addToDataClasses(sections[i])
                }
            }

            // Import Items as Data Elements and Groups as optional nested Data Classes

        }

        dataModelService.checkImportedDataModelAssociations(currentUser, dataModel)
        dataModel
    }

    MetadataAware addMetadataFromColumnValues(MetadataAware metadataAware, ProfileProviderService profileProviderService, Map<String, String> columnMap,
                                              Map<String, String> sheetValues) {
        String metadataNamespace = profileProviderService.getMetadataNamespace()
        Profile profile = profileProviderService.getNewProfile()
        columnMap.each {Map.Entry<String, String> entry ->
            profile.allFields.find {it.metadataPropertyName.equalsIgnoreCase(entry.value)}?.with {
                if (sheetValues[entry.key])
                    metadataAware.addToMetadata namespace: metadataNamespace, key: it.metadataPropertyName, value: sheetValues[entry.key]
            }
        }
        metadataAware
    }

    @Override
    List<DataModel> importModels(User currentUser, DataModelFileImporterProviderServiceParameters importerParameters) {
        throw new ApiNotYetImplementedException('OC304', 'importModels')
    }
}
