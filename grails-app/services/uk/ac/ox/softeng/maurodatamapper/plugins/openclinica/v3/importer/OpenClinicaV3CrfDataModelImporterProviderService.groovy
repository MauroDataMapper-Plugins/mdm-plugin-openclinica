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

package uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.importer

import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiInternalException
import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiNotYetImplementedException
import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiUnauthorizedException
import uk.ac.ox.softeng.maurodatamapper.core.model.facet.MetadataAware
import uk.ac.ox.softeng.maurodatamapper.core.provider.importer.parameter.FileParameter
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModelService
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModelType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClass
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataElement
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.DataType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.DataTypeService
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.importer.DataModelImporterProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.ColumnHeaders
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.OpenClinicaV3CrfProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.OpenClinicaV3CrfDefaultDataTypeProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.group.OpenClinicaV3CrfGroupProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.item.OpenClinicaV3CrfItemProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.section.OpenClinicaV3CrfSectionProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.importer.parameters.OpenClinicaV3CrfDataModelImporterParameters
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.workbook.WorkbookHandler
import uk.ac.ox.softeng.maurodatamapper.profile.object.Profile
import uk.ac.ox.softeng.maurodatamapper.profile.provider.ProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.security.User

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

class OpenClinicaV3CrfDataModelImporterProviderService extends DataModelImporterProviderService<OpenClinicaV3CrfDataModelImporterParameters> {

    static final String CONTENT_TYPE = 'application/vnd.ms-excel'

    DataModelService dataModelService
    DataTypeService dataTypeService
    OpenClinicaV3CrfProfileProviderService openClinicaV3CrfProfileProviderService
    OpenClinicaV3CrfSectionProfileProviderService openClinicaV3CrfSectionProfileProviderService
    OpenClinicaV3CrfGroupProfileProviderService openClinicaV3CrfGroupProfileProviderService
    OpenClinicaV3CrfItemProfileProviderService openClinicaV3CrfItemProfileProviderService
    OpenClinicaV3CrfDefaultDataTypeProviderService openClinicaV3CrfDefaultDataTypeProviderService

    WorkbookHandler workbookHandler = new WorkbookHandler()

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
    DataModel importModel(User currentUser, OpenClinicaV3CrfDataModelImporterParameters importerParameters) {
        if (!currentUser) throw new ApiUnauthorizedException('OC301', 'User must be logged in to import model')
        log.info('Importing {} as {}', importerParameters.importFile.fileName, currentUser.emailAddress)
        FileParameter importFile = importerParameters.importFile

        DataModel dataModel
        workbookHandler.loadWorkbookFromInputStream(importFile.fileName, importFile.inputStream).withCloseable {Workbook workbook ->
            // Import CRF as DataModel
            Sheet crfSheet = workbook.getSheet('CRF')
            if (!crfSheet) {
                throw new ApiInternalException('OC302', 'The CRF file must include a sheet "CRF"')
            }

            List<Map<String, String>> crfSheetValues = workbookHandler.getSheetValues(ColumnHeaders.CRF_SHEET_COLUMN_PATTERNS, crfSheet)
            Map<String, String> crfValues = crfSheetValues.findAll {it.crfName && it.version}.last()

            dataModel = new DataModel(label: crfValues.crfName, type: DataModelType.DATA_ASSET)
            dataTypeService.addDefaultListOfDataTypesToDataModel(dataModel, openClinicaV3CrfDefaultDataTypeProviderService.defaultListOfDataTypes)
            addMetadataFromColumnValues(dataModel, openClinicaV3CrfProfileProviderService, ColumnHeaders.CRF_SHEET_COLUMNS, crfValues)

            // Import Sections as tree of Data Classes
            Sheet sectionsSheet = workbook.getSheet('Sections')
            if (!sectionsSheet) {
                throw new ApiInternalException('OC303', 'The CRF file must include a sheet "Sections"')
            }

            List<Map<String, String>> sectionsSheetValues = workbookHandler.getSheetValues(ColumnHeaders.SECTIONS_SHEET_COLUMN_PATTERNS, sectionsSheet)

            List<DataClass> sections = []
            sectionsSheetValues.eachWithIndex {Map<String, String> sectionValues, Integer i ->
                sections << new DataClass(label: sectionValues.sectionLabel, description: sectionValues.sectionTitle, idx: i).tap {DataClass dataClass ->
                    addMetadataFromColumnValues(dataClass, openClinicaV3CrfSectionProfileProviderService, ColumnHeaders.SECTIONS_SHEET_COLUMNS, sectionValues)
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
            Sheet groupsSheet = workbook.getSheet('Groups')
            if (!groupsSheet) {
                throw new ApiInternalException('OC303', 'The CRF file must include a sheet "Groups"')
            }

            Sheet itemsSheet = workbook.getSheet('Items')
            if (!itemsSheet) {
                throw new ApiInternalException('OC303', 'The CRF file must include a sheet "Items"')
            }

            List<Map<String, String>> itemsSheetValues = workbookHandler.getSheetValues(ColumnHeaders.ITEMS_SHEET_COLUMN_PATTERNS, itemsSheet)
            List<Map<String, String>> groupsSheetValues = workbookHandler.getSheetValues(ColumnHeaders.GROUPS_SHEET_COLUMN_PATTERNS, groupsSheet)

            List<DataClass> groups = []
            DataType ocStringDataType = dataModel.dataTypes.find {it.label == 'ST'}
            itemsSheetValues.eachWithIndex {Map<String, String> itemValues, Integer i ->
                DataElement item = new DataElement(label: itemValues.itemName, description: itemValues.descriptionLabel,
                                                   dataType: dataModel.dataTypes.find {it.label == itemValues.dataType} ?: ocStringDataType, idx: i)
                    .tap {DataElement dataElement ->
                        addMetadataFromColumnValues(dataElement, openClinicaV3CrfItemProfileProviderService, ColumnHeaders.ITEMS_SHEET_COLUMNS, itemValues)
                    }
                DataClass parentClass
                if (itemValues.groupLabel) {
                    parentClass = groups.find {DataClass group ->
                        group.label == itemValues.groupLabel && sections.find {DataClass section ->
                            section.label == itemValues.sectionLabel && section.dataClasses?.contains(group)
                        }
                    }
                    if (!parentClass) {
                        DataClass sectionClass = sections.find {it.label == itemValues.sectionLabel}
                        parentClass = new DataClass(label: itemValues.groupLabel, description: itemValues.groupHeader, idx: i)
                        Map<String, String> groupValues = groupsSheetValues.find {Map<String, String> groupValues -> groupValues.groupLabel == itemValues.groupLabel}
                        addMetadataFromColumnValues(parentClass, openClinicaV3CrfGroupProfileProviderService, ColumnHeaders.GROUPS_SHEET_COLUMNS, groupValues)
                        groups << parentClass
                        sectionClass.addToDataClasses(parentClass)
                    }
                } else {
                    parentClass = sections.find {it.label == itemValues.sectionLabel}
                }

                parentClass.addToDataElements(item)
            }
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
    List<DataModel> importModels(User currentUser, OpenClinicaV3CrfDataModelImporterParameters importerParameters) {
        throw new ApiNotYetImplementedException('OC304', 'importModels')
    }
}
