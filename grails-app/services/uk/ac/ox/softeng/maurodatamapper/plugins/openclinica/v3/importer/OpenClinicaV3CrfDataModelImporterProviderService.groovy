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
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.EnumerationType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.PrimitiveType
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.importer.DataModelImporterProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.forms.datatype.FormDataTypeProvider
import uk.ac.ox.softeng.maurodatamapper.plugins.forms.question.FormQuestionProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.forms.section.FormSectionProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.ColumnHeaders
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.OpenClinicaV3CrfProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.OpenClinicaV3CrfDefaultDataTypeProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.group.OpenClinicaV3CrfGroupProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.item.OpenClinicaV3CrfItemProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.section.OpenClinicaV3CrfSectionProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.importer.parameters.OpenClinicaV3CrfDataModelImporterParameters
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.workbook.WorkbookHandler
import uk.ac.ox.softeng.maurodatamapper.profile.domain.ProfileField
import uk.ac.ox.softeng.maurodatamapper.profile.object.JsonProfile
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
    FormSectionProfileProviderService formSectionProfileProviderService
    FormQuestionProfileProviderService formQuestionProfileProviderService
    FormDataTypeProvider formDataTypeProvider

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
        if (!currentUser) throw new ApiUnauthorizedException('OC3I01', 'User must be logged in to import model')
        log.info('Importing {} as {}', importerParameters.importFile.fileName, currentUser.emailAddress)
        FileParameter importFile = importerParameters.importFile

        DataModel dataModel
        workbookHandler.loadWorkbookFromInputStream(importFile.fileName, importFile.inputStream).withCloseable {Workbook workbook ->
            // Import CRF as DataModel
            Sheet crfSheet = workbook.getSheet('CRF')
            if (!crfSheet) {
                throw new ApiInternalException('OC3I02', 'The CRF file must include a sheet "CRF"')
            }

            List<Map<String, String>> crfSheetValues = workbookHandler.getSheetValues(ColumnHeaders.CRF_SHEET_COLUMN_PATTERNS, crfSheet)
            Map<String, String> crfValues = crfSheetValues.findAll {it.crfName && it.version}.last()

            dataModel = new DataModel(label: crfValues.crfName, modelVersionTag: crfValues.version, type: DataModelType.DATA_ASSET)
            if (crfValues.versionDescription) dataModel.description = crfValues.versionDescription
            dataTypeService.addDefaultListOfDataTypesToDataModel(dataModel, formDataTypeProvider.defaultListOfDataTypes)
            addMetadataFromColumnValues(dataModel, openClinicaV3CrfProfileProviderService, ColumnHeaders.CRF_SHEET_COLUMNS, crfValues)

            // Import Sections as tree of Data Classes
            Sheet sectionsSheet = workbook.getSheet('Sections')
            if (!sectionsSheet) {
                throw new ApiInternalException('OC3I03', 'The CRF file must include a sheet "Sections"')
            }

            List<Map<String, String>> sectionsSheetValues =
                workbookHandler.getSheetValues(ColumnHeaders.SECTIONS_SHEET_COLUMN_PATTERNS, sectionsSheet).findAll {it.sectionLabel}

            List<DataClass> sections = []
            sectionsSheetValues.eachWithIndex {Map<String, String> sectionValues, Integer i ->
                sections << new DataClass(label: sectionValues.sectionLabel, idx: i).tap {DataClass dataClass ->
                    if (sectionValues.sectionTitle) description = sectionValues.sectionTitle
                    addMetadataFromColumnValues(dataClass, formSectionProfileProviderService, [instructions: 'instruction'], sectionValues)
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
                throw new ApiInternalException('OC3I03', 'The CRF file must include a sheet "Groups"')
            }

            Sheet itemsSheet = workbook.getSheet('Items')
            if (!itemsSheet) {
                throw new ApiInternalException('OC3I03', 'The CRF file must include a sheet "Items"')
            }

            List<Map<String, String>> itemsSheetValues = workbookHandler.getSheetValues(ColumnHeaders.ITEMS_SHEET_COLUMN_PATTERNS, itemsSheet).findAll {it.itemName}
            List<Map<String, String>> groupsSheetValues = workbookHandler.getSheetValues(ColumnHeaders.GROUPS_SHEET_COLUMN_PATTERNS, groupsSheet).findAll {it.groupLabel}

            getEnumerationTypes(itemsSheetValues).each {
                dataModel.addToDataTypes(it)
            }

            println sectionsSheetValues
            println groupsSheetValues
            println itemsSheetValues

            List<DataClass> groups = []
            itemsSheetValues.eachWithIndex {Map<String, String> itemValues, Integer i ->
                DataElement item = new DataElement(label: itemValues.itemName, idx: i, minMultiplicity: itemValues.required == '1' ? 1 : 0)
                    .tap {DataElement dataElement ->
                        if (itemValues.descriptionLabel) description = itemValues.descriptionLabel
                        dataType = dataModel.enumerationTypes.find {it.label == itemValues.responseLabel} ?: getFormDataType(dataModel.primitiveTypes, itemValues.dataType)
                        JsonProfile questionProfile = formQuestionProfileProviderService.getNewProfile()
                        questionProfile.allFields.find {it.metadataPropertyName == 'question_instruction'}.currentValue = itemValues.leftItemText
                        questionProfile.allFields.find {it.metadataPropertyName == 'units'}.currentValue = itemValues.units
                        questionProfile.allFields.find {it.metadataPropertyName == 'answer_instruction'}.currentValue = itemValues.rightItemText
                        questionProfile.allFields.find {it.metadataPropertyName == 'label'}.currentValue = itemValues.questionNumber
                        questionProfile.allFields.find {it.metadataPropertyName == 'default'}.currentValue = itemValues.defaultValue
                        if (itemValues.itemDisplayStatus?.equalsIgnoreCase('HIDE')) questionProfile.allFields.find {it.metadataPropertyName == 'style'}.currentValue = 'Hidden'
                        addProfileMetadata(dataElement, formQuestionProfileProviderService, questionProfile)
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
                        parentClass = new DataClass(label: itemValues.groupLabel, idx: i)
                        Map<String, String> groupValues = groupsSheetValues.find {Map<String, String> groupValues -> groupValues.groupLabel == itemValues.groupLabel}
                        if (groupValues.groupHeader) parentClass.description = groupValues.groupHeader
                        if (groupValues.groupRepeatMax) parentClass.maxMultiplicity = groupValues.groupRepeatMax.toInteger()
                        JsonProfile groupProfile = formSectionProfileProviderService.getNewProfile()
                        if (groupValues.groupLayout?.equalsIgnoreCase('GRID')) groupProfile.allFields.find {it.metadataPropertyName == 'style'}.currentValue = 'Tabular'
                        else if (groupValues.groupLayout?.equalsIgnoreCase('NON-REPEATING'))
                            groupProfile.allFields.find {it.metadataPropertyName == 'style'}.currentValue = 'Inline'
                        if (groupValues.itemDisplayStatus?.equalsIgnoreCase('HIDE')) groupProfile.allFields.find {it.metadataPropertyName == 'style'}.currentValue = 'Hidden'
                        addProfileMetadata(parentClass, formSectionProfileProviderService, groupProfile)
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
                if (sheetValues[entry.key]) {
                    metadataAware.addToMetadata namespace: metadataNamespace, key: it.metadataPropertyName, value: sheetValues[entry.key]
                }
            }
        }
        metadataAware
    }

    MetadataAware addProfileMetadata(MetadataAware metadataAware, ProfileProviderService providerService, Profile profile) {
        profile.allFields.each {ProfileField profileField ->
            if (profileField.currentValue) {
                metadataAware.addToMetadata namespace: providerService.metadataNamespace, key: profileField.metadataPropertyName, value: profileField.currentValue
            }
        }
        metadataAware
    }

    DataType getFormDataType(Set<PrimitiveType> defaultTypes, String ocDataType) {
        [
            'ST'   : defaultTypes.find {it.label == 'Text'},
            'INT'  : defaultTypes.find {it.label == 'Number'},
            'REAL' : defaultTypes.find {it.label == 'Decimal'},
            'DATE' : defaultTypes.find {it.label == 'Date'},
            'PDATE': defaultTypes.find {it.label == 'Date'},
            'FILE' : defaultTypes.find {it.label == 'File'}
        ][ocDataType.toUpperCase()] ?: defaultTypes.find {it.label == 'Text'}
    }

    List<EnumerationType> getEnumerationTypes(List<Map<String, String>> itemsSheetValues) {
        List<Map<String, String>> enumerationTypeRows = []
        itemsSheetValues
            .findAll {it.responseType.toLowerCase() in ['single-select', 'radio', 'multi-select', 'checkbox'] && it.responseLabel && it.responseOptionsText && it.
                responseValuesOrCalculations}.findAll {it.responseOptionsText.split(',').length == it.responseValuesOrCalculations.split(',').length}.each {
            if (!enumerationTypeRows.find {etRow -> it.responseLabel == etRow.responseLabel}) {
                enumerationTypeRows.add(it)
            }
        }
        enumerationTypeRows.collect {
            new EnumerationType(label: it.responseLabel).tap {et ->
                List<String> keys = it.responseOptionsText.split(',')
                List<String> values = it.responseValuesOrCalculations.split(',')
                keys.eachWithIndex {key, i ->
                    et.addToEnumerationValues(key: key, value: values[i], idx: i)
                }
            }
        }
    }

    @Override
    List<DataModel> importModels(User currentUser, OpenClinicaV3CrfDataModelImporterParameters importerParameters) {
        throw new ApiNotYetImplementedException('OC3I04', 'importModels')
    }
}
