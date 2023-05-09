package uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.exporter

import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiBadRequestException
import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiException
import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiNotYetImplementedException
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClass
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataElement
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.exporter.DataModelExporterProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.ColumnHeaders
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.OpenClinicaV3CrfProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.group.OpenClinicaV3CrfGroupProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.item.OpenClinicaV3CrfItemProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.section.OpenClinicaV3CrfSectionProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.profile.ProfileService
import uk.ac.ox.softeng.maurodatamapper.profile.object.Profile
import uk.ac.ox.softeng.maurodatamapper.security.User

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

class OpenClinicaV3CrfDataModelExporterProviderService extends DataModelExporterProviderService {

    static final String CONTENT_TYPE = 'application/vnd.ms-excel'

    OpenClinicaV3CrfProfileProviderService openClinicaV3CrfProfileProviderService
    OpenClinicaV3CrfSectionProfileProviderService openClinicaV3CrfSectionProfileProviderService
    OpenClinicaV3CrfGroupProfileProviderService openClinicaV3CrfGroupProfileProviderService
    OpenClinicaV3CrfItemProfileProviderService openClinicaV3CrfItemProfileProviderService
    ProfileService profileService

    @Override
    String getDisplayName() {
        'OpenClinica 3.x CRF (XLS) Exporter'
    }

    @Override
    String getVersion() {
        getClass().getPackage().getSpecificationVersion() ?: 'SNAPSHOT'
    }

    @Override
    String getNamespace() {
        'uk.ac.ox.softeng.maurodatamapper.plugins.excel.openclinica'
    }

    @Override
    String getFileExtension() {
        'xls'
    }

    @Override
    String getContentType() {
        CONTENT_TYPE
    }

    @Override
    ByteArrayOutputStream exportDataModel(User currentUser, DataModel crfModel, Map<String, Object> parameters) throws ApiException {
        new HSSFWorkbook().withCloseable {Workbook workbook ->
            if (!(profileService.getUsedProfileServices(crfModel, false).find {it instanceof OpenClinicaV3CrfProfileProviderService})) {
                throw new ApiBadRequestException('OC3E01', 'The DataModel to export must have an OpenClinica 3.x CRF profile applied')
            }

            // Create CRF front sheet
            Sheet crfSheet = workbook.createSheet('CRF')
            crfSheet.createRow(0).with {
                ColumnHeaders.CRF_SHEET_COLUMNS.values().eachWithIndex {String header, Integer i ->
                    Cell cell = createCell(i)
                    cell.setCellValue(header)
                }
            }
            crfSheet.createRow(1).with {
                Profile crfProfile = openClinicaV3CrfProfileProviderService.createProfileFromEntity(crfModel)
                ColumnHeaders.CRF_SHEET_COLUMNS.values().eachWithIndex {String header, Integer i ->
                    Cell cell = createCell(i)
                    crfProfile.allFields.find {it.metadataPropertyName.equalsIgnoreCase(header)}?.with {profileField ->
                        cell.setCellValue(profileField.currentValue)
                    }
                }
            }

            // Create Sections sheet
            Sheet sectionsSheet = workbook.createSheet('Sections')
            sectionsSheet.createRow(0).with {
                ColumnHeaders.SECTIONS_SHEET_COLUMNS.values().eachWithIndex {String header, Integer i ->
                    Cell cell = createCell(i)
                    cell.setCellValue(header)
                }
            }
            crfModel.childDataClasses.eachWithIndex {DataClass dataClass, Integer i ->
                if (profileService.getUsedProfileServices(dataClass, false).find {it instanceof OpenClinicaV3CrfSectionProfileProviderService}) {
                    sectionsSheet.createRow(i + 1).with {
                        Profile sectionProfile = openClinicaV3CrfSectionProfileProviderService.createProfileFromEntity(dataClass)
                        ColumnHeaders.SECTIONS_SHEET_COLUMNS.values().eachWithIndex {String header, Integer j ->
                            Cell cell = createCell(j)
                            sectionProfile.allFields.find {it.metadataPropertyName.equalsIgnoreCase(header)}?.with {profileField ->
                                cell.setCellValue(profileField.currentValue)
                            }
                        }
                    }
                }
            }

            // Create Groups sheet
            Sheet groupsSheet = workbook.createSheet('Groups')
            groupsSheet.createRow(0).with {
                ColumnHeaders.GROUPS_SHEET_COLUMNS.values().eachWithIndex {String header, Integer i ->
                    Cell cell = createCell(i)
                    cell.setCellValue(header)
                }
            }
            List<DataClass> groupClasses = []
            crfModel.childDataClasses.each {
                if (profileService.getUsedProfileServices(it, false).find {it instanceof OpenClinicaV3CrfGroupProfileProviderService}) {
                    groupClasses.add(it)
                } else {
                    it.dataClasses.sort().each {
                        if (profileService.getUsedProfileServices(it, false).find {it instanceof OpenClinicaV3CrfGroupProfileProviderService}) {
                            groupClasses.add(it)
                        }
                    }
                }
            }
            groupClasses.eachWithIndex {DataClass dataClass, Integer i ->
                if (profileService.getUsedProfileServices(dataClass, false).find {it instanceof OpenClinicaV3CrfGroupProfileProviderService}) {
                    groupsSheet.createRow(i + 1).with {
                        Profile groupProfile = openClinicaV3CrfGroupProfileProviderService.createProfileFromEntity(dataClass)
                        ColumnHeaders.GROUPS_SHEET_COLUMNS.values().eachWithIndex {String header, Integer j ->
                            Cell cell = createCell(j)
                            groupProfile.allFields.find {it.metadataPropertyName.equalsIgnoreCase(header)}?.with {profileField ->
                                cell.setCellValue(profileField.currentValue)
                            }
                        }
                    }
                }
            }

            // Create Items sheet
            Sheet itemsSheet = workbook.createSheet('Items')
            itemsSheet.createRow(0).with {
                ColumnHeaders.ITEMS_SHEET_COLUMNS.values().eachWithIndex {String header, Integer i ->
                    Cell cell = createCell(i)
                    cell.setCellValue(header)
                }
            }
            List<DataElement> itemElements = []
            crfModel.childDataClasses.each {
                // Ungrouped items
                it.dataElements.sort().each {
                    if (profileService.getUsedProfileServices(it, false).find {it instanceof OpenClinicaV3CrfItemProfileProviderService}) {
                        itemElements.add(it)
                    }
                }
                // Grouped items
                it.dataClasses.sort().each {
                    it.dataElements.sort().each {
                        if (profileService.getUsedProfileServices(it, false).find {it instanceof OpenClinicaV3CrfItemProfileProviderService}) {
                            itemElements.add(it)
                        }
                    }
                }
            }
            itemElements.eachWithIndex {DataElement dataElement, Integer i ->
                if (profileService.getUsedProfileServices(dataElement, false).find {it instanceof OpenClinicaV3CrfItemProfileProviderService}) {
                    itemsSheet.createRow(i + 1).with {
                        Profile itemProfile = openClinicaV3CrfItemProfileProviderService.createProfileFromEntity(dataElement)
                        ColumnHeaders.ITEMS_SHEET_COLUMNS.values().eachWithIndex {String header, Integer j ->
                            Cell cell = createCell(j)
                            itemProfile.allFields.find {it.metadataPropertyName.equalsIgnoreCase(header)}?.with {profileField ->
                                cell.setCellValue(profileField.currentValue)
                            }
                        }
                    }
                }
            }

            // Create Instructions sheet
            Sheet instructionsSheet = workbook.createSheet('Instructions')
            ['OpenClinica CRF Design Template v3.9',
             'This worksheet contains important additional information, tips, and best practices to help you better design your OpenClinica CRFs.',
             '',
             'Note: Each CRF version should be defined in its own Excel file.'].eachWithIndex {String value, Integer i ->
                instructionsSheet.createRow(i).with {
                    Cell cell = createCell(0)
                    cell.setCellValue(value)
                }
            }

            new ByteArrayOutputStream().tap {ByteArrayOutputStream exportStream ->
                workbook.write(exportStream)
            }
        }
    }

    @Override
    ByteArrayOutputStream exportDataModels(User currentUser, List<DataModel> dataModels, Map<String, Object> parameters) throws ApiException {
        throw new ApiNotYetImplementedException('OC3EXX', 'Importing multiple CRFs not supported')
    }

}
