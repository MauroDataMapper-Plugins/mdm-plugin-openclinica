package uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.exporter

import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiBadRequestException
import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiException
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.exporter.DataModelExporterProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.ColumnHeaders
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf.OpenClinicaV3CrfProfileProviderService
import uk.ac.ox.softeng.maurodatamapper.profile.ProfileService
import uk.ac.ox.softeng.maurodatamapper.security.User

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

class OpenClinicaV3CrfDataModelExporterProviderService extends DataModelExporterProviderService {

    static final String CONTENT_TYPE = 'application/vnd.ms-excel'

    OpenClinicaV3CrfProfileProviderService openClinicaV3CrfProfileProviderService
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

            new ByteArrayOutputStream().tap {ByteArrayOutputStream exportStream ->
                workbook.write(exportStream)
            }
        }
    }
}
