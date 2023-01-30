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

package uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.importer

import uk.ac.ox.softeng.maurodatamapper.core.container.Folder
import uk.ac.ox.softeng.maurodatamapper.core.provider.importer.parameter.FileParameter
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModelService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClass
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataElement
import uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.importer.parameters.OpenClinicaV3CrfDataModelImporterParameters
import uk.ac.ox.softeng.maurodatamapper.test.functional.BaseFunctionalSpec
import uk.ac.ox.softeng.maurodatamapper.test.integration.BaseIntegrationSpec

import grails.gorm.transactions.Rollback
import grails.testing.mixin.integration.Integration
import grails.util.BuildSettings
import grails.validation.ValidationException
import groovy.util.logging.Slf4j

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

@Slf4j
@Integration
@Rollback
class OpenClinicaV3CrfDataModelImporterProviderServiceSpec extends BaseIntegrationSpec {

    static final resourcePath = Paths.get(BuildSettings.BASE_DIR.absolutePath, 'src', 'integration-test', 'resources')

    OpenClinicaV3CrfDataModelImporterProviderService openClinicaV3CrfDataModelImporterProviderService
    DataModelService dataModelService

    @Override
    void setupDomainData() {
        folder = new Folder(label: 'OpenClinica', createdBy: admin.emailAddress)
        checkAndSave(folder)
    }

    void 'OC01 : Import OpenClinica V3 CRF without Groups'() {
        given:
        setupDomainData()

        when:
        DataModel crf = importAndValidateModel('Test_CRF_No_Groups.xls')

        then:
        crf
        crf.label == 'Test CRF'
        crf.dataClasses.size() == 1
        crf.allDataElements.size() == 2

        when:
        DataClass mainSection = crf.dataClasses.first()

        then:
        mainSection
        mainSection.label == 'Main'
        mainSection.dataElements.size() == 2

        when:
        DataElement item1 = mainSection.dataElements.find {it.label == 'Item 1'}

        then:
        item1
        item1.dataType.label == 'ST'
        item1.metadata.find {it.key == 'response_type'}.value == 'text'
    }

    OpenClinicaV3CrfDataModelImporterParameters createImportParameters(Path filePath) {
        new OpenClinicaV3CrfDataModelImporterParameters(importFile: new FileParameter(filePath.toString(), 'xls', Files.readAllBytes(resourcePath.resolve(filePath))),
                                                        folderId: folder.id)
    }

    DataModel importAndValidateModel(String filename) {
        Path filePath = Path.of(filename)
        OpenClinicaV3CrfDataModelImporterParameters parameters = createImportParameters(filePath)
        DataModel dataModel = openClinicaV3CrfDataModelImporterProviderService.importDomain(admin, parameters)
        validateAndSave(dataModel)
    }

    DataModel validateAndSave(DataModel dataModel) {
        assert dataModel
        dataModel.folder = folder
        dataModelService.validate(dataModel)
        if (dataModel.errors.hasErrors()) {
            GormUtils.outputDomainErrors(messageSource, dataModel)
            throw new ValidationException("Domain object is not valid. Has ${dataModel.errors.errorCount} errors", dataModel.errors)
        }
        dataModelService.saveModelWithContent(dataModel)
        dataModelService.get(dataModel.id)
    }

}
