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

package uk.ac.ox.softeng.maurodatamapper.plugins.openclinica.v3.crf

import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.PrimitiveType
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.DefaultDataTypeProvider
import uk.ac.ox.softeng.maurodatamapper.datamodel.rest.transport.DefaultDataType

class OpenClinicaV3CrfDefaultDataTypeProviderService implements DefaultDataTypeProvider {

    @Override
    String getDisplayName() {
        'OpenClinica 3.x CRF DataTypes'
    }

    @Override
    String getVersion() {
        getClass().getPackage().getSpecificationVersion() ?: 'SNAPSHOT'
    }

    List<DefaultDataType> getDefaultListOfDataTypes() {
        [
            [
                label      : 'ST',
                description: 'String.  Any characters can be provided for this data type.'
            ],
            [
                label      : 'INT',
                description: 'Integer.  Only numbers with no decimal places are allowed for this data type.'
            ],
            [
                label      : 'REAL',
                description: 'Numbers with decimal places are allowed for this data type.'
            ],
            [
                label      : 'DATE',
                description: 'Only full dates are allowed for this data type.  The default date format the user must provide the value in is DD-MMM-YYYY.'
            ],
            [
                label      : 'PDATE',
                description: 'Partial dates are allowed for this data type.  The default date format is DD-MMM-YYYY so users can provide either MMM-YYYY or YYYY values.'
            ],
            [
                label      : 'FILE',
                description: 'This data type allows files to be attached to the item.  It must be used in conjunction with a RESPONSE_TYPE of file.  The attached file is ' +
                             'saved to the server and a URL is displayed to the user viewing the form.'
            ]
        ].collect {Map<String, String> properties -> new DefaultDataType(new PrimitiveType(properties))}
    }
}