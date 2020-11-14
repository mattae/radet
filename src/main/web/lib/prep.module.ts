import {RouterModule, Routes} from '@angular/router';
import {PrepConverterComponent} from './components/prep/prep-converter.component';
import {NgModule} from '@angular/core';
import {CommonModule} from '@angular/common';
import {FormsModule} from '@angular/forms';
import {DatePickerModule, DateRangePickerModule} from '@syncfusion/ej2-angular-calendars';
import {DropDownListModule} from '@syncfusion/ej2-angular-dropdowns';
import {
    MatButtonModule,
    MatCardModule,
    MatCheckboxModule,
    MatDividerModule,
    MatIconModule,
    MatInputModule,
    MatListModule,
    MatProgressBarModule,
    MatSelectModule,
    MatTabsModule
} from '@angular/material';

export const PREP_ROUTES: Routes = [
    {
        path: '',
        data: {
            breadcrumb: 'PREP CONVERTER',
            title: 'PrEP Converter'
        },
        children: [
            {
                path: '',
                component: PrepConverterComponent,
                data: {
                    breadcrumb: 'PREP CONVERTER',
                    title: 'PrEP Converter'
                },
            }
        ]
    }
];

@NgModule({
    declarations: [
        PrepConverterComponent
    ],
    imports: [
        CommonModule,
        FormsModule,
        MatInputModule,
        MatIconModule,
        MatDividerModule,
        MatCardModule,
        MatSelectModule,
        MatButtonModule,
        MatTabsModule,
        RouterModule.forChild(PREP_ROUTES),
        MatProgressBarModule,
        MatListModule,
        MatCheckboxModule,
        DateRangePickerModule,
        DropDownListModule,
        DatePickerModule,
    ]
})
export class PrepModule {

}
