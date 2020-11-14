import {Routes} from '@angular/router';
import {RadetConverterComponent} from '../components/radet/radet-converter.component';


export const ROUTES: Routes = [
    {
        path: '',
        data: {
            title: 'Radet Converter',
            breadcrumb: 'RADET CONVERTER'
        },
        children: [
            {
                path: '',
                component: RadetConverterComponent,
                data: {
                    breadcrumb: 'RADET CONVERTER',
                    title: 'Radet Converter'
                },
            }
        ]
    }
];

