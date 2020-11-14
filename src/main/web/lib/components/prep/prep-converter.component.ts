import {Component, OnDestroy, OnInit} from '@angular/core';
import {RadetConverterService} from "../../services/radet-converter.service";
import {RxStompService} from "@stomp/ng2-stompjs";
import {Message} from '@stomp/stompjs';
import {Subscription} from "rxjs";
import {DomSanitizer} from "@angular/platform-browser";
import {saveAs} from 'file-saver';
import {DateRange} from '@syncfusion/ej2-calendars';

export interface Facility {
    id: number;
    name: string;
    selected: boolean;
}

@Component({
    selector: 'prep-converter',
    templateUrl: './prep-convert.component.html'
})
export class PrepConverterComponent implements OnInit, OnDestroy {
    private topicSubscription: Subscription;
    facilities: Facility[] = [];
    files: string[];
    running = false;
    message: any;
    finished = false;
    dateRange: DateRange = {
        start: new Date(1900, 0, 1),
        end: new Date()
    };
    reportingPeriod: Date = new Date();
    todaySelectable = true;
    today = new Date();
    current = false;

    constructor(private service: RadetConverterService, private stompService: RxStompService, private domSanitizer: DomSanitizer) {
    }

    ngOnInit() {
        this.service.listFacilities().subscribe(res => this.facilities = res);
        this.topicSubscription = this.stompService.watch("/topic/prep/status").subscribe((msg: Message) => {
            if (msg.body === 'start') {
                this.running = true
            } else if (msg.body === 'end') {
                this.running = false;
                this.message = "Conversion finished; download files from Download tab";
                this.finished = true;
                this.service.listFiles().subscribe(res => {
                    this.files = res;
                })
            } else {
                this.message = msg.body;
                this.running = true;
            }
        })
    }

    selected(): boolean {
        return this.facilities.filter(f => f.selected).length > 0
    }

    download(name: string) {
        this.service.downloadPrepFile(name).subscribe(res => {
            const file = new File([res], name + '_PrEP.xlsx', {type: 'application/octet-stream'});
            saveAs(file);
        });
    }

    tabChanged(event) {
        if (event.index === 1) {
            this.service.listPrepFiles().subscribe(res => {
                this.files = res;
            })
        }
    }

    monthChanged(month: Date) {
        this.todaySelectable = new Date().getMonth() === month.getMonth()
    }

    convert() {
        this.running = true;
        this.finished = false;
        let ids = this.facilities.filter(f => f.selected)
            .map(f => f.id);
        this.service.convertPrep(this.dateRange.start, this.dateRange.end, this.reportingPeriod, ids, this.current).subscribe()
    }

    ngOnDestroy(): void {
        this.topicSubscription.unsubscribe()
    }
}
