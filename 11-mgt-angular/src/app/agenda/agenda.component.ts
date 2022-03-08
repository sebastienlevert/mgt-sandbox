import { Component, OnInit, ViewChild } from '@angular/core';

@Component({
  selector: 'app-agenda',
  templateUrl: './agenda.component.html',
  styleUrls: ['./agenda.component.scss']
})
export class AgendaComponent implements OnInit {
  @ViewChild('agenda', {static: true}) agendaElement: any;
  
  constructor() { }

  ngOnInit() {
    this.agendaElement.nativeElement.templateContext = {
      openWebLink: (e: any, context: { event: { webLink: string | undefined; }; }, root: any) => {
          window.open(context.event.webLink, '_blank');
      }
    };
  }

}
