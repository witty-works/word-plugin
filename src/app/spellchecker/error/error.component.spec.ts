import { provideHttpClient, withXhr } from '@angular/common/http';
import { provideHttpClientTesting } from '@angular/common/http/testing';
import { CUSTOM_ELEMENTS_SCHEMA } from '@angular/core';
import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ErrorComponent } from './error.component';
import { SpellcheckerComponent } from '../spellchecker.component';

describe('ErrorComponent', () => {
  let component: ErrorComponent;
  let fixture: ComponentFixture<ErrorComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ ErrorComponent ],
      providers: [
        provideHttpClient(withXhr()),
        provideHttpClientTesting(),
        // ErrorComponent injects its parent component directly.
        { provide: SpellcheckerComponent, useValue: {} },
      ],
      schemas: [CUSTOM_ELEMENTS_SCHEMA],

    })
    .compileComponents();

    fixture = TestBed.createComponent(ErrorComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
