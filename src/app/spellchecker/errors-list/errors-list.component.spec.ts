import { provideHttpClient, withXhr } from '@angular/common/http';
import { provideHttpClientTesting } from '@angular/common/http/testing';
import { CUSTOM_ELEMENTS_SCHEMA } from '@angular/core';
import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ErrorsListComponent } from './errors-list.component';

describe('ErrorsListComponent', () => {
  let component: ErrorsListComponent;
  let fixture: ComponentFixture<ErrorsListComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ ErrorsListComponent ],
      providers: [provideHttpClient(withXhr()), provideHttpClientTesting()],
      schemas: [CUSTOM_ELEMENTS_SCHEMA],

    })
    .compileComponents();

    fixture = TestBed.createComponent(ErrorsListComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
