import { provideHttpClient, withXhr } from '@angular/common/http';
import { provideHttpClientTesting } from '@angular/common/http/testing';
import { CUSTOM_ELEMENTS_SCHEMA } from '@angular/core';
import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SpellcheckerComponent } from './spellchecker.component';

describe('SpellcheckerComponent', () => {
  let component: SpellcheckerComponent;
  let fixture: ComponentFixture<SpellcheckerComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ SpellcheckerComponent ],
      providers: [provideHttpClient(withXhr()), provideHttpClientTesting()],
      schemas: [CUSTOM_ELEMENTS_SCHEMA],

    })
    .compileComponents();

    fixture = TestBed.createComponent(SpellcheckerComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
