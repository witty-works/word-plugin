import { provideHttpClient, withXhr } from '@angular/common/http';
import { provideHttpClientTesting } from '@angular/common/http/testing';
import { TestBed } from '@angular/core/testing';

import { CheckingService } from './checking.service';

describe('SpellcheckerService', () => {
  let service: CheckingService;

  beforeEach(() => {
    TestBed.configureTestingModule({
      providers: [provideHttpClient(withXhr()), provideHttpClientTesting()],
    });
    service = TestBed.inject(CheckingService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
