import { provideHttpClient } from '@angular/common/http';
import { provideHttpClientTesting } from '@angular/common/http/testing';
import { TestBed } from '@angular/core/testing';

import { UserDictionaryService } from './user-dictionary.service';

describe('UserDictionaryService', () => {
  let service: UserDictionaryService;

  beforeEach(() => {
    TestBed.configureTestingModule({
      providers: [provideHttpClient(), provideHttpClientTesting()],
    });
    service = TestBed.inject(UserDictionaryService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
