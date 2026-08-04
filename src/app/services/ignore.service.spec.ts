import { provideHttpClient } from '@angular/common/http';
import { provideHttpClientTesting } from '@angular/common/http/testing';
import { TestBed } from "@angular/core/testing";

import { IgnoreService } from "./ignore.service";

describe("IgnoreService", () => {
  let service: IgnoreService;

  beforeEach(() => {
    TestBed.configureTestingModule({
      providers: [provideHttpClient(), provideHttpClientTesting()],
    });
    service = TestBed.inject(IgnoreService);
  });

  it("should be created", () => {
    expect(service).toBeTruthy();
  });
});
