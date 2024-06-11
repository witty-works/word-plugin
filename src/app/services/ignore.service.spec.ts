import { TestBed } from "@angular/core/testing";

import { IgnoreService } from "./ignore.service";

describe("IgnoreService", () => {
  let service: IgnoreService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(IgnoreService);
  });

  it("should be created", () => {
    expect(service).toBeTruthy();
  });
});
