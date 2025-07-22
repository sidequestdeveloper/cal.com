import prismock from "../../../../../tests/libs/__mocks__/prisma";
import "../__mocks__/getOfficeAppKeys";
import {
  fetcherMock,
  mockCalendarList,
  mockEvents,
  mockSubscription,
  mockBatchResponse,
  mockUserEndpoint,
  setFetcherMockResponse,
  resetMocks,
} from "../__mocks__/microsoft-graph";

import { expect, test, beforeEach, vi, describe } from "vitest";

import { CalendarCache } from "@calcom/features/calendar-cache/calendar-cache";
import { getTimeMax, getTimeMin } from "@calcom/features/calendar-cache/lib/datesForCache";
import { SelectedCalendarRepository } from "@calcom/lib/server/repository/selectedCalendar";
import type { BufferedBusyTime } from "@calcom/types/BufferedBusyTime";
import type { CredentialForCalendarServiceWithTenantId } from "@calcom/types/Credential";

import CalendarService from "../CalendarService";

// const log = logger.getSubLogger({ prefix: ["Office365CalendarService.test"] });

// Mock OAuthManager to intercept requests
vi.mock("../../_utils/oauth/OAuthManager", () => ({
  OAuthManager: vi.fn().mockImplementation(() => ({
    requestRaw: fetcherMock,
  })),
}));

const createCredentialForCalendarService = (userId: number): CredentialForCalendarServiceWithTenantId =>
  ({
    id: 1,
    type: "office365_calendar",
    key: {
      access_token: "fake_access_token",
      refresh_token: "fake_refresh_token",
      expires_at: Date.now() + 3600000, // 1 hour from now
    },
    userId,
    user: {
      email: "test@example.com",
    },
    teamId: null,
    appId: "office365-calendar",
    invalid: false,
    delegatedTo: null,
    delegationCredentialId: null,
  } as unknown as CredentialForCalendarServiceWithTenantId);

async function expectCacheToBeSet({
  credentialId,
  itemsInKey,
}: {
  credentialId: number;
  itemsInKey: { id: string }[];
}) {
  const caches = await prismock.calendarCache.findMany({
    where: {
      credentialId,
    },
  });
  expect(caches).toHaveLength(1);
  const cache = caches[0];
  const parsedKey = JSON.parse(cache.key);
  expect(parsedKey.items).toEqual(itemsInKey);
  return cache;
}

describe("Office365CalendarService", () => {
  beforeEach(() => {
    resetMocks();
    vi.clearAllMocks();
    prismock.calendarCache.deleteMany();
    prismock.selectedCalendar.deleteMany();
  });

  describe("getAvailability", () => {
    test("should fetch availability without cache when shouldServeCache is false", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      const dateFrom = "2024-01-01T00:00:00Z";
      const dateTo = "2024-01-02T00:00:00Z";
      const selectedCalendars = [
        {
          credentialId: 1,
          userId: 1,
          externalId: "calendar-1",
          integration: "office365_calendar",
          eventTypeId: null,
        },
      ];

      setFetcherMockResponse({
        ok: true,
        json: async () => mockBatchResponse,
      });

      const availability = await calendarService.getAvailability(
        dateFrom,
        dateTo,
        selectedCalendars,
        false // shouldServeCache
      );

      expect(fetcherMock).toHaveBeenCalled();
      expect(availability).toHaveLength(2);
      expect(availability[0]).toEqual({
        start: mockEvents[0].start.dateTime,
        end: mockEvents[0].end.dateTime,
        source: "office365_calendar",
      });
    });

    test("should use cache when available", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      const dateFrom = "2024-01-01T00:00:00Z";
      const dateTo = "2024-01-02T00:00:00Z";
      const selectedCalendars = [
        {
          credentialId: 1,
          userId: 1,
          externalId: "calendar-1",
          integration: "office365_calendar",
          eventTypeId: null,
        },
      ];

      // Set up cache
      const calendarCache = await CalendarCache.init(calendarService);
      const cacheData: BufferedBusyTime[] = [
        {
          start: "2024-01-01T10:00:00Z",
          end: "2024-01-01T11:00:00Z",
        },
      ];
      await calendarCache.upsertCachedAvailability({
        credentialId: credential.id,
        userId: credential.userId,
        args: {
          timeMin: getTimeMin(dateFrom),
          timeMax: getTimeMax(dateTo),
          items: [{ id: "calendar-1" }],
        },
        value: JSON.parse(JSON.stringify(cacheData)),
      });

      const availability = await calendarService.getAvailability(dateFrom, dateTo, selectedCalendars);

      // Should not call fetcher when cache is available
      expect(fetcherMock).not.toHaveBeenCalled();
      expect(availability).toHaveLength(1);
      expect(availability[0]).toEqual({
        start: cacheData[0].start,
        end: cacheData[0].end,
        source: "office365_calendar",
      });
    });
  });

  describe("listCalendars", () => {
    test("should list calendars successfully", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      setFetcherMockResponse({
        ok: true,
        json: async () => ({ value: mockCalendarList }),
      });

      const calendars = await calendarService.listCalendars();

      expect(fetcherMock).toHaveBeenCalledWith("/calendars", expect.any(Object));
      expect(calendars).toHaveLength(2);
      expect(calendars[0]).toEqual({
        externalId: "calendar-1",
        integration: "office365_calendar",
        name: "Primary Calendar",
        primary: true,
        readOnly: false,
        email: "test@example.com",
      });
    });
  });

  describe("watchCalendar", () => {
    beforeEach(() => {
      process.env.OFFICE365_WEBHOOK_CLIENT_STATE = "test-client-state";
      vi.stubEnv("NEXT_PUBLIC_WEBAPP_URL", "https://example.com");
    });

    test("should create a new subscription", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      setFetcherMockResponse({
        ok: true,
        json: async () => mockUserEndpoint,
      });

      setFetcherMockResponse({
        ok: true,
        statusText: "Created",
        json: async () => mockSubscription,
      });

      // Mock fetchAvailability for initial cache population
      setFetcherMockResponse({
        ok: true,
        json: async () => mockBatchResponse,
      });

      const result = await calendarService.watchCalendar({
        calendarId: "calendar-1",
        eventTypeIds: [null],
      });

      expect(fetcherMock).toHaveBeenCalledWith(
        "/subscriptions",
        expect.objectContaining({
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: expect.stringContaining("calendar-1"),
        })
      );

      expect(result).toEqual({
        id: mockSubscription.id,
        expirationDateTime: mockSubscription.expirationDateTime,
        clientState: mockSubscription.clientState,
      });

      // Verify that selected calendar was updated
      const selectedCalendars = await SelectedCalendarRepository.findMany({
        where: {
          credentialId: credential.id,
          externalId: "calendar-1",
        },
      });

      expect(selectedCalendars).toHaveLength(1);
      expect(selectedCalendars[0].office365SubscriptionId).toBe(mockSubscription.id);
    });

    test("should reuse existing subscription", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      // Create existing subscription in database
      await SelectedCalendarRepository.upsert({
        credentialId: credential.id,
        userId: credential.userId!,
        integration: "office365_calendar",
        externalId: "calendar-1",
        eventTypeId: null,
        office365SubscriptionId: "existing-subscription-id",
        office365SubscriptionExpiration: new Date(Date.now() + 24 * 60 * 60 * 1000),
        office365SubscriptionClientState: "test-client-state",
      });

      const result = await calendarService.watchCalendar({
        calendarId: "calendar-1",
        eventTypeIds: [1], // Different event type
      });

      // Should not call create subscription
      expect(fetcherMock).not.toHaveBeenCalledWith("/subscriptions", expect.any(Object));

      expect(result).toEqual({
        id: "existing-subscription-id",
        expirationDateTime: expect.any(String),
        clientState: "test-client-state",
      });
    });
  });

  describe("unwatchCalendar", () => {
    test("should delete subscription when no other calendars use it", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      // Create existing subscription in database
      await SelectedCalendarRepository.upsert({
        credentialId: credential.id,
        userId: credential.userId!,
        integration: "office365_calendar",
        externalId: "calendar-1",
        eventTypeId: null,
        office365SubscriptionId: "subscription-to-delete",
        office365SubscriptionExpiration: new Date(Date.now() + 24 * 60 * 60 * 1000),
        office365SubscriptionClientState: "test-client-state",
      });

      setFetcherMockResponse({
        ok: true,
      });

      await calendarService.unwatchCalendar({
        calendarId: "calendar-1",
        eventTypeIds: [null],
      });

      expect(fetcherMock).toHaveBeenCalledWith(
        "/subscriptions/subscription-to-delete",
        expect.objectContaining({
          method: "DELETE",
        })
      );

      // Verify that selected calendar was updated
      const selectedCalendars = await SelectedCalendarRepository.findMany({
        where: {
          credentialId: credential.id,
          externalId: "calendar-1",
        },
      });

      expect(selectedCalendars).toHaveLength(1);
      expect(selectedCalendars[0].office365SubscriptionId).toBeNull();
    });

    test("should not delete subscription when other calendars use it", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      // Create existing subscriptions in database
      await SelectedCalendarRepository.upsert({
        credentialId: credential.id,
        userId: credential.userId!,
        integration: "office365_calendar",
        externalId: "calendar-1",
        eventTypeId: null,
        office365SubscriptionId: "shared-subscription",
        office365SubscriptionExpiration: new Date(Date.now() + 24 * 60 * 60 * 1000),
        office365SubscriptionClientState: "test-client-state",
      });

      await SelectedCalendarRepository.upsert({
        credentialId: credential.id,
        userId: credential.userId!,
        integration: "office365_calendar",
        externalId: "calendar-1",
        eventTypeId: 1,
        office365SubscriptionId: "shared-subscription",
        office365SubscriptionExpiration: new Date(Date.now() + 24 * 60 * 60 * 1000),
        office365SubscriptionClientState: "test-client-state",
      });

      await calendarService.unwatchCalendar({
        calendarId: "calendar-1",
        eventTypeIds: [null],
      });

      // Should not call delete subscription
      expect(fetcherMock).not.toHaveBeenCalledWith(
        expect.stringContaining("/subscriptions/"),
        expect.objectContaining({
          method: "DELETE",
        })
      );

      // Verify that only the specific calendar was updated
      const selectedCalendars = await SelectedCalendarRepository.findMany({
        where: {
          credentialId: credential.id,
          externalId: "calendar-1",
        },
      });

      expect(selectedCalendars).toHaveLength(2);
      const nullEventTypeCalendar = selectedCalendars.find((sc) => sc.eventTypeId === null);
      const eventTypeCalendar = selectedCalendars.find((sc) => sc.eventTypeId === 1);

      expect(nullEventTypeCalendar?.office365SubscriptionId).toBeNull();
      expect(eventTypeCalendar?.office365SubscriptionId).toBe("shared-subscription");
    });
  });

  describe("fetchAvailabilityAndSetCache", () => {
    test("should fetch availability and set cache", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      const selectedCalendars = [
        {
          credentialId: 1,
          userId: 1,
          externalId: "calendar-1",
          integration: "office365_calendar",
          eventTypeId: null,
        },
      ];

      setFetcherMockResponse({
        ok: true,
        json: async () => mockBatchResponse,
      });

      await calendarService.fetchAvailabilityAndSetCache(selectedCalendars);

      // Verify cache was set
      await expectCacheToBeSet({
        credentialId: credential.id,
        itemsInKey: [{ id: "calendar-1" }],
      });
    });

    test("should handle multiple event types separately", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      const selectedCalendars = [
        {
          credentialId: 1,
          userId: 1,
          externalId: "calendar-1",
          integration: "office365_calendar",
          eventTypeId: null,
        },
        {
          credentialId: 1,
          userId: 1,
          externalId: "calendar-1",
          integration: "office365_calendar",
          eventTypeId: 1,
        },
      ];

      setFetcherMockResponse({
        ok: true,
        json: async () => mockBatchResponse,
      });

      setFetcherMockResponse({
        ok: true,
        json: async () => mockBatchResponse,
      });

      await calendarService.fetchAvailabilityAndSetCache(selectedCalendars);

      // Should have created 2 cache entries
      const caches = await prismock.calendarCache.findMany({
        where: {
          credentialId: credential.id,
        },
      });
      expect(caches).toHaveLength(2);
    });
  });

  describe("fetchAvailability", () => {
    test("should fetch availability directly from API", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      setFetcherMockResponse({
        ok: true,
        json: async () => mockBatchResponse,
      });

      const result = await calendarService.fetchAvailability({
        timeMin: "2024-01-01T00:00:00Z",
        timeMax: "2024-01-02T00:00:00Z",
        items: [{ id: "calendar-1" }],
      });

      expect(fetcherMock).toHaveBeenCalledWith(
        expect.stringContaining("/me/calendars/getSchedule"),
        expect.objectContaining({
          method: "POST",
        })
      );
      expect(result).toHaveLength(2);
    });

    test("should return empty array when no calendars", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      const result = await calendarService.fetchAvailability({
        timeMin: "2024-01-01T00:00:00Z",
        timeMax: "2024-01-02T00:00:00Z",
        items: [],
      });

      expect(fetcherMock).not.toHaveBeenCalled();
      expect(result).toEqual([]);
    });
  });

  describe("getFreeBusyResult", () => {
    test("should use cache when shouldServeCache is true", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      const calendarCache = await CalendarCache.init(calendarService);
      const cacheData: BufferedBusyTime[] = [
        {
          start: "2024-01-01T10:00:00Z",
          end: "2024-01-01T11:00:00Z",
        },
      ];

      await calendarCache.upsertCachedAvailability({
        credentialId: credential.id,
        userId: credential.userId,
        args: {
          timeMin: getTimeMin("2024-01-01T00:00:00Z"),
          timeMax: getTimeMax("2024-01-02T00:00:00Z"),
          items: [{ id: "calendar-1" }],
        },
        value: JSON.parse(JSON.stringify(cacheData)),
      });

      const result = await calendarService.getFreeBusyResult({
        timeMin: "2024-01-01T00:00:00Z",
        timeMax: "2024-01-02T00:00:00Z",
        items: [{ id: "calendar-1" }],
      });

      expect(fetcherMock).not.toHaveBeenCalled();
      expect(result).toEqual(cacheData);
    });

    test("should fetch from API when cache disabled", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      setFetcherMockResponse({
        ok: true,
        json: async () => mockBatchResponse,
      });

      const result = await calendarService.getFreeBusyResult(
        {
          timeMin: "2024-01-01T00:00:00Z",
          timeMax: "2024-01-02T00:00:00Z",
          items: [{ id: "calendar-1" }],
        },
        false // shouldServeCache
      );

      expect(fetcherMock).toHaveBeenCalled();
      expect(result).toHaveLength(2);
    });
  });

  describe("getCacheOrFetchAvailability", () => {
    test("should return EventBusyDate format", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      setFetcherMockResponse({
        ok: true,
        json: async () => mockBatchResponse,
      });

      const result = await calendarService.getCacheOrFetchAvailability({
        timeMin: "2024-01-01T00:00:00Z",
        timeMax: "2024-01-02T00:00:00Z",
        items: [{ id: "calendar-1" }],
      });

      expect(result[0]).toEqual({
        start: mockEvents[0].start.dateTime,
        end: mockEvents[0].end.dateTime,
        source: "office365_calendar",
      });
    });
  });

  describe("upsertSelectedCalendar", () => {
    test("should upsert selected calendar", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      const result = await calendarService.upsertSelectedCalendar({
        externalId: "calendar-1",
        eventTypeId: null,
        office365SubscriptionId: "sub-123",
        office365SubscriptionExpiration: new Date(),
        office365SubscriptionClientState: "client-state",
      });

      expect(result).toBeTruthy();
      expect(result?.externalId).toBe("calendar-1");
      expect(result?.office365SubscriptionId).toBe("sub-123");
    });

    test("should handle missing userId gracefully", async () => {
      const credential = { ...createCredentialForCalendarService(1), userId: null };
      const calendarService = new CalendarService(credential as any);

      const result = await calendarService.upsertSelectedCalendar({
        externalId: "calendar-1",
        eventTypeId: null,
      });

      expect(result).toBeUndefined();
    });
  });

  describe("upsertSelectedCalendarsForEventTypeIds", () => {
    test("should upsert for multiple event type IDs", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      await calendarService.upsertSelectedCalendarsForEventTypeIds(
        {
          externalId: "calendar-1",
          office365SubscriptionId: "sub-123",
          office365SubscriptionExpiration: new Date(),
          office365SubscriptionClientState: "client-state",
        },
        [null, 1, 2]
      );

      const calendars = await prismock.selectedCalendar.findMany({
        where: {
          credentialId: credential.id,
          externalId: "calendar-1",
        },
      });

      expect(calendars).toHaveLength(3);
      expect(calendars.map((c) => c.eventTypeId).sort()).toEqual([null, 1, 2]);
    });
  });

  describe("error handling", () => {
    test("should handle webhook creation failure", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      setFetcherMockResponse({
        ok: true,
        json: async () => mockUserEndpoint,
      });

      setFetcherMockResponse({
        ok: false,
        status: 400,
        statusText: "Bad Request",
        text: async () => "Invalid subscription",
      });

      await expect(
        calendarService.watchCalendar({
          calendarId: "calendar-1",
          eventTypeIds: [null],
        })
      ).rejects.toThrow("Failed to create subscription");
    });

    test("should handle unwatch errors gracefully", async () => {
      const credential = createCredentialForCalendarService(1);
      const calendarService = new CalendarService(credential);

      await SelectedCalendarRepository.upsert({
        credentialId: credential.id,
        userId: credential.userId!,
        integration: "office365_calendar",
        externalId: "calendar-1",
        eventTypeId: null,
        office365SubscriptionId: "sub-123",
      });

      setFetcherMockResponse({
        ok: false,
        status: 500,
        statusText: "Internal Server Error",
      });

      await expect(
        calendarService.unwatchCalendar({
          calendarId: "calendar-1",
          eventTypeIds: [null],
        })
      ).rejects.toThrow("Failed to delete subscription");
    });
  });
});
