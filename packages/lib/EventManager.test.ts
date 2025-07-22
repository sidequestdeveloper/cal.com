import type { DestinationCalendar } from "@prisma/client";
import { describe, expect, it, vi, beforeEach } from "vitest";

import { symmetricDecrypt } from "@calcom/lib/crypto";
import type { CredentialForCalendarService } from "@calcom/types/Credential";

import EventManager from "./EventManager";

vi.mock("@calcom/lib/crypto", () => ({
  symmetricDecrypt: vi.fn(),
}));

const mockedSymmetricDecrypt = vi.mocked(symmetricDecrypt);

function buildCalDAVCredential(data: {
  id: number;
  key: string;
  userId?: number;
}): CredentialForCalendarService {
  return {
    id: data.id,
    type: "caldav_calendar",
    key: data.key,
    userId: data.userId || 1,
    user: { email: "test@example.com" },
    teamId: null,
    appId: "caldav-calendar",
    invalid: false,
    delegatedTo: null,
    delegationCredentialId: null,
  };
}

function buildDestinationCalendar(data: {
  id: number;
  integration: string;
  externalId: string | null;
}): DestinationCalendar {
  return {
    id: data.id,
    integration: data.integration,
    externalId: data.externalId || "",
    userId: 1,
    eventTypeId: null,
    credentialId: null,
    primaryEmail: null,
    createdAt: new Date(),
    updatedAt: new Date(),
    delegationCredentialId: null,
    domainWideDelegationCredentialId: null,
  };
}

describe("EventManager CalDAV credential validation", () => {
  let eventManager: EventManager;

  beforeEach(() => {
    vi.clearAllMocks();
    eventManager = new EventManager({
      credentials: [],
      destinationCalendar: null,
    });
  });

  describe("extractServerUrlFromCredential", () => {
    it("should extract server URL from valid CalDAV credential", () => {
      const credential = buildCalDAVCredential({
        id: 1,
        key: "encrypted_key",
      });

      mockedSymmetricDecrypt.mockReturnValue(
        JSON.stringify({
          username: "user",
          password: "pass",
          url: "https://caldav.example.com/dav/calendars/user/",
        })
      );

      const result = (eventManager as any).extractServerUrlFromCredential(credential);
      expect(result).toBe("https://caldav.example.com");
      expect(mockedSymmetricDecrypt).toHaveBeenCalledWith(
        "encrypted_key",
        process.env.CALENDSO_ENCRYPTION_KEY || ""
      );
    });

    it("should return null for non-CalDAV credential", () => {
      const credential = {
        ...buildCalDAVCredential({ id: 1, key: "encrypted_key" }),
        type: "google_calendar",
      } as CredentialForCalendarService;

      const result = (eventManager as any).extractServerUrlFromCredential(credential);
      expect(result).toBeNull();
      expect(mockedSymmetricDecrypt).not.toHaveBeenCalled();
    });

    it("should return null when decryption fails", () => {
      const credential = buildCalDAVCredential({
        id: 1,
        key: "invalid_key",
      });

      mockedSymmetricDecrypt.mockImplementation(() => {
        throw new Error("Decryption failed");
      });

      const result = (eventManager as any).extractServerUrlFromCredential(credential);
      expect(result).toBeNull();
    });

    it("should return null when decrypted data has no URL", () => {
      const credential = buildCalDAVCredential({
        id: 1,
        key: "encrypted_key",
      });

      mockedSymmetricDecrypt.mockReturnValue(
        JSON.stringify({
          username: "user",
          password: "pass",
        })
      );

      const result = (eventManager as any).extractServerUrlFromCredential(credential);
      expect(result).toBeNull();
    });

    it("should return null when URL is invalid", () => {
      const credential = buildCalDAVCredential({
        id: 1,
        key: "encrypted_key",
      });

      mockedSymmetricDecrypt.mockReturnValue(
        JSON.stringify({
          username: "user",
          password: "pass",
          url: "invalid-url",
        })
      );

      const result = (eventManager as any).extractServerUrlFromCredential(credential);
      expect(result).toBeNull();
    });

    it("should handle URLs with different ports", () => {
      const credential = buildCalDAVCredential({
        id: 1,
        key: "encrypted_key",
      });

      mockedSymmetricDecrypt.mockReturnValue(
        JSON.stringify({
          username: "user",
          password: "pass",
          url: "https://caldav.example.com:8443/dav/calendars/user/",
        })
      );

      const result = (eventManager as any).extractServerUrlFromCredential(credential);
      expect(result).toBe("https://caldav.example.com:8443");
    });

    it("should handle HTTP URLs", () => {
      const credential = buildCalDAVCredential({
        id: 1,
        key: "encrypted_key",
      });

      mockedSymmetricDecrypt.mockReturnValue(
        JSON.stringify({
          username: "user",
          password: "pass",
          url: "http://caldav.internal.com/dav/calendars/user/",
        })
      );

      const result = (eventManager as any).extractServerUrlFromCredential(credential);
      expect(result).toBe("http://caldav.internal.com");
    });
  });

  describe("extractServerUrlFromDestination", () => {
    it("should extract server URL from CalDAV destination calendar", () => {
      const destination = buildDestinationCalendar({
        id: 1,
        integration: "caldav_calendar",
        externalId: "https://caldav.example.com/dav/calendars/user/calendar1/",
      });

      const result = (eventManager as any).extractServerUrlFromDestination(destination);
      expect(result).toBe("https://caldav.example.com");
    });

    it("should return null for non-CalDAV destination", () => {
      const destination = buildDestinationCalendar({
        id: 1,
        integration: "google_calendar",
        externalId: "calendar@gmail.com",
      });

      const result = (eventManager as any).extractServerUrlFromDestination(destination);
      expect(result).toBeNull();
    });

    it("should return null when externalId is null", () => {
      const destination = buildDestinationCalendar({
        id: 1,
        integration: "caldav_calendar",
        externalId: null,
      });

      const result = (eventManager as any).extractServerUrlFromDestination(destination);
      expect(result).toBeNull();
    });

    it("should return null when externalId is invalid URL", () => {
      const destination = buildDestinationCalendar({
        id: 1,
        integration: "caldav_calendar",
        externalId: "invalid-url",
      });

      const result = (eventManager as any).extractServerUrlFromDestination(destination);
      expect(result).toBeNull();
    });

    it("should handle URLs with different ports", () => {
      const destination = buildDestinationCalendar({
        id: 1,
        integration: "caldav_calendar",
        externalId: "https://caldav.example.com:8443/dav/calendars/user/calendar1/",
      });

      const result = (eventManager as any).extractServerUrlFromDestination(destination);
      expect(result).toBe("https://caldav.example.com:8443");
    });

    it("should handle HTTP URLs", () => {
      const destination = buildDestinationCalendar({
        id: 1,
        integration: "caldav_calendar",
        externalId: "http://caldav.internal.com/dav/calendars/user/calendar1/",
      });

      const result = (eventManager as any).extractServerUrlFromDestination(destination);
      expect(result).toBe("http://caldav.internal.com");
    });
  });

  describe("credentialMatchesDestination", () => {
    it("should return true for matching CalDAV server URLs", () => {
      const credential = buildCalDAVCredential({
        id: 1,
        key: "encrypted_key",
      });

      const destination = buildDestinationCalendar({
        id: 1,
        integration: "caldav_calendar",
        externalId: "https://caldav.example.com/dav/calendars/user/calendar1/",
      });

      mockedSymmetricDecrypt.mockReturnValue(
        JSON.stringify({
          username: "user",
          password: "pass",
          url: "https://caldav.example.com/dav/calendars/user/",
        })
      );

      const result = (eventManager as any).credentialMatchesDestination(credential, destination);
      expect(result).toBe(true);
    });

    it("should return false for non-matching CalDAV server URLs", () => {
      const credential = buildCalDAVCredential({
        id: 1,
        key: "encrypted_key",
      });

      const destination = buildDestinationCalendar({
        id: 1,
        integration: "caldav_calendar",
        externalId: "https://different.example.com/dav/calendars/user/calendar1/",
      });

      mockedSymmetricDecrypt.mockReturnValue(
        JSON.stringify({
          username: "user",
          password: "pass",
          url: "https://caldav.example.com/dav/calendars/user/",
        })
      );

      const result = (eventManager as any).credentialMatchesDestination(credential, destination);
      expect(result).toBe(false);
    });

    it("should return true for non-CalDAV credentials", () => {
      const credential = {
        ...buildCalDAVCredential({ id: 1, key: "encrypted_key" }),
        type: "google_calendar",
      } as CredentialForCalendarService;

      const destination = buildDestinationCalendar({
        id: 1,
        integration: "google_calendar",
        externalId: "calendar@gmail.com",
      });

      const result = (eventManager as any).credentialMatchesDestination(credential, destination);
      expect(result).toBe(true);
      expect(symmetricDecrypt).not.toHaveBeenCalled();
    });

    it("should return true when credential is CalDAV but destination is not", () => {
      const credential = buildCalDAVCredential({
        id: 1,
        key: "encrypted_key",
      });

      const destination = buildDestinationCalendar({
        id: 1,
        integration: "google_calendar",
        externalId: "calendar@gmail.com",
      });

      const result = (eventManager as any).credentialMatchesDestination(credential, destination);
      expect(result).toBe(true);
      expect(symmetricDecrypt).not.toHaveBeenCalled();
    });

    it("should return false when credential URL extraction fails", () => {
      const credential = buildCalDAVCredential({
        id: 1,
        key: "invalid_key",
      });

      const destination = buildDestinationCalendar({
        id: 1,
        integration: "caldav_calendar",
        externalId: "https://caldav.example.com/dav/calendars/user/calendar1/",
      });

      mockedSymmetricDecrypt.mockImplementation(() => {
        throw new Error("Decryption failed");
      });

      const result = (eventManager as any).credentialMatchesDestination(credential, destination);
      expect(result).toBe(false);
    });

    it("should return false when destination URL extraction fails", () => {
      const credential = buildCalDAVCredential({
        id: 1,
        key: "encrypted_key",
      });

      const destination = buildDestinationCalendar({
        id: 1,
        integration: "caldav_calendar",
        externalId: "invalid-url",
      });

      mockedSymmetricDecrypt.mockReturnValue(
        JSON.stringify({
          username: "user",
          password: "pass",
          url: "https://caldav.example.com/dav/calendars/user/",
        })
      );

      const result = (eventManager as any).credentialMatchesDestination(credential, destination);
      expect(result).toBe(false);
    });

    it("should handle URLs with different paths but same server", () => {
      const credential = buildCalDAVCredential({
        id: 1,
        key: "encrypted_key",
      });

      const destination = buildDestinationCalendar({
        id: 1,
        integration: "caldav_calendar",
        externalId: "https://caldav.example.com/different/path/calendar1/",
      });

      mockedSymmetricDecrypt.mockReturnValue(
        JSON.stringify({
          username: "user",
          password: "pass",
          url: "https://caldav.example.com/dav/calendars/user/",
        })
      );

      const result = (eventManager as any).credentialMatchesDestination(credential, destination);
      expect(result).toBe(true);
    });

    it("should handle URLs with different ports as non-matching", () => {
      const credential = buildCalDAVCredential({
        id: 1,
        key: "encrypted_key",
      });

      const destination = buildDestinationCalendar({
        id: 1,
        integration: "caldav_calendar",
        externalId: "https://caldav.example.com:8443/dav/calendars/user/calendar1/",
      });

      mockedSymmetricDecrypt.mockReturnValue(
        JSON.stringify({
          username: "user",
          password: "pass",
          url: "https://caldav.example.com:9443/dav/calendars/user/",
        })
      );

      const result = (eventManager as any).credentialMatchesDestination(credential, destination);
      expect(result).toBe(false);
    });

    it("should handle HTTP vs HTTPS as non-matching", () => {
      const credential = buildCalDAVCredential({
        id: 1,
        key: "encrypted_key",
      });

      const destination = buildDestinationCalendar({
        id: 1,
        integration: "caldav_calendar",
        externalId: "https://caldav.example.com/dav/calendars/user/calendar1/",
      });

      mockedSymmetricDecrypt.mockReturnValue(
        JSON.stringify({
          username: "user",
          password: "pass",
          url: "http://caldav.example.com/dav/calendars/user/",
        })
      );

      const result = (eventManager as any).credentialMatchesDestination(credential, destination);
      expect(result).toBe(false);
    });
  });
});

describe("EventManager Daily video room expiration", () => {
  let eventManager: EventManager;

  beforeEach(() => {
    eventManager = new EventManager({
      user: {
        id: 1,
        email: "test@example.com",
        username: "testuser",
        name: "Test User",
        credentials: [],
        destinationCalendar: null,
        allowDynamicBooking: false,
        lockedDestinationCalendar: false,
        profile: null,
        hasTeamPlan: false,
        organizationId: null,
      },
    });
    vi.clearAllMocks();
  });

  describe("updateAllCalendarEvents with Daily video room", () => {
    it("should recreate meeting when Daily video room is expired (14+ days)", async () => {
      const fifteenDaysAgo = new Date();
      fifteenDaysAgo.setDate(fifteenDaysAgo.getDate() - 15);

      const mockBooking = {
        id: 123,
        uid: "booking-123",
        title: "Test Meeting",
        startTime: fifteenDaysAgo,
        endTime: new Date(fifteenDaysAgo.getTime() + 60 * 60 * 1000), // 1 hour later
        userId: 1,
        attendees: [],
        location: "integrations:daily",
        references: [
          {
            id: 1,
            type: "daily_video",
            uid: "daily-room-123",
            meetingId: "daily-room-123",
            meetingPassword: null,
            meetingUrl: "https://team.daily.co/daily-room-123",
            bookingId: 123,
            externalCalendarId: null,
            deleted: null,
            credentialId: null,
            thirdPartyRecurringEventId: null,
          },
        ],
      };

      const mockEvent = {
        type: "conference" as const,
        id: 123,
        uid: "booking-123",
        title: "Test Meeting",
        startTime: fifteenDaysAgo.toISOString(),
        endTime: new Date(fifteenDaysAgo.getTime() + 60 * 60 * 1000).toISOString(),
        attendees: [],
        organizer: {
          name: "Test User",
          email: "test@example.com",
          timeZone: "UTC",
          language: { locale: "en" },
        },
        location: "integrations:daily",
        conferenceData: {
          createRequest: {
            requestId: "req-123",
          },
        },
        requiresConfirmation: false,
        destinationCalendar: null,
      };

      // Mock the updateLocation method to simulate new room creation
      const updateLocationSpy = vi.spyOn(eventManager as any, "updateLocation").mockResolvedValue({
        results: [
          {
            type: "daily_video",
            success: true,
            uid: "new-daily-room-123",
            createdEvent: {
              id: "new-daily-room-123",
              url: "https://team.daily.co/new-daily-room-123",
            },
          },
        ],
        referencesToCreate: [
          {
            type: "daily_video",
            uid: "new-daily-room-123",
            meetingId: "new-daily-room-123",
            meetingUrl: "https://team.daily.co/new-daily-room-123",
          },
        ],
      });

      const result = await eventManager.updateAllCalendarEvents(mockEvent, mockBooking as any);

      expect(updateLocationSpy).toHaveBeenCalled();
      expect(result.results).toHaveLength(1);
      expect(result.results[0].success).toBe(true);
      expect(result.referencesToCreate).toHaveLength(1);
      expect(result.referencesToCreate[0].meetingUrl).toContain("new-daily-room");
    });

    it("should not recreate meeting when Daily video room is not expired (< 14 days)", async () => {
      const tenDaysAgo = new Date();
      tenDaysAgo.setDate(tenDaysAgo.getDate() - 10);

      const mockBooking = {
        id: 123,
        uid: "booking-123",
        title: "Test Meeting",
        startTime: tenDaysAgo,
        endTime: new Date(tenDaysAgo.getTime() + 60 * 60 * 1000),
        userId: 1,
        attendees: [],
        location: "integrations:daily",
        references: [
          {
            id: 1,
            type: "daily_video",
            uid: "daily-room-123",
            meetingId: "daily-room-123",
            meetingPassword: null,
            meetingUrl: "https://team.daily.co/daily-room-123",
            bookingId: 123,
            externalCalendarId: null,
            deleted: null,
            credentialId: null,
            thirdPartyRecurringEventId: null,
          },
        ],
      };

      const mockEvent = {
        type: "conference" as const,
        id: 123,
        uid: "booking-123",
        title: "Test Meeting",
        startTime: tenDaysAgo.toISOString(),
        endTime: new Date(tenDaysAgo.getTime() + 60 * 60 * 1000).toISOString(),
        attendees: [],
        organizer: {
          name: "Test User",
          email: "test@example.com",
          timeZone: "UTC",
          language: { locale: "en" },
        },
        location: "integrations:daily",
        conferenceData: {
          createRequest: {
            requestId: "req-123",
          },
        },
        requiresConfirmation: false,
        destinationCalendar: null,
      };

      const updateLocationSpy = vi.spyOn(eventManager as any, "updateLocation");

      const result = await eventManager.updateAllCalendarEvents(mockEvent, mockBooking as any);

      // Should not call updateLocation when room is not expired
      expect(updateLocationSpy).not.toHaveBeenCalled();
      expect(result.results).toHaveLength(0);
      expect(result.referencesToCreate).toHaveLength(0);
    });

    it("should check expiration based on booking endTime, not startTime", async () => {
      const thirteenDaysAgo = new Date();
      thirteenDaysAgo.setDate(thirteenDaysAgo.getDate() - 13);

      const mockBooking = {
        id: 123,
        uid: "booking-123",
        title: "Test Meeting",
        startTime: new Date(thirteenDaysAgo.getTime() - 60 * 60 * 1000), // 1 hour before endTime
        endTime: thirteenDaysAgo, // Exactly 13 days ago - should not be expired (needs 14+)
        userId: 1,
        attendees: [],
        location: "integrations:daily",
        references: [
          {
            id: 1,
            type: "daily_video",
            uid: "daily-room-123",
            meetingId: "daily-room-123",
            meetingPassword: null,
            meetingUrl: "https://team.daily.co/daily-room-123",
            bookingId: 123,
            externalCalendarId: null,
            deleted: null,
            credentialId: null,
            thirdPartyRecurringEventId: null,
          },
        ],
      };

      const mockEvent = {
        type: "conference" as const,
        id: 123,
        uid: "booking-123",
        title: "Test Meeting",
        startTime: new Date(thirteenDaysAgo.getTime() - 60 * 60 * 1000).toISOString(),
        endTime: thirteenDaysAgo.toISOString(),
        attendees: [],
        organizer: {
          name: "Test User",
          email: "test@example.com",
          timeZone: "UTC",
          language: { locale: "en" },
        },
        location: "integrations:daily",
        conferenceData: {
          createRequest: {
            requestId: "req-123",
          },
        },
        requiresConfirmation: false,
        destinationCalendar: null,
      };

      const updateLocationSpy = vi.spyOn(eventManager as any, "updateLocation");

      await eventManager.updateAllCalendarEvents(mockEvent, mockBooking as any);

      // Should not recreate as it's exactly 13 days, not 14+
      expect(updateLocationSpy).not.toHaveBeenCalled();
    });

    it("should not check expiration for non-Daily video locations", async () => {
      const fifteenDaysAgo = new Date();
      fifteenDaysAgo.setDate(fifteenDaysAgo.getDate() - 15);

      const mockBooking = {
        id: 123,
        uid: "booking-123",
        title: "Test Meeting",
        startTime: fifteenDaysAgo,
        endTime: new Date(fifteenDaysAgo.getTime() + 60 * 60 * 1000),
        userId: 1,
        attendees: [],
        location: "integrations:zoom", // Not Daily video
        references: [],
      };

      const mockEvent = {
        type: "conference" as const,
        id: 123,
        uid: "booking-123",
        title: "Test Meeting",
        startTime: fifteenDaysAgo.toISOString(),
        endTime: new Date(fifteenDaysAgo.getTime() + 60 * 60 * 1000).toISOString(),
        attendees: [],
        organizer: {
          name: "Test User",
          email: "test@example.com",
          timeZone: "UTC",
          language: { locale: "en" },
        },
        location: "integrations:zoom",
        conferenceData: {
          createRequest: {
            requestId: "req-123",
          },
        },
        requiresConfirmation: false,
        destinationCalendar: null,
      };

      const updateLocationSpy = vi.spyOn(eventManager as any, "updateLocation");

      await eventManager.updateAllCalendarEvents(mockEvent, mockBooking as any);

      // Should not check expiration for non-Daily locations
      expect(updateLocationSpy).not.toHaveBeenCalled();
    });
  });
});
