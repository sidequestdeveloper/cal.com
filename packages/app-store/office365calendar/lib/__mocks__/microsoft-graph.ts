import { vi } from "vitest";

export const mockUserEndpoint = "https://graph.microsoft.com/v1.0/me";

export const mockCalendarList = [
  {
    id: "calendar-1",
    name: "Primary Calendar",
    isDefaultCalendar: true,
    canEdit: true,
    canShare: true,
    canViewPrivateItems: true,
    owner: {
      name: "Test User",
      address: "test@example.com",
    },
  },
  {
    id: "calendar-2",
    name: "Secondary Calendar",
    isDefaultCalendar: false,
    canEdit: true,
    canShare: true,
    canViewPrivateItems: true,
    owner: {
      name: "Test User",
      address: "test@example.com",
    },
  },
];

export const mockEvents = [
  {
    id: "event-1",
    subject: "Meeting 1",
    start: {
      dateTime: "2024-01-01T10:00:00",
      timeZone: "UTC",
    },
    end: {
      dateTime: "2024-01-01T11:00:00",
      timeZone: "UTC",
    },
    showAs: "busy",
    isAllDay: false,
    isCancelled: false,
    isReminderOn: true,
    location: {
      displayName: "Conference Room A",
    },
    body: {
      contentType: "text",
      content: "Meeting description",
    },
    organizer: {
      emailAddress: {
        name: "Test User",
        address: "test@example.com",
      },
    },
    attendees: [],
  },
  {
    id: "event-2",
    subject: "Meeting 2",
    start: {
      dateTime: "2024-01-01T14:00:00",
      timeZone: "UTC",
    },
    end: {
      dateTime: "2024-01-01T15:00:00",
      timeZone: "UTC",
    },
    showAs: "busy",
    isAllDay: false,
    isCancelled: false,
    isReminderOn: true,
    location: {
      displayName: "Virtual",
    },
    body: {
      contentType: "text",
      content: "Virtual meeting",
    },
    organizer: {
      emailAddress: {
        name: "Test User",
        address: "test@example.com",
      },
    },
    attendees: [],
  },
];

export const mockSubscription = {
  id: "subscription-123",
  resource: "me/calendars/calendar-1/events",
  changeType: "created,updated,deleted",
  clientState: "test-client-state",
  notificationUrl: "https://example.com/api/integrations/office365calendar/webhook",
  expirationDateTime: new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toISOString(),
};

export const mockBatchResponse = {
  responses: [
    {
      id: "0",
      status: 200,
      headers: {},
      body: {
        value: mockEvents,
      },
    },
  ],
};

export const fetcherMock = vi.fn();

export const setFetcherMockResponse = (response: any) => {
  fetcherMock.mockResolvedValueOnce(response);
};

export const setFetcherMockError = (error: any) => {
  fetcherMock.mockRejectedValueOnce(error);
};

export const resetMocks = () => {
  fetcherMock.mockReset();
};
