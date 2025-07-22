import prismock from "../../../../../tests/libs/__mocks__/prisma";
import getCalendarMock from "../../../tests/__mocks__/getCalendar";

import type { NextApiRequest, NextApiResponse } from "next";
import { expect, test, beforeEach, vi, describe } from "vitest";

import { HttpError } from "@calcom/lib/http-error";

import webhook from "../webhook";

// Mock handlers - we'll call the handlers directly
const webhookHandlers = webhook as any;

vi.mock("@calcom/lib/delegationCredential/server", () => ({
  getCredentialForCalendarCache: vi.fn().mockResolvedValue({
    id: 1,
    type: "office365_calendar",
    key: {
      access_token: "fake_access_token",
    },
    userId: 1,
  }),
}));

vi.mock("../../../_utils/getCalendar", () => ({
  getCalendar: getCalendarMock,
}));

// const log = logger.getSubLogger({ prefix: ["office365calendar", "webhook", "test"] });

const mockCalendarService = {
  fetchAvailabilityAndSetCache: vi.fn(),
};

getCalendarMock.mockResolvedValue(mockCalendarService);

const createMockRequest = (overrides: Partial<NextApiRequest> = {}): NextApiRequest => {
  return {
    method: "GET",
    query: {},
    body: {},
    headers: {},
    ...overrides,
  } as NextApiRequest;
};

const createMockResponse = (): NextApiResponse => {
  const res = {} as NextApiResponse;
  res.status = vi.fn().mockReturnValue(res);
  res.json = vi.fn().mockReturnValue(res);
  res.send = vi.fn().mockReturnValue(res);
  res.setHeader = vi.fn().mockReturnValue(res);
  return res;
};

describe("Office365 Calendar Webhook", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    process.env.OFFICE365_WEBHOOK_CLIENT_STATE = "test-client-state";
    prismock.selectedCalendar.deleteMany();
    prismock.credential.deleteMany();
  });

  describe("GET handler", () => {
    test("should validate subscription with validation token", async () => {
      const req = createMockRequest({
        method: "GET",
        query: {
          validationToken: "test-validation-token",
        },
      });
      const res = createMockResponse();

      const handler = await webhookHandlers.GET;
      await handler.default(req, res);

      expect(res.setHeader).toHaveBeenCalledWith("Content-Type", "text/plain");
      expect(res.status).toHaveBeenCalledWith(200);
      expect(res.send).toHaveBeenCalledWith("test-validation-token");
    });

    test("should throw error when validation token is missing", async () => {
      const req = createMockRequest({
        method: "GET",
        query: {},
      });
      const res = createMockResponse();

      const handler = await webhookHandlers.GET;
      await expect(handler.default(req, res)).rejects.toThrow(HttpError);
      await expect(handler.default(req, res)).rejects.toThrow("Missing validation token");
    });
  });

  describe("POST handler", () => {
    test("should handle validation token in POST request", async () => {
      const req = createMockRequest({
        method: "POST",
        query: {
          validationToken: "test-validation-token",
        },
      });
      const res = createMockResponse();

      const handler = await webhookHandlers.POST;
      await handler.default(req, res);

      expect(res.setHeader).toHaveBeenCalledWith("Content-Type", "text/plain");
      expect(res.status).toHaveBeenCalledWith(200);
      expect(res.send).toHaveBeenCalledWith("test-validation-token");
    });

    test("should process webhook notifications successfully", async () => {
      // Set up test data
      const credential = await prismock.credential.create({
        data: {
          id: 1,
          type: "office365_calendar",
          key: {
            access_token: "fake_access_token",
          },
          userId: 1,
          appId: "office365-calendar",
        },
      });

      await prismock.selectedCalendar.create({
        data: {
          credentialId: credential.id,
          userId: 1,
          integration: "office365_calendar",
          externalId: "calendar-1",
          office365SubscriptionId: "subscription-123",
        },
      });

      const notifications = [
        {
          subscriptionId: "subscription-123",
          clientState: "test-client-state",
          resource: "me/calendars/calendar-1/events",
          changeType: "created",
        },
      ];

      const req = createMockRequest({
        method: "POST",
        body: {
          value: notifications,
        },
      });
      const res = createMockResponse();

      const handler = await webhookHandlers.POST;
      const result = await handler.default(req, res);

      expect(result).toEqual({ message: "ok" });
      expect(mockCalendarService.fetchAvailabilityAndSetCache).toHaveBeenCalledTimes(1);
    });

    test("should reject notifications with invalid client state", async () => {
      const notifications = [
        {
          subscriptionId: "subscription-123",
          clientState: "invalid-client-state",
          resource: "me/calendars/calendar-1/events",
        },
      ];

      const req = createMockRequest({
        method: "POST",
        body: {
          value: notifications,
        },
      });
      const res = createMockResponse();

      const handler = await webhookHandlers.POST;
      await expect(handler.default(req, res)).rejects.toThrow(HttpError);
      await expect(handler.default(req, res)).rejects.toThrow("Invalid client state");
    });

    test("should throw error when webhook client state is not configured", async () => {
      delete process.env.OFFICE365_WEBHOOK_CLIENT_STATE;

      const req = createMockRequest({
        method: "POST",
        body: {
          value: [{ subscriptionId: "123" }],
        },
      });
      const res = createMockResponse();

      const handler = await webhookHandlers.POST;
      await expect(handler.default(req, res)).rejects.toThrow(HttpError);
      await expect(handler.default(req, res)).rejects.toThrow("Webhook not configured");
    });

    test("should throw error for invalid notification payload", async () => {
      const req = createMockRequest({
        method: "POST",
        body: {
          // Missing 'value' property
        },
      });
      const res = createMockResponse();

      const handler = await webhookHandlers.POST;
      await expect(handler.default(req, res)).rejects.toThrow(HttpError);
      await expect(handler.default(req, res)).rejects.toThrow("Invalid notification payload");
    });

    test("should continue processing other notifications when one fails", async () => {
      // Set up test data
      const credential1 = await prismock.credential.create({
        data: {
          id: 1,
          type: "office365_calendar",
          key: {
            access_token: "fake_access_token",
          },
          userId: 1,
          appId: "office365-calendar",
        },
      });

      const credential2 = await prismock.credential.create({
        data: {
          id: 2,
          type: "office365_calendar",
          key: {
            access_token: "fake_access_token",
          },
          userId: 2,
          appId: "office365-calendar",
        },
      });

      await prismock.selectedCalendar.create({
        data: {
          credentialId: credential1.id,
          userId: 1,
          integration: "office365_calendar",
          externalId: "calendar-1",
          office365SubscriptionId: "subscription-123",
        },
      });

      await prismock.selectedCalendar.create({
        data: {
          credentialId: credential2.id,
          userId: 2,
          integration: "office365_calendar",
          externalId: "calendar-2",
          office365SubscriptionId: "subscription-456",
        },
      });

      const notifications = [
        {
          subscriptionId: "subscription-123",
          clientState: "test-client-state",
          resource: "me/calendars/calendar-1/events",
        },
        {
          subscriptionId: "subscription-not-found", // This will not be found
          clientState: "test-client-state",
          resource: "me/calendars/unknown/events",
        },
        {
          subscriptionId: "subscription-456",
          clientState: "test-client-state",
          resource: "me/calendars/calendar-2/events",
        },
      ];

      const req = createMockRequest({
        method: "POST",
        body: {
          value: notifications,
        },
      });
      const res = createMockResponse();

      const handler = await webhookHandlers.POST;
      const result = await handler.default(req, res);

      expect(result).toEqual({ message: "ok" });
      // Should be called twice (for subscription-123 and subscription-456)
      expect(mockCalendarService.fetchAvailabilityAndSetCache).toHaveBeenCalledTimes(2);
    });

    test("should handle notifications without required fields gracefully", async () => {
      const notifications = [
        {
          clientState: "test-client-state",
          // Missing subscriptionId and resource
        },
        {
          subscriptionId: "subscription-123",
          clientState: "test-client-state",
          // Missing resource
        },
      ];

      const req = createMockRequest({
        method: "POST",
        body: {
          value: notifications,
        },
      });
      const res = createMockResponse();

      const handler = await webhookHandlers.POST;
      const result = await handler.default(req, res);

      expect(result).toEqual({ message: "ok" });
      expect(mockCalendarService.fetchAvailabilityAndSetCache).not.toHaveBeenCalled();
    });
  });
});
