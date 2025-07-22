import prismock from "../../../../tests/libs/__mocks__/prisma";

import { describe, expect, it, beforeEach } from "vitest";

import prisma from "@calcom/prisma";

import { SelectedCalendarRepository } from "./selectedCalendar";

describe("SelectedCalendarRepository", () => {
  beforeEach(() => {
    prismock.selectedCalendar.deleteMany();
  });

  describe("create", () => {
    it("should create a selected calendar", async () => {
      const data = {
        userId: 1,
        integration: "google_calendar",
        externalId: "test@gmail.com",
        credentialId: 1,
      };

      const result = await SelectedCalendarRepository.create(data);

      expect(result).toEqual(expect.objectContaining(data));
    });

    it("should throw error if we try to create a user-level calendar with same userId_integration_externalId as an existing user-level calendar", async () => {
      const data = {
        userId: 1,
        integration: "google_calendar",
        externalId: "test@gmail.com",
        credentialId: 1,
        eventTypeId: null,
      };

      await SelectedCalendarRepository.create(data);

      await expect(
        SelectedCalendarRepository.create({
          ...data,
          credentialId: 2,
        })
      ).rejects.toThrow("Selected calendar already exists");
    });

    it("should allow creating a user-level calendar with same userId_integration_externalId as an existing event-type level calendar", async () => {
      const data = {
        userId: 1,
        integration: "google_calendar",
        externalId: "test@gmail.com",
        credentialId: 1,
        eventTypeId: 1,
      };

      await SelectedCalendarRepository.create(data);

      const userLevelCalendarData = {
        ...data,
        eventTypeId: null,
      };
      const created = await SelectedCalendarRepository.create(userLevelCalendarData);

      expect(created).toEqual(expect.objectContaining(userLevelCalendarData));
    });
  });

  describe("update", () => {
    it("should update a selected calendar and return it", async () => {
      const calendarToUpdate = await SelectedCalendarRepository.create({
        userId: 1,
        integration: "google_calendar",
        externalId: "test@gmail.com",
        credentialId: 1,
      });

      const updatedCalendar = await SelectedCalendarRepository.update({
        where: { userId: 1, externalId: "test@gmail.com" },
        data: { integration: "office365_calendar" },
      });

      expect(updatedCalendar.id).toBe(calendarToUpdate.id);
      expect(updatedCalendar.integration).toBe("office365_calendar");
    });

    it("should throw error when trying to update multiple calendars", async () => {
      await SelectedCalendarRepository.create({
        userId: 1,
        integration: "google_calendar",
        externalId: "test1@gmail.com",
        credentialId: 1,
      });

      await SelectedCalendarRepository.create({
        userId: 1,
        integration: "google_calendar2",
        externalId: "test1@gmail.com",
        credentialId: 2,
      });

      await expect(
        SelectedCalendarRepository.update({
          where: { userId: 1, externalId: "test1@gmail.com" },
          data: { integration: "office365_calendar" },
        })
      ).rejects.toThrow(
        "Multiple SelectedCalendar records found to update. updateMany should be used instead"
      );
    });
  });

  describe("delete", () => {
    it("should delete a selected calendar and return it", async () => {
      const calendar = await SelectedCalendarRepository.create({
        userId: 1,
        integration: "google_calendar",
        externalId: "test@gmail.com",
        credentialId: 1,
      });

      const deleted = await SelectedCalendarRepository.delete({
        where: {
          userId: 1,
          integration: "google_calendar",
          externalId: "test@gmail.com",
          credentialId: 1,
        },
      });

      expect(deleted).toEqual(calendar);
      const result = await SelectedCalendarRepository.findFirst({
        where: { id: calendar.id },
      });

      expect(result).toBeNull();
    });

    it("should throw error when trying to delete non-existent calendar", async () => {
      await expect(
        SelectedCalendarRepository.delete({
          where: {
            userId: 999,
            integration: "google_calendar",
            externalId: "nonexistent@gmail.com",
            credentialId: 999,
          },
        })
      ).rejects.toThrow("SelectedCalendar not found");
    });
  });

  describe("findUserLevelUniqueOrThrow", () => {
    it("should find user level calendar", async () => {
      const calendar = await SelectedCalendarRepository.create({
        userId: 1,
        integration: "google_calendar",
        externalId: "test@gmail.com",
        credentialId: 1,
        eventTypeId: null, // User level calendar
      });

      const result = await SelectedCalendarRepository.findUserLevelUniqueOrThrow({
        where: { userId: 1, externalId: "test@gmail.com" },
      });

      expect(result).toEqual(
        expect.objectContaining({
          userId: calendar.userId,
          externalId: calendar.externalId,
        })
      );
    });

    it("should not find event type level calendar", async () => {
      await SelectedCalendarRepository.create({
        userId: 1,
        integration: "google_calendar",
        externalId: "test@gmail.com",
        credentialId: 1,
        eventTypeId: 1, // Event type level calendar
      });

      await expect(
        SelectedCalendarRepository.findUserLevelUniqueOrThrow({
          where: { userId: 1, externalId: "test@gmail.com" },
        })
      ).rejects.toThrow("SelectedCalendar not found");
    });
  });

  describe("upsert", () => {
    describe("User Level Calendar", () => {
      const eventTypeId = null;
      it("should update existing calendar as long as a record with same userId_integration_externalId is present for eventTypeId=null", async () => {
        const initialData = {
          userId: 1,
          integration: "google_calendar",
          externalId: "test@gmail.com",
          credentialId: 1,
          eventTypeId,
        };

        const existingCalendar = await SelectedCalendarRepository.create(initialData);

        const updatedData = {
          ...initialData,
          credentialId: 2,
          eventTypeId,
        };

        const result = await SelectedCalendarRepository.upsert(updatedData);

        expect(result.credentialId).toBe(2);
        expect(existingCalendar.id).toBe(result.id);
      });

      it("should create new calendar if no record with same userId_integration_externalId is present when eventTypeId=null", async () => {
        const initialData = {
          userId: 1,
          integration: "google_calendar",
          externalId: "test@gmail.com",
          credentialId: 1,
          eventTypeId,
        };

        const existingCalendar = await SelectedCalendarRepository.create(initialData);

        const updatedData = {
          ...initialData,
          externalId: "test2@gmail.com",
          credentialId: 2,
          eventTypeId,
        };

        const result = await SelectedCalendarRepository.upsert(updatedData);
        expect(await prisma.selectedCalendar.count()).toBe(2);
        expect(result).toEqual(expect.objectContaining(updatedData));
        expect(existingCalendar.id).not.toBe(result.id);
      });
    });

    describe("Event Type Level Calendar", () => {
      const eventTypeId = 101;
      it("should update existing calendar as long as record with same userId_integration_externalId_eventTypeId is present", async () => {
        const initialData = {
          userId: 1,
          integration: "google_calendar",
          externalId: "test@gmail.com",
          credentialId: 1,
          eventTypeId,
        };

        const existingCalendar = await SelectedCalendarRepository.create(initialData);

        const updatedData = {
          ...initialData,
          credentialId: 2,
          eventTypeId,
        };

        const result = await SelectedCalendarRepository.upsert(updatedData);

        expect(result.credentialId).toBe(2);
        expect(existingCalendar.id).toBe(result.id);
      });

      it("should create new calendar if no record with same userId_integration_externalId_eventTypeId is present", async () => {
        const initialData = {
          userId: 1,
          integration: "google_calendar",
          externalId: "test@gmail.com",
          credentialId: 1,
          eventTypeId,
        };

        const existingCalendar = await SelectedCalendarRepository.create(initialData);

        const updatedData = {
          ...initialData,
          credentialId: 2,
          externalId: "test2@gmail.com",
          eventTypeId,
        };

        const result = await SelectedCalendarRepository.upsert(updatedData);
        expect(await prisma.selectedCalendar.count()).toBe(2);
        expect(result).toEqual(expect.objectContaining(updatedData));
        expect(existingCalendar.id).not.toBe(result.id);
      });
    });
  });

  describe("Delegation Credential", () => {
    it("should create a selected calendar with delegationCredentialId", async () => {
      const data = {
        userId: 1,
        integration: "google_calendar",
        externalId: "test@gmail.com",
        delegationCredentialId: "delegationCredential-123",
      };

      const result = await SelectedCalendarRepository.create(data);

      expect(result).toEqual(expect.objectContaining(data));
    });

    describe("upsert", () => {
      describe("updation", () => {
        it("should update existing record with delegationCredentialId if credentialId is -1", async () => {
          const initialData = {
            userId: 1,
            integration: "google_calendar",
            externalId: "test@gmail.com",
            credentialId: 1,
            eventTypeId: null,
          };

          const existingCalendar = await SelectedCalendarRepository.create(initialData);

          const data = {
            userId: 1,
            integration: "google_calendar",
            externalId: "test@gmail.com",
            credentialId: -1,
            delegationCredentialId: "delegationCredential-123",
          };

          const result = await SelectedCalendarRepository.upsert(data);
          expect(result.id).not.toBe(null);
          expect(result.id).toBe(existingCalendar.id);
          expect(result.credentialId).toBe(null);
          expect(result.delegationCredentialId).toBe(data.delegationCredentialId);
        });

        it("should update existing record with credentialId if credentialId is valid(>0) even if delegationCredentialId is set", async () => {
          const initialData = {
            userId: 1,
            integration: "google_calendar",
            externalId: "test@gmail.com",
            credentialId: 1,
            eventTypeId: null,
          };

          const existingCalendar = await SelectedCalendarRepository.create(initialData);

          const data = {
            userId: 1,
            integration: "google_calendar",
            externalId: "test@gmail.com",
            credentialId: 2,
            delegationCredentialId: "delegationCredential-123",
          };
          const beforeDelegationCredentialId = data.delegationCredentialId;

          const result = await SelectedCalendarRepository.upsert(data);
          expect(result.id).not.toBe(null);
          expect(result.id).toBe(existingCalendar.id);
          expect(result.credentialId).toBe(data.credentialId);
          expect(result.delegationCredentialId).toBe(beforeDelegationCredentialId);
        });
      });

      describe("creation", () => {
        it("should create a new record with delegationCredentialId if credentialId is -1", async () => {
          const initialData = {
            userId: 1,
            integration: "google_calendar",
            externalId: "test@gmail.com",
            credentialId: 1,
            eventTypeId: null,
          };

          const existingCalendar = await SelectedCalendarRepository.create(initialData);

          const data = {
            userId: 1,
            integration: "google_calendar",
            externalId: "anotheremail@gmail.com",
            credentialId: -1,
            delegationCredentialId: "delegationCredential-123",
          };

          // It will create a new record because of unique constraint violation
          const result = await SelectedCalendarRepository.upsert(data);
          expect(result.id).not.toBe(null);
          expect(result.id).not.toBe(existingCalendar.id);
          expect(result.credentialId).toBe(null);
          expect(result.delegationCredentialId).toBe(data.delegationCredentialId);
        });
      });

      it("shouldnt update existing delegationCredentialId if upsert data doesn't have it", async () => {
        const initialData = {
          userId: 1,
          integration: "google_calendar",
          externalId: "test@gmail.com",
          eventTypeId: null,
          delegationCredentialId: "delegationCredential-123",
          credentialId: 1,
        };

        const beforeDelegationCredentialId = initialData.delegationCredentialId;
        const existingCalendar = await SelectedCalendarRepository.create(initialData);

        const data = {
          userId: 1,
          integration: "google_calendar",
          externalId: "test@gmail.com",
          credentialId: 1,
        };

        const result = await SelectedCalendarRepository.upsert(data);
        expect(result.id).toBe(existingCalendar.id);
        expect(result.credentialId).toBe(existingCalendar.credentialId);
        expect(result.delegationCredentialId).toBe(beforeDelegationCredentialId);
      });

      it("shouldnt update delegationCredentialId if it is undefined", async () => {
        const initialData = {
          userId: 1,
          integration: "google_calendar",
          externalId: "test@gmail.com",
          eventTypeId: null,
          delegationCredentialId: "delegationCredential-123",
        };

        const existingCalendar = await SelectedCalendarRepository.create(initialData);

        const data = {
          userId: 1,
          integration: "google_calendar",
          externalId: "test@gmail.com",
        };

        const result = await SelectedCalendarRepository.upsert(data);
        expect(result.id).toBe(existingCalendar.id);
        expect(result.credentialId).toBe(existingCalendar.credentialId);
        expect(result.delegationCredentialId).toBe(existingCalendar.delegationCredentialId);
      });
    });
  });

  describe("getNextBatchToWatch", () => {
    beforeEach(async () => {
      // Clean up
      await prismock.selectedCalendar.deleteMany();
      await prismock.user.deleteMany();
      await prismock.team.deleteMany();
      await prismock.membership.deleteMany();
      await prismock.feature.deleteMany();
    });

    it("should return Google calendars for users with team calendar-cache feature", async () => {
      // Create feature
      const feature = await prismock.feature.create({
        data: {
          slug: "calendar-cache",
          enabled: true,
          description: "Calendar Cache",
          type: "OPERATIONAL",
        },
      });

      // Create team with calendar-cache feature
      const team = await prismock.team.create({
        data: {
          name: "Test Team",
          slug: "test-team",
          features: {
            create: {
              featureId: feature.slug,
            },
          },
        },
      });

      // Create user
      const user = await prismock.user.create({
        data: {
          email: "test@example.com",
          username: "testuser",
        },
      });

      // Add user to team
      await prismock.membership.create({
        data: {
          userId: user.id,
          teamId: team.id,
          role: "MEMBER",
          accepted: true,
        },
      });

      // Create Google calendar without subscription
      await prismock.selectedCalendar.create({
        data: {
          userId: user.id,
          integration: "google_calendar",
          externalId: "google-cal-1",
          credentialId: 1,
        },
      });

      // Create Google calendar with expired subscription
      const expiredDate = new Date(Date.now() - 1000).toISOString();
      await prismock.selectedCalendar.create({
        data: {
          userId: user.id,
          integration: "google_calendar",
          externalId: "google-cal-2",
          credentialId: 1,
          googleChannelExpiration: expiredDate,
        },
      });

      const result = await SelectedCalendarRepository.getNextBatchToWatch();

      expect(result).toHaveLength(2);
      expect(result.every((cal) => cal.integration === "google_calendar")).toBe(true);
    });

    it("should return Office365 calendars without team requirement", async () => {
      const user = await prismock.user.create({
        data: {
          email: "test@example.com",
          username: "testuser",
        },
      });

      // Create Office365 calendar without subscription
      await prismock.selectedCalendar.create({
        data: {
          userId: user.id,
          integration: "office365_calendar",
          externalId: "office365-cal-1",
          credentialId: 1,
        },
      });

      // Create Office365 calendar with expired subscription
      const expiredDate = new Date(Date.now() - 1000);
      await prismock.selectedCalendar.create({
        data: {
          userId: user.id,
          integration: "office365_calendar",
          externalId: "office365-cal-2",
          credentialId: 1,
          office365SubscriptionExpiration: expiredDate,
        },
      });

      const result = await SelectedCalendarRepository.getNextBatchToWatch();

      expect(result).toHaveLength(2);
      expect(result.every((cal) => cal.integration === "office365_calendar")).toBe(true);
    });

    it("should not return calendars that have reached max watch attempts", async () => {
      const user = await prismock.user.create({
        data: {
          email: "test@example.com",
          username: "testuser",
        },
      });

      // Create calendar with max attempts reached
      await prismock.selectedCalendar.create({
        data: {
          userId: user.id,
          integration: "office365_calendar",
          externalId: "office365-cal-1",
          credentialId: 1,
          error: "Some error",
          watchAttempts: 10, // Assuming maxAttempts is 10
        },
      });

      const result = await SelectedCalendarRepository.getNextBatchToWatch();

      expect(result).toHaveLength(0);
    });
  });

  describe("getNextBatchToUnwatch", () => {
    beforeEach(async () => {
      await prismock.selectedCalendar.deleteMany();
      await prismock.user.deleteMany();
      await prismock.team.deleteMany();
      await prismock.membership.deleteMany();
      await prismock.feature.deleteMany();
    });

    it("should unwatch all calendars when global calendar-cache is disabled", async () => {
      // Create feature but disable it
      await prismock.feature.create({
        data: {
          slug: "calendar-cache",
          enabled: false,
          description: "Calendar Cache",
          type: "OPERATIONAL",
        },
      });

      const user = await prismock.user.create({
        data: {
          email: "test@example.com",
          username: "testuser",
        },
      });

      // Create Google calendar with active subscription
      await prismock.selectedCalendar.create({
        data: {
          userId: user.id,
          integration: "google_calendar",
          externalId: "google-cal-1",
          credentialId: 1,
          googleChannelExpiration: new Date(Date.now() + 100000).toISOString(),
        },
      });

      // Create Office365 calendar with active subscription
      await prismock.selectedCalendar.create({
        data: {
          userId: user.id,
          integration: "office365_calendar",
          externalId: "office365-cal-1",
          credentialId: 2,
          office365SubscriptionExpiration: new Date(Date.now() + 100000),
        },
      });

      const result = await SelectedCalendarRepository.getNextBatchToUnwatch();

      expect(result).toHaveLength(2);
      expect(result.some((cal) => cal.integration === "google_calendar")).toBe(true);
      expect(result.some((cal) => cal.integration === "office365_calendar")).toBe(true);
    });

    it("should only unwatch Google calendars for teams without calendar-cache feature when global is enabled", async () => {
      // Create feature and enable it
      await prismock.feature.create({
        data: {
          slug: "calendar-cache",
          enabled: true,
          description: "Calendar Cache",
          type: "OPERATIONAL",
        },
      });

      // Create team WITHOUT calendar-cache feature
      const team = await prismock.team.create({
        data: {
          name: "Test Team",
          slug: "test-team",
        },
      });

      const user = await prismock.user.create({
        data: {
          email: "test@example.com",
          username: "testuser",
        },
      });

      // Add user to team
      await prismock.membership.create({
        data: {
          userId: user.id,
          teamId: team.id,
          role: "MEMBER",
          accepted: true,
        },
      });

      // Create Google calendar with active subscription
      await prismock.selectedCalendar.create({
        data: {
          userId: user.id,
          integration: "google_calendar",
          externalId: "google-cal-1",
          credentialId: 1,
          googleChannelExpiration: new Date(Date.now() + 100000).toISOString(),
        },
      });

      // Create Office365 calendar with active subscription (should NOT be unwatched)
      await prismock.selectedCalendar.create({
        data: {
          userId: user.id,
          integration: "office365_calendar",
          externalId: "office365-cal-1",
          credentialId: 2,
          office365SubscriptionExpiration: new Date(Date.now() + 100000),
        },
      });

      const result = await SelectedCalendarRepository.getNextBatchToUnwatch();

      expect(result).toHaveLength(1);
      expect(result[0].integration).toBe("google_calendar");
    });

    it("should not return calendars that have reached max unwatch attempts", async () => {
      await prismock.feature.create({
        data: {
          slug: "calendar-cache",
          enabled: false,
          description: "Calendar Cache",
          type: "OPERATIONAL",
        },
      });

      const user = await prismock.user.create({
        data: {
          email: "test@example.com",
          username: "testuser",
        },
      });

      // Create calendar with max unwatch attempts reached
      await prismock.selectedCalendar.create({
        data: {
          userId: user.id,
          integration: "office365_calendar",
          externalId: "office365-cal-1",
          credentialId: 1,
          office365SubscriptionExpiration: new Date(Date.now() + 100000),
          error: "Some error",
          unwatchAttempts: 10, // Assuming maxAttempts is 10
        },
      });

      const result = await SelectedCalendarRepository.getNextBatchToUnwatch();

      expect(result).toHaveLength(0);
    });
  });

  describe("findFirstByOffice365SubscriptionId", () => {
    it("should find calendar by Office365 subscription ID", async () => {
      const user = await prismock.user.create({
        data: {
          email: "test@example.com",
          username: "testuser",
        },
      });

      const credential = await prismock.credential.create({
        data: {
          type: "office365_calendar",
          userId: user.id,
          key: {},
          appId: "office365-calendar",
        },
      });

      await prismock.selectedCalendar.create({
        data: {
          userId: user.id,
          integration: "office365_calendar",
          externalId: "office365-cal-1",
          credentialId: credential.id,
          office365SubscriptionId: "sub-123",
        },
      });

      await prismock.selectedCalendar.create({
        data: {
          userId: user.id,
          integration: "office365_calendar",
          externalId: "office365-cal-2",
          credentialId: credential.id,
          office365SubscriptionId: "sub-456",
        },
      });

      const result = await SelectedCalendarRepository.findFirstByOffice365SubscriptionId("sub-123");

      expect(result).toBeTruthy();
      expect(result?.credential).toBeTruthy();
      expect(result?.credential?.id).toBe(credential.id);
    });

    it("should return null when subscription ID not found", async () => {
      const result = await SelectedCalendarRepository.findFirstByOffice365SubscriptionId("non-existent");

      expect(result).toBeNull();
    });
  });

  describe("upsertManyForEventTypeIds", () => {
    it("should upsert calendars for multiple event type IDs", async () => {
      const user = await prismock.user.create({
        data: {
          email: "test@example.com",
          username: "testuser",
        },
      });

      const data = {
        userId: user.id,
        integration: "office365_calendar",
        externalId: "office365-cal-1",
        credentialId: 1,
        office365SubscriptionId: "sub-123",
      };

      const eventTypeIds = [null, 1, 2];

      const results = await SelectedCalendarRepository.upsertManyForEventTypeIds({
        data,
        eventTypeIds,
      });

      expect(results).toHaveLength(3);
      expect(results[0].eventTypeId).toBeNull();
      expect(results[1].eventTypeId).toBe(1);
      expect(results[2].eventTypeId).toBe(2);
      expect(results.every((r) => r.office365SubscriptionId === "sub-123")).toBe(true);
    });

    it("should update existing calendars when upserting", async () => {
      const user = await prismock.user.create({
        data: {
          email: "test@example.com",
          username: "testuser",
        },
      });

      // Create existing calendar
      await prismock.selectedCalendar.create({
        data: {
          userId: user.id,
          integration: "office365_calendar",
          externalId: "office365-cal-1",
          credentialId: 1,
          eventTypeId: 1,
          office365SubscriptionId: "old-sub",
        },
      });

      const data = {
        userId: user.id,
        integration: "office365_calendar",
        externalId: "office365-cal-1",
        credentialId: 1,
        office365SubscriptionId: "new-sub",
      };

      const eventTypeIds = [1, 2];

      const results = await SelectedCalendarRepository.upsertManyForEventTypeIds({
        data,
        eventTypeIds,
      });

      expect(results).toHaveLength(2);
      expect(results[0].office365SubscriptionId).toBe("new-sub");
      expect(results[1].office365SubscriptionId).toBe("new-sub");
    });
  });
});
