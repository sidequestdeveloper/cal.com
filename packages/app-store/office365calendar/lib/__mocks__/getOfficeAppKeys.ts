import { vi } from "vitest";

export const defaultOfficeKeys = {
  client_id: "FAKE_CLIENT_ID",
  client_secret: "FAKE_CLIENT_SECRET",
};

const getOfficeAppKeysMock = vi.fn().mockResolvedValue(defaultOfficeKeys);

vi.mock("../getOfficeAppKeys", () => {
  return {
    getOfficeAppKeys: getOfficeAppKeysMock,
  };
});

export { getOfficeAppKeysMock };
