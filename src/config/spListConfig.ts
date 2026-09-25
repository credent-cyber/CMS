/**
 * SharePoint Lists and Libraries Configuration
 * Centralized place to manage all list and library names
 * Update names here and they'll be used throughout the application
 */
/* eslint-disable @typescript-eslint/no-explicit-any */
export const LIST_CONFIG = {
  // Lists
  LISTS: {
    EMPLOYEE_MASTER: "EmployeeMaster",
    // EMPLOYEE_MASTER: "EmployeesDummy",
    BuddyAllocate: "Buddy Allocation",
    BUDDY_QUESTIONS: "Buddy Interaction Questions Master",
    MEETING_ANSWERS: "Buddy Interaction with NJ Data",
    MemberSuccessStories: "BuddyMemberSuccessStories",
    FEEDBACK_QUESTIONS_MASTER: "JoineeFeedbackQuestionsMaster",
    Holidays: "NeulandHolidays",
    Location_Master: "BuddyLocationMaster",
    BUDDY_HELPDESK_REQUEST: "BuddyHelpDeskRequest",
    IT_Helpdesk: "ITHelpDeskMaster",
    BuddyNewJoineeCriteria: "BuddyNewJoineeCriteria",
    BUDDY_DOC_CATEGORY_MASTER: "Buddy Doc Category Master",
  },

  // Libraries3
  LIBRARIES: {
    BuddyBannerTemplate: "BuddyBannerTemplates",
    BuddyLogo: "BuddyLogoImg",
    BuddyGallery: "BuddyGallery",
    HolidaysTemplate: "HolidayTemplate",
    UploadedHolidaysTemplate: "UploadedHolidaysTemplate",
    BuddyUserGuides: "Buddy User Guides",
  },

  // SharePoint library GUIDs. Keep the value blank until the library exists;
  // consumers will safely fall back to the configured library title.
  LIBRARY_GUIDS: {
    BuddyUserGuides: "Buddy User Guides", 
  },
  SITENAME: {
    // --- for local development site URL ---
    // BUDDY_EMP_MASTER_SITE: "https://neulandlaboratories.sharepoint.com/sites/AppsDev"

    // --- for Prod site URL ---
    // BUDDY_EMP_MASTER_SITE: "https://neulandlaboratories.sharepoint.com",

    // --- for Dev site URL ---
    BUDDY_EMP_MASTER_SITE: "https://credentinfotec.sharepoint.com/sites/NeulandLabs",
  },
};


export const SHAREPOINT_ID_WINDOW_SIZE = 2000;
const SHAREPOINT_READ_RETRY_COUNT = 3;

const executeSharePointRead = async <T>(
  operation: () => Promise<T>,
): Promise<T> => { 
  let lastError: any;

  for (let attempt = 0; attempt < SHAREPOINT_READ_RETRY_COUNT; attempt += 1) {
    try {
      return await operation();
    } catch (error) {
      lastError = error;
      const status = Number(
        (error as any)?.status || (error as any)?.response?.status || 0,
      );
      const message = String((error as any)?.message || "").toLowerCase();
      const retryable =
        status === 429 ||
        status === 503 ||
        status === 504 ||
        message.includes("throttl") ||
        message.includes("temporar");

      if (!retryable || attempt === SHAREPOINT_READ_RETRY_COUNT - 1)
        throw error;
      await new Promise((resolve) =>
        window.setTimeout(resolve, 400 * 2 ** attempt),
      );
    }
  }

  throw lastError;
};

export interface ThresholdSafeListReadOptions<T = any> {
  web: any;
  listTitle: string;
  listId?: string;
  select: string;
  filter?: string;
  expand?: string;
  includeItem?: (item: T) => boolean;
  sort?: (left: T, right: T) => number;
  maxItems?: number;
}

export const getThresholdSafeListItems = async <T = any>({
  web,
  listTitle,
  listId,
  select,
  filter,
  expand,
  includeItem,
  sort,
  maxItems,
}: ThresholdSafeListReadOptions<T>): Promise<T[]> => {
  const normalizedListId = String(listId || "")
    .trim()
    .replace(/[{}]/g, "");
  const hasValidListId =
    /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
      normalizedListId,
    );
  const list = hasValidListId
    ? web.lists.getById(normalizedListId)
    : web.lists.getByTitle(listTitle);
  const latest = await executeSharePointRead<Array<{ ID?: number | string }>>(
    () => list.items.select("ID").orderBy("ID", false).top(1).get(),
  );

  let upperId = Number(latest?.[0]?.ID || 0);
  const results: T[] = [];
  const businessFilter = String(filter || "").trim();

  while (upperId > 0) {
    const lowerId = Math.max(0, upperId - SHAREPOINT_ID_WINDOW_SIZE);
    const idWindow = `ID le ${upperId} and ID gt ${lowerId}`;
    const combinedFilter = businessFilter
      ? `${idWindow} and (${businessFilter})`
      : idWindow;

    let query: any = list.items
      .select(select)
      .filter(combinedFilter)
      .orderBy("ID", false)
      .top(SHAREPOINT_ID_WINDOW_SIZE);

    if (expand) query = query.expand(expand);

    const page: T[] =
      (await executeSharePointRead<T[]>(() => query.get())) || [];
    results.push(...(includeItem ? page.filter(includeItem) : page));
    if (!sort && typeof maxItems === "number" && results.length >= maxItems)
      break;
    upperId = lowerId;
  }

  const ordered = sort
    ? results.sort(sort)
    : results.sort((a: any, b: any) => Number(b?.ID || 0) - Number(a?.ID || 0));

  return typeof maxItems === "number" ? ordered.slice(0, maxItems) : ordered;
};

export default LIST_CONFIG;
