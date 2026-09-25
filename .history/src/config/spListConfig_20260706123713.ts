/**
 * SharePoint Lists and Libraries Configuration
 * Centralized place to manage all list and library names
 * Update names here and they'll be used throughout the application
 */

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
  },

  // Libraries3
  LIBRARIES: {
    BuddyBannerTemplate: "BuddyBannerTemplates",
    BuddyLogo: "BuddyLogoImg",
    BuddyGallery: "BuddyGallery",
    HolidaysTemplate: "HolidayTemplate",
    UploadedHolidaysTemplate: "UploadedHolidaysTemplate",
  },
  SITENAME: {
  //  BUDDY_EMP_MASTER_SITE: "https://neulandlaboratories.sharepoint.com/sites/AppsDev"
    // BUDDY_EMP_MASTER_SITE: "https://neulandlaboratories.sharepoint.com"
     BUDDY_EMP_MASTER_SITE: "https://credentinfotec.sharepoint.com/sites/NeulandLabs",
  },
};

export default LIST_CONFIG;
