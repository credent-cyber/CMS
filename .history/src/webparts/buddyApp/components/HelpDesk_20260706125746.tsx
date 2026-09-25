/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable no-void */
/* eslint-disable max-lines */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable @typescript-eslint/no-non-null-assertion */
import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Web } from "@pnp/sp/presets/all";
import {
  Button,
  Drawer,
  IconButton,
  InputAdornment,
  TextField,
} from "@mui/material";
import { useSnackbar } from "./Snackbar";
import { BuddyLoader } from "./Buddyloader";
import LIST_CONFIG from "../../../config/spListConfig";
import styles from "./HelpDesk.module.scss";
import type { IBuddyAppProps } from "./IBuddyAppProps";
import AdminPanelSettingsOutlinedIcon from "@mui/icons-material/AdminPanelSettingsOutlined";
import ArrowForwardIcon from "@mui/icons-material/ArrowForward";
import CheckCircleOutlineIcon from "@mui/icons-material/CheckCircleOutline";
import CloseIcon from "@mui/icons-material/Close";
import HeadsetMicOutlinedIcon from "@mui/icons-material/HeadsetMicOutlined";
import LocationOnOutlinedIcon from "@mui/icons-material/LocationOnOutlined";
import MailOutlineIcon from "@mui/icons-material/MailOutline";
import RefreshIcon from "@mui/icons-material/Refresh";
import SearchIcon from "@mui/icons-material/Search";
import { sp } from "@pnp/sp";

// import SupportAgentOutlinedIcon from "@mui/icons-material/SupportAgentOutlined";

interface HelpDeskUser {
  id: string;
  name: string;
  email: string;
  location: string;
  department: string;
  globalId: string;
}

interface LocationGroup {
  location: string;
  ldSiteAdmin: {
    name: string;
    email: string;
    globalId?: string;
  };
  helpDeskUsers: HelpDeskUser[];
}

interface HelpDeskRequest {
  helpdeskEmail: string;
  title: string;
  description: string;
  userEmail: string;
  userName: string;
  userLocation: string;
  submittedDate: string;
}

interface ITHelpDeskContact {
  id: string;
  name: string;
  email: string;
  globalId?: string;
}

interface SelectedContact {
  name: string;
  email: string;
  role: "L&D Site Admin" | "Help Desk Assistant" | "IT Help Desk";
  location?: string;
  department?: string;
  globalId?: string;
}

const escSP = (value?: string): string =>
  String(value || "")
    .replace(/'/g, "''")
    .trim();

const normalizeText = (value?: string): string =>
  String(value || "")
    .trim()
    .toLowerCase();

const getInitials = (name?: string): string => {
  const parts = String(name || "?")
    .trim()
    .split(/\s+/)
    .filter(Boolean);

  if (parts.length >= 2) {
    return `${parts[0][0]}${parts[parts.length - 1][0]}`.toUpperCase();
  }

  return (parts[0]?.[0] || "?").toUpperCase();
};

const HelpDesk: React.FC<IBuddyAppProps> = (props: IBuddyAppProps) => {
  const { showSnackbar } = useSnackbar();
  const isMountedRef = useRef(true);

  const [loading, setLoading] = useState(true);
  const [userReady, setUserReady] = useState(false);
  const [locationGroups, setLocationGroups] = useState<LocationGroup[]>([]);
  const [itHelpDeskContacts, setItHelpDeskContacts] = useState<ITHelpDeskContact[]>([]);
  const [superAdminUsers, setSuperAdminUsers] = useState<HelpDeskUser[]>([]);
  const [selectedUser, setSelectedUser] = useState<HelpDeskUser | null>(null);
  const [selectedLdAdmin, setSelectedLdAdmin] = useState<{
    name: string;
    email: string;
    location?: string;
    globalId?: string;
  } | null>(null);
  const [selectedItContact, setSelectedItContact] = useState<ITHelpDeskContact | null>(null);
  const [requestDrawerOpen, setRequestDrawerOpen] = useState(false);
  const [requestTitle, setRequestTitle] = useState("");
  const [requestDescription, setRequestDescription] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [currentUserRole, setCurrentUserRole] = useState("");
  const [currentUserLocation, setCurrentUserLocation] = useState("");
  const [searchText, setSearchText] = useState("");

  const currentUserName = props.userloginDetails?.name || props.userDisplayName || "";
  const currentUserEmail = props.userloginDetails?.email || "";

  const isElevatedUser =
    currentUserRole === "Super Admin";

  const selectedContact: SelectedContact | null = useMemo(() => {
    if (selectedUser) {
      return {
        name: selectedUser.name,
        email: selectedUser.email,
        role: "Help Desk Assistant",
        location: selectedUser.location,
        department: selectedUser.department,
        globalId: selectedUser.globalId,
      };
    }

    if (selectedLdAdmin) {
      return {
        name: selectedLdAdmin.name,
        email: selectedLdAdmin.email,
        role: "L&D Site Admin",
        location: selectedLdAdmin.location,
        globalId: selectedLdAdmin.globalId,
      };
    }

    if (selectedItContact) {
      return {
        name: selectedItContact.name,
        email: selectedItContact.email,
        role: "IT Help Desk",
      };
    }

    return null;
  }, [selectedItContact, selectedLdAdmin, selectedUser]);

//   const totalHelpDeskUsers = useMemo(
//     () => locationGroups.reduce((sum, group) => sum + group.helpDeskUsers.length, 0),
//     [locationGroups],
//   );

  const filteredGroups = useMemo(() => {
    const term = normalizeText(searchText);
    if (!term) return locationGroups;

    return locationGroups
      .map((group) => {
        const groupMatches =
          normalizeText(group.location).includes(term) ||
          normalizeText(group.ldSiteAdmin.name).includes(term) ||
          normalizeText(group.ldSiteAdmin.email).includes(term);

        const users = group.helpDeskUsers.filter(
          (user) =>
            normalizeText(user.name).includes(term) ||
            normalizeText(user.email).includes(term) ||
            normalizeText(user.department).includes(term) ||
            normalizeText(user.globalId).includes(term) ||
            normalizeText(user.location).includes(term),
        );

        return {
          ...group,
          helpDeskUsers: groupMatches ? group.helpDeskUsers : users,
          groupMatches,
        };
      })
      .filter((group) => group.groupMatches || group.helpDeskUsers.length > 0)
      .map(({ groupMatches, ...group }) => group);
  }, [locationGroups, searchText]);

  useEffect(() => {
    return () => {
      isMountedRef.current = false;
    };
  }, []);

  useEffect(() => {
    const loadCurrentUser = async () => {
      setLoading(true);
      setUserReady(false);

      try {
        if (!currentUserEmail) {
          setCurrentUserRole("");
          setCurrentUserLocation("");
          return;
        }

        const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
        const employees = await empMasterWeb.lists
          .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
          .items.select("ID", "Title", "EmployeeName", "EmailID", "Role", "Location", "GlobalID")
          .filter(`EmailID eq '${escSP(currentUserEmail)}'`)
          .top(1)
          .get();

        if (employees && employees.length > 0) {
          const emp = employees[0];
          setCurrentUserRole(emp.Role || "");
          setCurrentUserLocation(emp.Location || "");
        } else {
          setCurrentUserRole("");
          setCurrentUserLocation("");
        }
      } catch (err) {
        console.error("[HelpDesk] Failed to load current user:", err);
        showSnackbar("Failed to load user information", "error", 3000);
      } finally {
        if (isMountedRef.current) {
          setUserReady(true);
        }
      }
    };

    void loadCurrentUser();
  }, [currentUserEmail, showSnackbar]);

  const loadHelpDeskData = useCallback(async () => {
    try {
      setLoading(true);
      const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);

      const helpDeskUsers = await empMasterWeb.lists
        .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
        .items.select(
          "ID",
          "Title",
          "EmployeeName",
          "EmailID",
          "Location",
          "Department",
          "GlobalID",
          "IsHelpDeskExecutive",
          "Role",
          "EmployeeStatus",
        )
        .filter("IsHelpDeskExecutive eq 1 and Role eq 'Super Admin' and EmployeeStatus eq 'Active'")
        .orderBy("Location", true)
        .top(5000)
        .get();

      const ldAdmins = await empMasterWeb.lists
        .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
        .items.select("ID", "Title", "EmployeeName","EmailID", "Location", "GlobalID", "Role", "IsHelpDeskExecutive", "EmployeeStatus")
        .filter("Role eq 'L&D Site Admin' and IsHelpDeskExecutive eq 1 and EmployeeStatus eq 'Active'")
        .top(5000)
        .get();

      const itHelpDesk = await sp.web.lists
        .getByTitle(LIST_CONFIG.LISTS.IT_Helpdesk)
        .items.select("ID", "Title", "EmployeeName", "EmailID")
        .orderBy("Title", true)
        .top(5000)
        .get();

      const locationsMap = new Map<string, LocationGroup>();

      // Create set of L&D Admin emails to filter out duplicates
      const ldAdminEmails = new Set(ldAdmins.map((admin: any) => String(admin.EmailID || "").toLowerCase()));

      // First, initialize all locations with their L&D Site Admins
      ldAdmins.forEach((admin: any) => {
        const location = admin.Location || "Unknown";
        if (!locationsMap.has(location)) {
          locationsMap.set(location, {
            location,
            ldSiteAdmin: {
              name: admin.EmployeeName || admin.Title || "L&D Site Admin",
              email: admin.EmailID || "",
              globalId: admin.GlobalID || "",
            },
            helpDeskUsers: [],
          });
        }
      });

      // Then, add Help Desk Users to their respective locations
      helpDeskUsers.forEach((user: any) => {
        // Skip if user is also an L&D Site Admin (to avoid duplicate entries)
        if (ldAdminEmails.has(String(user.EmailID || "").toLowerCase())) {
          return;
        }

        const location = user.Location || "Unknown";
        const ldAdmin = ldAdmins.find((admin: any) => admin.Location === location);

        if (!locationsMap.has(location)) {
          locationsMap.set(location, {
            location,
            ldSiteAdmin: ldAdmin
              ? { name: ldAdmin.EmployeeName || ldAdmin.Title || "L&D Site Admin", email: ldAdmin.EmailID || "", globalId: ldAdmin.GlobalID || "" }
              : { name: "Not assigned", email: "", globalId: "" },
            helpDeskUsers: [],
          });
        }

        locationsMap.get(location)!.helpDeskUsers.push({
          id: String(user.ID),
          name: user.EmployeeName  || "Unnamed user",
          email: user.EmailID || "",
          location,
          department: user.Department || "Not specified",
          globalId: user.GlobalID || "",
        });
      });

      let groups = Array.from(locationsMap.values()).sort((a, b) =>
        a.location.localeCompare(b.location),
      );

      if (!isElevatedUser) {
        // Filter to user's location only, but show all if location is empty
        if (currentUserLocation) {
          groups = groups.filter((group) => group.location === currentUserLocation);
        }
      }

      // Extract all Super Admin users for global display (not location-specific)
      const allSuperAdmins: HelpDeskUser[] = [];
      Array.from(locationsMap.values()).forEach((group) => {
        allSuperAdmins.push(...group.helpDeskUsers);
      });

      const itContacts: ITHelpDeskContact[] = itHelpDesk.map((contact: any) => ({
        id: String(contact.ID),
        name: contact.EmployeeName || "Unnamed contact",
        email: contact.EmailID || "",
      }));

      if (isMountedRef.current) {
        setLocationGroups(groups);
        setItHelpDeskContacts(itContacts);
        setSuperAdminUsers(allSuperAdmins);
      }
    } catch (err) {
      console.error("[HelpDesk] Failed to load Help Desk data:", err);
      showSnackbar("Failed to load Help Desk data", "error", 3000);
      if (isMountedRef.current) {
        setLocationGroups([]);
        setItHelpDeskContacts([]);
        setSuperAdminUsers([]);
      }
    } finally {
      if (isMountedRef.current) {
        setLoading(false);
      }
    }
  }, [currentUserLocation, isElevatedUser, showSnackbar]);

  useEffect(() => {
    if (!userReady) return;
    void loadHelpDeskData();
  }, [loadHelpDeskData, userReady]);

  const resetRequestForm = () => {
    setRequestTitle("");
    setRequestDescription("");
  };

  const openRequestDrawer = (
    contact: HelpDeskUser | { name: string; email: string; location?: string } | ITHelpDeskContact,
    type: "user" | "admin" | "it",
  ) => {
    resetRequestForm();

    if (type === "user") {
      setSelectedUser(contact as HelpDeskUser);
      setSelectedLdAdmin(null);
      setSelectedItContact(null);
    } else if (type === "admin") {
      setSelectedLdAdmin(contact);
      setSelectedUser(null);
      setSelectedItContact(null);
    } else if (type === "it") {
      setSelectedItContact(contact as ITHelpDeskContact);
      setSelectedUser(null);
      setSelectedLdAdmin(null);
    }

    setRequestDrawerOpen(true);
  };

  const handleCloseDrawer = () => {
    if (submitting) return;
    setRequestDrawerOpen(false);
    setSelectedUser(null);
    setSelectedLdAdmin(null);
    setSelectedItContact(null);
    resetRequestForm();
  };

  const sendEmailNotification = async (request: HelpDeskRequest) => {
    try {
      // Email notification is expected to be handled by a Power Automate flow
      // triggered when a Buddy HelpDesk Request item is created.
      console.log(
        "[HelpDesk] Email notification should be sent to:",
        request.helpdeskEmail,
        "and helpdeskit@gmail.com",
      );
      console.log("[HelpDesk] Request:", request);
    } catch (err) {
      console.error("[HelpDesk] Email notification failed:", err);
      throw err;
    }
  };

  const handleSubmitRequest = async () => {
    if (!selectedContact?.email) {
      showSnackbar("Please select a Help Desk contact", "error", 2500);
      return;
    }

    if (!requestTitle.trim() || !requestDescription.trim()) {
      showSnackbar("Please fill in all required fields", "error", 2500);
      return;
    }

    try {
      setSubmitting(true);

      const request: HelpDeskRequest = {
        helpdeskEmail: selectedContact.email,
        title: requestTitle.trim(),
        description: requestDescription.trim(),
        userEmail: currentUserEmail,
        userName: currentUserName,
        userLocation: currentUserLocation,
        submittedDate: new Date().toISOString(),
      };

      // Use only the helpdesk contact information
      const helpdeskUserName = selectedContact?.name || "";
      const helpdeskEmail = selectedContact?.email || "";

      await sp.web.lists
        .getByTitle(LIST_CONFIG.LISTS.BUDDY_HELPDESK_REQUEST)
        
        .items.add({
          Title: request.title,
          HelpdeskEmail: helpdeskEmail,
          HelpdeskUserName: helpdeskUserName,
          Description: request.description,
          UserEmail: request.userEmail,
          UserName: request.userName,
          UserLocation: request.userLocation,
          SubmittedDate: request.submittedDate,
          Status: "Pending",
          RunWF:"Yes"
        });

      try {
        await sendEmailNotification(request);
      } catch (emailErr) {
        console.warn("[HelpDesk] Email notification failed:", emailErr);
      }

      showSnackbar("Help Desk request submitted successfully", "success", 3000);
      setRequestDrawerOpen(false);
      setSelectedUser(null);
      setSelectedLdAdmin(null);
      resetRequestForm();
    } catch (err) {
      console.error("[HelpDesk] Failed to submit request:", err);
      showSnackbar("Failed to submit Help Desk request", "error", 3000);
    } finally {
      if (isMountedRef.current) setSubmitting(false);
    }
  };

  const renderEmptyState = (title: string, subtitle: string) => (
    <div className={styles.noData}>
      <div className={styles.emptyIcon}>
        <HeadsetMicOutlinedIcon />
      </div>
      <h3>{title}</h3>
      <p>{subtitle}</p>
      <button className={styles.refreshBtn} onClick={() => void loadHelpDeskData()}>
        <RefreshIcon sx={{ fontSize: "1rem" }} />
        Refresh
      </button>
    </div>
  );

  return (
    <div className={styles.container}>
      <BuddyLoader
        visible={loading || submitting}
        variant={submitting ? "submit" : "default"}
        message={submitting ? "Submitting Help Desk request..." : "Loading Help Desk contacts..."}
      />

      <div className={styles.pageHeading}>
        <div className={styles.pageHeadingLeft}>
          <div className={styles.pageHeadingIcon}>
            <HeadsetMicOutlinedIcon />
          </div>
          <div>
            <p className={styles.pageTitle}>Help Desk Support</p>
            {/* <p className={styles.pageSubtitle}>
              Find the right L&D admin or Help Desk assistant and raise your support request.
            </p> */}
          </div>
        </div>
        <div className={styles.pageHeadingRight}>
          <TextField
            value={searchText}
            onChange={(e) => setSearchText(e.target.value)}
            placeholder="Search by location, name, email, department, or global ID"
            size="small"
            className={styles.searchInput}
            InputProps={{
              startAdornment: (
                <InputAdornment position="start">
                  <SearchIcon sx={{ fontSize: "1rem", color: "#4b5563" }} />
                </InputAdornment>
              ),
            }}
          />
        </div>
      </div>

      {itHelpDeskContacts.length > 0 && (
        <div className={styles.itHelpDeskSection}>
          <div className={styles.itHelpDeskContacts}>
            {itHelpDeskContacts.map((contact) => (
              <button
                type="button"
                key={contact.id}
                className={styles.itContactCard}
                onClick={() => openRequestDrawer(contact, "it")}
              >
                <div className={styles.contactInfo}>
                  <div className={styles.contactAvatar}>{getInitials(contact.name)}</div>
                  <div className={styles.contactDetails}>
                    <div className={styles.contactName}>{contact.name}</div>
                    <div className={styles.contactEmail}>{contact.email || "Email not available"}</div>
                  </div>
                </div>
                <span className={styles.cardAction}>
                  <MailOutlineIcon sx={{ fontSize: "1rem" }} />
                  Request
                  <ArrowForwardIcon sx={{ fontSize: "0.95rem" }} />
                </span>
              </button>
            ))}
          </div>
        </div>
      )}

      {superAdminUsers.length > 0 && (
        <div className={styles.itHelpDeskSection}>
          <div className={styles.itHelpDeskContacts}>
            {superAdminUsers.map((user) => (
              <div
                key={user.id}
                className={styles.itContactCard}
                style={{ cursor: "default" }}
              >
                <div className={styles.contactInfo}>
                  <div className={styles.contactAvatar}>{getInitials(user.name)}</div>
                  <div className={styles.contactDetails}>
                    <div className={styles.contactName}>{user.name}</div>
                    <div className={styles.contactEmail}>{user.email || "Email not available"}</div>
                    <div className={styles.contactEmail} style={{ fontSize: "0.85rem", color: "#666", marginTop: "0.25rem" }}>Super Admin</div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {itHelpDeskContacts.length === 0 && superAdminUsers.length === 0 && locationGroups.length === 0 && !loading ? (
        renderEmptyState(
          "No Help Desk contacts available",
          isElevatedUser
            ? "No active Help Desk executives were found in the employee master."
            : "No active Help Desk contacts were found for your location.",
        )
      ) : filteredGroups.length === 0 && !loading && locationGroups.length > 0 ? (
        renderEmptyState("No matching contacts", "Try a different name, location, email, department, or global ID.")
      ) : (
        <div className={styles.groupsContainer}>
          {filteredGroups.map((group) => (
            <section key={group.location} className={styles.locationCard}>
              <div className={styles.cardHeader}>
                <div className={styles.locationBadge}>
                  <LocationOnOutlinedIcon sx={{ fontSize: "1rem" }} />
                </div>
                <div className={styles.cardHeaderText}>
                  <h3>{group.location}</h3>
                  {/* <p>{group.helpDeskUsers.length} active Help Desk assistant(s)</p> */}
                </div>
              </div>

              <div className={styles.cardBody}>
                

                {group.ldSiteAdmin.email && (
                  <>
                    <div className={styles.helpDeskTitle} style={{ marginTop: "0.75rem"}}>
                      <AdminPanelSettingsOutlinedIcon sx={{ fontSize: "1rem" }} />
                      L&D Site Admin
                    </div>

                    <div className={styles.helpDeskUsers}>
                      <button
                        type="button"
                        className={styles.userCard}
                        onClick={() =>
                          openRequestDrawer(
                            { name: group.ldSiteAdmin.name, email: group.ldSiteAdmin.email, location: group.location, globalId: group.ldSiteAdmin.globalId },
                            "admin",
                          )
                        }
                      >
                        <div className={styles.userInfo}>
                          <div className={styles.userAvatar}>{getInitials(group.ldSiteAdmin.name)}</div>
                          <div className={styles.userDetails}>
                            <div className={styles.userName}>{group.ldSiteAdmin.name}</div>
                            <div className={styles.userEmail}>{group.ldSiteAdmin.email || "Email not available"}</div>
                            <div className={styles.userMeta}>
                              <span>L&D Administration</span>
                            </div>
                          </div>
                        </div>
                        <span className={styles.cardAction}>
                          <MailOutlineIcon sx={{ fontSize: "1rem" }} />
                          Request
                          <ArrowForwardIcon sx={{ fontSize: "0.95rem" }} />
                        </span>
                      </button>
                    </div>
                  </>
                )}
                {!group.ldSiteAdmin.email && (
                  <div className={styles.noData} style={{ padding: "1.5rem", textAlign: "center" }}>
                    <p style={{ fontSize: "0.9rem", color: "#666" }}>No Help Desk contacts available for this location</p>
                  </div>
                )}
              </div>
            </section>
          ))}
        </div>
      )}

      <Drawer
        anchor="right"
        open={requestDrawerOpen}
        onClose={handleCloseDrawer}
        sx={{
          zIndex: 99999,
          "& .MuiBackdrop-root": { zIndex: 99998 },
          "& .MuiDrawer-paper": {
            top: "0 !important",
            height: "100vh !important",
            maxHeight: "100vh !important",
            zIndex: 100000,
            display: "flex",
            flexDirection: "column",
            maxWidth: "520px",
            width: "100%",
            background: "#F5F2ED",
            fontFamily: "'Inter', sans-serif",
          },
        }}
      >
        <div className={styles.drawer}>
          <div className={styles.drawerHeader}>
            <div>
              {/* <p className={styles.drawerEyebrow}>New support request</p> */}
              <h2 className={styles.drawerTitle}>Submit Help Desk Request</h2>
              {/* <p className={styles.drawerSubtitle}>
                Share the issue details and the selected contact will be notified.
              </p> */}
            </div>
            <IconButton onClick={handleCloseDrawer} className={styles.closeBtn} size="small" disabled={submitting}>
              <CloseIcon fontSize="small" />
            </IconButton>
          </div>

          <div className={styles.drawerBody}>
            {/* {selectedContact && (
              <div className={styles.contactPreviewCard}>
                <div className={styles.contactPreviewTop}>
                  <div className={styles.contactAvatar}>{getInitials(selectedContact.name)}</div>
                  <div>
                    <span>{selectedContact.role}</span>
                    <strong>{selectedContact.name}</strong>
                    <p>{selectedContact.email}</p>
                  </div>
                </div>
                <div className={styles.contactPreviewMeta}>
                  {selectedContact.location && (
                    <span>
                      <LocationOnOutlinedIcon sx={{ fontSize: "0.9rem" }} />
                      {selectedContact.location}
                    </span>
                  )}
                  {selectedContact.department && (
                    <span>
                      <BadgeOutlinedIcon sx={{ fontSize: "0.9rem" }} />
                      {selectedContact.department}
                    </span>
                  )}
                </div>
              </div>
            )} */}

            <div className={styles.editForm}>
              <div className={styles.requesterGrid}>
                <div className={styles.editField}>
                  <label className={styles.editLabel}>Name</label>
                  <div className={styles.editValue}>{currentUserName || "Not available"}</div>
                </div>
                <div className={styles.editField}>
                  <label className={styles.editLabel}>Location</label>
                  <div className={styles.editValue}>{currentUserLocation || "Not available"}</div>
                </div>
              </div>
              <div className={styles.editField}>
                <label className={styles.editLabel}>
                  Title <span className={styles.required}>*</span>
                </label>
                <TextField
                  fullWidth
                  value={requestTitle}
                  onChange={(e) => setRequestTitle(e.target.value)}
                  variant="outlined"
                  size="small"
                  placeholder="Example: Unable to access onboarding form"
                  className={styles.select}
                  inputProps={{ maxLength: 40 }}
                />
                <div className={styles.fieldHint}>{requestTitle.length}/40s characters</div>
              </div>

              <div className={styles.editField}>
                <label className={styles.editLabel}>
                  Description <span className={styles.required}>*</span>
                </label>
                <TextField
                  fullWidth
                  value={requestDescription}
                  onChange={(e) => setRequestDescription(e.target.value)}
                  variant="outlined"
                  size="small"
                  placeholder="Describe the issue, expected support, and any important dates."
                  multiline
                  minRows={6}
                  className={styles.select}
                  inputProps={{ maxLength: 1200 }}
                />
                <div className={styles.fieldHint}>{requestDescription.length}/1200 characters</div>
              </div>
            </div>
          </div>

          <div className={styles.drawerFooter}>
            <Button onClick={handleCloseDrawer} variant="outlined" className={styles.cancelBtn} disabled={submitting}>
              Cancel
            </Button>
            <Button
              onClick={handleSubmitRequest}
              variant="contained"
              className={styles.saveBtn}
              disabled={submitting || !requestTitle.trim() || !requestDescription.trim()}
              startIcon={<CheckCircleOutlineIcon />}
            >
              {submitting ? "Submitting..." : "Submit Request"}
            </Button>
          </div>
        </div>
      </Drawer>
    </div>
  );
};

export default HelpDesk;
