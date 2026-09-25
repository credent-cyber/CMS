/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable max-lines */
/* eslint-disable no-void */

import * as React from "react";
import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import { sp } from "@pnp/sp/presets/all";
import {
  BookOpen,
  CheckCircle2,
  Clock3,
  Download,
  ExternalLink,
  File as FileIcon,
  FileSpreadsheet,
  FileText,
  FolderOpen,
  Image as ImageIcon,
  Pencil,
  RefreshCw,
  Search,
  ShieldCheck,
  Trash2,
  UploadCloud,
  Video,
  X,
} from "lucide-react";

import LIST_CONFIG, {
  getThresholdSafeListItems,
} from "../../../config/spListConfig";
import { useSnackbar } from "./Snackbar";
import styles from "./UserGuides.module.scss";

const USER_GUIDES_LIBRARY = LIST_CONFIG.LIBRARIES.BuddyUserGuides;
const BUDDY_DOC_CATEGORY_MASTER =
  (LIST_CONFIG.LISTS as any).BUDDY_DOC_CATEGORY_MASTER;
const APPROVED_EXTENSIONS = [
  "pdf",
  "ppt",
  "pptx",
  "doc",
  "docx",
  "xls",
  "xlsx",
  "csv",
  "txt",
  "jpg",
  "jpeg",
  "png",
  "gif",
  "webp",
  "mp4",
  "webm",
  "mov",
  "m4v",
] as const;
const OFFICE_VIEW_EXTENSIONS = [
  "ppt",
  "pptx",
  "doc",
  "docx",
  "xls",
  "xlsx",
  "csv",
] as const;
const IMAGE_EXTENSIONS = ["jpg", "jpeg", "png", "gif", "webp"] as const;
const VIDEO_EXTENSIONS = ["mp4", "webm", "mov", "m4v"] as const;
const MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024;
const CHUNKED_UPLOAD_THRESHOLD_BYTES = 10 * 1024 * 1024;
const SUPPORTED_FILE_TYPES_TEXT =
  "PDF, DOCX, TXT, PPT, Excel, Image, GIF, Video";
const OTHER_CATEGORY_VALUE = "__other__";

const isOtherCategory = (value?: string): boolean => {
  const normalized = String(value || "").trim().toLowerCase();
  return normalized === "other" || normalized === "others";
};

type GuideStatus = "Published" | "Draft" | "Archived";

interface IUserGuidesProps {
  context: any;
  isSuperAdmin: boolean;
}

interface IUserGuide {
  id: number;
  title: string;
  description: string;
  category: string;
  status: GuideStatus;
  fileName: string;
  fileRef: string;
  uniqueId: string;
  fileSize: number;
  modified: string;
  created: string;
}

interface IGuideForm {
  title: string;
  description: string;
  category: string;
  status: GuideStatus;
}

const EMPTY_FORM: IGuideForm = {
  title: "",
  description: "",
  category: "",
  status: "Published",
};

const getExtension = (fileName: string): string => {
  const parts = String(fileName || "").toLowerCase().split(".");
  return parts.length > 1 ? parts.pop() || "file" : "file";
};

const getFileTypeLabel = (extension: string): string => {
  if (extension === "pdf") return "PDF";
  if (["ppt", "pptx"].includes(extension)) return "PowerPoint";
  if (["doc", "docx"].includes(extension)) return "Word";
  if (["xls", "xlsx", "csv"].includes(extension)) return "Spreadsheet";
  if (extension === "txt") return "Text";
  if ((IMAGE_EXTENSIONS as readonly string[]).includes(extension)) return "Image";
  if ((VIDEO_EXTENSIONS as readonly string[]).includes(extension)) return "Video";
  return extension.toUpperCase();
};

const formatFileSize = (value: number): string => {
  if (!value) return "—";
  if (value < 1024) return `${value} B`;
  if (value < 1024 * 1024) return `${(value / 1024).toFixed(1)} KB`;
  return `${(value / (1024 * 1024)).toFixed(1)} MB`;
};

const formatDate = (value: string): string => {
  if (!value) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "—";
  return date.toLocaleDateString(undefined, {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
};

const sanitizeFileName = (fileName: string): string =>
  fileName.replace(/[~"#%&*:<>?/\\{|}]/g, "-").replace(/\s+/g, " ").trim();

const createUniqueFileName = (fileName: string): string => {
  const sanitizedName = sanitizeFileName(fileName);
  const extensionIndex = sanitizedName.lastIndexOf(".");
  const baseName = extensionIndex > 0
    ? sanitizedName.slice(0, extensionIndex)
    : sanitizedName;
  const extension = extensionIndex > 0
    ? sanitizedName.slice(extensionIndex)
    : "";
  const digitCount = 5 + Math.floor(Math.random() * 3);
  const minimum = 10 ** (digitCount - 1);
  const suffix = Math.floor(minimum + Math.random() * minimum * 9);

  return `${baseName}-${suffix}${extension}`;
};

const GuideTypeIcon: React.FC<{ extension: string }> = ({ extension }) => {
  if (["ppt", "pptx"].includes(extension)) return <BookOpen size={24} />;
  if (["xls", "xlsx", "csv"].includes(extension))
    return <FileSpreadsheet size={24} />;
  if (["pdf", "doc", "docx", "txt"].includes(extension))
    return <FileText size={24} />;
  if ((IMAGE_EXTENSIONS as readonly string[]).includes(extension))
    return <ImageIcon size={24} />;
  if ((VIDEO_EXTENSIONS as readonly string[]).includes(extension))
    return <Video size={24} />;
  return <FileIcon size={24} />;
};

export const UserGuides: React.FC<IUserGuidesProps> = ({
  context,
  isSuperAdmin,
}) => {
  const { showSnackbar } = useSnackbar();
  const inputRef = useRef<HTMLInputElement>(null);
  const isMountedRef = useRef(true);

  const [guides, setGuides] = useState<IUserGuide[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [loadError, setLoadError] = useState("");
  const [search, setSearch] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("All");
  const [categoryOptions, setCategoryOptions] = useState<string[]>([]);
  const [categoryLoading, setCategoryLoading] = useState(true);
  const [isDragging, setIsDragging] = useState(false);
  const [editorOpen, setEditorOpen] = useState(false);
  const [editingGuide, setEditingGuide] = useState<IUserGuide | null>(null);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [form, setForm] = useState<IGuideForm>(EMPTY_FORM);
  const [categoryChoice, setCategoryChoice] = useState("");
  const [customCategory, setCustomCategory] = useState("");
  const [deleteTarget, setDeleteTarget] = useState<IUserGuide | null>(null);
  const [downloadingGuideId, setDownloadingGuideId] = useState<number | null>(null);

  const webAbsoluteUrl = context.pageContext.web.absoluteUrl;

  useEffect(() => {
    isMountedRef.current = true;
    return () => {
      isMountedRef.current = false;
    };
  }, []);

  const loadCategoryOptions = useCallback(async (): Promise<void> => {
    setCategoryLoading(true);
    try {
      const rows = await getThresholdSafeListItems({
        web: sp.web,
        listTitle: BUDDY_DOC_CATEGORY_MASTER,
        select: "ID,Title",
      });
      const titles = Array.from(
        new Set(
          (rows || [])
            .map((item: any) => String(item.Title || "").trim())
            .filter((title) => Boolean(title) && !isOtherCategory(title)),
        ),
      ).sort((left, right) => left.localeCompare(right));
      if (isMountedRef.current) setCategoryOptions(titles);
    } catch (error) {
      console.error("Unable to load Buddy Doc Category Master:", error);
      if (isMountedRef.current) setCategoryOptions([]);
    } finally {
      if (isMountedRef.current) setCategoryLoading(false);
    }
  }, []);

  useEffect(() => {
    void loadCategoryOptions();
  }, [loadCategoryOptions]);

  const loadGuides = useCallback(async (): Promise<void> => {
    setLoading(true);
    setLoadError("");
    try {
      const guideRows = await getThresholdSafeListItems({
        web: sp.web,
        listTitle: USER_GUIDES_LIBRARY,
        select:
          "ID,Title,GuideDescription,GuideCategory,GuideStatus,FileLeafRef,FileRef,Modified,Created,FSObjType,File/Length,File/UniqueId,File/ServerRelativeUrl,File/Name",
        expand: "File",
        filter: isSuperAdmin
          ? "FSObjType eq 0"
          : "FSObjType eq 0 and GuideStatus eq 'Published'",
      });

      const nextGuides: IUserGuide[] = (guideRows || [])
        .map((item: any) => ({
          id: Number(item.ID),
          title: item.Title || item.FileLeafRef || "Untitled guide",
          description: item.GuideDescription || "",
          category: item.GuideCategory || "General",
          status: (item.GuideStatus || "Draft") as GuideStatus,
          fileName: item.File?.Name || item.FileLeafRef || "Guide",
          fileRef: item.File?.ServerRelativeUrl || item.FileRef || "",
          uniqueId: String(item.File?.UniqueId || "").replace(/[{}]/g, ""),
          fileSize: Number(item.File?.Length || 0),
          modified: item.Modified || "",
          created: item.Created || "",
        }))
        .sort(
          (left, right) =>
            new Date(right.modified).getTime() -
            new Date(left.modified).getTime(),
        );

      if (isMountedRef.current) setGuides(nextGuides);
    } catch (error) {
      console.error("Unable to load User Guides:", error);
      if (isMountedRef.current) {
        setLoadError(
          "We couldn't load the user guides right now. Please refresh the page or try again shortly.",
        );
      }
    } finally {
      if (isMountedRef.current) setLoading(false);
    }
  }, [isSuperAdmin]);

  useEffect(() => {
    loadGuides();
  }, [loadGuides]);

  const categories = useMemo(() => {
    const values = [
      ...categoryOptions,
      ...guides
        .map((guide) => String(guide.category || "").trim())
        .filter((category) => Boolean(category) && !isOtherCategory(category)),
    ];
    return [
      "All",
      ...Array.from(new Set(values)).sort((left, right) =>
        left.localeCompare(right),
      ),
    ];
  }, [categoryOptions, guides]);

  useEffect(() => {
    if (categoryFilter !== "All" && !categories.includes(categoryFilter)) {
      setCategoryFilter("All");
    }
  }, [categoryFilter, categories]);

  const filteredGuides = useMemo(() => {
    const term = search.trim().toLowerCase();
    return guides.filter((guide) => {
      const matchesCategory =
        categoryFilter === "All" || guide.category === categoryFilter;
      const matchesSearch =
        !term ||
        [
          guide.title,
          guide.description,
          guide.category,
          guide.fileName,
          guide.status,
        ]
          .join(" ")
          .toLowerCase()
          .includes(term);
      return matchesCategory && matchesSearch;
    });
  }, [guides, search, categoryFilter]);

  const publishedCount = useMemo(
    () => guides.filter((guide) => guide.status === "Published").length,
    [guides],
  );
  const draftCount = useMemo(
    () => guides.filter((guide) => guide.status === "Draft").length,
    [guides],
  );

  const resetEditor = useCallback(() => {
    setEditorOpen(false);
    setEditingGuide(null);
    setSelectedFile(null);
    setForm(EMPTY_FORM);
    setCategoryChoice("");
    setCustomCategory("");
    setIsDragging(false);
  }, []);

  const closeEditor = useCallback(() => {
    if (saving) return;
    resetEditor();
  }, [saving, resetEditor]);

  const openUploadEditor = useCallback(() => {
    setEditingGuide(null);
    setSelectedFile(null);
    setForm(EMPTY_FORM);
    setCategoryChoice("");
    setCustomCategory("");
    setEditorOpen(true);
  }, []);

  const openMetadataEditor = useCallback(
    (guide: IUserGuide) => {
      const matchedCategory = categoryOptions.find(
        (category) =>
          category.toLowerCase() === String(guide.category || "").toLowerCase(),
      );
      const usesCustomCategory =
        Boolean(guide.category) && !matchedCategory && !isOtherCategory(guide.category);

      setEditingGuide(guide);
      setSelectedFile(null);
      setCategoryChoice(
        usesCustomCategory
          ? OTHER_CATEGORY_VALUE
          : matchedCategory || (isOtherCategory(guide.category) ? OTHER_CATEGORY_VALUE : ""),
      );
      setCustomCategory(usesCustomCategory ? guide.category : "");
      setForm({
        title: guide.title,
        description: guide.description,
        category: usesCustomCategory
          ? guide.category
          : matchedCategory || "",
        status: guide.status,
      });
      setEditorOpen(true);
    },
    [categoryOptions],
  );

  const validateAndSelectFile = useCallback(
    (file?: File): void => {
      if (!file) return;
      const extension = getExtension(file.name);
      if (!APPROVED_EXTENSIONS.includes(extension as any)) {
        showSnackbar(
          `This file type (.${extension}) is not supported. Please select a document, image or video file.`,
          "warning",
          4500,
        );
        return;
      }
      if (file.size > MAX_FILE_SIZE_BYTES) {
        showSnackbar(
          "File size exceeds the 50 MB limit. Please select a file up to 50 MB.",
          "warning",
          4500,
        );
        return;
      }
      setSelectedFile(file);
      if (!form.title.trim()) {
        setForm((current) => ({
          ...current,
          title: file.name.replace(/\.[^.]+$/, ""),
        }));
      }
    },
    [form.title, showSnackbar],
  );

  const saveGuide = useCallback(async (): Promise<void> => {
    if (!form.title.trim()) {
      showSnackbar("Please enter a title for this guide before continuing.", "warning", 3500);
      return;
    }
    if (!editingGuide && !selectedFile) {
      showSnackbar("Please select a file to upload before continuing.", "warning", 3500);
      return;
    }

    setSaving(true);
    try {
      const resolvedCategory =
        categoryChoice === OTHER_CATEGORY_VALUE
          ? customCategory.trim() || OTHER_CATEGORY_VALUE
          : form.category.trim();

      const metadata = {
        Title: form.title.trim(),
        GuideDescription: form.description.trim(),
        GuideCategory: resolvedCategory,
        GuideStatus: form.status,
      };

      if (editingGuide) {
        await sp.web.lists
          .getByTitle(USER_GUIDES_LIBRARY)
          .items.getById(editingGuide.id)
          .update(metadata);
        showSnackbar("Guide details updated successfully.", "success", 3000);
      } else if (selectedFile) {
        const rootFolder = await sp.web.lists
          .getByTitle(USER_GUIDES_LIBRARY)
          .rootFolder.select("ServerRelativeUrl")
          .get();
        const safeFileName = createUniqueFileName(selectedFile.name);
        const folder = sp.web.getFolderByServerRelativeUrl(
          rootFolder.ServerRelativeUrl,
        );
        const folderFiles: any = folder.files;
        const uploadResult: any =
          selectedFile.size > CHUNKED_UPLOAD_THRESHOLD_BYTES &&
          typeof folderFiles.addChunked === "function"
            ? await folderFiles.addChunked(
                safeFileName,
                selectedFile,
                () => undefined,
                false,
              )
            : await folderFiles.add(safeFileName, selectedFile, false);
        const uploadedFile =
          uploadResult?.file ||
          sp.web.getFileByServerRelativeUrl(
            `${String(rootFolder.ServerRelativeUrl).replace(/\/$/, "")}/${safeFileName}`,
          );
        const uploadedItem = await uploadedFile.getItem();
        await uploadedItem.update(metadata);
        showSnackbar(
          form.status === "Published"
            ? "Guide uploaded and published successfully."
            : "Guide uploaded as draft.",
          "success",
          3500,
        );
      }

      resetEditor();
      await loadGuides();
    } catch (error) {
      console.error("Unable to save User Guide:", error);
      showSnackbar(
        "We couldn't save this guide. Please verify the document and try again.",
        "error",
        4000,
      );
    } finally {
      if (isMountedRef.current) setSaving(false);
    }
  }, [
    form,
    categoryChoice,
    customCategory,
    editingGuide,
    selectedFile,
    showSnackbar,
    resetEditor,
    loadGuides,
  ]);

  const changeStatus = useCallback(
    async (guide: IUserGuide): Promise<void> => {
      const nextStatus: GuideStatus =
        guide.status === "Published" ? "Draft" : "Published";
      try {
        await sp.web.lists
          .getByTitle(USER_GUIDES_LIBRARY)
          .items.getById(guide.id)
          .update({ GuideStatus: nextStatus });
        showSnackbar(
          nextStatus === "Published"
            ? "Guide activated and published successfully."
            : "Guide set to inactive draft successfully.",
          "success",
          3000,
        );
        await loadGuides();
      } catch (error) {
        console.error("Unable to update guide status:", error);
        showSnackbar("We couldn't update the guide status. Please try again.", "error", 4000);
      }
    },
    [showSnackbar, loadGuides],
  );

  const deleteGuide = useCallback(async (): Promise<void> => {
    if (!deleteTarget) return;
    setSaving(true);
    try {
      const deletedTitle = deleteTarget.title;
      await sp.web.getFileByServerRelativeUrl(deleteTarget.fileRef).recycle();
      showSnackbar(
        `"${deletedTitle}" was removed from BuddyPedia successfully.`,
        "success",
        4000,
      );
      setDeleteTarget(null);
      await loadGuides();
    } catch (error) {
      console.error("Unable to delete User Guide:", error);
      showSnackbar("We couldn't remove this guide. Please try again.", "error", 4000);
    } finally {
      if (isMountedRef.current) setSaving(false);
    }
  }, [deleteTarget, showSnackbar, loadGuides]);

  const openGuide = useCallback(
    (guide: IUserGuide): void => {
      if (!guide.fileRef) {
        showSnackbar(
          "This guide is temporarily unavailable. Please refresh the page and try again.",
          "warning",
          4000,
        );
        return;
      }

      const extension = getExtension(guide.fileName);
      const webUrl = String(webAbsoluteUrl || "").replace(/\/$/, "");
      const fileUrl = new URL(guide.fileRef, webUrl).toString();
      const usesOfficeViewer =
        OFFICE_VIEW_EXTENSIONS.includes(extension as any) && guide.uniqueId;
      const isMediaFile =
        (IMAGE_EXTENSIONS as readonly string[]).includes(extension) ||
        (VIDEO_EXTENSIONS as readonly string[]).includes(extension);
      const viewUrl = isMediaFile
        ? fileUrl
        : usesOfficeViewer
          ? `${webUrl}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(
              `{${guide.uniqueId}}`,
            )}&file=${encodeURIComponent(guide.fileName)}&action=embedview`
          : `${fileUrl}${fileUrl.includes("?") ? "&" : "?"}web=1`;
      const openedWindow = window.open(viewUrl, "_blank");

      if (openedWindow) {
        openedWindow.opener = null;
      } else {
        showSnackbar(
          "Your browser blocked the guide viewer. Please allow pop-ups for this site and try again.",
          "warning",
          4500,
        );
      }
    },
    [showSnackbar, webAbsoluteUrl],
  );

  const downloadGuide = useCallback(
    async (guide: IUserGuide): Promise<void> => {
      if (!isSuperAdmin || !guide.fileRef || downloadingGuideId !== null) return;

      setDownloadingGuideId(guide.id);
      try {
        const fileBlob = await sp.web
          .getFileByServerRelativeUrl(guide.fileRef)
          .getBlob();
        const objectUrl = window.URL.createObjectURL(fileBlob);
        const downloadLink = document.createElement("a");
        downloadLink.href = objectUrl;
        downloadLink.download = guide.fileName;
        downloadLink.style.display = "none";
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
        window.setTimeout(() => window.URL.revokeObjectURL(objectUrl), 1000);
      } catch (error) {
        console.error("Unable to download User Guide:", error);
        showSnackbar(
          "We couldn't download this guide. Please refresh the page and try again.",
          "error",
          4500,
        );
      } finally {
        if (isMountedRef.current) setDownloadingGuideId(null);
      }
    },
    [downloadingGuideId, isSuperAdmin, showSnackbar],
  );

  return (
    <section className={styles.userGuides}>
      <div className={styles.pageShell}>
        <header className={styles.heroHeader}>
          <div className={styles.heroIdentity}>
            <span className={styles.heroIcon}>
              <BookOpen size={22} />
            </span>
            <div>
              {/* <p className={styles.eyebrow}>Buddypedia Knowledge Hub</p> */}
              <h1>BuddyPedia</h1>
              <p className={styles.heroSubtitle}>
                Find trusted guides, templates and step-by-step resources.
              </p>
            </div>
          </div>
          <div className={styles.heroActions}>
            <button
              type="button"
              className={styles.refreshButton}
              onClick={loadGuides}
              disabled={loading}
              title="Refresh guides"
              style={{ cursor: loading ? "default" : "pointer" }}
            >
              <RefreshCw
                size={15}
                className={loading ? styles.spinning : ""}
                style={{ pointerEvents: "none" }}
              />
              <span>Refresh</span>
            </button>
            {isSuperAdmin && (
              <button
                type="button"
                className={styles.uploadButton}
                onClick={openUploadEditor}
              >
                <UploadCloud size={16} /> Upload Guide
              </button>
            )}
          </div>
        </header>

        <div className={styles.statGrid}>
          <article className={styles.statCard}>
            <span className={`${styles.statIcon} ${styles.statIconGreen}`}>
              <FolderOpen size={18} />
            </span>
            <div>
              <strong>{guides.length}</strong>
              <span>{isSuperAdmin ? "Total guides" : "Available guides"}</span>
            </div>
          </article>
          <article className={styles.statCard}>
            <span className={`${styles.statIcon} ${styles.statIconDark}`}>
              <CheckCircle2 size={18} />
            </span>
            <div>
              <strong>{publishedCount}</strong>
              <span>Published</span>
            </div>
          </article>
          {isSuperAdmin && (
            <article className={styles.statCard}>
              <span className={`${styles.statIcon} ${styles.statIconStone}`}>
                <Clock3 size={18} />
              </span>
              <div>
                <strong>{draftCount}</strong>
                <span>Drafts</span>
              </div>
            </article>
          )}
          <article className={styles.statCard}>
            <span className={`${styles.statIcon} ${styles.statIconMint}`}>
              <ShieldCheck size={18} />
            </span>
            <div>
              <strong>{Math.max(0, categories.length - 1)}</strong>
              <span>Categories</span>
            </div>
          </article>
        </div>

        <div className={styles.toolbar}>
          <label className={styles.searchBox}>
            <Search size={16} />
            <input
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder="Search guides, category, file name or status..."
              aria-label="Search user guides"
            />
            {search && (
              <button
                type="button"
                onClick={() => setSearch("")}
                aria-label="Clear search"
              >
                <X size={14} />
              </button>
            )}
          </label>
          <label className={styles.categoryFilterSelect}>
            <FolderOpen size={15} />
            <span>Category</span>
            <select
              value={categoryFilter}
              onChange={(event) => setCategoryFilter(event.target.value)}
              aria-label="Filter BuddyPedia by category"
              disabled={categoryLoading}
            >
              {categories.map((category) => (
                <option key={category} value={category}>
                  {category}
                </option>
              ))}
            </select>
          </label>
        </div>

        {loadError && (
          <div className={styles.errorBanner}>
            <FileIcon size={17} />
            <span>{loadError}</span>
            <button type="button" onClick={loadGuides}>
              Try again
            </button>
          </div>
        )}

        <div className={styles.guideViewport}>
          {loading ? (
            <div className={styles.loadingGrid}>
              {[1, 2, 3, 4, 5, 6].map((item) => (
                <div key={item} className={styles.skeletonCard} />
              ))}
            </div>
          ) : filteredGuides.length === 0 ? (
            <div className={styles.emptyState}>
              <span><BookOpen size={28} /></span>
              <h2>{search || categoryFilter !== "All" ? "No matching guides" : "No guides published yet"}</h2>
              <p>
                {search || categoryFilter !== "All"
                  ? "Try another search term or category."
                  : isSuperAdmin
                    ? "Upload the first guide to start building Buddypedia."
                    : "Published guides will appear here when they are available."}
              </p>
              {isSuperAdmin && !search && categoryFilter === "All" && (
                <button type="button" onClick={openUploadEditor}>
                  <UploadCloud size={15} /> Upload first guide
                </button>
              )}
            </div>
          ) : (
            <div className={styles.guideGrid}>
              {filteredGuides.map((guide) => {
                const extension = getExtension(guide.fileName);
                return (
                  <article key={guide.id} className={styles.guideCard}>
                    <div className={styles.guideCardTop}>
                      <span className={`${styles.fileTypeIcon} ${((styles as unknown) as Record<string, string>)[`type_${extension}`] || styles.type_default}`}>
                        <GuideTypeIcon extension={extension} />
                      </span>
                      <div className={styles.guideBadges}>
                        <span className={styles.fileTypeBadge}>{getFileTypeLabel(extension)}</span>
                        {isSuperAdmin && (
                          <span className={`${styles.statusBadge} ${styles[`status${guide.status}`]}`}>
                            {guide.status}
                          </span>
                        )}
                      </div>
                    </div>

                    <div className={styles.guideContent}>
                      <span className={styles.categoryLabel}>{guide.category}</span>
                      <h2 title={guide.title}>{guide.title}</h2>
                      <p title={guide.description}>
                        {guide.description || "Open this guide to view the complete resource."}
                      </p>
                    </div>

                    <div className={styles.guideMeta}>
                      <span>{formatFileSize(guide.fileSize)}</span>
                      <span>Updated {formatDate(guide.modified)}</span>
                    </div>

                    <div className={styles.guideActions}>
                      <button type="button" className={styles.openButton} onClick={() => openGuide(guide)} title="Open in view-only mode">
                        <ExternalLink size={13} /> View
                      </button>
                      {isSuperAdmin && (
                        <button
                          type="button"
                          className={styles.downloadButton}
                          onClick={() => downloadGuide(guide)}
                          title="Download guide"
                          disabled={downloadingGuideId !== null}
                          aria-label={`Download ${guide.title}`}
                          style={{ cursor: downloadingGuideId !== null ? "default" : "pointer" }}
                        >
                          {downloadingGuideId === guide.id
                            ? <RefreshCw
                                size={13}
                                className={styles.spinning}
                                style={{ pointerEvents: "none" }}
                              />
                            : <Download size={13} style={{ pointerEvents: "none" }} />}
                        </button>
                      )}
                      {isSuperAdmin && (
                        <div className={styles.adminActions}>
                          <button type="button" onClick={() => openMetadataEditor(guide)} title="Edit guide details">
                            <Pencil size={14} />
                          </button>
                          <button
                            type="button"
                            onClick={() => changeStatus(guide)}
                            title={guide.status === "Published"
                              ? "Set as inactive draft"
                              : "Activate and publish guide"}
                          >
                            {guide.status === "Published" ? <Clock3 size={14} /> : <CheckCircle2 size={14} />}
                          </button>
                          <button type="button" className={styles.deleteButton} onClick={() => setDeleteTarget(guide)} title="Remove guide">
                            <Trash2 size={14} />
                          </button>
                        </div>
                      )}
                    </div>
                  </article>
                );
              })}
            </div>
          )}
        </div>
      </div>

      {editorOpen && (
        <div className={styles.modalBackdrop} onMouseDown={closeEditor}>
          <div className={styles.editorModal} onMouseDown={(event) => event.stopPropagation()} role="dialog" aria-modal="true" aria-labelledby="guide-editor-title">
            <div className={styles.modalHeader}>
              <div>
                <span className={styles.modalHeaderIcon}>{editingGuide ? <Pencil size={18} /> : <UploadCloud size={18} />}</span>
                <div>
                  <h2 id="guide-editor-title">{editingGuide ? "Edit guide" : "Upload new guide"}</h2>
                  <p>{editingGuide ? "Update publishing details and metadata." : "Add a trusted resource to Buddypedia."}</p>
                </div>
              </div>
              <button type="button" onClick={closeEditor} disabled={saving} aria-label="Close editor"><X size={17} /></button>
            </div>

            <div className={styles.modalBody}>
              {!editingGuide && (
                <div
                  className={`${styles.dropZone} ${isDragging ? styles.dropZoneActive : ""} ${selectedFile ? styles.dropZoneSelected : ""}`}
                  onDragOver={(event) => { event.preventDefault(); setIsDragging(true); }}
                  onDragLeave={() => setIsDragging(false)}
                  onDrop={(event) => {
                    event.preventDefault();
                    setIsDragging(false);
                    validateAndSelectFile(event.dataTransfer.files?.[0]);
                  }}
                  onClick={() => inputRef.current?.click()}
                  role="button"
                  tabIndex={0}
                  onKeyDown={(event) => {
                    if (event.key === "Enter" || event.key === " ") inputRef.current?.click();
                  }}
                >
                  <input
                    ref={inputRef}
                    type="file"
                    accept={APPROVED_EXTENSIONS.map((extension) => `.${extension}`).join(",")}
                    onChange={(event) => validateAndSelectFile(event.target.files?.[0])}
                  />
                  {selectedFile ? (
                    <>
                      <span className={styles.selectedFileIcon}><GuideTypeIcon extension={getExtension(selectedFile.name)} /></span>
                      <strong>{selectedFile.name}</strong>
                      <small>{formatFileSize(selectedFile.size)} of 50 MB max - Click to replace</small>
                      <span className={styles.dropZoneTypes}>
                        <b>Supported:</b> {SUPPORTED_FILE_TYPES_TEXT}
                      </span>
                    </>
                  ) : (
                    <>
                      <span className={styles.dropZoneIcon}><UploadCloud size={25} /></span>
                      <strong>{isDragging ? "Drop the guide here" : "Drag & drop a file"}</strong>
                      <small>or click to browse - Maximum file size: 50 MB</small>
                      <span className={styles.dropZoneTypes}>
                        <b>Supported file types:</b> {SUPPORTED_FILE_TYPES_TEXT}
                      </span>
                    </>
                  )}
                </div>
              )}

              <div className={styles.formGrid}>
                <label className={styles.fullField}>
                  <span>Guide title <em>*</em></span>
                  <input value={form.title} maxLength={120} onChange={(event) => setForm((current) => ({ ...current, title: event.target.value }))} placeholder="e.g. Buddy Meeting Best Practices" />
                </label>
                <div
                  className={styles.categoryStatusRow}
                  data-has-other={categoryChoice === OTHER_CATEGORY_VALUE ? "true" : "false"}
                >
                  <label>
                    <span>Category</span>
                    <select
                      value={categoryChoice}
                      onChange={(event) => {
                        const value = event.target.value;
                        setCategoryChoice(value);
                        if (value === OTHER_CATEGORY_VALUE) {
                          setForm((current) => ({
                            ...current,
                            category: customCategory.trim(),
                          }));
                        } else {
                          setCustomCategory("");
                          setForm((current) => ({
                            ...current,
                            category: value,
                          }));
                        }
                      }}
                      disabled={categoryLoading}
                    >
                      <option value="">
                        {categoryLoading
                          ? "Loading categories..."
                          : categoryOptions.length === 0
                            ? "Select category "
                            : "Select category "}
                      </option>
                      {categoryOptions.map((category) => (
                        <option key={category} value={category}>
                          {category}
                        </option>
                      ))}
                      <option value={OTHER_CATEGORY_VALUE}>Other</option>
                    </select>
                  </label>

                  {categoryChoice === OTHER_CATEGORY_VALUE && (
                    <label className={styles.customCategoryField}>
                      <span>Other category</span>
                      <input
                        type="text"
                        value={customCategory}
                        maxLength={100}
                        onChange={(event) => {
                          const value = event.target.value;
                          setCustomCategory(value);
                          setForm((current) => ({
                            ...current,
                            category: value,
                          }));
                        }}
                        placeholder="Enter category name "
                      />
                    </label>
                  )}

                  <label>
                    <span>Publishing status</span>
                    <select value={form.status} onChange={(event) => setForm((current) => ({ ...current, status: event.target.value as GuideStatus }))}>
                      <option value="Published">Published</option>
                      <option value="Draft">Draft</option>
                      <option value="Archived">Archived</option>
                    </select>
                  </label>
                </div>
                <label className={styles.fullField}>
                  <span>Description</span>
                  <textarea value={form.description} maxLength={500} onChange={(event) => setForm((current) => ({ ...current, description: event.target.value }))} placeholder="Briefly explain what users will learn from this guide…" rows={4} />
                  <small>{form.description.length}/500</small>
                </label>
              </div>
            </div>

            <div className={styles.modalFooter}>
              <button type="button" className={styles.cancelButton} onClick={closeEditor} disabled={saving}>Cancel</button>
              <button type="button" className={styles.saveButton} onClick={saveGuide} disabled={saving}>
                {saving ? <span className={styles.buttonSpinner} /> : editingGuide ? <Pencil size={15} /> : <UploadCloud size={15} />}
                {saving ? "Saving…" : editingGuide ? "Save changes" : "Upload guide"}
              </button>
            </div>
          </div>
        </div>
      )}

      {deleteTarget && (
        <div className={styles.modalBackdrop} onMouseDown={() => !saving && setDeleteTarget(null)}>
          <div className={styles.confirmModal} onMouseDown={(event) => event.stopPropagation()} role="alertdialog" aria-modal="true" aria-labelledby="delete-guide-title">
            <span className={styles.confirmIcon}><Trash2 size={22} /></span>
            <h2 id="delete-guide-title">Remove this item?</h2>
            <p>
              &quot;{deleteTarget.title}&quot; will be removed from BuddyPedia and will no longer be available to users.
            </p>
            <div>
              <button type="button" onClick={() => setDeleteTarget(null)} disabled={saving}>Cancel</button>
              <button type="button" className={styles.confirmDeleteButton} onClick={deleteGuide} disabled={saving}>
                {saving ? "Removing..." : "Remove"}
              </button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
};

export default UserGuides;
