/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable react/self-closing-comp */

import * as React from "react";
import  { useState, useRef, useCallback } from "react";
import { Upload, AlertCircle } from "lucide-react";
import styles from "./BannerLogoUpload.module.scss";
import LIST_CONFIG from "../../../config/spListConfig";

interface BannerLogoUploadProps {
  context: any;
  onUploadSuccess?: () => void;
  showSnackbar: (message: string, type: "success" | "error" | "info", duration?: number) => void;
}

const BannerLogoUpload: React.FC<BannerLogoUploadProps> = ({
  context,
  onUploadSuccess,
  showSnackbar,
}): React.ReactElement => {
  const [activeTab, setActiveTab] = useState<"banner" | "logo">("banner");
  const [uploadFile, setUploadFile] = useState<File | null>(null);
  const [uploadPreview, setUploadPreview] = useState<string>("");
  const [isUploading, setIsUploading] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const dropRef = useRef<HTMLDivElement>(null);
  const isMountedRef = useRef(true);

  const IMAGE_EXTENSIONS = ["jpg", "jpeg", "png", "gif", "webp"];
  const IMAGE_MAX_SIZE = 10 * 1024 * 1024; 

  const getLibraryName = (): string =>
    activeTab === "banner"
      ? LIST_CONFIG.LIBRARIES.BuddyBannerTemplate
      : LIST_CONFIG.LIBRARIES.BuddyLogo;

  const getMaxWidth = (): string =>
    activeTab === "banner" ? "1920px" : "500px";

  const getMaxHeight = (): string =>
    activeTab === "banner" ? "600px" : "500px";

  const processFile = useCallback(
    (file: File) => {
      const ext = file.name.split(".").pop()?.toLowerCase() || "";
      const isImage = IMAGE_EXTENSIONS.includes(ext);

      if (!isImage) {
        showSnackbar("Only image files (WEBP, JPG, JPEG, PNG, GIF) are allowed.", "error", 4000);
        return;
      }

      if (file.size > IMAGE_MAX_SIZE) {
        const sizeMB = (file.size / (1024 * 1024)).toFixed(2);
        showSnackbar(
          `Image size (${sizeMB} MB) exceeds the 10 MB limit. Please choose a smaller image.`,
          "error",
          4000
        );
        return;
      }

      setUploadFile(file);
      const reader = new FileReader();
      reader.onload = (e) => {
        if (isMountedRef.current) {
          setUploadPreview(e.target?.result as string);
        }
      };
      reader.readAsDataURL(file);
    },
    [showSnackbar]
  );

  const handleFileInput = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const f = e.target.files?.[0];
      if (f) processFile(f);
      e.target.value = "";
    },
    [processFile]
  );

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);
      const f = e.dataTransfer.files[0];
      if (f) processFile(f);
    },
    [processFile]
  );

  const getRequestDigest = useCallback(async (): Promise<string> => {
    const digest = (context.pageContext as any).legacyPageContext?.formDigestValue;
    if (digest) return digest;
    const webUrl = context.pageContext.web.absoluteUrl;
    const res = await fetch(`${webUrl}/_api/contextinfo`, { method: "POST" });
    const data = await res.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  }, [context]);

  const handleUpload = useCallback(async () => {
    if (!uploadFile || !uploadPreview) {
      showSnackbar("Please select an image first", "error", 3000);
      return;
    }

    if (uploadFile.size > IMAGE_MAX_SIZE) {
      const sizeMB = (uploadFile.size / (1024 * 1024)).toFixed(2);
      showSnackbar(
        `Image size (${sizeMB} MB) exceeds the 10 MB limit. Please choose a smaller image.`,
        "error",
        4000
      );
      return;
    }

    setIsUploading(true);
    const webUrl = context.pageContext.web.absoluteUrl;
    const webServerRelativeUrl = context.pageContext.web.serverRelativeUrl;
    const libraryName = getLibraryName();

    try {
      const fileBuffer = await uploadFile.arrayBuffer();
      const randomNum = Math.floor(Math.random() * 900000) + 100000;
      const parts = uploadFile.name.split(".");
      const ext = parts.pop();
      const base = parts.join(".");
      const uniqueName = `${base}-${randomNum}.${ext}`;

      const folderPath = `${webServerRelativeUrl}/${libraryName}`;
      const digest = await getRequestDigest();
      const uploadUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(
        folderPath
      )}')/Files/add(url='${encodeURIComponent(uniqueName)}',overwrite=true)`;

      const uploadResp = await fetch(uploadUrl, {
        method: "POST",
        headers: {
          Accept: "application/json",
          "Content-Type": "application/octet-stream",
          "X-RequestDigest": digest,
        },
        body: fileBuffer,
      });

      if (!uploadResp.ok) {
        throw new Error(`Upload failed: ${uploadResp.status}`);
      }

      if (!isMountedRef.current) return;

      setUploadFile(null);
      setUploadPreview("");

      const libType = activeTab === "banner" ? "Banner" : "Logo";
      showSnackbar(`${libType} image uploaded successfully!`, "success", 3000);

      if (onUploadSuccess) {
        onUploadSuccess();
      }
    } catch (e) {
      console.error("Upload error:", e);
      if (isMountedRef.current) {
        showSnackbar("Failed to upload image. Please try again.", "error", 4000);
      }
    } finally {
      if (isMountedRef.current) {
        setIsUploading(false);
      }
    }
  }, [uploadFile, uploadPreview, context, getRequestDigest, activeTab, onUploadSuccess, showSnackbar]);

  const clearForm = useCallback(() => {
    setUploadFile(null);
    setUploadPreview("");
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  }, []);

  return (
    <div className={styles.Container}>
      <div className={styles.pageHeading}>
        <div className={styles.pageHeadingLeft}>
          <div className={styles.pageHeadingIcon}>
            <Upload size={18} />
          </div>
          <div>
            <p className={styles.pageTitle}>Image Management</p>
            <p className={styles.pageSubtitle}>Upload banner and logo images</p>
          </div>
        </div>
      </div>

      {/* Tabs */}
      <div className={styles.tabs}>
        <button
          className={`${styles.tab} ${activeTab === "banner" ? styles.activeTab : ""}`}
          onClick={() => {
            setActiveTab("banner");
            clearForm();
          }}
        >
          Banner Image
        </button>
        <button
          className={`${styles.tab} ${activeTab === "logo" ? styles.activeTab : ""}`}
          onClick={() => {
            setActiveTab("logo");
            clearForm();
          }}
        >
          Logo Image
        </button>
      </div>

      {/* Info Box */}
      <div className={styles.infoBox}>
        <AlertCircle size={16} />
        <div>
          <p className={styles.infoTitle}>
            {activeTab === "banner" ? "Banner Image Requirements" : "Logo Image Requirements"}
          </p>
          <p className={styles.infoText}>
            {activeTab === "banner"
              ? "Recommended: 1920×600px • Max: 10 MB • Formats: WEBP, JPG, JPEG, PNG, GIF"
              : "Recommended: 500×500px • Max: 10 MB • Formats: WEBP, JPG, JPEG, PNG, GIF"}
          </p>
        </div>
      </div>

      {/* Upload Area */}
      {!uploadPreview ? (
        <div
          ref={dropRef}
          className={`${styles.dropZone} ${isDragging ? styles.dragging : ""}`}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
        >
          <div className={styles.dropContent}>
            <Upload size={32} />
            <p className={styles.dropTitle}>Drag and drop your image here</p>
            <p className={styles.dropSubtitle}>or click to browse your computer</p>
          </div>
          <input
            ref={fileInputRef}
            type="file"
            accept={IMAGE_EXTENSIONS.map((ext) => `.${ext}`).join(",")}
            onChange={handleFileInput}
            style={{ display: "none" }}
          />
        </div>
      ) : (
        <div className={styles.previewContainer}>
          <div className={styles.previewBox}>
            <img
              src={uploadPreview}
              alt="Preview"
              className={styles.previewImage}
              style={{
                maxWidth: getMaxWidth(),
                maxHeight: getMaxHeight(),
              }}
            />
          </div>
          <div className={styles.previewActions}>
            <button
              className={styles.changeBtnOutlined}
              onClick={() => fileInputRef.current?.click()}
              disabled={isUploading}
            >
              Change Image
            </button>
            <button
              className={styles.uploadBtnPrimary}
              onClick={handleUpload}
              disabled={isUploading}
            >
              {isUploading ? "Uploading..." : "Upload Image"}
            </button>
          </div>
          <input
            ref={fileInputRef}
            type="file"
            accept={IMAGE_EXTENSIONS.map((ext) => `.${ext}`).join(",")}
            onChange={handleFileInput}
            style={{ display: "none" }}
          />
        </div>
      )}
    </div>
  );
};

export default BannerLogoUpload;
