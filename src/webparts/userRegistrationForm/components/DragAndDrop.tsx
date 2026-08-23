import * as React from "react";
import { useDropzone } from "react-dropzone";

import { SPFI, spfi, SPFx } from "@pnp/sp";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";

interface IResumeUploadProps {
  context: any;

  // ID of the UserRegistrationDetails item
  userRegistrationItemId: number;
}

interface IResumeDocument {
  id: number;
  fileName: string;
  serverRelativeUrl: string;
}

const ResumeUpload: React.FC<IResumeUploadProps> = ({
  context,
  userRegistrationItemId
}) => {

  const [resumeFile, setResumeFile] =
    React.useState<File | null>(null);

  const [existingResume, setExistingResume] =
    React.useState<IResumeDocument | null>(null);

  const [loading, setLoading] =
    React.useState<boolean>(true);

  const [saving, setSaving] =
    React.useState<boolean>(false);

  const [error, setError] =
    React.useState<string>("");


  // ---------------------------------------
  // Initialize PnP
  // ---------------------------------------

  const sp: SPFI = React.useMemo(() => {

    return spfi().using(
      SPFx(context)
    );

  }, [context]);


  // ---------------------------------------
  // Load existing resume
  // ---------------------------------------

  const loadExistingResume = React.useCallback(
    async () => {

      try {

        setLoading(true);
        setError("");

        // 1. Get Resume lookup ID
        const item = await sp.web
          .getList(
            "/sites/RootSite/Lists/UserRegisterationDetailsList"
          )
          .items
          .getById(userRegistrationItemId)
          .select(
            "Id",
            "ResumeId"
          )();

        console.log("Resume lookup ID:", item.ResumeId);

        // No resume uploaded yet
        if (!item.ResumeId) {
          setExistingResume(null);
          return;
        }

        // 2. Get document library item
        const resumeItem = await sp.web
          .getList(
            "/sites/RootSite/Shared Documents"
          )
          .items
          .getById(item.ResumeId)
          .select(
            "Id",
            "FileLeafRef",
            "FileRef"
          )();

        console.log("Resume document:", resumeItem);

        // 3. Store document information
        setExistingResume({
          id: resumeItem.Id,
          fileName: resumeItem.FileLeafRef,
          serverRelativeUrl: resumeItem.FileRef
        });

      } catch (error) {

        console.error(
          "Error loading existing resume:",
          error
        );

        setError(
          "Unable to load existing resume."
        );

      } finally {

        setLoading(false);

      }

    },
    [
      sp,
      userRegistrationItemId
    ]
  );


  // ---------------------------------------
  // Load on component mount
  // ---------------------------------------

  React.useEffect(() => {

    loadExistingResume();

  }, [loadExistingResume]);


  // ---------------------------------------
  // Drag & Drop
  // ---------------------------------------

  const onDrop = React.useCallback(
    (acceptedFiles: File[]) => {

      if (!acceptedFiles.length) {
        return;
      }

      const file = acceptedFiles[0];

      setResumeFile(file);

    },
    []
  );


  const {
    getRootProps,
    getInputProps,
    isDragActive
  } = useDropzone({

    onDrop,

    multiple: false,

    accept: {
      "application/pdf": [".pdf"],

      "application/msword": [".doc"],

      "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        [".docx"]
    },

    maxSize: 10 * 1024 * 1024

  });


  // ---------------------------------------
  // Upload new resume
  // ---------------------------------------

  const uploadResume = async (
    file: File
  ): Promise<number> => {

    const folderPath =
      "/sites/RootSite/Shared Documents/Resume";

    const fileName =
      `${Date.now()}_${file.name}`;

    const folder =
      sp.web.getFolderByServerRelativePath(
        folderPath
      );

    const uploadResult: any =
      await folder.files.addUsingPath(
        fileName,
        file,
        {
          Overwrite: true
        }
      );

    const uploadedFile = uploadResult;

    const uploadedItem: any = await sp.web
      .getFileByServerRelativePath(uploadedFile.ServerRelativeUrl)
      .getItem();

    return uploadedItem.Id;
  };


  // ---------------------------------------
  // Save
  // ---------------------------------------

  const handleSave = async () => {

    if (!resumeFile) {

      alert(
        "Please select a resume first."
      );

      return;
    }

    try {

      setSaving(true);
      setError("");


      // -----------------------------------
      // 1. Upload new resume
      // -----------------------------------

      const documentItemId =
        await uploadResume(resumeFile);


      console.log(
        "Uploaded document ID:",
        documentItemId
      );


      // -----------------------------------
      // 2. Update lookup
      // -----------------------------------

      await sp.web
        .getList(
          "/sites/RootSite/Lists/UserRegisterationDetailsList"
        )
        .items
        .getById(userRegistrationItemId)
        .update({

          ResumeId: documentItemId

        });


      // -----------------------------------
      // 3. Clear selected file
      // -----------------------------------

      setResumeFile(null);


      // -----------------------------------
      // 4. Reload existing document
      // -----------------------------------

      await loadExistingResume();


      alert(
        "Resume saved successfully!"
      );

    } catch (err) {

      console.error(
        "Resume save error:",
        err
      );

      setError(
        "Failed to save resume."
      );

    } finally {

      setSaving(false);

    }
  };


  // ---------------------------------------
  // Remove selected NEW file
  // ---------------------------------------

  const removeSelectedFile = () => {

    setResumeFile(null);

  };


  // ---------------------------------------
  // UI
  // ---------------------------------------

  if (loading) {

    return (
      <div>
        Loading resume...
      </div>
    );

  }


  return (

    <div>

      <h3>Resume</h3>


      {/* -------------------------------- */}
      {/* Existing Resume */}
      {/* -------------------------------- */}

      {existingResume && !resumeFile && (

        <div
          style={{
            padding: "12px",
            border: "1px solid #ddd",
            marginBottom: "15px"
          }}
        >

          <div>
            <strong>
              Current Resume
            </strong>
          </div>

          <a
            href={existingResume.serverRelativeUrl}
            target="_blank"
            rel="noreferrer"
          >
            {existingResume.fileName}
          </a>

        </div>

      )}


      {/* -------------------------------- */}
      {/* New Selected File */}
      {/* -------------------------------- */}

      {resumeFile && (

        <div
          style={{
            padding: "12px",
            border: "1px solid #ddd",
            marginBottom: "15px"
          }}
        >

          <strong>
            New Resume:
          </strong>

          <span>
            {" "}
            {resumeFile.name}
          </span>

          <button
            type="button"
            onClick={removeSelectedFile}
            style={{
              marginLeft: "10px"
            }}
          >
            Remove
          </button>

        </div>

      )}


      {/* -------------------------------- */}
      {/* Drop Zone */}
      {/* -------------------------------- */}

      <div
        {...getRootProps()}
        style={{
          border: "2px dashed #0078d4",
          borderRadius: "6px",
          padding: "30px",
          textAlign: "center",
          cursor: "pointer",
          background: isDragActive
            ? "#f3f9fd"
            : "#fff"
        }}
      >

        <input {...getInputProps()} />

        {isDragActive ? (

          <p>
            Drop your resume here...
          </p>

        ) : (

          <p>
            Drag & drop your resume here
            <br />
            or click to select
            <br />
            <small>
              PDF, DOC or DOCX — maximum 10 MB
            </small>
          </p>

        )}

      </div>


      {/* -------------------------------- */}
      {/* Error */}
      {/* -------------------------------- */}

      {error && (

        <div
          style={{
            color: "red",
            marginTop: "10px"
          }}
        >
          {error}
        </div>

      )}


      {/* -------------------------------- */}
      {/* Save */}
      {/* -------------------------------- */}

      <button
        type="button"
        onClick={handleSave}
        disabled={
          saving ||
          !resumeFile
        }
        style={{
          marginTop: "15px"
        }}
      >

        {saving
          ? "Saving..."
          : existingResume
            ? "Replace Resume"
            : "Save Resume"
        }

      </button>

    </div>

  );
};

export default ResumeUpload;