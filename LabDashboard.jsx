import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import "./Lashboard.css";

const LabDashboard = () => {
  const [role, setRole] = useState(
    localStorage.getItem("role") || ""
  );

  const [showSidebar, setShowSidebar] = useState(false);
  const [sidInput, setSidInput] = useState("");
  const [selectedYear, setSelectedYear] = useState("2026");
  const [selectedTest, setSelectedTest] = useState("");
  const [searchTerm, setSearchTerm] = useState("");
  const [currentProgress, setCurrentProgress] = useState(-1);
  const [excelDatabase, setExcelDatabase] = useState({});
  const [patientList, setPatientList] = useState([]);
  const [selectedPatient, setSelectedPatient] = useState("");

  const [patientInfo, setPatientInfo] = useState({
    name: "",
    age: "",
    referral: "",
    contact: "",
    photo: "",
  });

  const stepNames = [
    "Billing",
    "Sample Collection",
    "Sample Processing",
    "Sample Outsourced",
    "Result Received",
    "Result Entry",
    "Authorisation",
    "2nd Level Authorisation",
    "Report Print",
    "Report Sent",
  ];

  const isOutsourced =
    sidInput &&
    excelDatabase[sidInput]?.tests[selectedTest]?.isOutsourced;

  const handleLogin = (selectedRole) => {
    setRole(selectedRole);
    localStorage.setItem("role", selectedRole);
  };

  const handleLogout = () => {
    localStorage.removeItem("role");
    setRole("");
  };
  
  const parseJsonCell = (value) => {
    if (!value) return {};
    if (typeof value === "object") return value;

    if (typeof value === "string") {
      const trimmed = value.trim();

      try {
        return JSON.parse(trimmed);
      } catch (error) {
        const parsed = {};
        const normalized = trimmed.replace(/\r?\n/g, ";");
        normalized.split(/;+/).forEach((part) => {
          const [key, ...rest] = part.split(":");
          if (!key || rest.length === 0) return;
          const valueText = rest.join(":").trim();
          const label = key.trim();
          if (label && valueText) parsed[label] = valueText;
        });

        if (Object.keys(parsed).length) {
          return parsed;
        }

        console.warn("StepDetails parse failed", { value: trimmed, error });
        return {};
      }
    }

    return {};
  };

  // FETCH EXCEL
  useEffect(() => {
    const fetchExcelData = async () => {
      try {
        const response = await fetch("/data.xlsx");
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const worksheet = workbook.Sheets[selectedYear];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        const formattedData = {};
        const patients = [];

        jsonData.forEach((row) => {
          let parsedTests = {};

          try {
            parsedTests =
              typeof row.Tests === "string"
                ? JSON.parse(row.Tests)
                : row.Tests;
          } catch {
            parsedTests = {
              "General Test": 0,
            };
          }

          const sid = String(row.SID).padStart(6, "0");

          // PARSE STEP DETAILS FOR MULTIPLE TESTS
          const stepDetailsForAllTests = {};

          stepNames.forEach((stepName) => {
            // Get the column value for this step
            let columnValue = row[stepName];

            if (!columnValue) return;

            // Parse the JSON that contains test-specific data
            const parsedStepData = parseJsonCell(columnValue);

            if (!parsedStepData || Object.keys(parsedStepData).length === 0)
              return;

            // Now parsedStepData should be: { "Blood Test": {...}, "Urine Test": {...} }
            stepDetailsForAllTests[stepName] = parsedStepData;
          });

          formattedData[sid] = {
            name: row.Name,
            age: row.Age,
            referral: row.Referral,
            contact: row.Contact,
            photo: row.Photo,
            tests:
              parsedTests && typeof parsedTests === "object"
                ? parsedTests
                : {},
            // Store all step details with test-specific breakdown
            stepDetails: stepDetailsForAllTests,
          };

          patients.push({
            sid: sid,
            name: row.Name,
            contact: row.Contact,
            referral: row.Referral,
          });

          console.log("FORMATTED DATA:", formattedData[sid]);
        });

        setExcelDatabase(formattedData);
        setPatientList(patients);
      } catch (error) {
        console.error("Error fetching Excel data:", error);
      }
    };

    fetchExcelData();
  }, [selectedYear]);

  // SEARCH
  const filteredPatients = patientList.filter((patient) =>
    patient.sid.includes(searchTerm) ||
    patient.name?.toLowerCase().includes(searchTerm.toLowerCase()) ||
    String(patient.contact)?.includes(searchTerm) ||
    patient.referral?.toLowerCase().includes(searchTerm.toLowerCase())
  );

  // SID INPUT
  const handleSidChange = (e) => {
    const value = e.target.value.replace(/\D/g, "").slice(0, 6);
    setSidInput(value);

    if (value.length === 6) {
      const data = excelDatabase[value];

      if (data) {
        setPatientInfo(data);

        const firstTest = data.tests
          ? Object.keys(data.tests)[0]
          : "";

        setSelectedTest(firstTest);
        setCurrentProgress(data.tests[firstTest]);
      }
    }
  };

  const handleTestChange = (e) => {
    const testName = e.target.value;
    setSelectedTest(testName);

    const data = excelDatabase[sidInput];

    if (data && data.tests[testName] !== undefined) {
      setCurrentProgress(data.tests[testName]);
    }
  };

  const stepTitles = [
    "Billing",
    "Sample Collection",
    "Sample Processing",
    "Sample Outsourced",
    "Result Received",
    "Result Entry",
    "Authorisation",
    "2nd Level Authorisation",
    "Report Print",
    "Report Sent",
  ];

  const getStepStatus = (index) => {
    if (currentProgress === -1) return "pending";

    if (currentProgress === 9) {
      return "completed";
    }

    if (index < currentProgress) return "completed";

    if (index === currentProgress) return "in-progress";

    return "pending";
  };

  const trackingSteps = stepTitles.map((title, idx) => ({
    title,
    status: getStepStatus(idx),
  }));

  const getStatusConfig = (status) => {
    switch (status) {
      case "completed":
        return {
          color: "#57c595",
          bg: "rgba(87,197,149,0.15)",
          label: "Completed",
        };

      case "in-progress":
        return {
          color: "#00d9ff",
          bg: "rgba(0,217,255,0.15)",
          label: "Processing",
        };

      default:
        return {
          color: "#6b7280",
          bg: "rgba(255,255,255,0.05)",
          label: "Pending",
        };
    }
  };

  // IMPROVED HOVER CONTENT FUNCTION
  const getHoverContent = (stepIndex) => {

  const stepName = stepNames[stepIndex];

  const patientData =
    excelDatabase[sidInput];

  if (!patientData || !selectedTest) {

    return {
      left: [],
      right: [],
    };
  }

  const stepData =
    patientData.stepDetails?.[
      stepName
    ]?.[selectedTest];

  if (!stepData) {

    return {
      left: [],
      right: [],
    };
  }

  // BILLING

  if (stepIndex === 0) {

    return {

  left: [
    {
      label: "Company/Lab",
      value: stepData.company,
    },
  ],

  right: [
    {
      label: "Date",
      value: stepData.date,
    },

    {
      label: "Time",
      value: stepData.time,
    },
  ],
};
  }

  // SAMPLE COLLECTION

  if (stepIndex === 1) {

    return {

  left: [
    {
      label: "Outsourced To",
      value: stepData.company,
    },
  ],

  right: [
    {
      label: "Date",
      value: stepData.date,
    },

    {
      label: "Time",
      value: stepData.time,
    },
  ],
};
  }

  // SAMPLE PROCESSING

  if (stepIndex === 2) {

    return {

  left: [
    {
      label: "Company/Lab",
      value: stepData.company,
    },
  ],

  right: [
    {
      label: "Date",
      value: stepData.date,
    },

    {
      label: "Time",
      value: stepData.time,
    },
  ],
};
  }

  // SAMPLE OUTSOURCED

  if (stepIndex === 3) {

    return {

  left: [
    {
      label: "Outsourced To",
      value: stepData.company,
    },
  ],

  right: [
    {
      label: "Date",
      value: stepData.date,
    },

    {
      label: "Time",
      value: stepData.time,
    },
  ],
};
  }

  // RESULT RECEIVED

  if (stepIndex === 4) {

    return {

      left: [],

      right: [
        {
          label: "Date",
          value: stepData.date,
        },

        {
          label: "Time",
          value: stepData.time,
        },
      ],
    };
  }

  // RESULT ENTRY

  if (stepIndex === 5) {

    return {

      left: [
        {
          label: "Staff ID",
          value: stepData.staffId,
        },

        {
          label: "Staff Name",
          value: stepData.staffName,
        },
      ],

      right: [
        {
          label: "Date",
          value: stepData.date,
        },

        {
          label: "Time",
          value: stepData.time,
        },
      ],
    };
  }

  // AUTHORISATION

  if (stepIndex === 6) {

    return {

  left: [
    {
      label: "Authoriser ID",
      value: stepData.authoriserId,
    },

    {
      label: "Authoriser Name",
      value: stepData.authoriserName,
    },
  ],

  right: [
    {
      label: "Date",
      value: stepData.date,
    },

    {
      label: "Time",
      value: stepData.time,
    },
  ],
};
  }

  // 2ND LEVEL AUTHORISATION

  if (stepIndex === 7) {

    return {

  left: [
    {
      label: "Authoriser ID",
      value: stepData.authoriserId,
    },

    {
      label: "Authoriser Name",
      value: stepData.authoriserName,
    },
  ],

  right: [
    {
      label: "Date",
      value: stepData.date,
    },

    {
      label: "Time",
      value: stepData.time,
    },
  ],
};
  }

  // REPORT PRINT

  if (stepIndex === 8) {

    return {

  left: [
    {
      label: "Staff ID",
      value: stepData.staffId,
    },

    {
      label: "Staff Name",
      value: stepData.staffName,
    },
  ],

  right: [
    {
      label: "Date",
      value: stepData.date,
    },

    {
      label: "Time",
      value: stepData.time,
    },
  ],
};
  }

  // REPORT SENT

  if (stepIndex === 9) {

    return {

  left: [
    {
      label: "Received By",
      value: stepData.receivedBy,
    },
  ],

  right: [
    {
      label: "Date",
      value: stepData.date,
    },

    {
      label: "Time",
      value: stepData.time,
    },
  ],
};
  }

  return {
    left: [],
    right: [],
  };
};
  const coords =
    window.innerWidth <= 768
      ? [
          { t: "9%", l: "15%" },
          { t: "9%", l: "45%" },
          { t: "9%", l: "75%" },
          { t: "40%", l: "90%" },
          { t: "40%", l: "65%" },
          { t: "40%", l: "35%" },
          { t: "40%", l: "10%" },
          { t: "75%", l: "25%" },
          { t: "75%", l: "55%" },
          { t: "75%", l: "90%" },
        ]
      : [
          { t: "9%", l: "15%" },
          { t: "9%", l: "45%" },
          { t: "9%", l: "75%" },
          { t: "40%", l: "90%" },
          { t: "40%", l: "65%" },
          { t: "40%", l: "35%" },
          { t: "40%", l: "10%" },
          { t: "75%", l: "25%" },
          { t: "75%", l: "55%" },
          { t: "75%", l: "90%" },
        ];

  const createCurvedPath = (points) => {

  if (!points.length) return "";

  let path = `M ${parseFloat(points[0].l)} ${parseFloat(points[0].t)}`;

  for (let i = 1; i < points.length; i++) {

    const prev = points[i - 1];
    const curr = points[i];

    const prevX = parseFloat(prev.l);
    const prevY = parseFloat(prev.t);

    const currX = parseFloat(curr.l);
    const currY = parseFloat(curr.t);

    const midX = (prevX + currX) / 2;
    const midY = (prevY + currY) / 2;

    path += ` Q ${prevX} ${prevY} ${midX} ${midY}`;
  }
  const lastPoint = points[points.length - 1];

  path += ` T ${parseFloat(lastPoint.l)} ${parseFloat(lastPoint.t)}`;

  return path;
};

const svgPath = createCurvedPath(coords);
let progressPath = "";

if (currentProgress >= 0) {

  const progressCoords = [];

  // FINAL STEP COMPLETED

  if (
    currentProgress >=
    coords.length - 1
  ) {

    progressCoords.push(...coords);

  } else {

    // ADD ALL COMPLETED STEPS
    // INCLUDING CURRENT STEP

    for (
      let i = 0;
      i <= currentProgress;
      i++
    ) {

      progressCoords.push(coords[i]);
    }

    // MOVE TOWARDS NEXT STEP

    const current =
      coords[currentProgress];

    const next =
      coords[currentProgress + 1];

    const currentX =
      parseFloat(current.l);

    const currentY =
      parseFloat(current.t);

    const nextX =
      parseFloat(next.l);

    const nextY =
      parseFloat(next.t);
  }

  // CREATE SVG PATH

progressPath = createCurvedPath(progressCoords);
}
  const processEstimates = [
    { hours: "15 mins", days: 0 },
    { hours: "30 mins", days: 0 },
    { hours: "2 Hours", days: 0 },
    { hours: "5 Hours", days: 1 },
    { hours: "4 Hours", days: 0 },
    { hours: "5 Hours", days: 0 },
    { hours: "6 Hours", days: 0 },
    { hours: "8 Hours", days: 1 },
    { hours: "10 Hours", days: 1 },
    { hours: "Completed", days: 0 },
  ];

  const overallProgress =
    sidInput && excelDatabase[sidInput]
      ? Math.round(
          (Object.values(excelDatabase[sidInput].tests).reduce(
            (a, b) => a + b,
            0
          ) /
            (Object.keys(excelDatabase[sidInput].tests).length * 9)) *
            100
        )
      : 0;

  const currentEstimate =
    currentProgress >= 0
      ? processEstimates[currentProgress]
      : { hours: "--", days: 0 };

  const estimatedDate = new Date();

  estimatedDate.setDate(
    estimatedDate.getDate() + currentEstimate.days
  );

  const formattedDate = estimatedDate.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });

  const currentData =
    sidInput && excelDatabase[sidInput] ? excelDatabase[sidInput] : {};

  if (!role) {
    return (
      <div className="login-page">
        <div className="login-card">
          <h1>Select Your Role</h1>

          <button
            className="login-btn"
            onClick={() => handleLogin("staff")}
          >
            Staff
          </button>

          <button
            className="login-btn patient-btn"
            onClick={() => handleLogin("patient")}
          >
            Patient
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="main-container">
      {role === "staff" && (
        <div className={`sidebar ${showSidebar ? "show" : ""}`}>
          <div className="sidebar-header">
            <button
              className="sidebar-close-btn"
              onClick={() => setShowSidebar(false)}
            >
              ✕
            </button>
          </div>

          <div className="client-logo">
            <img
              src="https://sukraa.in/wp-content/uploads/2025/02/Sukraa-Logo-111.png"
              alt="logo"
            />
          </div>

          <div className="search-box">
            <input
              type="text"
              placeholder="Search Patient..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="search-input"
            />

            {searchTerm && (
              <button
                className="clear-btn"
                onClick={() => setSearchTerm("")}
              >
                ✕
              </button>
            )}
          </div>

          <div className="patient-list">
            {filteredPatients.length > 0 ? (
              filteredPatients.map((patient, index) => {
                const data = excelDatabase[patient.sid];

                const isCompleted =
                  data &&
                  Object.values(data.tests).every((value) => value === 9);

                return (
                  <div
                    key={index}
                    className={`patient-card ${
                      selectedPatient === patient.sid
                        ? "active-patient"
                        : ""
                    }`}
                    style={{
                      borderLeft: isCompleted
                        ? "6px solid #22c55e"
                        : "6px solid #2563eb",
                    }}
                  >
                    <div
                      className="patient-title"
                      onClick={() => {
                        const data = excelDatabase[patient.sid];

                        setSidInput(patient.sid);
                        setSelectedPatient(patient.sid);
                        setPatientInfo(data);

                        const firstTest = Object.keys(data.tests)[0];

                        setSelectedTest(firstTest);
                        setCurrentProgress(data.tests[firstTest]);

                        setShowSidebar(false);
                      }}
                    >
                      Patient {index + 1}
                      <p className="patient-title2">
                        {patient.name}
                        {" - "}
                        {patient.sid}
                      </p>
                    </div>
                  </div>
                );
              })
            ) : (
              <p style={{ padding: "20px", color: "#000000" }}>
                No patients found
              </p>
            )}
          </div>
        </div>
      )}

      <div className="content-container">
        <button className="small-logout-btn" onClick={handleLogout}>
          Logout
        </button>
        <div className="topbar">
          {role === "staff" && (
            <button
              className="menu-btn"
              onClick={() => setShowSidebar(true)}
            >
              ☰
            </button>
          )}

          <div className="left-controls">
            <select
              value={selectedYear}
              onChange={(e) => setSelectedYear(e.target.value)}
              className="input"
            >
              <option>2026</option>
              <option>2027</option>
              <option>2028</option>
            </select>

            <select
              value={sidInput}
              onChange={(e) => {
                const value = e.target.value;
                setSidInput(value);

                const data = excelDatabase[value];

                if (data) {
                  setPatientInfo(data);

                  const firstTest = Object.keys(data.tests)[0];

                  setSelectedTest(firstTest);
                  setCurrentProgress(data.tests[firstTest]);
                }
              }}
              className="input"
            >
              <option value="">Select SID</option>

              {patientList.map((patient) => (
                <option key={patient.sid} value={patient.sid}>
                  {patient.sid}
                </option>
              ))}
            </select>
          </div>

          <div className="profile-container">
            <img
              src={
                patientInfo.photo ||
                "https://www.w3schools.com/howto/img_avatar.png"
              }
              alt="profile"
              className="profile-image"
            />

            <div>
              <p className="label">NAME</p>
              <h3 className="value">{patientInfo.name || "---"}</h3>
            </div>
            <div>
              <p className="label">AGE</p>
              <h3 className="value">{patientInfo.age || "--"}</h3>
            </div>
            <div>
              <p className="label">REFERRAL</p>
              <h3 className="value">{patientInfo.referral || "---"}</h3>
            </div>
          </div>
        </div>

        <div className="dashboard-top-section">
          <div className="estimation-container">
            <p>
              <b>Estimated time:</b> {currentEstimate.hours}
            </p>

            <p>
              <b>Estimated date:</b> {formattedDate}
            </p>
          </div>

          <div className="progress-card">
            <div className="progress-header">
              <span className="progress-title">Overall Progress</span>

              <span className="progress-percent">{overallProgress}%</span>
            </div>

            <div className="progress-bg">
              <div
                className="progress-fill"
                style={{
                  width: `${overallProgress}%`,
                }}
              ></div>
            </div>
          </div>

          <div className="status-selection-container">
            <p className="text">
              Current Status :{" "}
              {trackingSteps.find((s) => s.status === "in-progress")?.title ||
                "---"}
            </p>

            <select
              value={selectedTest}
              onChange={handleTestChange}
              className="left-controls2"
            >
              <option>Select Test</option>

              {sidInput &&
                excelDatabase[sidInput] &&
                Object.keys(excelDatabase[sidInput].tests).map((test) => (
                  <option key={test}>{test}</option>
                ))}
            </select>
          </div>
        </div>

        <div className="tracking-area">
          <svg
              viewBox="0 0 100 100"
              className="svg-style"
              preserveAspectRatio="none"
            >

              {/* GREY BACKGROUND PATH */}

              <path
                d={svgPath}
                className="path-bg"
              />

              {/* GREEN PROGRESS PATH */}

              {progressPath && (
                <path
                  d={progressPath}
                  className="path-fill"
                />
              )}

          </svg>

          {trackingSteps.map((step, idx) => {
            const config = getStatusConfig(step.status);

            return (
              <div
                key={idx}
                style={{
                  position: "absolute",
                  top: coords[idx].t,
                  left: coords[idx].l,
                  transform: "translate(-50%, -50%)",
                }}
              >
                <div
                  className={`step-image-container ${step.status}`}
                >
                  <img
                    src={`/steps/step${idx + 1}.png`}
                    alt={step.title}
                    className={`step-image 
                      ${idx === 9 ? "step10-image" : ""}
                      ${
                        step.status === "in-progress"
                          ? "active-step"
                          : ""
                      }
                    `}
                  />

                  
              <div
                  className={`

                    hover-info-card

                    ${
                      role === "staff"
                        ? "staff-hover-card"
                        : "patient-hover-card"
                    }

                  `}
                >

                  {/* STAFF VIEW */}

                  {role === "staff" ? (

                    <div className="hover-grid">

                      {/* LEFT */}

                      <div className="hover-column">

                        <div className="hover-heading">
                          Done By
                        </div>

                        {getHoverContent(idx)?.left?.map(
                          (item, i) => (

                            <div
                              className="hover-item"
                              key={i}
                            >
                              <span className="hover-label">
                                {item.label}
                              </span>

                              <span className="hover-value">
                                {item.value || "--"}
                              </span>
                            </div>
                          )
                        )}

                      </div>

                      {/* RIGHT */}

                      <div className="hover-column">

                        <div className="hover-heading">
                          Done On
                        </div>

                        {getHoverContent(idx)?.right?.map(
                          (item, i) => (

                            <div
                              className="hover-item"
                              key={i}
                            >
                              <span className="hover-label">
                                {item.label}
                              </span>

                              <span className="hover-value">
                                {item.value || "--"}
                              </span>
                            </div>
                          )
                        )}

                      </div>

                    </div>

                  ) : (

                    /* PATIENT VIEW */

                    <div className="patient-hover-view">

                      {getHoverContent(idx)?.right?.map(
                        (item, i) => (

                          <div
                            className="hover-item"
                            key={i}
                          >
                            <span className="hover-label">
                              {item.label}
                            </span>

                            <span className="hover-value">
                              {item.value || "--"}
                            </span>
                          </div>
                        )
                      )}

                    </div>

                  )}

                </div>


                </div>
                </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};

export default LabDashboard;
