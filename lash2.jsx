  import React, { useState, useEffect } from "react";
  import * as XLSX from "xlsx";
  import "./lashboard2.css";
  const lash2 = () => {

    const [role, setRole] = useState(
      localStorage.getItem("role") || ""
    );

    const [showSidebar, setShowSidebar] =
      useState(false);

    const [sidInput, setSidInput] =
      useState("");

    const [selectedYear, setSelectedYear] =
      useState("2026");

    const [selectedTest, setSelectedTest] =
      useState("");

    const [searchTerm, setSearchTerm] =
      useState("");

    const [currentProgress, setCurrentProgress] =
      useState(-1);

    const [excelDatabase, setExcelDatabase] =
      useState({});

    const [patientList, setPatientList] =
      useState([]);

    const [patientInfo, setPatientInfo] =
      useState({
        name: "",
        age: "",
        referral: "",
        contact: "",
        photo: "",
      });
const [selectedPatient, setSelectedPatient] =
  useState("");
    const handleLogin = (selectedRole) => {
      setRole(selectedRole);

      localStorage.setItem(
        "role",
        selectedRole
      );
    };
    
    const handleLogout = () => {
      localStorage.removeItem("role");
      setRole("");
    };

    // FETCH EXCEL

    useEffect(() => {

      const fetchExcelData = async () => {

        try {

          const response =
            await fetch("/data.xlsx");

          const arrayBuffer =
            await response.arrayBuffer();

          const data =
            new Uint8Array(arrayBuffer);

          const workbook =
            XLSX.read(data, {
              type: "array",
            });

          const worksheet =
            workbook.Sheets[selectedYear];

          const jsonData =
            XLSX.utils.sheet_to_json(worksheet);

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

            const sid =
              String(row.SID).padStart(6, "0");

            formattedData[sid] = {

              name: row.Name,
              age: row.Age,
              referral: row.Referral,
              contact: row.Contact,
              photo: row.Photo,
              tests: parsedTests,

            };

            patients.push({

              sid,
              name: row.Name,
              contact: row.Contact,
              referral: row.Referral,

            });

          });

          setExcelDatabase(formattedData);
          setPatientList(patients);

        } catch (error) {

          console.error(error);

        }

      };

      fetchExcelData();

    }, [selectedYear]);

    // SEARCH

    const filteredPatients =
      patientList.filter((patient) =>

        patient.sid.includes(searchTerm) ||

        patient.name
          ?.toLowerCase()
          .includes(searchTerm.toLowerCase()) ||

        String(patient.contact)
          ?.includes(searchTerm) ||

        patient.referral
          ?.toLowerCase()
          .includes(searchTerm.toLowerCase())
      );

    // SID INPUT

    const handleSidChange = (e) => {

      const value =
        e.target.value
          .replace(/\D/g, "")
          .slice(0, 6);

      setSidInput(value);

      if (value.length === 6) {

        const data =
          excelDatabase[value];

        if (data) {

          setPatientInfo(data);

          const firstTest =
            Object.keys(data.tests)[0];

          setSelectedTest(firstTest);

          setCurrentProgress(
            data.tests[firstTest]
          );

        }

      }

    };

    const handleTestChange = (e) => {

      const testName = e.target.value;

      setSelectedTest(testName);

      const data =
        excelDatabase[sidInput];

      if (
        data &&
        data.tests[testName] !== undefined
      ) {

        setCurrentProgress(
          data.tests[testName]
        );

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

  if (currentProgress === -1)
    return "pending";

  /* FINAL STAGE */
  if (currentProgress === 8) {
    return "completed";
  }

  if (index < currentProgress)
    return "completed";

  if (index === currentProgress)
    return "in-progress";

  return "pending";

};

    const trackingSteps =
      stepTitles.map((title, idx) => ({
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

   const coords = window.innerWidth <= 768

  ? [

      { t: "15%", l: "15%" },
      { t: "10%", l: "35%" },
      { t: "25%", l: "55%" },
      { t: "10%", l: "75%" },
      { t: "45%", l: "60%" },
      { t: "60%", l: "40%" },
      { t: "75%", l: "20%" },
      { t: "88%", l: "45%" },
      { t: "75%", l: "75%" },

    ]

  : [

      { t: "30%", l: "15%" },
      { t: "20%", l: "35%" },
      { t: "20%", l: "55%" },
      { t: "19%", l: "75%" },
      { t: "40%", l: "90%" },
      { t: "55%", l: "60%" },
      { t: "70%", l: "40%" },
      { t: "85%", l: "20%" },
      { t: "90%", l: "45%" },
      { t: "85%", l: "75%" },

    ];

    const svgPath = coords
      .map((point, index) => {

        const x = parseFloat(point.l);
        const y = parseFloat(point.t);

        return `${
          index === 0 ? "M" : "L"
        } ${x} ${y}`;

      })
      .join(" ");
    const processEstimates = [
    { hours: "15 mins", days: 0 },
    { hours: "30 mins", days: 0 },
    { hours: "2 Hours", days: 0 },
    { hours: "4 Hours", days: 0 },
    { hours: "5 Hours", days: 0 },
    { hours: "6 Hours", days: 0 },
    { hours: "8 Hours", days: 1 },
    { hours: "10 Hours", days: 1 },
    { hours: "Completed", days: 0 },
  ];

      const overallProgress = sidInput && 
    excelDatabase[sidInput] ?   
    Math.round( (Object.values( 
      excelDatabase[sidInput].tests ).reduce((a, b) => a + b, 0) / (Object.keys( 
        excelDatabase[sidInput].tests ).length * 8)) * 100 ) : 0;
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
    const currentActiveStep =
      trackingSteps.find((s) => s.status === "in-progress") ||
      (currentProgress === 8
        ? trackingSteps[8]
        : { title: "NO SID ENTERED" });


    if (!role) {

      return (

        <div className="login-page">

          <div className="login-card">

            <h1>Select Your Role</h1>

            <button
              className="login-btn"
              onClick={() =>
                handleLogin("staff")
              }
            >
              Staff
            </button>

            <button
              className="login-btn patient-btn"
              onClick={() =>
                handleLogin("patient")
              }
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

  <div
    className={`sidebar ${
      showSidebar ? "show" : ""
    }`}
  >

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
        onChange={(e) =>
          setSearchTerm(e.target.value)
        }
        className="search-input"
      />

      {searchTerm && (

        <button
          className="clear-btn"
          onClick={() =>
            setSearchTerm("")
          }
        >
          ✕
        </button>

      )}

    </div>

    {/* SCROLLABLE PATIENT LIST */}

    <div className="patient-list">
          
      {filteredPatients.map((patient, index) => {

  const data =
    excelDatabase[patient.sid];

  const isCompleted =
    data &&
    Object.values(data.tests).every(
      (value) => value === 8
    );

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

          const data =
            excelDatabase[patient.sid];

          setSidInput(patient.sid);
          setSelectedPatient(patient.sid);

          setPatientInfo(data);

          const firstTest =
            Object.keys(data.tests)[0];

          setSelectedTest(firstTest);

          setCurrentProgress(
            data.tests[firstTest]
          );

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

})}
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
                onClick={() =>
                  setShowSidebar(true)
                }
              >
                ☰
              </button>

            )}

            <div className="left-controls">

              <select
                value={selectedYear}
                onChange={(e) =>
                  setSelectedYear(
                    e.target.value
                  )
                }
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

    const data =
      excelDatabase[value];

    if (data) {

      setPatientInfo(data);

      const firstTest =
        Object.keys(data.tests)[0];

      setSelectedTest(firstTest);

      setCurrentProgress(
        data.tests[firstTest]
      );

    }

  }}
  className="input"
>

  <option value="">
    Select SID
  </option>

  {patientList.map((patient) => (

    <option
      key={patient.sid}
      value={patient.sid}
    >
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

                <p className="label">
                  NAME
                </p>

                <h3 className="value">
                  {patientInfo.name || "---"}
                </h3>

              </div>
  <div> <p className="label">AGE</p> 
  <h3 className="value"> {patientInfo.age || "--"} 
    </h3> 
    </div> <div> <p className="label"> REFERRAL </p> <h3 className="value"> {patientInfo.referral || "---"} </h3> </div>
            </div>

          </div>
<div className="dashboard-top-section">

  {/* LEFT SIDE */}

  <div className="estimation-container">

    <p>
      <b>Estimated time:</b>
      {" "}
      {currentEstimate.hours}
    </p>

    <p>
      <b>Estimated date:</b>
      {" "}
      {formattedDate}
    </p>

  </div>

  {/* CENTER */}

  <div className="progress-card">

    <div className="progress-header">

      <span className="progress-title">
        Overall Progress
      </span>

      <span className="progress-percent">
        {overallProgress}%
      </span>

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

  {/* RIGHT SIDE */}

  <div className="status-selection-container">

    <p className="text">

      Current Status :

      {" "}

      {trackingSteps.find(
        (s) => s.status === "in-progress"
      )?.title || "No SID Entered"}

    </p>

    <select
      value={selectedTest}
      onChange={handleTestChange}
      className="left-controls2"
    >

      <option>Select Test</option>

      {sidInput &&
        excelDatabase[sidInput] &&
        Object.keys(
          excelDatabase[sidInput].tests
        ).map((test) => (

          <option key={test}>
            {test}
          </option>

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
    <path
      d={svgPath}
      fill="none"
      stroke="#94a3b8"
      /* Ippo 1-nu irukuradha 0.5-ku mathunga */
      strokeWidth="0.5" 
      strokeDasharray="3 2"
    />
  </svg>

              {trackingSteps.map(
                (step, idx) => {

                  const config =
                    getStatusConfig(
                      step.status
                    );

                  return (

                    <div
                      key={idx}
                      style={{
                        position: "absolute",
                        top: coords[idx].t,
                        left: coords[idx].l,
                        transform:
                          "translate(-50%, -50%)",
                      }}
                    >

                      <div
                        className="step-card"
                        style={{
                          background:
                            config.bg,
                          border:
                            `1px solid ${config.color}`,
                        }}
                      >

                        <div
                          className="step-circle"
                          style={{
                            background:
                              config.color,
                          }}
                        >
                          {idx + 1}
                        </div>

                        <div className="step-title">
                          {step.title}
                        </div>

                        <div
                          className="step-status"
                          style={{
                            color:
                              config.color,
                          }}
                        >
                          {config.label}
                        </div>

                      </div>

                    </div>

                  );

                }
              )}

            </div>

          </div>

        </div>


    );

  };

  export default lash2;

