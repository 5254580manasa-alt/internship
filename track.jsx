import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import "./Lashboard.css";
import React from "react";
import ReactDOM from "react-dom/client";
import {
  BrowserRouter,
  Routes,
  Route,
} from "react-router-dom";

const track = () => {

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

  const [selectedPatient, setSelectedPatient] =
    useState("");

  const [patientInfo, setPatientInfo] =
    useState({
      name: "",
      age: "",
      referral: "",
      contact: "",
      photo: "",
    });

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
  XLSX.utils.sheet_to_json(worksheet, {
    defval: "",
  });

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
              "General Test": {
                progress: 0,
              },
            };

          }
          console.log("ROW:", row);
          if (!row.SID) return;

            const sid =
              String(row.SID).trim().padStart(6, "0");

          formattedData[sid] = {

            name: row.Name,
            age: row.Age,
            referral: row.Referral,
            contact: row.Contact,
            photo: row.Photo,
            tests: parsedTests,

          };

          patients.push({
            sid: sid,
            name: row.Name || "Unknown",
            contact: row.Contact || "",
            referral: row.Referral || "",
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
          data.tests[firstTest].progress
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
        data.tests[testName].progress
      );

    }

  };

  const stepTitles = [

    "Billing",
    "Sample Collection",
    "Sample Processing",
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

  const coords = [

    { t: "15%", l: "15%" },
    { t: "10%", l: "35%" },
    { t: "25%", l: "55%" },
    { t: "10%", l: "75%" },
    { t: "45%", l: "60%" },
    { t: "60%", l: "40%" },
    { t: "75%", l: "20%" },
    { t: "88%", l: "45%" },
    { t: "75%", l: "75%" },

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

  const currentTestData =
    sidInput &&
    excelDatabase[sidInput] &&
    selectedTest
      ? excelDatabase[sidInput].tests[selectedTest]
      : null;

  return (

    <div className="main-container">

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

                    {
                      [
                        "💳",
                        "🧪",
                        "⚙",
                        "📥",
                        "💻",
                        "✔",
                        "✅",
                        "🖨",
                        "📄",
                      ][idx]
                    }

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

                  <div className="step-extra-info">

                    {idx === 0 &&
                      currentTestData?.billing && (
                        <>
                          <p>
                            {
                              currentTestData
                                .billing
                                .staffName
                            }
                          </p>

                          <p>
                            ID:
                            {" "}
                            {
                              currentTestData
                                .billing
                                .staffId
                            }
                          </p>
                        </>
                      )}

                    {idx === 1 &&
                      currentTestData?.sampleCollection && (
                        <>
                          <p>
                            {
                              currentTestData
                                .sampleCollection
                                .staffName
                            }
                          </p>

                          <p>
                            ID:
                            {" "}
                            {
                              currentTestData
                                .sampleCollection
                                .staffId
                            }
                          </p>
                        </>
                      )}

                    {idx === 2 &&
                      currentTestData?.sampleProcessing && (
                        <>
                          <p>
                            {
                              currentTestData
                                .sampleProcessing
                                .labName
                            }
                          </p>

                          {
                            currentTestData
                              .sampleProcessing
                              .outsourced && (
                                <p>
                                  To:
                                  {" "}
                                  {
                                    currentTestData
                                      .sampleProcessing
                                      .outsourcedTo
                                  }
                                </p>
                              )
                          }
                        </>
                      )}

                    {idx === 3 &&
                      currentTestData?.resultReceived && (
                        <>
                          <p>
                            From:
                            {" "}
                            {
                              currentTestData
                                .resultReceived
                                .from
                            }
                          </p>

                          <p>
                            To:
                            {" "}
                            {
                              currentTestData
                                .resultReceived
                                .to
                            }
                          </p>
                        </>
                      )}

                    {idx === 4 &&
                      currentTestData?.resultEntry && (
                        <>
                          <p>
                            {
                              currentTestData
                                .resultEntry
                                .staffName
                            }
                          </p>

                          <p>
                            ID:
                            {" "}
                            {
                              currentTestData
                                .resultEntry
                                .staffId
                            }
                          </p>
                        </>
                      )}

                    {idx === 5 &&
                      currentTestData?.authorisation && (
                        <>
                          <p>
                            {
                              currentTestData
                                .authorisation
                                .staffName
                            }
                          </p>

                          <p>
                            ID:
                            {" "}
                            {
                              currentTestData
                                .authorisation
                                .staffId
                            }
                          </p>
                        </>
                      )}

                    {idx === 6 &&
                      currentTestData?.secondAuthorisation && (
                        <>
                          <p>
                            {
                              currentTestData
                                .secondAuthorisation
                                .staffName
                            }
                          </p>

                          <p>
                            ID:
                            {" "}
                            {
                              currentTestData
                                .secondAuthorisation
                                .staffId
                            }
                          </p>
                        </>
                      )}

                    {idx === 7 &&
                      currentTestData?.reportPrint && (
                        <>
                          <p>
                            {
                              currentTestData
                                .reportPrint
                                .staffName
                            }
                          </p>

                          <p>
                            ID:
                            {" "}
                            {
                              currentTestData
                                .reportPrint
                                .staffId
                            }
                          </p>
                        </>
                      )}

                    {idx === 8 &&
                      currentTestData?.reportSent && (
                        <>
                          <p>
                            Received by:
                            {" "}
                            {
                              currentTestData
                                .reportSent
                                .receivedBy
                            }
                          </p>
                        </>
                      )}

                  </div>

                </div>

              </div>

            );

          }
        )}

      </div>

    </div>

  );

};

export default track;