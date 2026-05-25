import React, { useState } from "react";
import {
  Search,
  LogOut,
} from "lucide-react";
import dummyprocessdata from "../Data/dummyprocessdata";
import dummyLabs from "../Data/dummylabs";
import "./OutSourcing_Process.css";
import { Link } from "react-router-dom";
const OutSourcingProcess = () => {

  const [labs] = useState(dummyLabs);

  const [processData, setProcessData] =
    useState(dummyprocessdata);

  const [selectedLab, setSelectedLab] =
    useState("");

  const [searchTerm, setSearchTerm] =
    useState("");

  const [activeTab, setActiveTab] =
    useState("received");

  const [selectedTests, setSelectedTests] =
    useState([]);

  const filteredLabs = labs.filter((lab) =>
    lab.name
      .toLowerCase()
      .includes(searchTerm.toLowerCase())
  );

  const filteredProcessData =
    processData.filter(
      (item) =>
        item.status === activeTab &&
        (!selectedLab ||
          item.lab === selectedLab)
    );

  const toggleSelection = (id) => {

    if (selectedTests.includes(id)) {

      setSelectedTests(
        selectedTests.filter(
          (item) => item !== id
        )
      );

    } else {

      setSelectedTests([
        ...selectedTests,
        id,
      ]);

    }

  };

  const handleSelectAll = () => {

    const allIds =
      filteredProcessData.map(
        (item) => item.id
      );

    if (
      selectedTests.length ===
      filteredProcessData.length
    ) {

      setSelectedTests([]);

    } else {

      setSelectedTests(allIds);

    }

  };

  const moveToProcessing = () => {

    const updated =
      processData.map((item) => {

        if (
          selectedTests.includes(item.id)
        ) {

          return {
            ...item,
            status: "processing",
          };

        }

        return item;

      });

    setProcessData(updated);

    setSelectedTests([]);

  };

  const moveToCompleted = () => {

    const updated =
      processData.map((item) => {

        if (
          selectedTests.includes(item.id)
        ) {

          return {
            ...item,
            status: "completed",
          };

        }

        return item;

      });

    setProcessData(updated);

    setSelectedTests([]);

  };

  const moveToSent = () => {

    const updated =
      processData.map((item) => {

        if (
          selectedTests.includes(item.id)
        ) {

          return {
            ...item,
            status: "sent",
          };

        }

        return item;

      });

    setProcessData(updated);

    setSelectedTests([]);

  };

  return (

    <div className="dashboard-container1">

      {/* TOPBAR */}

      <div className="topbar">

        {/* LEFT */}

        <div className="topbar-left">

          <img
            className="logo"
            src="https://uploads.onecompiler.io/44neydbch/44nfp3w34/Sukraa-Logo-111.png"
            alt="logo"
          />

        </div>

        {/* CENTER */}

        <div className="topbar-center">

        <Link to="/outsourcing_assign"> 
            <button className="nav-btn"> 
                Outsourcing Assign 
            </button> 
        </Link>

          <Link to="/outsourcing_process"> 
            <button className="nav-btn active-nav"> 
                Outsourcing Process
            </button> 
        </Link>

        </div>

        {/* RIGHT */}

        <div className="topbar-right">

          <div className="staff-details">

            <h3>Lavanya</h3>

            <p>STAFF ID : STF-102</p>

          </div>

          <button className="logout-btn">

            Logout

            <LogOut size={16} />

          </button>

        </div>

      </div>

      {/* MAIN GRID */}

      <div className="dashboard-grid">

        {/* LEFT PANEL */}

        <div className="sub-panel left-panel">

          <div className="search-box">

            <Search size={18} />

            <input
              type="text"
              placeholder="Search Labs..."
              value={searchTerm}
              onChange={(e) =>
                setSearchTerm(e.target.value)
              }
            />

          </div>

          <h2 className="panel-title">
            RECEIVED FROM LABS
          </h2>

          <div className="scroll-area">

            {filteredLabs.map((lab) => (

              <div
                key={lab.id}
                className={`lab-btn ${
                  selectedLab === lab.name
                    ? "active-lab"
                    : ""
                }`}
                onClick={() =>
                  setSelectedLab(lab.name)
                }
              >

                <h3>{lab.name}</h3>

                <p className="supported-tests">
                  {lab.supportedTests.join(", ")}
                </p>

                <div className="lab-stats">

                  <span>
                    Received:
                    {
                      processData.filter(
                        (item) =>
                          item.lab === lab.name
                      ).length
                    }
                  </span>

                  <span>
                    Pending:
                    {
                      processData.filter(
                        (item) =>
                          item.lab === lab.name &&
                          item.status !== "sent"
                      ).length
                    }
                  </span>

                </div>

              </div>

            ))}

          </div>

        </div>

        {/* RIGHT PANEL */}

        <div className="sub-panel tests-panel">

          {/* TABS */}

          <div className="process-tabs">

            <button
              className={`nav-btn ${
  activeTab === "received"
    ? "active-tab"
    : ""
}`}
              onClick={() =>
                setActiveTab("received")
              }
            >
              Received Samples
            </button>

            <button
              className={`nav-btn ${
  activeTab === "yet"
    ? "active-tab"
    : ""
}`}
              onClick={() =>
                setActiveTab("yet")
              }
            >
              Yet To Receive
            </button>

            <button
              className={`nav-btn ${
  activeTab === "processing"
    ? "active-tab"
    : ""
}`}
              onClick={() =>
                setActiveTab("processing")
              }
            >
              Currently Processing
            </button>

            <button
              className={`nav-btn ${
  activeTab === "completed"
    ? "active-tab"
    : ""
}`}
              onClick={() =>
                setActiveTab("completed")
              }
            >
              Completed
            </button>

            <button
              className={`nav-btn ${
  activeTab === "sent"
    ? "active-tab"
    : ""
}`}
              onClick={() =>
                setActiveTab("sent")
              }
            >
              Sent
            </button>

          </div>

          {/* BULK ACTIONS */}

          {activeTab !== "yet" &&
            activeTab !== "sent" && (

            <div className="bulk-actions">

              <label className="select-all">

                <input
                  type="checkbox"
                  checked={
                    filteredProcessData.length > 0 &&
                    selectedTests.length ===
                    filteredProcessData.length
                  }
                  onChange={handleSelectAll}
                />

                Select All

              </label>

              {activeTab === "received" && (

                <button
                  className="assign-btn"
                  onClick={moveToProcessing}
                >
                  Process
                </button>

              )}

              {activeTab === "processing" && (

                <button
                  className="assign-btn"
                  onClick={moveToCompleted}
                >
                  Complete
                </button>

              )}

              {activeTab === "completed" && (

                <button
                  className="assign-btn"
                  onClick={moveToSent}
                >
                  Send
                </button>

              )}

            </div>

          )}

          {/* DATA */}

          <div className="scroll-area">

            {filteredProcessData.length === 0 ? (

              <div className="empty-message">
                No Data Available
              </div>

            ) : (

              filteredProcessData.map((item) => (

                <div
                  className="test-card single-line-card"
                  key={item.id}
                >

                  {activeTab !== "yet" &&
                    activeTab !== "sent" && (

                    <input
                      type="checkbox"
                      checked={selectedTests.includes(
                        item.id
                      )}
                      onChange={() =>
                        toggleSelection(item.id)
                      }
                    />

                  )}

                  <h4>
                    {item.name} :
                    [{item.patient} - {item.sid}]
                  </h4>

                </div>

              ))

            )}

          </div>

        </div>

      </div>

    </div>

  );
};

export default OutSourcingProcess;

