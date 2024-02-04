import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import "./styles.css";
import {
  FaEdit,
  FaTrash,
  FaPlus,
  FaUndo,
  FaRedo,
  FaSync,
  FaDownload,
  FaCloudUploadAlt,
  FaTimes,
} from "react-icons/fa";
import { debounce } from "lodash";

export default function App() {
  const [selectedFile, setSelectedFile] = useState(null);
  const [workbook, setWorkbook] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [currentSheetIndex, setCurrentSheetIndex] = useState(0);
  const [jsonData, setJsonData] = useState([]);
  const [sheetData, setSheetData] = useState({});
  const [history, setHistory] = useState([]);
  const [currentStep, setCurrentStep] = useState(0);
  const [draggedIndex, setDraggedIndex] = useState(null);
  const [draggedType, setDraggedType] = useState(null);
  const [isMobile, setIsMobile] = useState(false);

  useEffect(() => {
    const handleResize = () => {
      setIsMobile(window.innerWidth <= 768);
    };

    handleResize();

    window.addEventListener("resize", handleResize);

    return () => {
      window.removeEventListener("resize", handleResize);
    };
  }, []);

  const handleFileUpload = (event) => {
    setSelectedFile(event.target.files[0]);
    setWorkbook(null);
    setSheetNames([]);
    setCurrentSheetIndex(0);
    setJsonData([]);
    setSheetData({});
    setHistory([]);
    setCurrentStep(0);
    convertToWorkbook(event.target.files[0]);
  };

  const convertToWorkbook = (file) => {
    const fileReader = new FileReader();
    fileReader.onload = (event) => {
      const data = event.target.result;
      const workbook = XLSX.read(data, { type: "binary" });
      setWorkbook(workbook);
      const sheetNames = workbook.SheetNames;
      setSheetNames(sheetNames);
      setCurrentSheetIndex(0);

      const maxColumns = sheetNames.reduce((max, sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        const maxRowColumns = Math.max(...json.map((row) => row.length));
        return Math.max(max, maxRowColumns);
      }, 0);

      const filledData = sheetNames.reduce((acc, sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        if (json.length === 0 || (json.length === 1 && json[0].length === 0)) {
          const emptyRows = Array.from({ length: 50 }, () =>
            Array(50).fill("")
          );
          acc[sheetName] = emptyRows;
        } else {
          const filledRows = json.map((row) => {
            const emptyCells = Array(maxColumns - row.length).fill("");
            return [...row, ...emptyCells];
          });
          acc[sheetName] = filledRows;
        }
        return acc;
      }, {});

      setSheetData(filledData);
      setJsonData(filledData[sheetNames[currentSheetIndex]]);
      addToHistory(filledData[sheetNames[currentSheetIndex]]);
    };

    fileReader.readAsBinaryString(file);
  };

  const convertSheetToJSON = (workbook, sheetNames) => {
    const data = {};
    sheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      data[sheetName] = json;
    });
    setSheetData(data);
    setJsonData(data[sheetNames[currentSheetIndex]]);
    addToHistory(data[sheetNames[currentSheetIndex]]);
  };

  const downloadExcel = () => {
    const combinedWorkbook = XLSX.utils.book_new();

    sheetNames.forEach((sheetName) => {
      const jsonDataForSheet = sheetData[sheetName];
      const worksheetForSheet = XLSX.utils.json_to_sheet(jsonDataForSheet);

      XLSX.utils.book_append_sheet(
        combinedWorkbook,
        worksheetForSheet,
        sheetName
      );
    });

    const excelBinaryData = XLSX.write(combinedWorkbook, {
      bookType: "xlsx",
      type: "binary",
    });

    const blob = new Blob([s2ab(excelBinaryData)], {
      type: "application/octet-stream",
    });

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "combined_sheets_export.xlsx";
    link.click();

    URL.revokeObjectURL(url);
  };

  function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
  }

  const handleCellChange = (rowIndex, colIndex, value) => {
    const updatedData = JSON.parse(JSON.stringify(jsonData));
    updatedData[rowIndex][colIndex] = value;
    setJsonData(updatedData);
    addToHistory(updatedData);
    setSheetData({
      ...sheetData,
      [sheetNames[currentSheetIndex]]: updatedData,
    });
  };

  const deleteRow = (rowIndex) => {
    const updatedData = [...jsonData];
    updatedData.splice(rowIndex, 1);
    setJsonData(updatedData);
    addToHistory(updatedData);
    setSheetData({
      ...sheetData,
      [sheetNames[currentSheetIndex]]: updatedData,
    });
  };

  const deleteColumn = (colIndex) => {
    const updatedData = jsonData.map((row) => {
      const newRow = [...row];
      newRow.splice(colIndex, 1);
      return newRow;
    });
    setJsonData(updatedData);
    addToHistory(updatedData);
    setSheetData({
      ...sheetData,
      [sheetNames[currentSheetIndex]]: updatedData,
    });
  };

  const switchSheet = (sheetIndex) => {
    setCurrentSheetIndex(sheetIndex);
    const sheetName = sheetNames[sheetIndex];
    setJsonData(sheetData[sheetName]);
  };

  useEffect(() => {
    if (workbook) {
      convertSheetToJSON(workbook, sheetNames);
    }
  }, [workbook, sheetNames]);

  const undo = () => {
    if (currentStep > 0) {
      setCurrentStep((prevStep) => prevStep - 1);
    }
  };

  const redo = () => {
    if (currentStep < history.length - 1) {
      setCurrentStep((prevStep) => prevStep + 1);
    }
  };

  const reset = () => {
    setCurrentStep(0);
    setJsonData(history[0]);
  };

  const addToHistory = (data) => {
    const updatedHistory = history.slice(0, currentStep + 1);
    updatedHistory.push(data);
    setHistory(updatedHistory);
    setCurrentStep(updatedHistory.length - 1);
  };

  useEffect(() => {
    if (currentStep === 0) {
      setJsonData(history[0]);
    } else {
      setJsonData(history[currentStep]);
    }
  }, [currentStep, history]);

  useEffect(() => {
    const initializeDefaultSheet = () => {
      const defaultSheetName = "DefaultSheet";
      const defaultRows = Array.from({ length: 50 }, () => Array(50).fill(""));

      setSheetData({
        [defaultSheetName]: defaultRows,
      });

      setJsonData(defaultRows);
      addToHistory(defaultRows);

      setSheetNames([defaultSheetName]);

      const defaultWorkbook = XLSX.utils.book_new();
      const worksheetForDefaultSheet = XLSX.utils.aoa_to_sheet(
        defaultRows.map((row) => row.map((cell) => ""))
      );
      XLSX.utils.book_append_sheet(
        defaultWorkbook,
        worksheetForDefaultSheet,
        defaultSheetName
      );
      setWorkbook(defaultWorkbook);
    };

    if (Object.keys(sheetData).length === 0) {
      initializeDefaultSheet();
    }
  }, [sheetData]);

  const handleDragStart = (event, rowIndex, colIndex, dragType) => {
    event.dataTransfer.setData("text/plain", "");
    setDraggedIndex(dragType === "row" ? rowIndex : colIndex);
    setDraggedType(dragType);
  };

  const handleDragOver = (event) => {
    event.preventDefault();
    event.dataTransfer.dropEffect = "move";
  };
  const handleDrop = (event, targetIndex, type) => {
    event.preventDefault();
    if (draggedIndex === targetIndex || draggedType !== type) {
      return;
    }
    const updatedData = [...jsonData];
    if (type === "row") {
      const [draggedRow] = updatedData.splice(draggedIndex, 1);
      updatedData.splice(targetIndex, 0, draggedRow);
    } else if (type === "col") {
      for (let i = 0; i < updatedData.length; i++) {
        const removedCell = updatedData[i].splice(draggedIndex, 1)[0];
        updatedData[i].splice(targetIndex, 0, removedCell);
      }
    }
    setJsonData(updatedData);
    addToHistory(updatedData);
    setSheetData({
      ...sheetData,
      [sheetNames[currentSheetIndex]]: updatedData,
    });
  };

  const handleScroll = (event) => {
    const target = event.target;

    if (target.scrollTop + target.clientHeight + 1 >= target.scrollHeight) {
      addRowsOnScroll();
    }

    if (target.scrollLeft + target.clientWidth + 1 >= target.scrollWidth) {
      addColumnsOnScroll();
    }
  };

  const addRowsOnScroll = () => {
    const updatedData = [...jsonData];
    for (let i = 0; i < 50; i++) {
      const newRow = Array(jsonData[0].length).fill("");
      updatedData.push(newRow);
    }

    setJsonData(updatedData);
    addToHistory(updatedData);
    setSheetData({
      ...sheetData,
      [sheetNames[currentSheetIndex]]: updatedData,
    });
  };

  const addColumnsOnScroll = () => {
    const updatedData = jsonData.map((row) => [...row, ...Array(50).fill("")]);

    setJsonData(updatedData);
    addToHistory(updatedData);
    setSheetData({
      ...sheetData,
      [sheetNames[currentSheetIndex]]: updatedData,
    });
  };

  const addSheet = () => {
    const createEmptySheet = (name) => {
      const emptyRows = Array.from({ length: 50 }, () => Array(50).fill(""));
      const updatedSheetData = { ...sheetData, [name]: emptyRows };
      const updatedSheetNames = [...sheetNames, name];

      setSheetData(updatedSheetData);
      setSheetNames(updatedSheetNames);
      setCurrentSheetIndex(updatedSheetNames.length - 1);
      setJsonData([]);
      addToHistory([]);

      const updatedWorkbook = XLSX.utils.book_new();
      updatedSheetNames.forEach((sheetName) => {
        const jsonDataForSheet = updatedSheetData[sheetName];
        const worksheetForSheet = XLSX.utils.aoa_to_sheet(
          jsonDataForSheet.map((row) => row.map((cell) => ""))
        );
        XLSX.utils.book_append_sheet(
          updatedWorkbook,
          worksheetForSheet,
          sheetName
        );
      });
      setWorkbook(updatedWorkbook);
    };

    let newSheetName;
    do {
      newSheetName = `Sheet ${sheetNames.length + 1}`;
    } while (sheetNames.includes(newSheetName));

    createEmptySheet(newSheetName);
  };

  const deleteSheet = (index) => {
    const updatedSheetData = { ...sheetData };
    const updatedSheetNames = [...sheetNames];

    delete updatedSheetData[sheetNames[index]];
    updatedSheetNames.splice(index, 1);

    setSheetData(updatedSheetData);
    setSheetNames(updatedSheetNames);

    if (index === currentSheetIndex) {
      setCurrentSheetIndex(0);
      setJsonData(updatedSheetData[updatedSheetNames[0]]);
    }
  };

  const editSheetName = (index) => {
    const sheetNameToEdit = sheetNames[index];
    const newSheetName = prompt(
      "Enter a new name for the sheet:",
      sheetNameToEdit
    );

    if (newSheetName && newSheetName !== sheetNameToEdit) {
      const updatedSheetData = { ...sheetData };
      const updatedSheetNames = [...sheetNames];

      updatedSheetData[newSheetName] = updatedSheetData[sheetNameToEdit];
      delete updatedSheetData[sheetNameToEdit];

      updatedSheetNames[index] = newSheetName;

      setSheetData(updatedSheetData);
      setSheetNames(updatedSheetNames);
      setCurrentSheetIndex(index);

      const updatedWorkbook = XLSX.utils.book_new();
      updatedSheetNames.forEach((name) => {
        const jsonDataForSheet = updatedSheetData[name];
        const worksheetForSheet = XLSX.utils.json_to_sheet(jsonDataForSheet);
        XLSX.utils.book_append_sheet(updatedWorkbook, worksheetForSheet, name);
      });
      setWorkbook(updatedWorkbook);
    }
  };

  const emptyCurrentSheet = () => {
    const emptyRows = Array.from({ length: 50 }, () => Array(50).fill(""));
    setSheetData({
      ...sheetData,
      [sheetNames[currentSheetIndex]]: emptyRows,
    });
    setJsonData(emptyRows);
    addToHistory(emptyRows);
  };

  const [mobileScreenHeight, setMobileScreenHeight] = useState(0);

  useEffect(() => {
    function getMobileScreenHeight() {
      const isMobile = /Mobi/i.test(window.navigator.userAgent);
      const hasOuterHeight = "outerHeight" in window;

      if (isMobile && hasOuterHeight) {
        return window.outerHeight;
      }

      return window.innerHeight;
    }

    function handleResize() {
      setMobileScreenHeight(getMobileScreenHeight());
    }

    setMobileScreenHeight(getMobileScreenHeight());

    window.addEventListener("resize", handleResize);

    return () => {
      window.removeEventListener("resize", handleResize);
    };
  }, []);

  const adjustedHeight = mobileScreenHeight - 125;

  const debouncedHandleScroll = debounce(handleScroll, 100);

  return (
    <div className="App">
      <div className="container">
        <div className="download-selectFile">
          <label htmlFor="file-upload" className="file-select-button">
            <input
              type="file"
              id="file-upload"
              accept=".xls,.xlsx"
              onChange={handleFileUpload}
            />
            <span className="file-icon">
              <FaCloudUploadAlt />
            </span>
            <span className="file-text">
              {selectedFile ? selectedFile.name : "Select File"}
            </span>
          </label>
          {}
          <button onClick={undo} disabled={currentStep === 0}>
            <FaUndo />
          </button>
          <button onClick={redo} disabled={currentStep === history.length - 1}>
            <FaRedo />
          </button>
          <button onClick={reset} disabled={currentStep === 0}>
            <FaSync />
          </button>
          <button
            onClick={downloadExcel}
            disabled={jsonData?.length === 0}
            className="download-button-excel"
          >
            <FaDownload />
          </button>
        </div>
        {jsonData?.length > 0 && (
          <div className="json-table-container">
            <table
              className="json-table"
              onScroll={debouncedHandleScroll}
              style={{ height: `${adjustedHeight}px` }}
            >
              <thead>
                <tr>
                  <th className="top-left-corner"></th>
                  {jsonData[0].map((_, colIndex) => (
                    <th
                      key={colIndex}
                      draggable
                      onDragStart={(event) =>
                        handleDragStart(event, null, colIndex, "col")
                      }
                      onDragOver={handleDragOver}
                      onDrop={(event) => handleDrop(event, colIndex, "col")}
                    >
                      {isMobile ? (
                        <span className="mobile-screen">{colIndex + 1}</span>
                      ) : (
                        <span className="large-screen">
                          Column {colIndex + 1}
                        </span>
                      )}
                      <button
                        onClick={() => deleteColumn(colIndex)}
                        className="delete-column-button"
                      >
                        &times;
                      </button>
                    </th>
                  ))}
                  <th>
                    <button onClick={emptyCurrentSheet} className="empty-sheet">
                      <span className="icon">
                        <FaTimes />
                      </span>
                    </button>
                  </th>
                </tr>
              </thead>
              <tbody>
                {jsonData.map((row, rowIndex) => (
                  <tr
                    key={rowIndex}
                    draggable
                    onDragStart={(event) =>
                      handleDragStart(event, rowIndex, null, "row")
                    }
                    onDragOver={handleDragOver}
                    onDrop={(event) => handleDrop(event, rowIndex, "row")}
                  >
                    {isMobile ? (
                      <th className="mobile-screen">{rowIndex + 1}</th>
                    ) : (
                      <th className="large-screen">Row {rowIndex + 1}</th>
                    )}
                    {row.map((value, colIndex) => (
                      <td key={colIndex}>
                        <input
                          type="text"
                          className="Edit-input"
                          value={value}
                          onChange={(e) =>
                            handleCellChange(rowIndex, colIndex, e.target.value)
                          }
                        />
                      </td>
                    ))}
                    <td>
                      <button onClick={() => deleteRow(rowIndex)}>
                        <span className="icon">
                          <FaTimes />
                        </span>
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        {workbook && (
          <div className="sheet-tabs">
            {sheetNames.map((name, index) => (
              <div
                className={`sheet-tab ${
                  index === currentSheetIndex ? "active" : ""
                }`}
              >
                <button onClick={() => switchSheet(index)}>{name}</button>
                <button className="edit" onClick={() => editSheetName(index)}>
                  <FaEdit />
                </button>
                <button className="delete" onClick={() => deleteSheet(index)}>
                  <FaTrash />
                </button>
              </div>
            ))}
            <button onClick={addSheet} className="add-sheet">
              <FaPlus />
            </button>
          </div>
        )}
      </div>
    </div>
  );
}
