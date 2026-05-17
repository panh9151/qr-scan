import { useEffect, useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import { useRef } from "react";

function App() {
  const [inputValue, setInputValue] = useState("");
  const [rows, setRows] = useState([]);
  const fileInputRef = useRef(null);
  const [locationValue, setLocationValue] = useState(
    localStorage.getItem("default-location") || "",
  );
  const [simplifyLocation, setSimplifyLocation] = useState(true);
  const [viewMode, setViewMode] = useState("table");

  const getDisplayLocation = (location) => {
    if (!simplifyLocation) return location;

    return location?.slice(-5) || "";
  };
  const toggleShelf = (shelfName) => {
    setExpandedShelves((prev) => {
      if (prev.includes(shelfName)) {
        return prev.filter((item) => item !== shelfName);
      }

      return [...prev, shelfName];
    });
  };

  const [expandedShelves, setExpandedShelves] = useState([]);

  const handleAddRow = () => {
    const splitData = inputValue.split("*");

    const newRow = {
      materialCode: splitData[0] || "",
      vendorCode: splitData[1] || "",
      lotNo: splitData[2] || "",
      quantity: splitData[3] || "",
      reelId: splitData[4] || "",
      location: splitData[7] || locationValue || "",
      qrCode: inputValue || "",
      done: false,
    };

    const updatedRows = [newRow, ...rows];

    setRows(updatedRows);

    setInputValue("");
  };

  const handleDeleteRow = (index) => {
    const updatedRows = rows.filter((_, i) => i !== index);

    setRows(updatedRows);
  };

  const handleChange = (index, field, value) => {
    const updatedRows = [...rows];

    updatedRows[index][field] = value;

    setRows(updatedRows);
  };

  const handleExportExcel = () => {
  const exportData = rows.map((row) => ({
    "Mã liệu": row.materialCode,
    "Mã NCC": row.vendorCode,
    "Lot no": row.lotNo,
    Quantity: row.quantity,
    "Reel ID": row.reelId,
    "Vị trí": row.location,

    // hide QR code
    // "QR Code": row.qrCode,

    // export checkbox-like value
    Done: row.done ? "TRUE" : "FALSE",
  }));

  const worksheet =
    XLSX.utils.json_to_sheet(exportData);

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "QR Data"
  );

  const excelBuffer = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array",
  });

  const data = new Blob([excelBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8",
  });

  // fixed filename
  saveAs(
    data,
    "storage-management.xlsx"
  );
};
  const handleImportExcel = (event) => {
    const file = event.target.files[0];

    if (!file) return;

    const reader = new FileReader();

    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);

      const workbook = XLSX.read(data, { type: "array" });

      const worksheet = workbook.Sheets[workbook.SheetNames[0]];

      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      const formattedData = jsonData.map((item) => ({
        materialCode: item["Mã liệu"] || "",
        vendorCode: item["Mã NCC"] || "",
        lotNo: item["Lot no"] || "",
        quantity: item["Quantity"] || "",
        reelId: item["Reel ID"] || "",
        location: item["Vị trí"] || "",
        qrCode: item["QR Code"] || "",
        done: item["Done"] || false,
      }));

      setRows(formattedData);
      setViewMode("table")
    };

    reader.readAsArrayBuffer(file);
  };

  const handleLocationChange = (value) => {
    setLocationValue(value);

    localStorage.setItem("default-location", value);
  };

  /* =========================
   SHELF VIEW DATA
========================= */

const groupedShelves = rows.reduce((acc, row) => {
  const shelf = row.location || "UNKNOWN";

  if (!acc[shelf]) {
    acc[shelf] = [];
  }

  acc[shelf].push(row);

  return acc;
}, {});

/* =========================
   HELPERS
========================= */

const groupMaterialsInShelf = (shelfRows) => {
  const grouped = {};

  shelfRows.forEach((row) => {
    const key = row.materialCode || "UNKNOWN";

    if (!grouped[key]) {
      grouped[key] = {
        materialCode: row.materialCode,
        vendorCode: row.vendorCode,
        totalQuantity: 0,
        reels: [],
        quantityMap: {},
      };
    }

    const qty = Number(row.quantity || 0);

    grouped[key].totalQuantity += qty;

    grouped[key].reels.push({
      reelId: row.reelId,
      quantity: qty,
    });

    if (!grouped[key].quantityMap[qty]) {
      grouped[key].quantityMap[qty] = 0;
    }

    grouped[key].quantityMap[qty] += 1;
  });

  return Object.values(grouped);
};

const getOtherShelves = (
  materialCode,
  currentShelf
) => {
  const sameMaterials = rows.filter(
    (item) =>
      item.materialCode === materialCode &&
      item.location &&
      item.location !== currentShelf
  );

  const grouped = sameMaterials.reduce(
    (acc, item) => {
      if (!acc[item.location]) {
        acc[item.location] = 0;
      }

      acc[item.location] += Number(
        item.quantity || 0
      );

      return acc;
    },
    {}
  );

  return Object.entries(grouped);
};

const formatGroupedQuantity = (
  totalQty,
  quantityMap
) => {
  const detail = Object.entries(quantityMap)
    .map(
      ([qty, count]) =>
        `${qty} pcs × ${count}`
    )
    .join(" + ");

  return `${totalQty.toLocaleString()} pcs (${detail})`;
};

  return (
    <div className="min-h-screen bg-slate-100 p-6">
      <div className="mx-auto max-w-7xl">
        {/* Header */}
        <div className="mb-6">
          <h1 className="text-3xl font-bold text-slate-800">
            QR Material Table
          </h1>

          <p className="mt-2 text-sm text-slate-500"></p>
        </div>

        <button
          onClick={() => fileInputRef.current.click()}
          className="
              rounded-xl bg-blue-600 px-5 py-3
              text-sm font-semibold text-white
              transition-all hover:bg-blue-700
              active:scale-[0.98]
            "
        >
          Import Excel
        </button>

        <button
          onClick={handleExportExcel}
          className="
              rounded-xl bg-emerald-600 px-5 py-3
              text-sm font-semibold text-white
              transition-all hover:bg-emerald-700
              active:scale-[0.98]
            "
        >
          Xuất Excel
        </button>

        {/* Input */}
        <div className="mb-6 rounded-2xl bg-white p-5 shadow-sm">
          <div className="flex flex-col gap-4 lg:flex-row">
            <input
              type="text"
              value={inputValue}
              onChange={(e) => setInputValue(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  handleAddRow();
                }
              }}
              placeholder="Input QR data..."
              className="
    h-12 flex-1 rounded-xl border border-slate-200
    px-4 text-sm outline-none transition-all
    focus:border-blue-500 focus:ring-4 focus:ring-blue-100
  "
            />
            <input
              type="text"
              value={locationValue}
              onChange={(e) => handleLocationChange(e.target.value)}
              placeholder="Default location..."
              className="
              h-12 rounded-xl border border-slate-200
              px-4 text-sm outline-none transition-all
              focus:border-blue-500 focus:ring-4 focus:ring-blue-100
            "
            />
            <button
              onClick={handleAddRow}
              className="
                h-12 rounded-xl bg-blue-600 px-6
                text-sm font-semibold text-white
                transition-all hover:bg-blue-700
                active:scale-[0.98]
              "
            >
              Add new
            </button>
          </div>
        </div>

        <div className="mb-4 flex flex-wrap items-center gap-6 rounded-2xl bg-white p-4 shadow-sm">
          <label className="flex items-center gap-2 text-sm font-medium text-slate-700">
            <input
              type="checkbox"
              checked={simplifyLocation}
              onChange={(e) => setSimplifyLocation(e.target.checked)}
              className="h-4 w-4"
            />
            Simplify location
          </label>

          <div className="flex items-center gap-4">
            <label className="flex items-center gap-2 text-sm font-medium text-slate-700">
              <input
                type="radio"
                name="view-mode"
                checked={viewMode === "table"}
                onChange={() => setViewMode("table")}
                className="h-4 w-4"
              />
              Table view
            </label>

            <label className="flex items-center gap-2 text-sm font-medium text-slate-700">
              <input
                type="radio"
                name="view-mode"
                checked={viewMode === "shelf"}
                onChange={() => setViewMode("shelf")}
                className="h-4 w-4"
              />
              Location view
            </label>
          </div>
        </div>

        {/* Table */}
        {viewMode === "table" && (
          <div className="overflow-hidden rounded-2xl bg-white shadow-sm">
            <div className="overflow-x-auto">
              <table className="min-w-full table-fixed border-collapse">
                <thead className="bg-slate-50">
                  <tr>
                    <th className="w-[70px] px-2 py-1 text-center text-sm font-semibold text-slate-600">
                      No.
                    </th>

                    <th className="px-2 py-1 text-left text-sm font-semibold text-slate-600">
                      Material Code
                    </th>

                    <th className="px-2 py-1 text-left text-sm font-semibold text-slate-600">
                      Vendor Code
                    </th>

                    <th className="px-2 py-1 text-left text-sm font-semibold text-slate-600">
                      Lot No
                    </th>

                    <th className="px-2 py-1 text-left text-sm font-semibold text-slate-600">
                      Quantity
                    </th>

                    <th className="px-2 py-1 text-left text-sm font-semibold text-slate-600">
                      Reel ID
                    </th>

                    <th className="px-2 py-1 text-left text-sm font-semibold text-slate-600">
                      Location
                    </th>

                    <th className="px-2 py-1 text-left text-sm font-semibold text-slate-600">
                      QR Code
                    </th>

                    <th className="w-[90px] px-2 py-1 text-center text-sm font-semibold text-slate-600">
                      Done
                    </th>
                    <th className="px-2 py-1 text-left text-sm font-semibold text-slate-600"></th>
                  </tr>
                </thead>

                <tbody>
                  {rows.map((row, index) => (
                    <tr
                      key={index}
                      className="border-t border-slate-100 hover:bg-slate-50"
                    >
                      <td className="px-2 py-1 text-center text-sm font-medium text-slate-700">
                        {index + 1}
                      </td>

                      <td className="px-2 py-1">
                        <input
                          value={row.materialCode}
                          onChange={(e) =>
                            handleChange(index, "materialCode", e.target.value)
                          }
                          className="
                          w-full rounded-lg border border-slate-200
                          bg-white px-3 py-2 text-sm
                          outline-none transition-all
                          focus:border-blue-500
                          focus:ring-4 focus:ring-blue-100
                        "
                        />
                      </td>

                      <td className="px-2 py-1">
                        <input
                          value={row.vendorCode}
                          onChange={(e) =>
                            handleChange(index, "vendorCode", e.target.value)
                          }
                          className="
                          w-full rounded-lg border border-slate-200
                          bg-white px-3 py-2 text-sm
                          outline-none transition-all
                          focus:border-blue-500
                          focus:ring-4 focus:ring-blue-100
                        "
                        />
                      </td>

                      <td className="px-2 py-1">
                        <input
                          value={row.quantity}
                          onChange={(e) =>
                            handleChange(index, "quantity", e.target.value)
                          }
                          className="
                          w-full rounded-lg border border-slate-200
                          bg-white px-3 py-2 text-sm
                          outline-none transition-all
                          focus:border-blue-500
                          focus:ring-4 focus:ring-blue-100
                        "
                        />
                      </td>

                      <td className="px-2 py-1">
                        <input
                          value={row.lotNo}
                          onChange={(e) =>
                            handleChange(index, "lotNo", e.target.value)
                          }
                          className="
                          w-full rounded-lg border border-slate-200
                          bg-white px-3 py-2 text-sm
                          outline-none transition-all
                          focus:border-blue-500
                          focus:ring-4 focus:ring-blue-100
                        "
                        />
                      </td>

                      <td className="px-2 py-1">
                        <input
                          value={row.reelId}
                          onChange={(e) =>
                            handleChange(index, "reelId", e.target.value)
                          }
                          className="
                          w-full rounded-lg border border-slate-200
                          bg-white px-3 py-2 text-sm
                          outline-none transition-all
                          focus:border-blue-500
                          focus:ring-4 focus:ring-blue-100
                        "
                        />
                      </td>

                      <td className="px-2 py-1">
                        <input
                          value={getDisplayLocation(row.location)}
                          onChange={(e) =>
                            handleChange(index, "location", e.target.value)
                          }
                          className="
                          w-full rounded-lg border border-slate-200
                          bg-white px-3 py-2 text-sm
                          outline-none transition-all
                          focus:border-blue-500
                          focus:ring-4 focus:ring-blue-100
                        "
                        />
                      </td>

                      <td className="px-2 py-1">
                        <input
                          value={row.qrCode}
                          onChange={(e) =>
                            handleChange(index, "qrCode", e.target.value)
                          }
                          className="
                          w-full rounded-lg border border-slate-200
                          bg-white px-3 py-2 text-sm
                          outline-none transition-all
                          focus:border-blue-500
                          focus:ring-4 focus:ring-blue-100
                        "
                        />
                      </td>

                      <td className="px-2 py-1 text-center">
                        <input
                          type="checkbox"
                          checked={row.done}
                          onChange={(e) =>
                            handleChange(index, "done", e.target.checked)
                          }
                          className="h-5 w-5 cursor-pointer"
                        />
                      </td>
                      <td className="w-[80px] px-2 py-1">
                        <button
                          onClick={() => handleDeleteRow(index)}
                          className="
                          rounded-md bg-red-100 px-2 py-1
                          text-xs font-medium text-red-600
                          transition-all hover:bg-red-200
                          "
                        >
                          Delete
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="mt-4 flex justify-end"></div>
          </div>
        )}
        {viewMode === "shelf" && (
  <div className="space-y-4">

    {/* Shelf list */}
    <div className="space-y-4">
      {Object.entries(groupedShelves).map(
        ([shelfName, shelfRows]) => {
          const groupedMaterials =
            groupMaterialsInShelf(shelfRows);

          return (
            <div
              key={shelfName}
              className="
                rounded-2xl bg-white
                p-4 shadow-sm
              "
            >
              {/* Shelf header */}
              <button
                onClick={() =>
                  toggleShelf(shelfName)
                }
                className="
                  flex w-full items-center
                  justify-between rounded-xl
                  border border-slate-200
                  bg-slate-50 px-4 py-3
                  transition-all duration-200
                  hover:bg-slate-100
                "
              >
                <div className="flex items-center gap-2">
                  <div
                    className="
                      text-sm font-bold
                      text-slate-800
                    "
                  >
                    {shelfName}
                  </div>

                  <div
                    className="
                      rounded-md bg-white
                      px-2 py-1 text-[11px]
                      font-medium text-slate-600
                    "
                  >
                    {
                      groupedMaterials.length
                    }{" "}
                    materials
                  </div>
                </div>

                <div
                  className={`
                    text-xs text-slate-500
                    transition-transform duration-300
                    ${
                      expandedShelves.includes(
                        shelfName
                      )
                        ? "rotate-180"
                        : ""
                    }
                  `}
                >
                  ▼
                </div>
              </button>

              {/* Expand */}
              <div
                className={`
                  overflow-hidden
                  transition-all duration-300 ease-in-out
                  ${
                    expandedShelves.includes(
                      shelfName
                    )
                      ? "mt-3 max-h-[3000px] opacity-100"
                      : "max-h-0 opacity-0"
                  }
                `}
              >
                <div
                  className="
                    overflow-hidden rounded-xl
                    border border-slate-200
                  "
                >
                  <table className="min-w-full border-collapse">
                    <thead className="bg-slate-100">
                      <tr>
                        <th
                          className="
                            px-3 py-2 text-left
                            text-xs font-semibold
                            text-slate-600
                          "
                        >
                          Material / Reel ID
                        </th>

                        <th
                          className="
                            w-[320px]
                            px-3 py-2 text-left
                            text-xs font-semibold
                            text-slate-600
                          "
                        >
                          Quantity
                        </th>

                        <th
                          className="
                            w-[120px]
                            px-3 py-2 text-left
                            text-xs font-semibold
                            text-slate-600
                          "
                        >
                          Vendor
                        </th>

                        <th
                          className="
                            px-3 py-2 text-left
                            text-xs font-semibold
                            text-slate-600
                          "
                        >
                          Other Shelves
                        </th>
                      </tr>
                    </thead>

                    <tbody>
                      {groupedMaterials.map(
                        (
                          material,
                          index
                        ) => {
                          const otherShelves =
                            getOtherShelves(
                              material.materialCode,
                              shelfName
                            );

                          return (
                            <tr
                              key={index}
                              className="
                                border-t
                                border-slate-100
                              "
                            >
                              {/* Material */}
                              <td className="px-3 py-2">
                                <div className="space-y-1">
                                  <div
                                    className="
                                      text-xs
                                      font-semibold
                                      text-slate-800
                                    "
                                  >
                                    {
                                      material.materialCode
                                    }
                                  </div>

                                  <div
                                    className="
                                      flex flex-wrap gap-1
                                    "
                                  >
                                    {material.reels.map(
                                      (
                                        reel,
                                        reelIndex
                                      ) => (
                                        <div
                                          key={
                                            reelIndex
                                          }
                                          className="
                                            rounded-md
                                            bg-slate-100
                                            px-2 py-[2px]
                                            text-[10px]
                                            text-slate-500
                                          "
                                        >
                                          {
                                            reel.reelId
                                          }
                                        </div>
                                      )
                                    )}
                                  </div>
                                </div>
                              </td>

                              {/* Quantity */}
                              <td
                                className="
                                  px-3 py-2
                                  text-xs
                                  text-slate-700
                                "
                              >
                                <div className="space-y-1">
                                  <div className="font-semibold">
                                    {material.totalQuantity.toLocaleString()}{" "}
                                    pcs
                                  </div>

                                  <div
                                    className="
                                      text-[11px]
                                      text-slate-400
                                    "
                                  >
                                    {formatGroupedQuantity(
                                      material.totalQuantity,
                                      material.quantityMap
                                    )}
                                  </div>
                                </div>
                              </td>

                              {/* Vendor */}
                              <td
                                className="
                                  px-3 py-2
                                  text-xs
                                  text-slate-500
                                "
                              >
                                {
                                  material.vendorCode
                                }
                              </td>

                              {/* Other shelves */}
                              <td className="px-3 py-2">
                                <div
                                  className="
                                    flex flex-wrap gap-2
                                  "
                                >
                                  {otherShelves.length >
                                  0 ? (
                                    otherShelves.map(
                                      (
                                        [
                                          otherShelf,
                                          qty,
                                        ],
                                        idx
                                      ) => (
                                        <div
                                          key={idx}
                                          className="
                                            rounded-md
                                            bg-amber-50
                                            px-2 py-1
                                            text-[11px]
                                            text-amber-700
                                          "
                                        >
                                          {
                                            otherShelf
                                          }{" "}
                                          ({qty})
                                        </div>
                                      )
                                    )
                                  ) : (
                                    <div
                                      className="
                                        text-[11px]
                                        text-slate-400
                                      "
                                    >
                                      No duplicate
                                    </div>
                                  )}
                                </div>
                              </td>
                            </tr>
                          );
                        }
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          );
        }
      )}
    </div>
  </div>
)}
        <div className="mt-4 flex justify-end gap-3">
          <input
            type="file"
            accept=".xlsx,.xls"
            ref={fileInputRef}
            onChange={handleImportExcel}
            className="hidden"
          />
        </div>
      </div>
    </div>
  );
}

export default App;

// Mã QR phải rõ nét,
// Import hàng loạt
// Thêm mới hàng trống
// Lưu lại qr code
// check box
// Lưu vị trí mặc định
// bỏ lưu trê localstorage

// Thống kê scroll doơn
// Tên file
// Search