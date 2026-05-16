import { useEffect, useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

function App() {
  const [inputValue, setInputValue] = useState("");
  const [rows, setRows] = useState([]);

  useEffect(() => {
    const storedData = localStorage.getItem("qr-data");

    if (storedData) {
      setRows(JSON.parse(storedData));
    }
  }, []);

  const saveToLocalStorage = (data) => {
    localStorage.setItem("qr-data", JSON.stringify(data));
  };

  const handleAddRow = () => {
    if (!inputValue.trim()) return;

    const splitData = inputValue.split("*");

    const newRow = {
      materialCode: splitData[0] || "",
      vendorCode: splitData[1] || "",
      lotNo: splitData[2] || "",
      quantity: splitData[3] || "",
      reelId: splitData[4] || "",
      location: splitData[7] || "",
      rawQr: inputValue,
    };

    const updatedRows = [newRow, ...rows];

    setRows(updatedRows);

    saveToLocalStorage(updatedRows);

    setInputValue("");
  };

  const handleDeleteRow = (index) => {
    const updatedRows = rows.filter((_, i) => i !== index);

    setRows(updatedRows);

    saveToLocalStorage(updatedRows);
  };

  const handleChange = (index, field, value) => {
    const updatedRows = [...rows];

    updatedRows[index][field] = value;

    setRows(updatedRows);

    saveToLocalStorage(updatedRows);
  };
  
  const handleExportExcel = () => {
  const exportData = rows.map((row) => ({
    "Mã liệu": row.materialCode,
    "Mã NCC": row.vendorCode,
    "Lot no": row.lotNo,
    Quantity: row.quantity,
    "Reel ID": row.reelId,
    "Vị trí": row.location,
  }));

  const worksheet = XLSX.utils.json_to_sheet(exportData);

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workbook, worksheet, "QR Data");

  const excelBuffer = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array",
  });

  const data = new Blob([excelBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8",
  });

  saveAs(data, `qr-data-${Date.now()}.xlsx`);
};

  return (
    <div className="min-h-screen bg-slate-100 p-6">
      <div className="mx-auto max-w-7xl">
        {/* Header */}
        <div className="mb-6">
          <h1 className="text-3xl font-bold text-slate-800">
            QR Material Table
          </h1>

          <p className="mt-2 text-sm text-slate-500">
            Editable table bằng React + Tailwind
          </p>
        </div>

        {/* Input */}
        <div className="mb-6 rounded-2xl bg-white p-5 shadow-sm">
          <div className="flex flex-col gap-4 md:flex-row">
            <input
              type="text"
              value={inputValue}
              onChange={(e) => setInputValue(e.target.value)}
              placeholder="Scan hoặc nhập dữ liệu QR..."
              className="
                h-12 flex-1 rounded-xl border border-slate-200
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
              Thêm mới
            </button>
          </div>
        </div>

        {/* Table */}
        <div className="overflow-hidden rounded-2xl bg-white shadow-sm">
          <div className="overflow-x-auto">
            <table className="min-w-full table-fixed border-collapse">
              <thead className="bg-slate-50">
                <tr>
                  <th className="w-[80px] px-2 py-1"></th>

                  <th className="px-2 py-1 text-left text-sm font-semibold text-slate-600">
                    Mã liệu
                  </th>

                  <th className="px-2 py-1 text-left text-sm font-semibold text-slate-600">
                    Mã NCC
                  </th>

                  

                  <th className="px-2 py-1 text-left text-sm font-semibold text-slate-600">
                    Quantity
                  </th>

                  <th className="px-2 py-1 text-left text-sm font-semibold text-slate-600">
                    Reel ID
                  </th>

                  <th className="px-2 py-1 text-left text-sm font-semibold text-slate-600">
                    Mã lô
                  </th>

                  <th className="px-2 py-1 text-left text-sm font-semibold text-slate-600">
                    Vị trí
                  </th>
                </tr>
              </thead>

              <tbody>
                {rows.map((row, index) => (
                  <tr
                    key={index}
                    className="border-t border-slate-100 hover:bg-slate-50"
                  >
                    <td className="w-[80px] px-2 py-1">
                      <button
                        onClick={() => handleDeleteRow(index)}
                        className="
                          rounded-md bg-red-100 px-2 py-1
                          text-xs font-medium text-red-600
                          transition-all hover:bg-red-200
                        "
                      >
                        Xóa
                      </button>
                    </td>

                    <td className="px-2 py-1">
                      <input
                        value={row.materialCode}
                        onChange={(e) =>
                          handleChange(
                            index,
                            "materialCode",
                            e.target.value
                          )
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
                        value={row.location}
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
                    
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="mt-4 flex justify-end">
</div>
        </div>
  <button
    onClick={handleExportExcel}
    className="
      rounded-xl bg-emerald-600 px-5 py-3
      text-sm font-semibold text-white
      transition-all hover:bg-emerald-700
      active:scale-[0.98] mt-4
    "
  >
    Xuất Excel
  </button>
      </div>
    </div>
  );
}

export default App;

// Import hàng loạt 
// Đồng bộ
// Mã QR phải rõ nét, 