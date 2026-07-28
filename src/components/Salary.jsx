/**
 * Salary.jsx — standalone payroll calculator utility
 * NOTE: The main app uses the integrated Payroll system in App.jsx (HRM → Payroll tab).
 * This component is kept as a reference/utility only and is NOT imported by App.jsx.
 */
import { useState } from "react";

export default function Salary() {
  const [employees, setEmployees] = useState([]);
  const [name, setName] = useState("");
  const [basicSalary, setBasicSalary] = useState("");
  const [duty, setDuty] = useState("");

  const handleAdd = (e) => {
    e.preventDefault();
    if (!name || !basicSalary) return;

    const salaryNum = parseFloat(basicSalary) || 0;
    const dutyNum = parseFloat(duty) || 0;
    const pension = salaryNum * 0.07;
    const net = salaryNum - pension + dutyNum;

    setEmployees(prev => [
      ...prev,
      { id: Date.now(), name, basicSalary: salaryNum, duty: dutyNum, pension, net },
    ]);

    setName("");
    setBasicSalary("");
    setDuty("");
  };

  const handleDelete = (id) => {
    setEmployees(prev => prev.filter(e => e.id !== id));
  };

  const totalNet = employees.reduce((s, e) => s + e.net, 0);

  return (
    <div style={{ padding: "20px", fontFamily: "Arial, sans-serif" }}>
      <h2 style={{ marginBottom: 16 }}>የሰራተኞች ደሞዝ ሂሳብ (Salary Calculator)</h2>

      <form
        onSubmit={handleAdd}
        style={{ marginBottom: "20px", display: "flex", gap: "10px", flexWrap: "wrap" }}
      >
        <input
          placeholder="የሰራተኛ ስም *"
          value={name}
          onChange={(e) => setName(e.target.value)}
          required
          style={{ padding: "8px 12px", borderRadius: 6, border: "1px solid #cbd5e1", minWidth: 160 }}
        />
        <input
          type="number"
          placeholder="የወር ደሞዝ (Birr) *"
          value={basicSalary}
          onChange={(e) => setBasicSalary(e.target.value)}
          required
          min={0}
          style={{ padding: "8px 12px", borderRadius: 6, border: "1px solid #cbd5e1", width: 140 }}
        />
        <input
          type="number"
          placeholder="ዲውቲ (Birr)"
          value={duty}
          onChange={(e) => setDuty(e.target.value)}
          min={0}
          style={{ padding: "8px 12px", borderRadius: 6, border: "1px solid #cbd5e1", width: 120 }}
        />
        <button
          type="submit"
          style={{
            padding: "8px 18px", borderRadius: 6, border: "none",
            background: "#2563eb", color: "white", fontWeight: 700, cursor: "pointer",
          }}
        >
          + መዝግብ
        </button>
      </form>

      {employees.length === 0 ? (
        <p style={{ color: "#94a3b8" }}>ምንም ሰራተኛ አልተጨመረም።</p>
      ) : (
        <>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 14 }}>
            <thead>
              <tr style={{ background: "#1e3a5f", color: "white" }}>
                <th style={{ padding: "10px 12px", textAlign: "left" }}>ስም</th>
                <th style={{ padding: "10px 12px", textAlign: "right" }}>ደሞዝ</th>
                <th style={{ padding: "10px 12px", textAlign: "right" }}>ዲውቲ</th>
                <th style={{ padding: "10px 12px", textAlign: "right", color: "#fca5a5" }}>7% ጡረታ</th>
                <th style={{ padding: "10px 12px", textAlign: "right", color: "#4ade80" }}>የሚከፈል (Net)</th>
                <th style={{ padding: "10px 12px" }}></th>
              </tr>
            </thead>
            <tbody>
              {employees.map((emp, i) => (
                <tr key={emp.id} style={{ background: i % 2 === 0 ? "#f8fafc" : "white", borderBottom: "1px solid #e2e8f0" }}>
                  <td style={{ padding: "9px 12px", fontWeight: 600 }}>{emp.name}</td>
                  <td style={{ padding: "9px 12px", textAlign: "right" }}>{emp.basicSalary.toLocaleString()} ETB</td>
                  <td style={{ padding: "9px 12px", textAlign: "right" }}>{emp.duty.toLocaleString()} ETB</td>
                  <td style={{ padding: "9px 12px", textAlign: "right", color: "#ef4444" }}>
                    -{emp.pension.toFixed(2)} ETB
                  </td>
                  <td style={{ padding: "9px 12px", textAlign: "right", fontWeight: 700, color: "#16a34a" }}>
                    {emp.net.toFixed(2)} ETB
                  </td>
                  <td style={{ padding: "9px 12px", textAlign: "center" }}>
                    <button
                      onClick={() => handleDelete(emp.id)}
                      style={{ background: "#dc2626", color: "white", border: "none", borderRadius: 4, padding: "4px 10px", cursor: "pointer", fontSize: 12 }}
                    >
                      ሰርዝ
                    </button>
                  </td>
                </tr>
              ))}
              {/* Totals row */}
              <tr style={{ background: "#0f172a", color: "white", fontWeight: 700 }}>
                <td colSpan={4} style={{ padding: "10px 12px" }}>ጠቅላላ ({employees.length} ሰራተኞች)</td>
                <td style={{ padding: "10px 12px", textAlign: "right", color: "#4ade80", fontSize: 16 }}>
                  {totalNet.toFixed(2)} ETB
                </td>
                <td></td>
              </tr>
            </tbody>
          </table>
          <p style={{ fontSize: 12, color: "#94a3b8", marginTop: 8 }}>
            ⚠️ ይህ ጊዚያዊ ሂሳቢያ ነው። ለቋሚ ፔሮል HRM → Payroll ይጠቀሙ።
          </p>
        </>
      )}
    </div>
  );
}
