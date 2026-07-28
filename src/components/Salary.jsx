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
    const total = salaryNum - pension + dutyNum;

    setEmployees([
      ...employees,
      { id: Date.now(), name, basicSalary: salaryNum, duty: dutyNum, pension, total },
    ]);

    setName("");
    setBasicSalary("");
    setDuty("");
  };

  return (
    <div style={{ padding: "20px" }}>
      <h2>የሰራተኞች ደሞዝ (Salary)</h2>

      <form onSubmit={handleAdd} style={{ marginBottom: "20px", display: "flex", gap: "10px" }}>
        <input
          placeholder="የሰራተኛ ስም"
          value={name}
          onChange={(e) => setName(e.target.value)}
          required
        />
        <input
          type="number"
          placeholder="የወር ደሞዝ"
          value={basicSalary}
          onChange={(e) => setBasicSalary(e.target.value)}
          required
        />
        <input
          type="number"
          placeholder="ዲውቲ"
          value={duty}
          onChange={(e) => setDuty(e.target.value)}
        />
        <button type="submit">መዝግብ</button>
      </form>

      <table border="1" cellPadding="8" style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ background: "#f0f0f0" }}>
            <th>ስም</th>
            <th>ደሞዝ</th>
            <th>ዲውቲ</th>
            <th>7% ጡረታ</th>
            <th>Total (የሚከፈል)</th>
          </tr>
        </thead>
        <tbody>
          {employees.map((emp) => (
            <tr key={emp.id}>
              <td>{emp.name}</td>
              <td>{emp.basicSalary} ETB</td>
              <td>{emp.duty} ETB</td>
              <td style={{ color: "red" }}>-{emp.pension.toFixed(2)} ETB</td>
              <td style={{ fontWeight: "bold", color: "green" }}>{emp.total.toFixed(2)} ETB</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
