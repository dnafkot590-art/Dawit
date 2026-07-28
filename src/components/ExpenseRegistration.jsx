export default function ExpenseRegistration({ form, banks, onChange, onSubmit }) {
  return (
    <form className="card" onSubmit={onSubmit}>
      <h3>2. Expense Registration</h3>
      <div className="form-grid">
        <input
          type="date"
          value={form.date}
          onChange={(e) => onChange("date", e.target.value)}
        />
        <select
          value={form.bankId}
          onChange={(e) => onChange("bankId", e.target.value)}
          required
        >
          <option value="">Select Bank Account</option>
          {banks.map(bank => (
            <option key={bank.id} value={bank.id}>{bank.name} ({bank.accountNumber})</option>
          ))}
        </select>
        <select
          value={form.category}
          onChange={(e) => onChange("category", e.target.value)}
        >
          <option value="Electricity">Electricity</option>
          <option value="Water">Water</option>
          <option value="Waste">Waste</option>
          <option value="Rent">Rent</option>
          <option value="Salary">Salary</option>
          <option value="Telecom">Telecom</option>
          <option value="Pharmacy">Pharmacy</option>
          <option value="Lab">Lab</option>
          <option value="Medicine">Medicine</option>
          <option value="Misc Expenses">Misc Expenses</option>
        </select>
        <input
          type="number"
          placeholder="Amount"
          value={form.amount}
          onChange={(e) => onChange("amount", e.target.value)}
        />
        <textarea
          placeholder="Notes"
          value={form.notes}
          onChange={(e) => onChange("notes", e.target.value)}
        />
      </div>
      <button className="btn" type="submit">Save Expense</button>
    </form>
  );
}
