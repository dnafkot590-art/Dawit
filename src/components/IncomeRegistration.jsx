export default function IncomeRegistration({ form, banks, onChange, onSubmit }) {
  return (
    <form className="card" onSubmit={onSubmit}>
      <h3>1. Income Registration</h3>
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
          <option value="">Select category</option>
          <option value="Clinic income">Clinic income</option>
          <option value="Medicine sales">Medicine sales</option>
          <option value="Service fee">Service fee</option>
          <option value="Other income">Other income</option>
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
      <button className="btn" type="submit">Save Income</button>
    </form>
  );
}

