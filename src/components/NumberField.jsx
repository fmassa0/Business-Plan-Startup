export default function NumberField({ label, name, value, step = 1, onChange }) {
  return (
    <div>
      <label htmlFor={name}>{label}</label>
      <input
        id={name}
        type="number"
        step={step}
        value={value ?? 0}
        onChange={(e) => onChange(name, e.target.value)}
      />
    </div>
  );
}
