import { InputHTMLAttributes } from 'react'

interface InputProps extends InputHTMLAttributes<HTMLInputElement> {
  label?: string
  error?: string
}

export function Input({ label, error, id, className = '', ...props }: InputProps) {
  return (
    <div>
      {label && (
        <label
          htmlFor={id}
          className="block text-xs font-semibold tracking-wider uppercase text-navy mb-2"
        >
          {label}
          {props.required && <span className="text-gold ml-0.5">*</span>}
        </label>
      )}
      <input
        id={id}
        {...props}
        className={`w-full rounded-md border bg-[#F5F6F8] px-3.5 py-2.5 text-sm text-slate transition-all focus:outline-none focus:ring-2 focus:ring-teal focus:border-transparent ${
          error ? 'border-[#E05040]' : 'border-[#E2E8F0]'
        } ${className}`}
      />
      {error && <p className="text-xs text-[#E05040] mt-1.5">{error}</p>}
    </div>
  )
}
