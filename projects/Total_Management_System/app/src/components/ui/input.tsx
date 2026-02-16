import { cn } from '@/lib/utils/cn';
import { forwardRef, type InputHTMLAttributes } from 'react';

interface InputProps extends InputHTMLAttributes<HTMLInputElement> {
  label?: string;
  error?: string;
}

export const Input = forwardRef<HTMLInputElement, InputProps>(
  ({ className, label, error, id, ...props }, ref) => (
    <div>
      {label && (
        <label htmlFor={id} className="block text-sm font-medium text-neutral-700 mb-1">
          {label}
        </label>
      )}
      <input
        ref={ref}
        id={id}
        className={cn(
          'w-full h-10 px-3 rounded-lg border bg-warm-ivory text-indigo-black placeholder:text-neutral-400',
          'focus:outline-none focus:ring-2 focus:ring-terracotta/40 focus:border-terracotta transition',
          error ? 'border-error' : 'border-neutral-200',
          className
        )}
        {...props}
      />
      {error && <p className="mt-1 text-xs text-error">{error}</p>}
    </div>
  )
);
Input.displayName = 'Input';
