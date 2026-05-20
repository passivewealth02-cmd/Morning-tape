/** Maintena rounded-M monogram. Uses currentColor so it inherits text color. */
export function MaintenaMark({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 120 120" fill="none" className={className} aria-hidden="true">
      <path
        d="M 24 96 L 24 30 L 60 70 L 96 30 L 96 96"
        stroke="currentColor"
        strokeWidth={16}
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

/** Boxed logo lockup: indigo tile with the white monogram, plus optional wordmark. */
export function MaintenaLogo({
  className = '',
  wordmark = true,
  tileClassName = 'w-7 h-7',
  textClassName = 'text-sm font-semibold text-gray-900',
}: {
  className?: string
  wordmark?: boolean
  tileClassName?: string
  textClassName?: string
}) {
  return (
    <span className={`inline-flex items-center gap-2 ${className}`}>
      <span className={`rounded-md bg-indigo-600 flex items-center justify-center ${tileClassName}`}>
        <MaintenaMark className="w-1/2 h-1/2 text-white" />
      </span>
      {wordmark && <span className={textClassName}>Maintena</span>}
    </span>
  )
}
