interface SectionDividerProps {
  section: string
}

export function SectionDivider({ section }: SectionDividerProps) {
  return (
    <div className="my-8 flex items-center justify-center">
      <div className="flex-1 rule-single" />
      <span className="section-marker px-6">
        {section}
      </span>
      <div className="flex-1 rule-single" />
    </div>
  )
}
