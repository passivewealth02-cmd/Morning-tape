interface SparklineProps {
  data: number[]
  width?: number
  height?: number
  trend: 'up' | 'down'
}

export function Sparkline({ data, width = 60, height = 20, trend }: SparklineProps) {
  if (!data || data.length < 2) {
    return null
  }

  const min = Math.min(...data)
  const max = Math.max(...data)
  const range = max - min || 1

  // Normalize data to fit within height
  const normalizedData = data.map(
    (value) => height - ((value - min) / range) * height
  )

  // Create SVG path
  const stepX = width / (data.length - 1)
  const pathData = normalizedData
    .map((y, i) => {
      const x = i * stepX
      return i === 0 ? `M ${x} ${y}` : `L ${x} ${y}`
    })
    .join(' ')

  return (
    <svg
      width={width}
      height={height}
      viewBox={`0 0 ${width} ${height}`}
      className="inline-block"
    >
      <path
        d={pathData}
        fill="none"
        strokeWidth={1.5}
        strokeLinecap="round"
        strokeLinejoin="round"
        className={trend === 'up' ? 'sparkline-up' : 'sparkline-down'}
      />
    </svg>
  )
}
