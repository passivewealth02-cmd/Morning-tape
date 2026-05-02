interface StarRatingProps {
  rating: 1 | 2 | 3
}

export function StarRating({ rating }: StarRatingProps) {
  const filled = '★'.repeat(rating)
  const empty = '☆'.repeat(3 - rating)
  
  return (
    <span className="star-rating text-sm">
      <span>{filled}</span>
      <span className="opacity-30">{empty}</span>
    </span>
  )
}
