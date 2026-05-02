export function Footer() {
  const currentYear = new Date().getFullYear()
  
  return (
    <footer className="py-12 border-t border-border">
      <div className="max-w-4xl mx-auto px-6">
        <div className="rule-double pb-8 mb-8" />
        
        <div className="text-center">
          <p className="font-serif text-xl tracking-wider uppercase mb-4">
            The Morning Tape
          </p>
          <p className="text-sm text-muted-foreground mb-6">
            Daily market intelligence, delivered.
          </p>
          
          <div className="flex justify-center gap-8 text-sm text-muted-foreground mb-8">
            <a href="#" className="editorial-link">About</a>
            <a href="#" className="editorial-link">Contact</a>
            <a href="#" className="editorial-link">Privacy</a>
            <a href="#" className="editorial-link">Terms</a>
          </div>
          
          <p className="text-xs text-muted-foreground">
            &copy; {currentYear} The Morning Tape. All rights reserved.
          </p>
          <p className="text-xs text-muted-foreground mt-2">
            Market data is for informational purposes only and does not constitute financial advice.
          </p>
        </div>
      </div>
    </footer>
  )
}
