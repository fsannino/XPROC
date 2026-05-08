export default function DashboardLoading() {
  return (
    <div className="animate-pulse space-y-6">
      <div className="space-y-3">
        <div className="h-3 bg-teal/10 rounded w-24" />
        <div className="h-8 bg-navy/10 rounded w-64" />
        <div className="h-3 bg-[#E2E8F0] rounded w-96 max-w-full" />
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-5">
        {[...Array(6)].map((_, i) => (
          <div
            key={i}
            className="bg-white rounded-lg border border-[#E2E8F0] border-t-4 border-t-[#E2E8F0] p-6"
          >
            <div className="w-11 h-11 rounded-lg bg-[#F5F6F8] mb-5" />
            <div className="h-7 bg-navy/10 rounded w-20 mb-2" />
            <div className="h-3 bg-[#E2E8F0] rounded w-32 mb-2" />
            <div className="h-3 bg-[#F5F6F8] rounded w-40" />
          </div>
        ))}
      </div>
    </div>
  )
}
