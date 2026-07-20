'use client'

import Sidebar from './Sidebar'
import Topbar from './Topbar'

export default function MainWrapper({ children }: { children: React.ReactNode }) {
  return (
    <div className="min-h-screen bg-slate-50 flex flex-col">
      {/* Right Sidebar on Desktop */}
      <Sidebar />

      {/* Content wrapper with RTL offsets */}
      <div className="flex-1 flex flex-col lg:mr-64 min-w-0">
        <Topbar />
        
        <main className="flex-1 p-4 sm:p-6 lg:p-8 space-y-6 max-w-7xl w-full mx-auto">
          {children}
        </main>
      </div>
    </div>
  )
}
