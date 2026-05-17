'use client'

import { useState } from 'react'
import Link from 'next/link'
import type { MaintenanceTicket, TicketMessage, ActivityLog, Vendor, TicketStatus } from '@/lib/db'
import { UrgencyBadge } from './urgency-badge'
import { formatDistanceToNow, format } from 'date-fns'
import { ChevronLeft, Zap, AlertTriangle, Clock, User, Building2, MessageSquare, Activity } from 'lucide-react'

const STATUS_OPTIONS: { value: TicketStatus; label: string }[] = [
  { value: 'new', label: 'New' },
  { value: 'assigned', label: 'Assigned' },
  { value: 'in_progress', label: 'In Progress' },
  { value: 'waiting', label: 'Waiting' },
  { value: 'completed', label: 'Completed' },
  { value: 'cancelled', label: 'Cancelled' },
]

interface TicketDetailProps {
  ticket: MaintenanceTicket
  messages: TicketMessage[]
  activityLogs: ActivityLog[]
  vendors: Vendor[]
  currentUserId: string
}

export function TicketDetail({ ticket: initial, messages: initialMessages, activityLogs, vendors, currentUserId }: TicketDetailProps) {
  const [ticket, setTicket] = useState(initial)
  const [messages, setMessages] = useState(initialMessages)
  const [newMessage, setNewMessage] = useState('')
  const [isInternal, setIsInternal] = useState(false)
  const [saving, setSaving] = useState(false)
  const [sendingMsg, setSendingMsg] = useState(false)
  const [activeTab, setActiveTab] = useState<'messages' | 'activity'>('messages')

  const updateTicket = async (changes: Partial<MaintenanceTicket>) => {
    setSaving(true)
    const updated = { ...ticket, ...changes }
    setTicket(updated)
    await fetch(`/api/tickets/${ticket.id}`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(changes),
    })
    setSaving(false)
  }

  const sendMessage = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!newMessage.trim()) return
    setSendingMsg(true)

    const res = await fetch(`/api/tickets/${ticket.id}/messages`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ message: newMessage, is_internal: isInternal, sender_type: 'manager' }),
    })

    if (res.ok) {
      const msg = await res.json()
      setMessages(prev => [...prev, msg])
      setNewMessage('')
    }
    setSendingMsg(false)
  }

  const assignedVendor = vendors.find(v => v.id === ticket.assigned_vendor_id)

  return (
    <div className="max-w-5xl">
      {/* Back */}
      <Link href="/dashboard/tickets" className="inline-flex items-center gap-1 text-sm text-gray-500 hover:text-gray-700 mb-4">
        <ChevronLeft className="w-4 h-4" />
        Back to tickets
      </Link>

      <div className="grid grid-cols-3 gap-6">
        {/* Main column */}
        <div className="col-span-2 space-y-4">
          {/* Header */}
          <div className="bg-white rounded-xl border border-gray-200 p-6">
            <div className="flex items-start gap-3 mb-4">
              <div className="flex-1">
                <h1 className="text-lg font-semibold text-gray-900 mb-2">{ticket.title}</h1>
                <div className="flex items-center gap-2 flex-wrap">
                  <UrgencyBadge urgency={ticket.urgency} />
                  {ticket.ai_category && (
                    <span className="text-xs bg-indigo-50 text-indigo-700 px-2 py-0.5 rounded-full capitalize">
                      <Zap className="w-3 h-3 inline mr-1" />
                      {ticket.ai_category.replace('_', ' ')}
                    </span>
                  )}
                  {ticket.ai_escalation_risk && (
                    <span className="text-xs bg-red-50 text-red-700 px-2 py-0.5 rounded-full flex items-center gap-1">
                      <AlertTriangle className="w-3 h-3" />
                      Escalation risk
                    </span>
                  )}
                </div>
              </div>
              {saving && <span className="text-xs text-gray-400">Saving...</span>}
            </div>

            {ticket.ai_summary && (
              <div className="bg-indigo-50 rounded-lg p-3 mb-4">
                <p className="text-xs font-medium text-indigo-600 mb-1">AI Summary</p>
                <p className="text-sm text-indigo-900">{ticket.ai_summary}</p>
                {ticket.ai_vendor_type && (
                  <p className="text-xs text-indigo-600 mt-1">Recommended: {ticket.ai_vendor_type}</p>
                )}
              </div>
            )}

            <div className="bg-gray-50 rounded-lg p-4">
              <p className="text-sm text-gray-700 whitespace-pre-wrap leading-relaxed">{ticket.description}</p>
            </div>
          </div>

          {/* Messages / Activity tabs */}
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden">
            <div className="border-b border-gray-100 flex">
              {[
                { id: 'messages', label: 'Messages', icon: MessageSquare, count: messages.length },
                { id: 'activity', label: 'Activity', icon: Activity, count: activityLogs.length },
              ].map(tab => (
                <button
                  key={tab.id}
                  onClick={() => setActiveTab(tab.id as 'messages' | 'activity')}
                  className={`flex items-center gap-2 px-4 py-3 text-sm font-medium border-b-2 transition-colors ${
                    activeTab === tab.id
                      ? 'border-indigo-600 text-indigo-600'
                      : 'border-transparent text-gray-500 hover:text-gray-700'
                  }`}
                >
                  <tab.icon className="w-4 h-4" />
                  {tab.label}
                  <span className="text-xs bg-gray-100 text-gray-500 px-1.5 py-0.5 rounded-full">{tab.count}</span>
                </button>
              ))}
            </div>

            {activeTab === 'messages' ? (
              <div>
                <div className="divide-y divide-gray-50 max-h-80 overflow-y-auto">
                  {messages.length === 0 ? (
                    <p className="p-6 text-sm text-gray-400 text-center">No messages yet</p>
                  ) : (
                    messages.map(msg => (
                      <div key={msg.id} className={`p-4 ${msg.is_internal ? 'bg-yellow-50' : ''}`}>
                        <div className="flex items-center gap-2 mb-1">
                          <span className="text-xs font-medium text-gray-700">{msg.sender_name || msg.sender_type}</span>
                          {msg.is_internal && (
                            <span className="text-xs bg-yellow-100 text-yellow-700 px-1.5 py-0.5 rounded-full">Internal</span>
                          )}
                          <span className="text-xs text-gray-400 ml-auto">
                            {format(new Date(msg.created_at), 'MMM d, h:mm a')}
                          </span>
                        </div>
                        <p className="text-sm text-gray-700">{msg.message}</p>
                      </div>
                    ))
                  )}
                </div>

                <form onSubmit={sendMessage} className="p-4 border-t border-gray-100">
                  <textarea
                    value={newMessage}
                    onChange={e => setNewMessage(e.target.value)}
                    placeholder="Add a note..."
                    rows={2}
                    className="w-full text-sm border border-gray-200 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-indigo-500 resize-none"
                    disabled={sendingMsg}
                  />
                  <div className="flex items-center justify-between mt-2">
                    <label className="flex items-center gap-2 text-xs text-gray-500 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={isInternal}
                        onChange={e => setIsInternal(e.target.checked)}
                        className="rounded"
                      />
                      Internal note (not visible to tenant/vendor)
                    </label>
                    <button
                      type="submit"
                      disabled={sendingMsg || !newMessage.trim()}
                      className="text-xs bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 text-white px-3 py-1.5 rounded-md"
                    >
                      {sendingMsg ? 'Sending...' : 'Send'}
                    </button>
                  </div>
                </form>
              </div>
            ) : (
              <div className="divide-y divide-gray-50 max-h-96 overflow-y-auto">
                {activityLogs.length === 0 ? (
                  <p className="p-6 text-sm text-gray-400 text-center">No activity yet</p>
                ) : (
                  activityLogs.map(log => (
                    <div key={log.id} className="px-4 py-3 flex items-start gap-3">
                      <div className="w-1.5 h-1.5 rounded-full bg-indigo-400 mt-2 shrink-0" />
                      <div className="flex-1 min-w-0">
                        <p className="text-sm text-gray-700">{log.description}</p>
                        <p className="text-xs text-gray-400 mt-0.5">
                          {log.user_name && <span>{log.user_name} · </span>}
                          {format(new Date(log.created_at), 'MMM d, h:mm a')}
                        </p>
                      </div>
                    </div>
                  ))
                )}
              </div>
            )}
          </div>
        </div>

        {/* Sidebar */}
        <div className="space-y-4">
          {/* Status */}
          <div className="bg-white rounded-xl border border-gray-200 p-4 space-y-4">
            <div>
              <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wide mb-2">Status</label>
              <select
                value={ticket.status}
                onChange={e => updateTicket({ status: e.target.value as TicketStatus })}
                className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white"
              >
                {STATUS_OPTIONS.map(opt => (
                  <option key={opt.value} value={opt.value}>{opt.label}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wide mb-2">Assigned Vendor</label>
              <select
                value={ticket.assigned_vendor_id || ''}
                onChange={e => updateTicket({ assigned_vendor_id: e.target.value || null })}
                className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white"
              >
                <option value="">Unassigned</option>
                {vendors.map(v => (
                  <option key={v.id} value={v.id}>
                    {v.name} ({v.trade_type})
                  </option>
                ))}
              </select>
            </div>
          </div>

          {/* Vendor info */}
          {assignedVendor && (
            <div className="bg-white rounded-xl border border-gray-200 p-4">
              <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wide mb-3">Vendor</label>
              <div className="space-y-2">
                <p className="text-sm font-medium text-gray-900">{assignedVendor.name}</p>
                <p className="text-xs text-gray-500 capitalize">{assignedVendor.trade_type}</p>
                {assignedVendor.phone && (
                  <a href={`tel:${assignedVendor.phone}`} className="text-xs text-indigo-600 hover:underline block">
                    {assignedVendor.phone}
                  </a>
                )}
                {assignedVendor.email && (
                  <a href={`mailto:${assignedVendor.email}`} className="text-xs text-indigo-600 hover:underline block">
                    {assignedVendor.email}
                  </a>
                )}
                <div className="flex items-center gap-1">
                  <span className="text-xs text-gray-500">Rating:</span>
                  <span className="text-xs font-medium">{assignedVendor.rating}/5</span>
                </div>
              </div>
            </div>
          )}

          {/* Tenant info */}
          {(ticket.tenant_name || ticket.tenant_email) && (
            <div className="bg-white rounded-xl border border-gray-200 p-4">
              <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wide mb-3">
                <User className="w-3 h-3 inline mr-1" />
                Tenant
              </label>
              <div className="space-y-1.5">
                {ticket.tenant_name && <p className="text-sm font-medium text-gray-900">{ticket.tenant_name}</p>}
                {ticket.tenant_email && (
                  <a href={`mailto:${ticket.tenant_email}`} className="text-xs text-indigo-600 hover:underline block">
                    {ticket.tenant_email}
                  </a>
                )}
                {ticket.tenant_phone && (
                  <a href={`tel:${ticket.tenant_phone}`} className="text-xs text-indigo-600 hover:underline block">
                    {ticket.tenant_phone}
                  </a>
                )}
              </div>
            </div>
          )}

          {/* Property */}
          {ticket.property_name && (
            <div className="bg-white rounded-xl border border-gray-200 p-4">
              <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wide mb-3">
                <Building2 className="w-3 h-3 inline mr-1" />
                Location
              </label>
              <p className="text-sm font-medium text-gray-900">{ticket.property_name}</p>
              {ticket.unit_number && <p className="text-xs text-gray-500">Unit {ticket.unit_number}</p>}
              {ticket.property_address && <p className="text-xs text-gray-400 mt-1">{ticket.property_address}</p>}
            </div>
          )}

          {/* Timestamps */}
          <div className="bg-white rounded-xl border border-gray-200 p-4">
            <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wide mb-3">
              <Clock className="w-3 h-3 inline mr-1" />
              Timeline
            </label>
            <div className="space-y-2 text-xs">
              <div className="flex justify-between">
                <span className="text-gray-500">Created</span>
                <span className="text-gray-700">{formatDistanceToNow(new Date(ticket.created_at), { addSuffix: true })}</span>
              </div>
              {ticket.assigned_at && (
                <div className="flex justify-between">
                  <span className="text-gray-500">Assigned</span>
                  <span className="text-gray-700">{formatDistanceToNow(new Date(ticket.assigned_at), { addSuffix: true })}</span>
                </div>
              )}
              {ticket.completed_at && (
                <div className="flex justify-between">
                  <span className="text-gray-500">Completed</span>
                  <span className="text-gray-700">{formatDistanceToNow(new Date(ticket.completed_at), { addSuffix: true })}</span>
                </div>
              )}
              {ticket.sla_due_at && (
                <div className="flex justify-between">
                  <span className="text-gray-500">SLA Due</span>
                  <span className={new Date(ticket.sla_due_at) < new Date() ? 'text-red-600 font-medium' : 'text-gray-700'}>
                    {formatDistanceToNow(new Date(ticket.sla_due_at), { addSuffix: true })}
                  </span>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}
