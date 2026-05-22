/**
 * db.js — All Supabase database operations for CPE Refurb Manager
 */
import { supabase } from './supabase'

// ─────────────────────────────────────────────────────────────────────────────
// DEVICES
// ─────────────────────────────────────────────────────────────────────────────

export async function fetchDevices() {
  return supabase
    .from('devices')
    .select('*')
    .order('created_at', { ascending: false })
}

export async function createDevice(device) {
  return supabase.from('devices').insert([device]).select().single()
}

export async function updateDevice(id, updates) {
  return supabase
    .from('devices')
    .update({ ...updates, updated_at: new Date().toISOString() })
    .eq('id', id)
    .select()
    .single()
}

export async function deleteDevice(id) {
  return supabase.from('devices').delete().eq('id', id)
}

export async function bulkInsertDevices(devices) {
  return supabase.from('devices').insert(devices).select()
}

// ─────────────────────────────────────────────────────────────────────────────
// USERS
// ─────────────────────────────────────────────────────────────────────────────

export async function fetchUsers() {
  return supabase.from('users').select('*').order('created_at', { ascending: true })
}

export async function createUser(user) {
  return supabase.from('users').insert([user]).select().single()
}

export async function updateUser(id, updates) {
  return supabase.from('users').update(updates).eq('id', id).select().single()
}

export async function deleteUser(id) {
  return supabase.from('users').delete().eq('id', id)
}

export async function findUserByCredentials(username, password) {
  return supabase
    .from('users')
    .select('*')
    .eq('username', username)
    .eq('password', password)
    .single()
}

// ─────────────────────────────────────────────────────────────────────────────
// UPLOAD LOGS
// ─────────────────────────────────────────────────────────────────────────────

export async function fetchUploadLogs() {
  return supabase
    .from('upload_logs')
    .select('*')
    .order('created_at', { ascending: false })
    .limit(100)
}

export async function createUploadLog(log) {
  return supabase.from('upload_logs').insert([log]).select().single()
}
