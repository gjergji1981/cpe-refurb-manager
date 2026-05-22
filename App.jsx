import { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";

// ── Supabase stub (preview mode — replace with real calls in production) ──
const findUserByCredentials = async () => ({ data: null, error: true });

// ─────────────────────────────────────────────
// CONSTANTS & HELPERS
// ─────────────────────────────────────────────
const DEVICE_TYPES = ["Router/Modem", "Set-top Box", "ONT/OLT"];
const STAGES = ["Intake","Triage","Refurbishment","QC Check","Stock","Scrap","Escalated","ECUS","In Transit","Confirmed"];
const TYPE_ICON = { "Router/Modem":"⬡", "Set-top Box":"▦", "ONT/OLT":"◈" };
const STAGE_COLOR = {
  Intake:        { bg:"#EEF2FF", text:"#3730A3", dot:"#6366F1" },
  Triage:        { bg:"#FFF7ED", text:"#9A3412", dot:"#F97316" },
  Refurbishment: { bg:"#FFFBEB", text:"#92400E", dot:"#F59E0B" },
  "QC Check":    { bg:"#EFF6FF", text:"#1E40AF", dot:"#3B82F6" },
  Stock:         { bg:"#F0FDF4", text:"#166534", dot:"#22C55E" },
  Scrap:         { bg:"#FEF2F2", text:"#991B1B", dot:"#EF4444" },
  Escalated:     { bg:"#FDF4FF", text:"#7E22CE", dot:"#A855F7" },
  ECUS:          { bg:"#FFF9EB", text:"#92600A", dot:"#F59E0B" },
  "In Transit":  { bg:"#EFF6FF", text:"#1D4ED8", dot:"#3B82F6" },
  "Confirmed":   { bg:"#F0FDF4", text:"#166534", dot:"#22C55E" },
};

// ─────────────────────────────────────────────────────────────────────────────
// OUI DATABASE  —  3-layer lookup: local extended table → API → "Unknown"
// Sources: IEEE MA-L registry + common CPE/ISP vendor additions
// ─────────────────────────────────────────────────────────────────────────────

// Layer 1: Extended local OUI table (200+ CPE-relevant vendors)
const OUI = {
  // ── Huawei ──────────────────────────────────────────────────
  "E8:65:D4":"Huawei","54:89:98":"Huawei","70:72:CF":"Huawei","00:9A:CD":"Huawei",
  "04:C0:6F":"Huawei","28:31:52":"Huawei","48:46:FB":"Huawei","AC:44:F2":"Huawei",
  "20:F3:A3":"Huawei","D4:6A:A8":"Huawei","F4:9F:F3":"Huawei","B4:15:13":"Huawei",
  "00:E0:FC":"Huawei","04:75:03":"Huawei","0C:37:DC":"Huawei","18:59:36":"Huawei",
  "20:08:ED":"Huawei","24:09:95":"Huawei","24:4C:07":"Huawei","28:6E:D4":"Huawei",
  "2C:AB:00":"Huawei","34:6A:C2":"Huawei","38:37:8B":"Huawei","40:4D:8E":"Huawei",
  "48:AD:08":"Huawei","4C:1F:CC":"Huawei","50:9F:27":"Huawei","58:2A:F7":"Huawei",
  "5C:C3:07":"Huawei","64:A2:F9":"Huawei","68:89:C1":"Huawei","6C:8D:C1":"Huawei",
  "70:54:F5":"Huawei","74:A5:28":"Huawei","78:1D:BA":"Huawei","7C:C3:A1":"Huawei",
  "80:D0:9B":"Huawei","88:3F:D3":"Huawei","8C:0D:76":"Huawei","98:28:A6":"Huawei",
  "9C:28:EF":"Huawei","A8:CA:7B":"Huawei","AC:85:3D":"Huawei","B0:75:0F":"Huawei",
  "C4:07:2F":"Huawei","C8:51:95":"Huawei","CC:53:B5":"Huawei","D0:7A:B5":"Huawei",
  "D4:A1:48":"Huawei","E4:68:A3":"Huawei","E8:08:8B":"Huawei","EC:23:3D":"Huawei",
  "F0:79:59":"Huawei","F4:5C:89":"Huawei","F8:36:9B":"Huawei","FC:48:EF":"Huawei",
  // ── ZTE ─────────────────────────────────────────────────────
  "00:19:CB":"ZTE","00:26:ED":"ZTE","34:4B:50":"ZTE","8C:A6:DF":"ZTE",
  "BC:F6:12":"ZTE","C8:6C:87":"ZTE","E4:63:DA":"ZTE","FC:AF:6A":"ZTE",
  "00:15:EB":"ZTE","04:90:42":"ZTE","0C:12:62":"ZTE","14:75:90":"ZTE",
  "18:68:CB":"ZTE","1C:6A:7A":"ZTE","20:F4:1B":"ZTE","28:5F:DB":"ZTE",
  "2C:95:7F":"ZTE","30:D3:2D":"ZTE","34:97:F6":"ZTE","38:22:9D":"ZTE",
  "3C:EA:EF":"ZTE","44:EA:D8":"ZTE","48:22:54":"ZTE","4C:09:B4":"ZTE",
  "50:8F:4C":"ZTE","54:C0:6A":"ZTE","58:05:99":"ZTE","5C:E6:CF":"ZTE",
  "60:DE:44":"ZTE","64:6E:69":"ZTE","68:23:B5":"ZTE","6C:E8:73":"ZTE",
  "70:AE:01":"ZTE","74:84:21":"ZTE","78:D9:0D":"ZTE","7C:DE:E8":"ZTE",
  "80:3C:14":"ZTE","84:D2:70":"ZTE","88:8B:39":"ZTE","90:F0:52":"ZTE",
  "94:C0:14":"ZTE","98:6C:F5":"ZTE","9C:6B:00":"ZTE","A0:E4:CB":"ZTE",
  "A4:3B:FA":"ZTE","A8:70:69":"ZTE","B4:0F:3B":"ZTE","B8:08:CF":"ZTE",
  "BC:E6:76":"ZTE","C0:9F:42":"ZTE","C4:02:89":"ZTE","C8:B2:B2":"ZTE",
  "D0:15:4A":"ZTE","D4:C6:55":"ZTE","D8:E9:DF":"ZTE","DC:02:8E":"ZTE",
  // ── Sagemcom ────────────────────────────────────────────────
  "00:24:E8":"Sagemcom","7C:4C:A5":"Sagemcom","C8:D7:19":"Sagemcom","E0:AB:39":"Sagemcom",
  "44:E9:DD":"Sagemcom","78:32:1B":"Sagemcom","48:EE:0C":"Sagemcom","B4:A5:EF":"Sagemcom",
  "00:13:C8":"Sagemcom","3C:36:E4":"Sagemcom","50:13:DA":"Sagemcom","88:57:6D":"Sagemcom",
  // ── Technicolor / Thomson ────────────────────────────────────
  "00:1A:C1":"Technicolor","00:24:D4":"Technicolor","28:C6:8E":"Technicolor",
  "50:E5:49":"Technicolor","78:E4:00":"Technicolor","A0:21:95":"Technicolor",
  "C0:25:2F":"Technicolor","DC:A6:32":"Technicolor","00:17:EE":"Technicolor",
  "2C:FE:8A":"Technicolor","34:7F:A7":"Technicolor","44:E9:DD":"Technicolor",
  "5C:49:79":"Technicolor","70:FC:8F":"Technicolor","88:03:55":"Technicolor",
  // ── Cisco / Cisco-Linksys ────────────────────────────────────
  "00:1A:2B":"Cisco","00:1B:2F":"Cisco","00:21:A0":"Cisco","00:26:99":"Cisco",
  "48:57:54":"Cisco","00:13:19":"Cisco","00:14:69":"Cisco","00:15:2B":"Cisco",
  "00:17:0E":"Cisco","00:18:BA":"Cisco","00:19:AA":"Cisco","00:1B:8F":"Cisco",
  "00:1C:10":"Cisco","00:1D:45":"Cisco","00:1E:14":"Cisco","00:1F:6C":"Cisco",
  "00:21:1B":"Cisco","00:22:BD":"Cisco","00:23:33":"Cisco","00:24:13":"Cisco",
  "00:25:2E":"Cisco","00:26:0A":"Cisco","2C:3E:CF":"Cisco","34:DB:FD":"Cisco",
  "44:E4:D9":"Cisco","54:78:1A":"Cisco","58:97:BD":"Cisco","5C:5A:C7":"Cisco",
  "60:5C:2A":"Cisco","64:F6:9D":"Cisco","6C:41:0E":"Cisco","70:10:5C":"Cisco",
  "74:26:AC":"Cisco","80:AC:AC":"Cisco","84:B8:02":"Cisco","88:5A:92":"Cisco",
  "8C:29:37":"Cisco","90:E2:BA":"Cisco","94:D4:69":"Cisco","98:90:96":"Cisco",
  "9C:AF:CA":"Cisco","A0:55:4F":"Cisco","A4:56:30":"Cisco","A8:B1:D4":"Cisco",
  "AC:F2:C5":"Cisco","B0:AA:77":"Cisco","B4:14:89":"Cisco","B8:38:61":"Cisco",
  "BC:16:65":"Cisco","C0:62:6B":"Cisco","C4:7D:4F":"Cisco","C8:9C:1D":"Cisco",
  "CC:D8:C1":"Cisco","D0:57:4C":"Cisco","D4:8C:B5":"Cisco","D8:24:DB":"Cisco",
  "DC:A4:CA":"Cisco","E0:1C:41":"Cisco","E4:AA:5D":"Cisco","E8:90:AF":"Cisco",
  "EC:1D:8B":"Cisco","F0:25:72":"Cisco","F4:CF:E2":"Cisco","F8:4F:57":"Cisco",
  "FC:FB:FB":"Cisco","B8:62:1F":"Cisco","D0:D0:FD":"Cisco","30:E4:DB":"Cisco",
  // ── Nokia ───────────────────────────────────────────────────
  "54:A7:03":"Nokia","9C:97:26":"Nokia","00:25:9C":"Nokia","D4:CA:6D":"Nokia",
  "B0:5B:67":"Nokia","18:ED:74":"Nokia","4C:F9:5D":"Nokia","30:B5:C2":"Nokia",
  "00:1E:9F":"Nokia","04:18:D6":"Nokia","14:0A:B9":"Nokia","2C:8A:72":"Nokia",
  "34:BB:26":"Nokia","50:4A:76":"Nokia","60:E3:AC":"Nokia","6C:1F:F7":"Nokia",
  "70:B3:D5":"Nokia","7C:E9:D3":"Nokia","84:74:2A":"Nokia","8C:F5:A3":"Nokia",
  // ── Arris ───────────────────────────────────────────────────
  "00:1D:7E":"Arris","00:21:5C":"Arris","00:26:B8":"Arris","18:1B:EB":"Arris",
  "3C:DF:A9":"Arris","70:2A:D5":"Arris","AC:20:2D":"Arris","E0:18:54":"Arris",
  "00:17:EF":"Arris","04:4E:5A":"Arris","08:61:95":"Arris","0C:BF:E7":"Arris",
  "18:68:1D":"Arris","20:46:F5":"Arris","24:8A:07":"Arris","2C:05:47":"Arris",
  "34:1F:E0":"Arris","38:D8:2F":"Arris","40:ED:00":"Arris","44:D9:E7":"Arris",
  "48:F8:B3":"Arris","50:39:55":"Arris","58:17:0C":"Arris","5C:50:15":"Arris",
  "60:67:20":"Arris","64:70:02":"Arris","68:C4:4D":"Arris","6C:70:9F":"Arris",
  // ── Netgear ─────────────────────────────────────────────────
  "00:17:10":"Netgear","00:1E:2A":"Netgear","20:E5:2A":"Netgear","2C:B0:5D":"Netgear",
  "84:1B:5E":"Netgear","A0:04:60":"Netgear","C4:04:15":"Netgear","00:14:6C":"Netgear",
  "00:18:4D":"Netgear","00:1F:33":"Netgear","00:22:3F":"Netgear","00:24:B2":"Netgear",
  "00:26:F2":"Netgear","04:A1:51":"Netgear","08:02:8E":"Netgear","0C:98:38":"Netgear",
  "10:0C:6B":"Netgear","14:59:C0":"Netgear","18:1E:78":"Netgear","1C:AF:F7":"Netgear",
  "20:4E:7F":"Netgear","28:80:88":"Netgear","30:46:9A":"Netgear","34:31:C4":"Netgear",
  "3C:37:86":"Netgear","40:4A:03":"Netgear","44:94:FC":"Netgear","4C:60:DE":"Netgear",
  "50:6A:03":"Netgear","54:04:A6":"Netgear","6C:B0:CE":"Netgear","70:4F:57":"Netgear",
  "74:44:01":"Netgear","78:D2:94":"Netgear","80:37:73":"Netgear","84:6A:ED":"Netgear",
  "9C:3D:CF":"Netgear","A0:40:A0":"Netgear","A4:2B:8C":"Netgear","B0:39:56":"Netgear",
  "B0:7F:B9":"Netgear","C0:3F:0E":"Netgear","C4:3D:C7":"Netgear","C4:E9:84":"Netgear",
  // ── TP-Link ──────────────────────────────────────────────────
  "30:46:9A":"TP-Link","50:C7:BF":"TP-Link","98:DE:D0":"TP-Link","E8:48:B8":"TP-Link",
  "00:27:22":"TP-Link","14:CC:20":"TP-Link","B0:BE:76":"TP-Link","EC:08:6B":"TP-Link",
  "10:FE:ED":"TP-Link","18:D6:C7":"TP-Link","1C:3B:F3":"TP-Link","20:DC:E6":"TP-Link",
  "24:69:A5":"TP-Link","28:28:5D":"TP-Link","2C:D0:5A":"TP-Link","30:B5:C2":"TP-Link",
  "40:8D:5C":"TP-Link","44:94:FC":"TP-Link","4C:E1:73":"TP-Link","50:FA:84":"TP-Link",
  "54:C8:0F":"TP-Link","58:EF:68":"TP-Link","5C:89:9A":"TP-Link","60:32:B1":"TP-Link",
  "64:70:02":"TP-Link","6C:5A:B0":"TP-Link","70:4F:57":"TP-Link","74:EA:3A":"TP-Link",
  "78:A1:06":"TP-Link","7C:8B:CA":"TP-Link","80:EA:07":"TP-Link","84:16:F9":"TP-Link",
  "88:DC:96":"TP-Link","8C:21:0A":"TP-Link","90:F6:52":"TP-Link","94:0C:6D":"TP-Link",
  "98:25:4A":"TP-Link","A0:F3:C1":"TP-Link","A4:2B:B0":"TP-Link","AC:84:C9":"TP-Link",
  "B0:95:8E":"TP-Link","B4:B0:24":"TP-Link","B8:27:EB":"TP-Link","BC:46:99":"TP-Link",
  "C0:4A:00":"TP-Link","C4:E9:84":"TP-Link","C8:0E:14":"TP-Link","D8:07:B6":"TP-Link",
  "DC:FE:18":"TP-Link","E0:28:6D":"TP-Link","E8:DE:27":"TP-Link","EC:17:2F":"TP-Link",
  "F4:EC:38":"TP-Link","F8:1A:67":"TP-Link","FC:EC:DA":"TP-Link",
  // ── D-Link ───────────────────────────────────────────────────
  "00:05:5D":"D-Link","00:0D:88":"D-Link","00:0F:3D":"D-Link","00:11:95":"D-Link",
  "00:13:46":"D-Link","00:15:E9":"D-Link","00:17:9A":"D-Link","00:19:5B":"D-Link",
  "00:1B:11":"D-Link","00:1C:F0":"D-Link","00:1E:58":"D-Link","00:1F:3A":"D-Link",
  "00:21:91":"D-Link","00:22:B0":"D-Link","00:24:01":"D-Link","00:26:5A":"D-Link",
  "1C:7E:E5":"D-Link","28:10:7B":"D-Link","2C:B0:5D":"D-Link","34:08:04":"D-Link",
  "5C:D9:98":"D-Link","78:54:2E":"D-Link","84:C9:B2":"D-Link","90:94:E4":"D-Link",
  "AC:F1:DF":"D-Link","B8:A3:86":"D-Link","C8:BE:19":"D-Link","CC:B2:55":"D-Link",
  "E4:6F:13":"D-Link","EC:43:F6":"D-Link","F0:B4:D2":"D-Link","FC:75:16":"D-Link",
  // ── Zyxel ────────────────────────────────────────────────────
  "00:13:49":"Zyxel","00:19:CB":"Zyxel","00:A0:C5":"Zyxel","1C:74:0D":"Zyxel",
  "20:CF:30":"Zyxel","28:28:5D":"Zyxel","38:72:C0":"Zyxel","40:4A:03":"Zyxel",
  "48:8F:5A":"Zyxel","4C:09:B4":"Zyxel","50:67:F0":"Zyxel","58:8B:F3":"Zyxel",
  "64:0B:B0":"Zyxel","6C:4B:90":"Zyxel","78:44:FD":"Zyxel","84:AA:9C":"Zyxel",
  "88:36:6C":"Zyxel","90:C7:D8":"Zyxel","98:D6:BB":"Zyxel","A0:E4:CB":"Zyxel",
  "AC:3B:77":"Zyxel","B8:D5:0B":"Zyxel","C0:3F:0E":"Zyxel","C8:6C:87":"Zyxel",
  "D8:5D:E2":"Zyxel","E8:37:7A":"Zyxel","FC:F5:28":"Zyxel",
  // ── Alcatel-Lucent / Nokia (legacy) ──────────────────────────
  "00:E0:6F":"Alcatel-Lucent","08:00:03":"Alcatel-Lucent","44:2B:03":"Alcatel-Lucent",
  "2C:FA:A2":"Alcatel-Lucent","40:61:86":"Alcatel-Lucent","78:19:F7":"Alcatel-Lucent",
  "84:6C:C1":"Alcatel-Lucent","8C:04:FF":"Alcatel-Lucent","A4:69:D3":"Alcatel-Lucent",
  "D4:CA:6D":"Alcatel-Lucent","E0:D0:45":"Alcatel-Lucent","F4:3E:61":"Alcatel-Lucent",
  // ── Ericsson ─────────────────────────────────────────────────
  "00:01:B2":"Ericsson","00:09:23":"Ericsson","00:0B:5A":"Ericsson","00:0F:A3":"Ericsson",
  "00:12:EE":"Ericsson","00:18:DE":"Ericsson","00:21:BC":"Ericsson","00:24:7F":"Ericsson",
  "28:8A:1C":"Ericsson","44:45:53":"Ericsson","5C:A4:8A":"Ericsson","6C:BF:B5":"Ericsson",
  "80:E0:1D":"Ericsson","9C:64:4A":"Ericsson","BC:3F:8F":"Ericsson","D0:84:B0":"Ericsson",
  // ── Calix ────────────────────────────────────────────────────
  "00:30:48":"Calix","00:E0:D8":"Calix","70:B3:D5":"Calix","B0:91:22":"Calix",
  "DC:19:64":"Calix","E4:D3:F1":"Calix",
  // ── Sercomm ──────────────────────────────────────────────────
  "00:09:01":"Sercomm","00:19:3E":"Sercomm","00:1A:AE":"Sercomm","00:90:D0":"Sercomm",
  "1C:40:04":"Sercomm","28:2C:B2":"Sercomm","5C:F4:AB":"Sercomm","6C:B0:CE":"Sercomm",
  "C4:27:95":"Sercomm","D8:38:FC":"Sercomm","E4:6C:29":"Sercomm",
  // ── Ubiquiti ─────────────────────────────────────────────────
  "00:15:6D":"Ubiquiti","00:27:22":"Ubiquiti","04:18:D6":"Ubiquiti","18:E8:29":"Ubiquiti",
  "24:A4:3C":"Ubiquiti","44:D9:E7":"Ubiquiti","68:72:51":"Ubiquiti","78:8A:20":"Ubiquiti",
  "80:2A:A8":"Ubiquiti","DC:9F:DB":"Ubiquiti","F0:9F:C2":"Ubiquiti","FC:EC:DA":"Ubiquiti",
  // ── Motorola / Zoom ──────────────────────────────────────────
  "00:1C:DF":"Motorola","AC:DE:48":"Motorola","58:93:96":"Motorola","00:15:6D":"Motorola",
  "20:FF:BD":"Motorola","30:91:8F":"Motorola","40:B4:CD":"Motorola","58:1F:AA":"Motorola",
  "74:2F:68":"Motorola","84:10:0D":"Motorola","9C:4F:DA":"Motorola","AC:3A:7A":"Motorola",
  // ── ASUS ─────────────────────────────────────────────────────
  "00:18:01":"ASUS","04:92:26":"ASUS","2C:FD:A1":"ASUS","50:46:5D":"ASUS",
  "10:BF:48":"ASUS","14:DA:E9":"ASUS","1C:87:2C":"ASUS","20:CF:30":"ASUS",
  "2C:4D:54":"ASUS","30:85:A9":"ASUS","38:2C:4A":"ASUS","40:16:7E":"ASUS",
  "48:5B:39":"ASUS","4C:ED:FB":"ASUS","50:3E:AA":"ASUS","5C:FF:35":"ASUS",
  "60:A4:4C":"ASUS","70:8B:CD":"ASUS","74:D0:2B":"ASUS","84:A9:C4":"ASUS",
  "88:D7:F6":"ASUS","90:E6:BA":"ASUS","94:DE:80":"ASUS","A8:5E:45":"ASUS",
  "AC:9E:17":"ASUS","B0:6E:BF":"ASUS","BC:AE:C5":"ASUS","C8:60:00":"ASUS",
  "D8:50:E6":"ASUS","E0:3F:49":"ASUS","E4:70:B8":"ASUS","F8:32:E4":"ASUS",
  // ── CommScope (ex-ARRIS/Ruckus) ──────────────────────────────
  "00:90:9E":"CommScope","3C:9A:59":"CommScope","68:A3:C4":"CommScope",
  "78:72:5D":"CommScope","80:2B:F9":"CommScope","98:31:E5":"CommScope",
  "B4:75:0E":"CommScope","C0:56:E3":"CommScope","D4:6A:35":"CommScope",
  // ── Genexis / Frontier ───────────────────────────────────────
  "00:1D:9A":"Genexis","2C:39:C1":"Genexis","40:4E:36":"Genexis",
  "6C:72:20":"Genexis","A8:C5:4F":"Genexis","D8:9E:3F":"Genexis",
  // ── AVM (FRITZ!Box) ──────────────────────────────────────────
  "00:04:0E":"AVM","3C:A6:2F":"AVM","C4:93:D9":"AVM","D4:21:22":"AVM",
  "DC:39:6F":"AVM","E8:40:F2":"AVM","F0:B0:14":"AVM",
  // ── Comtrend ─────────────────────────────────────────────────
  "00:12:BF":"Comtrend","00:17:9C":"Comtrend","00:D0:0B":"Comtrend",
  "34:1C:F0":"Comtrend","8C:04:FF":"Comtrend","A8:D2:36":"Comtrend",
  // ── Actiontec ────────────────────────────────────────────────
  "00:18:01":"Actiontec","00:26:B8":"Actiontec","60:C5:47":"Actiontec",
  "80:B2:C8":"Actiontec","A4:48:30":"Actiontec","D4:05:98":"Actiontec",
  // ── Hitron ───────────────────────────────────────────────────
  "00:21:80":"Hitron","2C:30:33":"Hitron","40:F0:1C":"Hitron",
  "68:1C:A2":"Hitron","A4:08:F5":"Hitron","C8:D5:FE":"Hitron","F8:08:4F":"Hitron",
  // ── Sagem (legacy) ───────────────────────────────────────────
  "00:01:26":"Sagem","00:11:BC":"Sagem","00:1A:70":"Sagem","00:E0:92":"Sagem",
  // ── Pace ─────────────────────────────────────────────────────
  "00:1D:D9":"Pace","00:24:12":"Pace","2C:30:33":"Pace","44:D7:74":"Pace",
  "5C:49:79":"Pace","68:5B:35":"Pace","88:5B:DD":"Pace","C4:3D:C7":"Pace",
  // ── Samsung (gateways) ───────────────────────────────────────
  "00:12:47":"Samsung","00:15:B9":"Samsung","00:1D:25":"Samsung","00:23:39":"Samsung",
  "00:26:37":"Samsung","10:1D:C0":"Samsung","14:7D:C5":"Samsung","18:46:17":"Samsung",
  "1C:62:B8":"Samsung","20:13:E0":"Samsung","28:BA:B5":"Samsung","2C:54:CF":"Samsung",
  "30:96:FB":"Samsung","34:31:11":"Samsung","38:AA:3C":"Samsung","3C:BD:D8":"Samsung",
  "40:0E:85":"Samsung","44:5C:E9":"Samsung","48:44:F7":"Samsung","4C:3C:16":"Samsung",
  "50:01:BB":"Samsung","54:88:0E":"Samsung","5C:49:7D":"Samsung","60:A1:0A":"Samsung",
  "64:B8:53":"Samsung","68:EB:C5":"Samsung","6C:83:36":"Samsung","70:77:81":"Samsung",
  "78:1F:DB":"Samsung","7C:0B:C6":"Samsung","80:18:A7":"Samsung","84:25:19":"Samsung",
  "88:32:9B":"Samsung","8C:77:12":"Samsung","90:18:7C":"Samsung","94:35:0A":"Samsung",
  "98:52:B1":"Samsung","9C:02:98":"Samsung","A0:0B:BA":"Samsung","A4:84:31":"Samsung",
};

// Layer 2: Session-level API cache (avoids duplicate requests)
const _ouiCache = new Map();

// Layer 3: Live API lookup (macvendors.com — free, no key required)
async function lookupOUIRemote(oui) {
  if (_ouiCache.has(oui)) return _ouiCache.get(oui);
  try {
    const res = await fetch(`https://api.macvendors.com/${oui}`, {
      signal: AbortSignal.timeout(3000),
    });
    if (!res.ok) { _ouiCache.set(oui, null); return null; }
    const vendor = (await res.text()).trim();
    _ouiCache.set(oui, vendor);
    return vendor;
  } catch {
    _ouiCache.set(oui, null);
    return null;
  }
}

// Extract OUI prefix from any MAC format
function extractOUI(mac) {
  if (!mac) return null;
  const c = mac.toUpperCase().replace(/[^0-9A-F]/g, "");
  if (c.length < 6) return null;
  return `${c.slice(0,2)}:${c.slice(2,4)}:${c.slice(4,6)}`;
}

// Sync lookup (local table only) — instant, used for badges/display
function lookupOUI(mac) {
  const oui = extractOUI(mac);
  return oui ? (OUI[oui] || null) : null;
}

// Async lookup (local → API fallback) — used when operator types a MAC
async function lookupOUIFull(mac) {
  const oui = extractOUI(mac);
  if (!oui) return null;
  // Check local first
  if (OUI[oui]) return OUI[oui];
  // Check cache
  if (_ouiCache.has(oui)) return _ouiCache.get(oui);
  // Hit the API
  return lookupOUIRemote(oui);
}

function formatMac(raw) {
  const c = raw.replace(/[^0-9a-fA-F]/g,"").slice(0,12).toUpperCase();
  const parts = c.match(/.{1,2}/g) || [];
  return parts.join(":");
}
let _id = 13;
function genId() { return `D-${String(_id++).padStart(4,"0")}`; }
function today() { return new Date().toISOString().slice(0,10); }
function useIsMobile() {
  const [mobile, setMobile] = useState(window.innerWidth < 768);
  useEffect(() => {
    const h = () => setMobile(window.innerWidth < 768);
    window.addEventListener("resize", h);
    return () => window.removeEventListener("resize", h);
  }, []);
  return mobile;
}

// Persists state to localStorage — survives page refresh
function useLocalStorage(key, initialValue) {
  const [value, setValue] = useState(() => {
    try {
      const stored = localStorage.getItem(key);
      return stored ? JSON.parse(stored) : initialValue;
    } catch {
      return initialValue;
    }
  });
  const set = (newValue) => {
    try {
      const toStore = typeof newValue === "function" ? newValue(value) : newValue;
      setValue(toStore);
      localStorage.setItem(key, JSON.stringify(toStore));
    } catch {
      setValue(newValue);
    }
  };
  return [value, set];
}

// ─────────────────────────────────────────────
// SEED DATA
// ─────────────────────────────────────────────
const SEED = [
  { id:"D-0001", serial:"SN-88421", mac:"E8:65:D4:11:22:33", model:"Huawei HG8245H",        type:"Router/Modem", stage:"Stock",         outcome:"Working",     received:"2025-05-01", notes:"Firmware reset",      sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0002", serial:"SN-33190", mac:"28:C6:8E:44:55:66", model:"Technicolor TC7200",     type:"Set-top Box",  stage:"Stock",         outcome:"Working",     received:"2025-05-02", notes:"HDMI port replaced",  sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0003", serial:"SN-77042", mac:"9C:97:26:77:88:99", model:"Nokia G-010G-P",         type:"ONT/OLT",      stage:"Stock",         outcome:"Working",     received:"2025-05-03", notes:"Clean & reconfigure", sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0004", serial:"SN-55610", mac:"34:4B:50:AA:BB:CC", model:"ZTE ZXHN H108N",         type:"Router/Modem", stage:"Scrap",         outcome:"Scrap",       received:"2025-05-04", notes:"PCB damage",          sentToPartner:false, partnerOutcome:null, partnerNotes:"" },
  { id:"D-0005", serial:"SN-12983", mac:"",                  model:"",                       type:"Set-top Box",  stage:"Scrap",         outcome:"Scrap",       received:"2025-05-05", notes:"Burned PSU",          sentToPartner:false, partnerOutcome:null, partnerNotes:"" },
  { id:"D-0006", serial:"SN-66101", mac:"54:89:98:DD:EE:FF", model:"Huawei EchoLife EG8145", type:"ONT/OLT",      stage:"Refurbishment", outcome:null,          received:"2025-05-06", notes:"In repair",           sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0007", serial:"SN-29344", mac:"C8:D7:19:12:34:56", model:"Sagemcom F@ST 5366",     type:"Router/Modem", stage:"Refurbishment", outcome:null,          received:"2025-05-07", notes:"Awaiting partner",    sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0008", serial:"SN-84720", mac:"",                  model:"",                       type:"Set-top Box",  stage:"Triage",        outcome:null,          received:"2025-05-08", notes:"",                    sentToPartner:false, partnerOutcome:null, partnerNotes:"" },
  { id:"D-0009", serial:"SN-39011", mac:"BC:F6:12:AB:CD:EF", model:"ZTE F660",               type:"ONT/OLT",      stage:"Escalated",     outcome:"Not Working", received:"2025-05-09", notes:"Failed QC twice",     sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0010", serial:"SN-47832", mac:"00:1A:2B:33:44:55", model:"Cisco DPC3825",          type:"Router/Modem", stage:"Stock",         outcome:"Working",     received:"2025-05-10", notes:"",                    sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0011", serial:"SN-90012", mac:"7C:4C:A5:66:77:88", model:"Sagemcom FAST 3686",     type:"Set-top Box",  stage:"Refurbishment", outcome:null,          received:"2025-05-11", notes:"Screen replaced",     sentToPartner:true,  partnerOutcome:null, partnerNotes:"" },
  { id:"D-0012", serial:"SN-55123", mac:"",                  model:"",                       type:"ONT/OLT",      stage:"Triage",        outcome:null,          received:"2025-05-12", notes:"",                    sentToPartner:false, partnerOutcome:null, partnerNotes:"" },
];

// ─────────────────────────────────────────────
// DESIGN TOKENS
// ─────────────────────────────────────────────
const C = {
  indigo:"#6366F1", indigoDark:"#4338CA", indigoLight:"#EEF2FF",
  purple:"#7C3AED", purpleLight:"#FDF4FF",
  amber:"#F59E0B",  amberLight:"#FFFBEB",
  green:"#22C55E",  greenLight:"#F0FDF4",  greenDark:"#166534",
  red:"#EF4444",    redLight:"#FEF2F2",    redDark:"#991B1B",
  slate:"#0F172A",  slate2:"#1E293B",      slate3:"#475569",
  slate4:"#94A3B8", slate5:"#CBD5E1",      slate6:"#E2E8F0",
  slate7:"#F1F5F9", slate8:"#F8FAFC",
  white:"#fff",
};

// ─────────────────────────────────────────────
// SHARED COMPONENTS
// ─────────────────────────────────────────────
function Badge({ stage }) {
  const c = STAGE_COLOR[stage] || { bg:C.slate7, text:C.slate3, dot:C.slate4 };
  return (
    <span style={{ background:c.bg, color:c.text, fontSize:11, fontWeight:700,
      padding:"3px 9px", borderRadius:20, display:"inline-flex", alignItems:"center", gap:4, whiteSpace:"nowrap" }}>
      <span style={{ width:6, height:6, borderRadius:"50%", background:c.dot, flexShrink:0 }}/>
      {stage}
    </span>
  );
}

function Btn({ children, onClick, variant="default", size="md", full=false, disabled=false, style:overrideStyle={} }) {
  const v = {
    default: { bg:C.slate8,     color:C.slate3,   border:`1px solid ${C.slate6}` },
    primary: { bg:C.indigo,     color:C.white,     border:"none" },
    success: { bg:C.greenLight, color:C.greenDark, border:`1px solid #BBF7D0` },
    danger:  { bg:C.redLight,   color:C.redDark,   border:`1px solid #FECACA` },
    purple:  { bg:C.purpleLight,color:C.purple,    border:`1px solid #E9D5FF` },
    amber:   { bg:C.amberLight, color:"#92400E",   border:`1px solid #FCD34D` },
    ghost:   { bg:"transparent",color:C.slate3,    border:`1px solid ${C.slate6}` },
    dark:    { bg:C.slate,      color:C.white,     border:"none" },
  }[variant] || {};
  const pad = size==="sm" ? "5px 12px" : size==="lg" ? "12px 24px" : "8px 16px";
  const fs  = size==="sm" ? 12 : size==="lg" ? 15 : 13;
  return (
    <button onClick={onClick} disabled={disabled}
      style={{ ...v, padding:pad, borderRadius:9, fontSize:fs, fontWeight:700,
        cursor:disabled?"not-allowed":"pointer", opacity:disabled?.5:1,
        whiteSpace:"nowrap", width:full?"100%":"auto",
        display:"inline-flex", alignItems:"center", justifyContent:"center", gap:6, transition:"opacity .15s", ...overrideStyle }}>
      {children}
    </button>
  );
}

function Card({ children, style={}, accent }) {
  return (
    <div style={{ background:C.white, border:`1.5px solid ${C.slate6}`, borderRadius:14,
      borderLeft: accent ? `4px solid ${accent}` : undefined,
      padding:16, ...style }}>
      {children}
    </div>
  );
}

function SectionTitle({ children }) {
  return <h3 style={{ margin:"0 0 14px", fontSize:14, fontWeight:700, color:C.slate }}>{children}</h3>;
}

function Label({ children }) {
  return <label style={{ fontSize:11, fontWeight:700, color:C.slate4, letterSpacing:".06em", display:"block", marginBottom:4 }}>{children}</label>;
}

const iStyle = (extra={}) => ({
  width:"100%", padding:"9px 12px", borderRadius:9, border:`1.5px solid ${C.slate6}`,
  fontSize:13, outline:"none", boxSizing:"border-box", background:C.white,
  WebkitAppearance:"none", ...extra
});

function StatCard({ label, value, sub, accent, badge }) {
  return (
    <Card accent={accent} style={{ display:"flex", flexDirection:"column", gap:4 }}>
      <span style={{ fontSize:11, fontWeight:700, color:C.slate4, textTransform:"uppercase", letterSpacing:".07em" }}>{label}</span>
      <div style={{ display:"flex", alignItems:"baseline", gap:8 }}>
        <span style={{ fontSize:28, fontWeight:800, color:C.slate, lineHeight:1 }}>{value}</span>
        {badge && <span style={{ fontSize:11, fontWeight:700, padding:"2px 7px", borderRadius:20,
          background:badge.bg, color:badge.color }}>{badge.text}</span>}
      </div>
      {sub && <span style={{ fontSize:12, color:C.slate4 }}>{sub}</span>}
    </Card>
  );
}

// Mobile-friendly device card (replaces table rows on mobile)
function DeviceCard({ d, actions, fields }) {
  return (
    <div style={{ border:`1px solid ${C.slate6}`, borderRadius:12, padding:14, background:C.white, marginBottom:10 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:8 }}>
        <div>
          <span style={{ fontWeight:800, color:C.indigo, fontSize:13 }}>{d.id}</span>
          <span style={{ fontFamily:"monospace", fontSize:12, color:C.slate2, marginLeft:8 }}>{d.serial}</span>
        </div>
        <Badge stage={d.stage}/>
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6, marginBottom:10 }}>
        {fields.map(([k,v])=>(
          <div key={k}>
            <span style={{ fontSize:10, fontWeight:700, color:C.slate4, textTransform:"uppercase", letterSpacing:".05em" }}>{k}</span>
            <div style={{ fontSize:12, color:C.slate2, marginTop:1, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{v||"—"}</div>
          </div>
        ))}
      </div>
      {actions && <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>{actions}</div>}
    </div>
  );
}

// Alert banner
function Alert({ children, type="info" }) {
  const colors = {
    info:    { bg:"#EEF2FF", border:`1.5px solid #C7D2FE`, color:"#3730A3" },
    warning: { bg:"#FFFBEB", border:`1.5px solid #FCD34D`, color:"#92400E" },
    success: { bg:C.greenLight, border:`1.5px solid #BBF7D0`, color:C.greenDark },
    danger:  { bg:C.redLight,   border:`1.5px solid #FECACA`,  color:C.redDark },
  }[type];
  return (
    <div style={{ ...colors, borderRadius:10, padding:"10px 14px", fontSize:13, display:"flex", alignItems:"center", gap:8 }}>
      {children}
    </div>
  );
}

// Tab switcher
function Tabs({ tabs, active, onChange }) {
  return (
    <div style={{ display:"flex", background:C.slate7, borderRadius:10, padding:3, gap:2 }}>
      {tabs.map(t=>(
        <button key={t.id} onClick={()=>onChange(t.id)}
          style={{ flex:1, padding:"7px 12px", border:"none", borderRadius:8, fontSize:12, fontWeight:700,
            cursor:"pointer", background:active===t.id?C.white:"transparent",
            color:active===t.id?C.slate:C.slate3, transition:"all .15s",
            boxShadow:active===t.id?"0 1px 3px rgba(0,0,0,.08)":undefined, whiteSpace:"nowrap" }}>
          {t.label}
        </button>
      ))}
    </div>
  );
}

// ─────────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────────
function Dashboard({ devices, isMobile }) {
  const total       = devices.length;
  const inFlow      = devices.filter(d=>!["Stock","Scrap"].includes(d.stage)).length;
  const stock       = devices.filter(d=>d.stage==="Stock").length;
  const scrap       = devices.filter(d=>d.stage==="Scrap").length;
  const atPartner   = devices.filter(d=>d.sentToPartner && d.stage==="Refurbishment" && !d.partnerOutcome).length;
  const ecusPending = devices.filter(d=>d.stage==="ECUS").length;
  const pendingConf = devices.filter(d=>d.partnerOutcome && d.stage==="Refurbishment").length;
  const byStage     = STAGES.map(s=>({ stage:s, count:devices.filter(d=>d.stage===s).length }));
  const byType      = DEVICE_TYPES.map(t=>({ type:t, total:devices.filter(d=>d.type===t).length, stock:devices.filter(d=>d.type===t&&d.stage==="Stock").length, scrap:devices.filter(d=>d.type===t&&d.stage==="Scrap").length }));
  const maxStage    = Math.max(...byStage.map(s=>s.count), 1);
  const done        = devices.filter(d=>d.outcome==="Working"||d.outcome==="Not Working");
  const working     = done.filter(d=>d.outcome==="Working").length;
  const rate        = done.length ? Math.round(working/done.length*100) : 0;

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      <div>
        <h2 style={{ margin:"0 0 2px", fontSize:isMobile?18:22, fontWeight:800, color:C.slate }}>Dashboard</h2>
        <p style={{ margin:0, color:C.slate4, fontSize:13 }}>Live pipeline overview</p>
      </div>

      <DeviceSearchBar devices={devices} isMobile={isMobile}/>
      {pendingConf>0 && <Alert type="warning">⚠ <strong>{pendingConf}</strong> partner outcome{pendingConf>1?"s":""} awaiting your confirmation in Refurbishment</Alert>}

      <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(7,1fr)", gap:10 }}>
        <StatCard label="Total" value={total} sub="All devices" accent={C.indigo}/>
        <StatCard label="Pipeline" value={inFlow} sub="In progress" accent={C.amber}/>
        <StatCard label="Stock" value={stock} sub="Ready" accent={C.green}/>
        <StatCard label="Scrapped" value={scrap} sub="Disposed" accent={C.red}/>
        <StatCard label="At Partner" value={atPartner}
          badge={pendingConf?{text:`${pendingConf} pending`,bg:"#FFFBEB",color:"#92400E"}:undefined}
          sub="Awaiting results" accent={C.purple}/>
        <StatCard label="ECUS" value={ecusPending} sub="Held pending" accent="#F59E0B"/>
      </div>

      <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr":"1fr 1fr", gap:16 }}>
        <Card>
          <SectionTitle>Devices by stage</SectionTitle>
          {byStage.map(({ stage, count })=>{
            const sc = STAGE_COLOR[stage] || { dot:C.slate4 };
            return (
              <div key={stage} style={{ display:"flex", alignItems:"center", gap:8, marginBottom:8 }}>
                <span style={{ width:isMobile?80:96, fontSize:12, fontWeight:600, color:C.slate3, textAlign:"right", flexShrink:0 }}>{stage}</span>
                <div style={{ flex:1, background:C.slate7, borderRadius:6, height:20, overflow:"hidden" }}>
                  <div style={{ width:`${(count/maxStage)*100}%`, minWidth:count?28:0, height:"100%",
                    background:sc.dot, borderRadius:6, display:"flex", alignItems:"center", paddingLeft:6, transition:"width .4s" }}>
                    {count>0 && <span style={{ fontSize:11, fontWeight:800, color:C.white }}>{count}</span>}
                  </div>
                </div>
                {count===0 && <span style={{ fontSize:11, color:C.slate5, width:14 }}>0</span>}
              </div>
            );
          })}
        </Card>

        <Card>
          <SectionTitle>By device type</SectionTitle>
          {byType.map(r=>(
            <div key={r.type} style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
              padding:"10px 0", borderBottom:`1px solid ${C.slate7}` }}>
              <span style={{ fontSize:13, fontWeight:600, color:C.slate }}>{TYPE_ICON[r.type]} {r.type}</span>
              <div style={{ display:"flex", gap:16 }}>
                <span style={{ fontSize:13, fontWeight:700, color:C.slate }}>{r.total}</span>
                <span style={{ fontSize:13, fontWeight:700, color:C.greenDark }}>↑{r.stock}</span>
                <span style={{ fontSize:13, fontWeight:700, color:C.redDark }}>↓{r.scrap}</span>
              </div>
            </div>
          ))}
          <div style={{ marginTop:14 }}>
            <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
              <span style={{ fontSize:12, color:C.slate3 }}>Refurb success rate</span>
              <span style={{ fontSize:12, fontWeight:800, color:C.greenDark }}>{rate}%</span>
            </div>
            <div style={{ background:C.slate7, borderRadius:8, height:10, overflow:"hidden" }}>
              <div style={{ width:`${rate}%`, height:"100%", background:C.green, borderRadius:8, transition:"width .4s" }}/>
            </div>
            <div style={{ fontSize:11, color:C.slate4, marginTop:4 }}>{working} of {done.length} tested units passed</div>
          </div>
        </Card>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// INTAKE & TRIAGE
// ─────────────────────────────────────────────
function IntakeTriage({ devices, setDevices, isMobile }) {
  const ACTIONS = ["Refurbishment", "Scrap", "ECUS"];

  const [tab, setTab]               = useState("manual");
  const [form, setForm]             = useState({ serial:"", type:DEVICE_TYPES[0], mac:"", model:"", notes:"", action:"Refurbishment" });
  const [ok, setOk]                 = useState(null);
  const [formError, setFormError]   = useState("");
  const [dragOver, setDragOver]     = useState(false);
  const [processing, setProcessing] = useState(false);
  const [uploadResult, setUploadResult] = useState(null);
  const [uploadLogs, setUploadLogs]     = useState([]);
  const [editingId, setEditingId]   = useState(null);
  const [editBuf, setEditBuf]       = useState({});
  const [deleteId, setDeleteId]     = useState(null);
  const [expandedId, setExpandedId] = useState(null);
  const [showConfirm, setShowConfirm] = useState(false);

  // Queue = devices in Intake/Triage that have a pendingAction set
  const queue = devices.filter(d => ["Intake","Triage"].includes(d.stage));

  // Counts for confirmation popup
  const toRefurb = queue.filter(d => d.pendingAction === "Refurbishment").length;
  const toScrap  = queue.filter(d => d.pendingAction === "Scrap").length;
  const toEcus   = queue.filter(d => d.pendingAction === "ECUS").length;
  const unassigned = queue.filter(d => !d.pendingAction).length;
  const readyCount = toRefurb + toScrap + toEcus;

  // ── MAC / OUI helper ──
  // Sync MAC change — updates field instantly, triggers OUI lookup separately
  function handleMacChange(raw, setBuf) {
    const mac = formatMac(raw);
    setBuf(f => ({ ...f, mac }));
    // Trigger async OUI lookup without blocking onChange
    const oui = extractOUI(mac);
    if (!oui) return;
    lookupOUIFull(mac).then(mfr => {
      if (!mfr) return;
      setBuf(f => {
        const wasAutoFilled = f.model && f.model.endsWith(" - ");
        return (!f.model || wasAutoFilled) ? { ...f, model: `${mfr} - ` } : f;
      });
    }).catch(() => {});
  }

  // ── Save single device to queue ──
  function saveToQueue() {
    if (!form.serial.trim()) { setFormError("Serial number is required."); return; }
    const duplicate = devices.find(d => d.serial.toLowerCase() === form.serial.trim().toLowerCase());
    if (duplicate) { setFormError("This serial number already exists."); return; }
    setFormError("");
    const d = {
      id: genId(), serial: form.serial.trim(), type: form.type,
      mac: form.mac, model: form.model, stage: "Triage",
      outcome: null, received: today(), notes: form.notes,
      sentToPartner: false, partnerOutcome: null, partnerNotes: "",
      pendingAction: form.action,
    };
    // Persist to Supabase (fire-and-forget; UI updates immediately)
    createDevice(deviceToRow(d)).catch(e => console.warn("DB write failed:", e));
    setDevices(p => [d, ...p]);
    setOk(d.id);
    setForm({ serial:"", type:DEVICE_TYPES[0], mac:"", model:"", notes:"", action:"Refurbishment" });
    setTimeout(() => setOk(null), 2500);
  }

  // ── Set pending action on a queued device ──
  function setPendingAction(id, action) {
    setDevices(p => p.map(d => d.id === id ? { ...d, pendingAction: action } : d));
  }

  // ── Edit / Delete ──
  function startEdit(d) {
    setEditingId(d.id);
    setEditBuf({ serial:d.serial, type:d.type, mac:d.mac||"", model:d.model||"", notes:d.notes||"", pendingAction:d.pendingAction||"Refurbishment" });
  }
  function saveEdit(id) {
    setDevices(p => p.map(d => d.id === id ? { ...d, ...editBuf } : d));
    setEditingId(null);
  }
  function deleteDevice(id) {
    deleteDevice_db(id).catch(e => console.warn("DB delete failed:", e));
    setDevices(p => p.filter(d => d.id !== id));
    setDeleteId(null);
  }

  // ── Execute queue ──
  function executeQueue() {
    setDevices(p => p.map(d => {
      if (!["Intake","Triage"].includes(d.stage)) return d;
      if (!d.pendingAction) return d;
      const isScrap = d.pendingAction === "Scrap";
      const isEcus  = d.pendingAction === "ECUS";
      const updated = {
        ...d,
        stage:        isScrap ? "Scrap" : isEcus ? "ECUS" : "Refurbishment",
        outcome:      isScrap ? "Scrap" : null,
        sentToPartner:!isScrap && !isEcus,
        pendingAction: null,
      };
      updateDevice(d.id, {
        stage: updated.stage, outcome: updated.outcome,
        sent_to_partner: updated.sentToPartner, pending_action: null
      }).catch(e => console.warn("DB update failed:", e));
      return updated;
    }));
    setShowConfirm(false);
  }

  // ── Bulk upload helpers ──
  function resolveType(raw="") {
    const v = raw.toLowerCase();
    if (v.includes("router")||v.includes("modem")) return "Router/Modem";
    if (v.includes("set-top")||v.includes("stb")||v.includes("settop")) return "Set-top Box";
    if (v.includes("ont")||v.includes("olt")) return "ONT/OLT";
    return null;
  }
  function resolveAction(raw="") {
    const v = raw.toLowerCase();
    if (v.includes("scrap")) return "Scrap";
    if (v.includes("ecus"))  return "ECUS";
    if (v.includes("refurb")||v.includes("repair")) return "Refurbishment";
    return null;
  }
  function normalizeRow(row) {
    const o={};
    Object.keys(row).forEach(k => { o[k.trim().toLowerCase().replace(/\s+/g,"_")] = String(row[k]||"").trim(); });
    return o;
  }
  function applyRows(rows) {
    const norm = rows.map(normalizeRow).filter(r => r.serial_number||r.serial);
    const added=[], skipped=[], warnings=[];
    const existing = new Set(devices.map(d => d.serial.toLowerCase()));
    const newDevs = [];
    norm.forEach(row => {
      const serial = (row.serial_number||row.serial||"").trim();
      if (!serial) { skipped.push("(empty)"); return; }
      if (existing.has(serial.toLowerCase())) { skipped.push(serial); return; }
      const rawType = row.device_type||row.type||"";
      const type = resolveType(rawType) || DEVICE_TYPES[0];
      if (rawType && !resolveType(rawType)) warnings.push(`${serial}: unknown type → Router/Modem`);
      const mac    = row.mac_address||row.mac||"";
      const mfr    = lookupOUI(mac);
      const model  = row.model||row.device_model||(mfr?`${mfr} - `:"");
      const rawAction = row.action||row.routing||"";
      const pendingAction = resolveAction(rawAction) || null; // null = unassigned in queue
      if (rawAction && !resolveAction(rawAction)) warnings.push(`${serial}: unknown action "${rawAction}" — unassigned in queue`);
      newDevs.push({
        id: genId(), serial, type, mac, model, stage:"Triage", outcome:null,
        received: row.received_date||row.date||today(), notes:row.notes||"",
        sentToPartner:false, partnerOutcome:null, partnerNotes:"", pendingAction,
      });
      existing.add(serial.toLowerCase());
      added.push(serial);
    });
    if (newDevs.length > 0) {
      bulkInsertDevices(newDevs.map(deviceToRow)).catch(e=>console.warn("DB bulk intake:",e));
    }
    setDevices(p => [...newDevs, ...p]);
    const log = {
      id:Date.now(), timestamp:new Date().toLocaleString(), total:norm.length,
      added:added.length, skipped:skipped.length, warnings:warnings.length,
      skippedSerials:skipped, warningMessages:warnings,
    };
    createUploadLog({ log_type:"intake", total_rows:norm.length, added:added.length, skipped:skipped.length })
      .catch(e=>console.warn("DB upload log:",e));
    setUploadLogs(p => [log, ...p]);
    setUploadResult(log);
  }
  function parseFile(file) {
    setProcessing(true); setUploadResult(null);
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext==="csv") {
      Papa.parse(file,{ header:true, skipEmptyLines:true,
        complete:(r)=>{ applyRows(r.data); setProcessing(false); }, error:()=>setProcessing(false) });
    } else if (ext==="xlsx"||ext==="xls") {
      const rd = new FileReader();
      rd.onload = e => { const wb=XLSX.read(e.target.result,{type:"array"}); applyRows(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""})); setProcessing(false); };
      rd.readAsArrayBuffer(file);
    } else setProcessing(false);
  }
  function downloadTemplate() {
    const rows=[
      ["serial_number","device_type","mac_address","model","received_date","action","notes"],
      ["SN-10001","Router/Modem","E8:65:D4:11:22:33","Huawei HG8245H","2025-05-20","Refurbishment","Customer return"],
      ["SN-10002","Set-top Box","","","2025-05-20","Scrap","Physical damage"],
      ["SN-10003","ONT/OLT","","","2025-05-20","Refurbishment",""],
    ];
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([rows.map(r=>r.join(",")).join("\n")],{type:"text/csv"}));
    a.download = "intake_template.csv"; a.click();
  }

  const tabStyle = active => ({
    padding:"8px 20px", border:"none", borderRadius:8, fontSize:13, fontWeight:700,
    cursor:"pointer", background:active?C.indigo:"transparent",
    color:active?C.white:C.slate3, transition:"all .15s"
  });

  // ── Action badge style ──
  function actionBadge(action) {
    if (!action)          return { bg:"#F1F5F9", color:C.slate3,  label:"— Unassigned" };
    if (action==="Scrap") return { bg:C.redLight, color:C.redDark, label:"🗑 Scrap" };
    if (action==="ECUS")  return { bg:"#FFF9EB",  color:"#92600A", label:"⏸ ECUS" };
    return { bg:"#EEF2FF", color:"#1E40AF", label:"🔧 Refurbishment" };
  }

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      <div>
        <h2 style={{ margin:"0 0 2px", fontSize:isMobile?18:22, fontWeight:800, color:C.slate }}>Intake & Triage</h2>
        <p style={{ margin:0, color:C.slate4, fontSize:13 }}>Register returns, assign actions, then execute the queue</p>
      </div>

      {/* ── Register card ── */}
      <DeviceSearchBar devices={devices} isMobile={isMobile}/>

      <Card>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
          <SectionTitle>Register returned devices</SectionTitle>
          <div style={{ display:"flex", background:C.slate7, borderRadius:10, padding:3, gap:2 }}>
            <button style={tabStyle(tab==="manual")} onClick={()=>setTab("manual")}>✏️ Manual</button>
            <button style={tabStyle(tab==="bulk")}   onClick={()=>setTab("bulk")}>📂 Bulk</button>
          </div>
        </div>

        {/* ── MANUAL ENTRY ── */}
        {tab==="manual" && (
          <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
            <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(5,1fr)", gap:10 }}>
              <div>
                <Label>SERIAL NUMBER *</Label>
                <input value={form.serial} onChange={e=>{ setFormError(""); setForm(f=>({...f,serial:e.target.value})); }}
                  onKeyDown={e=>e.key==="Enter"&&saveToQueue()} placeholder="SN-12345"
                  style={iStyle({ borderColor: formError && !form.serial ? C.red : C.slate6 })}/>
              </div>
              <div>
                <Label>DEVICE TYPE</Label>
                <select value={form.type} onChange={e=>setForm(f=>({...f,type:e.target.value}))} style={iStyle()}>
                  {DEVICE_TYPES.map(t=><option key={t}>{t}</option>)}
                </select>
              </div>
              <div>
                <Label>MAC ADDRESS</Label>
                <input value={form.mac} onChange={e=>handleMacChange(e.target.value, setForm)}
                  placeholder="E8:65:D4:11:22:33" style={iStyle({fontFamily:"monospace"})} maxLength={17}/>
              </div>
              <div>
                <Label>DEVICE MODEL</Label>
                <input value={form.model} onChange={e=>setForm(f=>({...f,model:e.target.value}))}
                  placeholder={lookupOUI(form.mac)?`e.g. ${lookupOUI(form.mac)} HG8245H`:"Will auto-fill from MAC"}
                  style={iStyle({ background: lookupOUI(form.mac)&&!form.model?"#FFFBEB":C.white })}/>
              </div>
              <div>
                <Label>ACTION *</Label>
                <select value={form.action} onChange={e=>setForm(f=>({...f,action:e.target.value}))}
                  style={iStyle({ fontWeight:700,
                    background: form.action==="Scrap" ? "#FFF0F0" : form.action==="ECUS" ? "#FFF9EB" : "#EFF6FF",
                    color: form.action==="Scrap" ? C.redDark : form.action==="ECUS" ? "#92600A" : "#1E40AF",
                    borderColor: form.action==="Scrap" ? "#FECACA" : form.action==="ECUS" ? "#FCD34D" : "#BFDBFE" })}>
                  <option value="Refurbishment">🔧 Refurbishment</option>
                  <option value="Scrap">🗑 Scrap</option>
                  <option value="ECUS">⏸ ECUS</option>
                </select>
              </div>
            </div>
            {(lookupOUI(form.mac) || _ouiCache.get(extractOUI(form.mac)||"")) && (
              <Alert type="success">🔍 OUI match: <strong>{lookupOUI(form.mac) || _ouiCache.get(extractOUI(form.mac)||"")}</strong> — complete the model number above</Alert>
            )}
            <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr":"1fr auto", gap:10, alignItems:"end" }}>
              <div>
                <Label>NOTES</Label>
                <input value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} placeholder="Optional notes" style={iStyle()}/>
              </div>
              <Btn onClick={saveToQueue} size={isMobile?"lg":"md"} full={isMobile} style={{ background:"#16A34A", color:"#fff", border:"none", fontWeight:800, boxShadow:"0 2px 8px rgba(22,163,74,.3)" }}>
                💾 Save to Queue
              </Btn>
            </div>
            {formError && <Alert type="danger">⚠ {formError}</Alert>}
            {ok && <Alert type="success">✓ Device {ok} saved to triage queue with action: {form.action||"Refurbishment"}</Alert>}
          </div>
        )}

        {/* ── BULK UPLOAD ── */}
        {tab==="bulk" && (
          <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8 }}>
              <span style={{ fontSize:13, color:C.slate3 }}>
                Upload CSV or Excel. Include an <code style={{ background:C.slate7, padding:"1px 6px", borderRadius:4, fontSize:12 }}>action</code> column
                (Refurbishment / Scrap) — or leave it blank to assign in the queue. Duplicates skipped automatically.
              </span>
              <Btn onClick={downloadTemplate} variant="purple" size="sm">↓ Template</Btn>
            </div>
            <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)}
              onDrop={e=>{e.preventDefault();setDragOver(false);const f=e.dataTransfer.files[0];if(f)parseFile(f);}}
              onClick={()=>document.getElementById("intake-file").click()}
              style={{ border:`2px dashed ${dragOver?C.indigo:C.slate5}`, borderRadius:12, padding:"28px 16px",
                textAlign:"center", background:dragOver?C.indigoLight:C.slate8, transition:"all .2s", cursor:"pointer" }}>
              <input id="intake-file" type="file" accept=".csv,.xlsx,.xls"
                onChange={e=>{const f=e.target.files[0];if(f)parseFile(f);e.target.value="";}} style={{ display:"none" }}/>
              <div style={{ fontSize:28, marginBottom:6 }}>{processing?"⏳":"📋"}</div>
              <div style={{ fontSize:13, fontWeight:700, color:C.slate2, marginBottom:4 }}>
                {processing?"Processing…":"Drop file here or tap to browse"}
              </div>
              <div style={{ fontSize:11, color:C.slate4 }}>
                serial_number · device_type · mac_address · model · action · received_date · notes
              </div>
            </div>

            {/* Column reference */}
            <div style={{ background:C.slate8, borderRadius:8, padding:"10px 14px" }}>
              <div style={{ fontSize:10, fontWeight:700, color:C.slate4, letterSpacing:".06em", marginBottom:8 }}>ACCEPTED COLUMNS</div>
              <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr":"1fr 1fr", gap:6 }}>
                {[
                  ["serial_number","Required · must be unique"],
                  ["device_type","Router/Modem · Set-top Box · ONT/OLT"],
                  ["action","Refurbishment or Scrap (optional — can assign in queue)"],
                  ["mac_address","Optional · triggers OUI manufacturer lookup"],
                  ["model","Optional · device model name"],
                  ["received_date","Optional · defaults to today"],
                  ["notes","Optional · free text"],
                ].map(([col,desc])=>(
                  <div key={col} style={{ fontSize:12 }}>
                    <code style={{ background:C.slate7, padding:"1px 6px", borderRadius:4, fontSize:11 }}>{col}</code>
                    <span style={{ color:C.slate4, marginLeft:6 }}>{desc}</span>
                  </div>
                ))}
              </div>
            </div>

            {uploadResult && (
              <div style={{ border:`1.5px solid ${C.slate6}`, borderRadius:12, padding:14 }}>
                <div style={{ fontSize:13, fontWeight:700, color:C.slate, marginBottom:10 }}>
                  Upload complete · {uploadResult.timestamp}
                </div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:8, marginBottom:10 }}>
                  {[[uploadResult.added,"Added",C.greenDark,C.greenLight],[uploadResult.skipped,"Skipped","#92400E",C.amberLight],[uploadResult.warnings,"Warnings",C.purple,C.purpleLight]].map(([n,l,c,bg])=>(
                    <div key={l} style={{ background:bg, borderRadius:8, padding:10, textAlign:"center" }}>
                      <div style={{ fontSize:22, fontWeight:800, color:c }}>{n}</div>
                      <div style={{ fontSize:11, color:c, fontWeight:700 }}>{l}</div>
                    </div>
                  ))}
                </div>
                {uploadResult.skippedSerials.length>0 && <Alert type="warning">Skipped: {uploadResult.skippedSerials.join(", ")}</Alert>}
                {uploadResult.warningMessages.length>0 && (
                  <div style={{ marginTop:8 }}>
                    {uploadResult.warningMessages.map((w,i)=><Alert key={i} type="warning">⚠ {w}</Alert>)}
                  </div>
                )}
                {unassigned>0 && <Alert type="info" style={{ marginTop:8 }}>ℹ {unassigned} device{unassigned>1?"s":""} have no action assigned — set them in the queue below before executing.</Alert>}
              </div>
            )}
          </div>
        )}
      </Card>

      {/* ── Triage Queue ── */}
      <Card>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <SectionTitle>Triage queue ({queue.length})</SectionTitle>
            {unassigned>0 && (
              <span style={{ background:C.amberLight, color:"#92400E", fontSize:11, fontWeight:700,
                padding:"2px 9px", borderRadius:20 }}>{unassigned} unassigned</span>
            )}
          </div>
          <Btn onClick={()=>setShowConfirm(true)} size={isMobile?"lg":"md"} disabled={readyCount===0}
            style={{ background:readyCount>0?"#16A34A":"#D1FAE5", color:readyCount>0?"#fff":"#6EE7B7",
              border:"none", fontWeight:800, letterSpacing:".02em",
              boxShadow:readyCount>0?"0 2px 8px rgba(22,163,74,.35)":"none",
              transition:"all .2s" }}>
            ▶ Execute Queue {readyCount>0?`(${readyCount})`:""}
          </Btn>
        </div>

        {queue.length===0
          ? <p style={{ color:C.slate4, fontSize:13, margin:0 }}>No devices awaiting triage. Register devices above to get started.</p>
          : isMobile
            ? queue.map(d => {
                const isEditing  = editingId===d.id;
                const isDeleting = deleteId===d.id;
                const expanded   = expandedId===d.id;
                const ab = actionBadge(d.pendingAction);

                if (isEditing) return (
                  <div key={d.id} style={{ border:`2px solid ${C.indigo}`, borderRadius:12, padding:14, marginBottom:10, background:C.indigoLight }}>
                    <div style={{ fontSize:12, fontWeight:700, color:C.indigo, marginBottom:10 }}>Editing {d.id}</div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginBottom:8 }}>
                      <div><Label>SERIAL</Label><input value={editBuf.serial} onChange={e=>setEditBuf(b=>({...b,serial:e.target.value}))} style={iStyle()}/></div>
                      <div><Label>TYPE</Label>
                        <select value={editBuf.type} onChange={e=>setEditBuf(b=>({...b,type:e.target.value}))} style={iStyle()}>
                          {DEVICE_TYPES.map(t=><option key={t}>{t}</option>)}
                        </select>
                      </div>
                      <div><Label>MAC</Label><input value={editBuf.mac} onChange={e=>{ const mac=formatMac(e.target.value); const mfr=lookupOUI(mac); setEditBuf(b=>({...b,mac,model:mfr&&!b.model?`${mfr} - `:b.model})); }} style={iStyle({fontFamily:"monospace"})} maxLength={17}/></div>
                      <div><Label>MODEL</Label><input value={editBuf.model} onChange={e=>setEditBuf(b=>({...b,model:e.target.value}))} style={iStyle()}/></div>
                      <div><Label>ACTION</Label>
                        <select value={editBuf.pendingAction||"Refurbishment"} onChange={e=>setEditBuf(b=>({...b,pendingAction:e.target.value}))} style={iStyle()}>
                          <option value="Refurbishment">🔧 Refurbishment</option>
                          <option value="Scrap">🗑 Scrap</option>
                          <option value="ECUS">⏸ ECUS</option>
                        </select>
                      </div>
                      <div><Label>NOTES</Label><input value={editBuf.notes} onChange={e=>setEditBuf(b=>({...b,notes:e.target.value}))} style={iStyle()}/></div>
                    </div>
                    <div style={{ display:"flex", gap:8 }}>
                      <Btn onClick={()=>saveEdit(d.id)} variant="success" full>✓ Save</Btn>
                      <Btn onClick={()=>setEditingId(null)} variant="ghost" full>Cancel</Btn>
                    </div>
                  </div>
                );
                if (isDeleting) return (
                  <div key={d.id} style={{ border:`2px solid ${C.red}`, borderRadius:12, padding:14, marginBottom:10, background:C.redLight }}>
                    <p style={{ margin:"0 0 10px", fontSize:13, fontWeight:600, color:C.redDark }}>Delete {d.serial}? This cannot be undone.</p>
                    <div style={{ display:"flex", gap:8 }}>
                      <Btn onClick={()=>deleteDevice(d.id)} variant="danger" full>Delete</Btn>
                      <Btn onClick={()=>setDeleteId(null)} variant="ghost" full>Cancel</Btn>
                    </div>
                  </div>
                );
                return (
                  <div key={d.id} style={{ border:`1px solid ${C.slate6}`, borderRadius:12, padding:14, marginBottom:10,
                    borderLeft:`3px solid ${d.pendingAction==="Scrap"?C.red:d.pendingAction?C.indigo:C.slate5}` }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}>
                      <div>
                        <span style={{ fontWeight:800, color:C.indigo, fontSize:12 }}>{d.id}</span>
                        <span style={{ fontFamily:"monospace", fontSize:12, color:C.slate2, marginLeft:8 }}>{d.serial}</span>
                      </div>
                      <button onClick={()=>setExpandedId(expanded?null:d.id)}
                        style={{ background:"none", border:"none", cursor:"pointer", fontSize:14, color:C.slate4 }}>
                        {expanded?"▲":"▼"}
                      </button>
                    </div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6, marginBottom:10 }}>
                      <div><span style={{ fontSize:10, fontWeight:700, color:C.slate4, textTransform:"uppercase" }}>Type</span><div style={{ fontSize:12 }}>{TYPE_ICON[d.type]} {d.type}</div></div>
                      <div><span style={{ fontSize:10, fontWeight:700, color:C.slate4, textTransform:"uppercase" }}>Received</span><div style={{ fontSize:12, color:C.slate3 }}>{d.received}</div></div>
                      {expanded && <>
                        <div><span style={{ fontSize:10, fontWeight:700, color:C.slate4, textTransform:"uppercase" }}>MAC</span><div style={{ fontSize:11, fontFamily:"monospace" }}>{d.mac||"—"}</div></div>
                        <div><span style={{ fontSize:10, fontWeight:700, color:C.slate4, textTransform:"uppercase" }}>Model</span><div style={{ fontSize:12 }}>{d.model||"—"}</div></div>
                        {d.notes && <div style={{ gridColumn:"1/-1" }}><span style={{ fontSize:10, fontWeight:700, color:C.slate4, textTransform:"uppercase" }}>Notes</span><div style={{ fontSize:12, color:C.slate3 }}>{d.notes}</div></div>}
                      </>}
                    </div>
                    {/* Action status badge — read-only, editable via ✏️ Edit */}
                    <div style={{ marginBottom:10 }}>
                      <span style={{ fontSize:10, fontWeight:700, color:C.slate4, textTransform:"uppercase", letterSpacing:".05em" }}>ACTION</span>
                      <div style={{ marginTop:4 }}>
                        {d.pendingAction
                          ? <span style={{ display:"inline-flex", alignItems:"center", gap:5, padding:"5px 12px", borderRadius:20, fontSize:12, fontWeight:700,
                              background: d.pendingAction==="Scrap" ? C.redLight : C.indigoLight,
                              color: d.pendingAction==="Scrap" ? C.redDark : C.indigoDark,
                              border: `1px solid ${d.pendingAction==="Scrap" ? "#FECACA" : "#C7D2FE"}` }}>
                              {d.pendingAction==="Scrap" ? "🗑 Scrap" : d.pendingAction==="ECUS" ? "⏸ ECUS" : "🔧 Refurbishment"}
                            </span>
                          : <span style={{ display:"inline-flex", alignItems:"center", gap:5, padding:"5px 12px", borderRadius:20, fontSize:12, fontWeight:700,
                              background:C.slate7, color:C.slate4, border:`1px solid ${C.slate6}` }}>
                              — Unassigned
                            </span>
                        }
                      </div>
                    </div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6 }}>
                      <Btn onClick={()=>startEdit(d)} variant="ghost" size="sm" full>✏️ Edit</Btn>
                      <Btn onClick={()=>setDeleteId(d.id)} variant="ghost" size="sm" full>🗑️ Delete</Btn>
                    </div>
                  </div>
                );
              })
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead>
                    <tr style={{ borderBottom:`1.5px solid ${C.slate6}`, background:C.slate8 }}>
                      {["ID","Serial","Type","MAC","Model","Notes","Action","Controls"].map(h => (
                        <th key={h} style={{ padding:"8px 10px", textAlign:"left", fontSize:11,
                          fontWeight:700, color:C.slate4, letterSpacing:".05em", whiteSpace:"nowrap" }}>
                          {h.toUpperCase()}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {queue.map(d => {
                      const isEditing  = editingId===d.id;
                      const isDeleting = deleteId===d.id;
                      const ab = actionBadge(d.pendingAction);
                      return (
                        <tr key={d.id} style={{ borderBottom:`1px solid ${C.slate7}`,
                          background: isEditing?C.indigoLight : isDeleting?C.redLight : "transparent",
                          borderLeft:`3px solid ${d.pendingAction==="Scrap"?C.red:d.pendingAction?C.indigo:C.slate5}` }}>
                          <td style={{ padding:"9px 10px", fontWeight:700, color:C.indigo, whiteSpace:"nowrap" }}>{d.id}</td>

                          {isEditing ? (
                            <>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.serial} onChange={e=>setEditBuf(b=>({...b,serial:e.target.value}))} style={{ ...iStyle(), width:110 }}/></td>
                              <td style={{ padding:"5px 6px" }}><select value={editBuf.type} onChange={e=>setEditBuf(b=>({...b,type:e.target.value}))} style={{ ...iStyle(), width:120 }}>{DEVICE_TYPES.map(t=><option key={t}>{t}</option>)}</select></td>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.mac} onChange={e=>{ const mac=formatMac(e.target.value); const mfr=lookupOUI(mac); setEditBuf(b=>({...b,mac,model:mfr&&!b.model?`${mfr} - `:b.model})); }} style={{ ...iStyle({fontFamily:"monospace"}), width:130 }} maxLength={17}/></td>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.model} onChange={e=>setEditBuf(b=>({...b,model:e.target.value}))} style={{ ...iStyle(), width:130 }}/></td>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.notes} onChange={e=>setEditBuf(b=>({...b,notes:e.target.value}))} style={{ ...iStyle(), width:110 }}/></td>
                              <td style={{ padding:"5px 6px" }}>
                                <select value={editBuf.pendingAction||"Refurbishment"} onChange={e=>setEditBuf(b=>({...b,pendingAction:e.target.value}))} style={{ ...iStyle(), width:130 }}>
                                  <option value="Refurbishment">🔧 Refurbishment</option>
                                  <option value="Scrap">🗑 Scrap</option>
                                  <option value="ECUS">⏸ ECUS</option>
                                </select>
                              </td>
                              <td style={{ padding:"5px 6px" }}>
                                <div style={{ display:"flex", gap:5 }}>
                                  <Btn onClick={()=>saveEdit(d.id)} variant="success" size="sm">✓ Save</Btn>
                                  <Btn onClick={()=>setEditingId(null)} variant="ghost" size="sm">Cancel</Btn>
                                </div>
                              </td>
                            </>
                          ) : isDeleting ? (
                            <>
                              <td colSpan={5} style={{ padding:"9px 10px", color:C.redDark, fontWeight:600 }}>
                                Delete <span style={{ fontFamily:"monospace" }}>{d.serial}</span>? This cannot be undone.
                              </td>
                              <td colSpan={2} style={{ padding:"5px 6px" }}>
                                <div style={{ display:"flex", gap:5 }}>
                                  <Btn onClick={()=>deleteDevice(d.id)} variant="danger" size="sm">Delete</Btn>
                                  <Btn onClick={()=>setDeleteId(null)} variant="ghost" size="sm">Cancel</Btn>
                                </div>
                              </td>
                            </>
                          ) : (
                            <>
                              <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:12 }}>{d.serial}</td>
                              <td style={{ padding:"9px 10px", whiteSpace:"nowrap" }}>{TYPE_ICON[d.type]} {d.type}</td>
                              <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:11, color:C.slate3 }}>{d.mac||"—"}</td>
                              <td style={{ padding:"9px 10px", fontSize:12 }}>
                                {d.model ? d.model : d.mac&&lookupOUI(d.mac)
                                  ? <span style={{ color:C.amber, fontStyle:"italic" }}>{lookupOUI(d.mac)} - ?</span>
                                  : <span style={{ color:C.slate5 }}>—</span>}
                              </td>
                              <td style={{ padding:"9px 10px", color:C.slate3, maxWidth:120, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{d.notes||"—"}</td>
                              {/* Action status — read-only badge */}
                              <td style={{ padding:"9px 10px" }}>
                                {d.pendingAction
                                  ? <span style={{ display:"inline-flex", alignItems:"center", gap:4, padding:"3px 10px", borderRadius:20, fontSize:11, fontWeight:700, whiteSpace:"nowrap",
                                      background: d.pendingAction==="Scrap" ? C.redLight : C.indigoLight,
                                      color: d.pendingAction==="Scrap" ? C.redDark : C.indigoDark,
                                      border:`1px solid ${d.pendingAction==="Scrap"?"#FECACA":"#C7D2FE"}` }}>
                                      {d.pendingAction==="Scrap" ? "🗑 Scrap" : d.pendingAction==="ECUS" ? "⏸ ECUS" : "🔧 Refurbishment"}
                                    </span>
                                  : <span style={{ display:"inline-flex", alignItems:"center", padding:"3px 10px", borderRadius:20, fontSize:11, fontWeight:600,
                                      background:C.slate7, color:C.slate4, border:`1px solid ${C.slate6}` }}>
                                      — Unassigned
                                    </span>
                                }
                              </td>
                              <td style={{ padding:"5px 8px" }}>
                                <div style={{ display:"flex", gap:4 }}>
                                  <Btn onClick={()=>startEdit(d)} variant="ghost" size="sm">✏️</Btn>
                                  <Btn onClick={()=>setDeleteId(d.id)} variant="ghost" size="sm">🗑️</Btn>
                                </div>
                              </td>
                            </>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )
        }

        {/* Execute Queue button at bottom */}
        {queue.length > 0 && (
          <div style={{ marginTop:16, paddingTop:14, borderTop:`1px solid ${C.slate7}`,
            display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:10 }}>
            <div style={{ fontSize:13, color:C.slate3 }}>
              {readyCount>0
                ? <span><strong style={{ color:C.indigoDark }}>{toRefurb}</strong> → Refurb &nbsp;·&nbsp; <strong style={{ color:C.redDark }}>{toScrap}</strong> → Scrap &nbsp;·&nbsp; <strong style={{ color:"#92600A" }}>{toEcus}</strong> → ECUS {unassigned>0 && <span style={{ color:"#92400E" }}>· <strong>{unassigned}</strong> unassigned</span>}</span>
                : <span style={{ color:C.slate4 }}>Assign actions to devices above, then execute.</span>
              }
            </div>
            <Btn onClick={()=>setShowConfirm(true)} disabled={readyCount===0}
              style={{ background:readyCount>0?"#16A34A":"#D1FAE5", color:readyCount>0?"#fff":"#6EE7B7",
                border:"none", fontWeight:800, letterSpacing:".02em",
                boxShadow:readyCount>0?"0 2px 8px rgba(22,163,74,.35)":"none",
                transition:"all .2s" }}>
              ▶ Execute Queue {readyCount>0?`(${readyCount})`:""}
            </Btn>
          </div>
        )}
      </Card>

      {/* ── CONFIRMATION POPUP ── */}
      {showConfirm && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.55)", zIndex:100,
          display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
          <div style={{ background:C.white, borderRadius:16, padding:28, maxWidth:440, width:"100%",
            boxShadow:"0 20px 60px rgba(0,0,0,.25)" }}>
            <div style={{ fontSize:24, marginBottom:8 }}>⚙️</div>
            <h3 style={{ margin:"0 0 6px", fontSize:18, fontWeight:800, color:C.slate }}>Execute Triage Queue</h3>
            <p style={{ margin:"0 0 20px", fontSize:13, color:C.slate3 }}>
              This will dispatch all assigned devices to their target stages. This action cannot be undone.
            </p>

            {/* Summary */}
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:20 }}>
              <div style={{ background:C.indigoLight, borderRadius:10, padding:"14px 16px",
                border:`1.5px solid #C7D2FE`, textAlign:"center" }}>
                <div style={{ fontSize:28, fontWeight:800, color:C.indigo }}>{toRefurb}</div>
                <div style={{ fontSize:12, fontWeight:700, color:C.indigoDark }}>🔧 → Refurbishment</div>
                <div style={{ fontSize:11, color:C.slate4, marginTop:2 }}>Sent to partner</div>
              </div>
              <div style={{ background:C.redLight, borderRadius:10, padding:"14px 16px",
                border:`1.5px solid #FECACA`, textAlign:"center" }}>
                <div style={{ fontSize:28, fontWeight:800, color:C.red }}>{toScrap}</div>
                <div style={{ fontSize:12, fontWeight:700, color:C.redDark }}>🗑 → Scrap</div>
                <div style={{ fontSize:11, color:C.slate4, marginTop:2 }}>Removed from circulation</div>
              </div>
            </div>

            {unassigned>0 && (
              <Alert type="warning">
                ⚠ <strong>{unassigned}</strong> device{unassigned>1?"s":""} have no action assigned and will stay in the queue.
              </Alert>
            )}

            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginTop:16 }}>
              <Btn onClick={()=>setShowConfirm(false)} variant="ghost" full size="lg">Cancel</Btn>
              <Btn onClick={executeQueue} full size="lg" disabled={readyCount===0}
                style={{ background:"#16A34A", color:"#fff", border:"none", fontWeight:800,
                  boxShadow:"0 2px 10px rgba(22,163,74,.4)", transition:"all .2s" }}>
                ✓ Confirm & Execute
              </Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
// ─────────────────────────────────────────────
// REFURB TRACKING
// ─────────────────────────────────────────────
function RefurbTracking({ devices, setDevices, isMobile }) {
  const inRefurb    = devices.filter(d=>d.stage==="Refurbishment" && !d.partnerOutcome);
  const pendingConf = devices.filter(d=>d.stage==="Refurbishment" && d.partnerOutcome);
  const inQC        = devices.filter(d=>d.stage==="QC Check");
  const escalated   = devices.filter(d=>d.stage==="Escalated");

  // In Refurbishment row state
  const [editingId, setEditingId] = useState(null);
  const [editBuf,   setEditBuf]   = useState({});
  const [deleteId,  setDeleteId]  = useState(null);
  const [returnId,  setReturnId]  = useState(null);

  function setOutcome(id, outcome) {
    const newStage = outcome==="Working" ? "Stock" : "Escalated";
    setDevices(p=>p.map(d=>d.id===id?{...d,stage:newStage,outcome}:d));
    updateDevice(id, { stage:newStage, outcome }).catch(e=>console.warn("DB setOutcome:",e));
  }
  function escalateToScrap(id) {
    setDevices(p=>p.map(d=>d.id===id?{...d,stage:"Scrap",outcome:"Scrap"}:d));
    updateDevice(id, { stage:"Scrap", outcome:"Scrap" }).catch(e=>console.warn("DB escalateToScrap:",e));
  }
  function requeue(id) {
    setDevices(p=>p.map(d=>d.id===id?{...d,stage:"Refurbishment",outcome:null,partnerOutcome:null,partnerNotes:""}:d));
    updateDevice(id, { stage:"Refurbishment", outcome:null, partner_outcome:null, partner_notes:"" }).catch(e=>console.warn("DB requeue:",e));
  }

  function confirmPartner(id) {
    // Confirmed devices go to "Confirmed" stage (awaiting physical return)
    // They will move to Stock or Scrap only after physical receipt is confirmed
    setDevices(p=>p.map(d=>{
      if (d.id!==id) return d;
      const newNotes = d.partnerNotes
        ? (d.notes ? d.notes+" | Partner: "+d.partnerNotes : "Partner: "+d.partnerNotes)
        : d.notes;
      updateDevice(id, { outcome:d.partnerOutcome, stage:"Confirmed",
        notes:newNotes, partner_outcome:null, partner_notes:"" }
      ).catch(e=>console.warn("DB confirmPartner:",e));
      return { ...d, outcome:d.partnerOutcome, stage:"Confirmed",
        notes:newNotes, partnerOutcome:null, partnerNotes:"" };
    }));
  }

  function rejectPartner(id) {
    setDevices(p=>p.map(d=>d.id===id?{...d,partnerOutcome:null,partnerNotes:""}:d));
    updateDevice(id, { partner_outcome:null, partner_notes:"" }).catch(e=>console.warn("DB rejectPartner:",e));
  }

  // Return device to Triage queue (unassigned, ready to re-triage)
  function returnToTriage(id) {
    setDevices(p=>p.map(d=>d.id===id
      ? { ...d, stage:"Triage", outcome:null, sentToPartner:false,
          partnerOutcome:null, partnerNotes:"", pendingAction:null }
      : d));
    updateDevice(id, { stage:"Triage", outcome:null, sent_to_partner:false,
      partner_outcome:null, partner_notes:"", pending_action:null
    }).catch(e=>console.warn("DB returnToTriage:",e));
    setReturnId(null);
  }

  function saveEdit(id) {
    setDevices(p=>p.map(d=>d.id===id ? { ...d,...editBuf } : d));
    updateDevice(id, deviceToRow({...editBuf, id})).catch(e=>console.warn("DB saveEdit:",e));
    setEditingId(null);
  }

  function deleteDevice(id) {
    setDevices(p=>p.filter(d=>d.id!==id));
    setDeleteId(null);
  }

  function DevSection({ title, items, accent, renderActions, extra }) {
    return (
      <div style={{ border:`1.5px solid ${accent}22`, borderTop:`3px solid ${accent}`, borderRadius:14, padding:16, background:C.white }}>
        <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:14 }}>
          <h3 style={{ margin:0, fontSize:14, fontWeight:700, color:C.slate }}>{title}</h3>
          <span style={{ fontSize:12, color:C.slate4 }}>({items.length})</span>
        </div>
        {extra}
        {items.length===0
          ? <p style={{ color:C.slate4, fontSize:13, margin:0 }}>None at this stage.</p>
          : isMobile
            ? items.map(d=>(
                <DeviceCard key={d.id} d={d}
                  fields={[["Model",d.model],["Type",`${TYPE_ICON[d.type]} ${d.type}`],["Received",d.received],["Notes",d.notes]]}
                  actions={renderActions(d)}/>
              ))
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead><tr style={{ borderBottom:`1.5px solid ${C.slate6}` }}>
                    {["ID","Serial","Model","Type","Received","Notes","Action"].map(h=>(
                      <th key={h} style={{ padding:"6px 10px", textAlign:"left", fontSize:11, fontWeight:700, color:C.slate4, letterSpacing:".05em", whiteSpace:"nowrap" }}>{h.toUpperCase()}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {items.map(d=>(
                      <tr key={d.id} style={{ borderBottom:`1px solid ${C.slate7}` }}>
                        <td style={{ padding:"9px 10px", fontWeight:700, color:C.indigo, whiteSpace:"nowrap" }}>{d.id}</td>
                        <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:12 }}>{d.serial}</td>
                        <td style={{ padding:"9px 10px", fontSize:12 }}>{d.model||"—"}</td>
                        <td style={{ padding:"9px 10px", whiteSpace:"nowrap" }}>{TYPE_ICON[d.type]} {d.type}</td>
                        <td style={{ padding:"9px 10px", color:C.slate3, whiteSpace:"nowrap" }}>{d.received}</td>
                        <td style={{ padding:"9px 10px", color:C.slate3, maxWidth:140, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{d.notes||"—"}</td>
                        <td style={{ padding:"9px 10px" }}><div style={{ display:"flex", gap:6 }}>{renderActions(d)}</div></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
        }
      </div>
    );
  }

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      <div>
        <h2 style={{ margin:"0 0 2px", fontSize:isMobile?18:22, fontWeight:800, color:C.slate }}>Refurbishment</h2>
        <p style={{ margin:0, color:C.slate4, fontSize:13 }}>Track devices through repair, QC, and confirmation</p>
      </div>

      {/* Pending partner confirmation */}
      {pendingConf.length>0 && (
        <div style={{ border:`2px solid #FCD34D`, borderTop:`3px solid ${C.amber}`, borderRadius:14, padding:16, background:C.white }}>
          <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:14 }}>
            <span style={{ fontSize:16 }}>⚠️</span>
            <h3 style={{ margin:0, fontSize:14, fontWeight:700, color:"#92400E" }}>Partner Results — Pending Confirmation ({pendingConf.length})</h3>
          </div>
          {isMobile
            ? pendingConf.map(d=>(
                <div key={d.id} style={{ border:`1px solid #FCD34D`, borderRadius:12, padding:12, marginBottom:10, background:"#FFFBEB" }}>
                  <div style={{ display:"flex", justifyContent:"space-between", marginBottom:8 }}>
                    <span style={{ fontWeight:800, color:C.indigo, fontSize:12 }}>{d.id} · {d.serial}</span>
                    <span style={{ fontSize:12, fontWeight:700, color:d.partnerOutcome==="Working"?C.greenDark:C.redDark,
                      background:d.partnerOutcome==="Working"?C.greenLight:C.redLight, padding:"2px 8px", borderRadius:20 }}>
                      {d.partnerOutcome==="Working"?"✓ Working":"✗ Not Working"}
                    </span>
                  </div>
                  {d.partnerNotes && <p style={{ margin:"0 0 8px", fontSize:12, color:C.slate3 }}>Notes: {d.partnerNotes}</p>}
                  <div style={{ display:"flex", gap:6 }}>
                    <Btn onClick={()=>confirmPartner(d.id)} variant="success" size="sm" full>✓ Confirm</Btn>
                    <Btn onClick={()=>rejectPartner(d.id)} variant="danger" size="sm" full>✕ Reject</Btn>
                  </div>
                </div>
              ))
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead><tr style={{ borderBottom:`1.5px solid #FCD34D` }}>
                    {["ID","Serial","Model","Type","Partner Outcome","Partner Notes","Action"].map(h=>(
                      <th key={h} style={{ padding:"6px 10px", textAlign:"left", fontSize:11, fontWeight:700, color:"#92400E", letterSpacing:".05em", whiteSpace:"nowrap" }}>{h.toUpperCase()}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {pendingConf.map(d=>(
                      <tr key={d.id} style={{ borderBottom:`1px solid #FEF3C7` }}>
                        <td style={{ padding:"9px 10px", fontWeight:700, color:C.indigo }}>{d.id}</td>
                        <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:12 }}>{d.serial}</td>
                        <td style={{ padding:"9px 10px", fontSize:12 }}>{d.model||"—"}</td>
                        <td style={{ padding:"9px 10px" }}>{TYPE_ICON[d.type]} {d.type}</td>
                        <td style={{ padding:"9px 10px" }}>
                          <span style={{ background:d.partnerOutcome==="Working"?C.greenLight:C.redLight,
                            color:d.partnerOutcome==="Working"?C.greenDark:C.redDark,
                            fontSize:12, fontWeight:700, padding:"3px 10px", borderRadius:20 }}>
                            {d.partnerOutcome==="Working"?"✓ Working":"✗ Not Working"}
                          </span>
                        </td>
                        <td style={{ padding:"9px 10px", color:C.slate3, maxWidth:140, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{d.partnerNotes||"—"}</td>
                        <td style={{ padding:"9px 10px" }}>
                          <div style={{ display:"flex", gap:6 }}>
                            <Btn onClick={()=>confirmPartner(d.id)} variant="success" size="sm">✓ Confirm</Btn>
                            <Btn onClick={()=>rejectPartner(d.id)} variant="danger" size="sm">✕ Reject</Btn>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
          }
        </div>
      )}

      {/* ── IN REFURBISHMENT — custom section with return/edit/delete ── */}
      <div style={{ border:`1.5px solid ${C.amber}22`, borderTop:`3px solid ${C.amber}`, borderRadius:14, padding:16, background:C.white }}>
        <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:14 }}>
          <h3 style={{ margin:0, fontSize:14, fontWeight:700, color:C.slate }}>In Refurbishment</h3>
          <span style={{ fontSize:12, color:C.slate4 }}>({inRefurb.length})</span>
        </div>
        {inRefurb.length===0
          ? <p style={{ color:C.slate4, fontSize:13, margin:0 }}>None at this stage.</p>
          : isMobile
            ? inRefurb.map(d=>{
                const isEditing  = editingId===d.id;
                const isDeleting = deleteId===d.id;
                const isReturning= returnId===d.id;
                if (isEditing) return (
                  <div key={d.id} style={{ border:`2px solid ${C.indigo}`, borderRadius:12, padding:14, marginBottom:10, background:C.indigoLight }}>
                    <div style={{ fontSize:12, fontWeight:700, color:C.indigo, marginBottom:10 }}>Editing {d.id}</div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginBottom:8 }}>
                      <div><Label>SERIAL</Label><input value={editBuf.serial} onChange={e=>setEditBuf(b=>({...b,serial:e.target.value}))} style={iStyle()}/></div>
                      <div><Label>TYPE</Label><select value={editBuf.type} onChange={e=>setEditBuf(b=>({...b,type:e.target.value}))} style={iStyle()}>{DEVICE_TYPES.map(t=><option key={t}>{t}</option>)}</select></div>
                      <div><Label>MAC</Label><input value={editBuf.mac} onChange={e=>{ const mac=formatMac(e.target.value); const mfr=lookupOUI(mac); setEditBuf(b=>({...b,mac,model:mfr&&!b.model?`${mfr} - `:b.model})); }} style={iStyle({fontFamily:"monospace"})} maxLength={17}/></div>
                      <div><Label>MODEL</Label><input value={editBuf.model} onChange={e=>setEditBuf(b=>({...b,model:e.target.value}))} style={iStyle()}/></div>
                    </div>
                    <div style={{ marginBottom:8 }}><Label>NOTES</Label><input value={editBuf.notes} onChange={e=>setEditBuf(b=>({...b,notes:e.target.value}))} style={iStyle()}/></div>
                    <div style={{ display:"flex", gap:8 }}>
                      <Btn onClick={()=>saveEdit(d.id)} variant="success" full>✓ Save</Btn>
                      <Btn onClick={()=>setEditingId(null)} variant="ghost" full>Cancel</Btn>
                    </div>
                  </div>
                );
                if (isDeleting) return (
                  <div key={d.id} style={{ border:`2px solid ${C.red}`, borderRadius:12, padding:14, marginBottom:10, background:C.redLight }}>
                    <p style={{ margin:"0 0 10px", fontSize:13, fontWeight:600, color:C.redDark }}>Delete {d.serial}? This cannot be undone.</p>
                    <div style={{ display:"flex", gap:8 }}>
                      <Btn onClick={()=>deleteDevice(d.id)} variant="danger" full>Delete</Btn>
                      <Btn onClick={()=>setDeleteId(null)} variant="ghost" full>Cancel</Btn>
                    </div>
                  </div>
                );
                if (isReturning) return (
                  <div key={d.id} style={{ border:`2px solid ${C.amber}`, borderRadius:12, padding:14, marginBottom:10, background:C.amberLight }}>
                    <p style={{ margin:"0 0 10px", fontSize:13, fontWeight:600, color:"#92400E" }}>Return {d.serial} to Triage queue? It will be unassigned and need re-triaging.</p>
                    <div style={{ display:"flex", gap:8 }}>
                      <Btn onClick={()=>returnToTriage(d.id)} variant="amber" full>↩ Confirm Return</Btn>
                      <Btn onClick={()=>setReturnId(null)} variant="ghost" full>Cancel</Btn>
                    </div>
                  </div>
                );
                return (
                  <DeviceCard key={d.id} d={d}
                    fields={[["Model",d.model],["Type",`${TYPE_ICON[d.type]} ${d.type}`],["Received",d.received],["Notes",d.notes]]}
                    actions={[
                      <Btn key="ret" onClick={()=>setReturnId(d.id)}  variant="amber"  size="sm">↩ Return to Triage</Btn>,
                      <Btn key="ed"  onClick={()=>{ setEditingId(d.id); setEditBuf({serial:d.serial,type:d.type,mac:d.mac||"",model:d.model||"",notes:d.notes||""}); }} variant="ghost" size="sm">✏️ Edit</Btn>,
                      <Btn key="del" onClick={()=>setDeleteId(d.id)}  variant="ghost"  size="sm">🗑️ Delete</Btn>,
                    ]}/>
                );
              })
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead>
                    <tr style={{ borderBottom:`1.5px solid ${C.slate6}`, background:C.slate8 }}>
                      {["ID","Serial","Model","Type","MAC","Received","Notes","Actions"].map(h=>(
                        <th key={h} style={{ padding:"8px 10px", textAlign:"left", fontSize:11, fontWeight:700, color:C.slate4, letterSpacing:".05em", whiteSpace:"nowrap" }}>{h.toUpperCase()}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {inRefurb.map(d=>{
                      const isEditing  = editingId===d.id;
                      const isDeleting = deleteId===d.id;
                      const isReturning= returnId===d.id;
                      return (
                        <tr key={d.id} style={{ borderBottom:`1px solid ${C.slate7}`,
                          background: isEditing?C.indigoLight : isDeleting?C.redLight : isReturning?"#FFFBEB" : "transparent" }}>
                          <td style={{ padding:"9px 10px", fontWeight:700, color:C.indigo, whiteSpace:"nowrap" }}>{d.id}</td>
                          {isEditing ? (
                            <>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.serial} onChange={e=>setEditBuf(b=>({...b,serial:e.target.value}))} style={{ ...iStyle(), width:100 }}/></td>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.model}  onChange={e=>setEditBuf(b=>({...b,model:e.target.value}))}  style={{ ...iStyle(), width:110 }}/></td>
                              <td style={{ padding:"5px 6px" }}><select value={editBuf.type} onChange={e=>setEditBuf(b=>({...b,type:e.target.value}))} style={{ ...iStyle(), width:120 }}>{DEVICE_TYPES.map(t=><option key={t}>{t}</option>)}</select></td>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.mac} onChange={e=>{ const mac=formatMac(e.target.value); const mfr=lookupOUI(mac); setEditBuf(b=>({...b,mac,model:mfr&&!b.model?`${mfr} - `:b.model})); }} style={{ ...iStyle({fontFamily:"monospace"}), width:130 }} maxLength={17}/></td>
                              <td style={{ padding:"5px 6px" }}>{d.received}</td>
                              <td style={{ padding:"5px 6px" }}><input value={editBuf.notes} onChange={e=>setEditBuf(b=>({...b,notes:e.target.value}))} style={{ ...iStyle(), width:120 }}/></td>
                              <td style={{ padding:"5px 6px" }}>
                                <div style={{ display:"flex", gap:5 }}>
                                  <Btn onClick={()=>saveEdit(d.id)} variant="success" size="sm">✓ Save</Btn>
                                  <Btn onClick={()=>setEditingId(null)} variant="ghost" size="sm">Cancel</Btn>
                                </div>
                              </td>
                            </>
                          ) : isDeleting ? (
                            <>
                              <td colSpan={5} style={{ padding:"9px 10px", color:C.redDark, fontWeight:600 }}>
                                Delete <span style={{ fontFamily:"monospace" }}>{d.serial}</span>? This cannot be undone.
                              </td>
                              <td colSpan={2} style={{ padding:"5px 6px" }}>
                                <div style={{ display:"flex", gap:5 }}>
                                  <Btn onClick={()=>deleteDevice(d.id)} variant="danger" size="sm">Delete</Btn>
                                  <Btn onClick={()=>setDeleteId(null)} variant="ghost" size="sm">Cancel</Btn>
                                </div>
                              </td>
                            </>
                          ) : isReturning ? (
                            <>
                              <td colSpan={5} style={{ padding:"9px 10px", color:"#92400E", fontWeight:600 }}>
                                Return <span style={{ fontFamily:"monospace" }}>{d.serial}</span> to Triage queue? It will be unassigned.
                              </td>
                              <td colSpan={2} style={{ padding:"5px 6px" }}>
                                <div style={{ display:"flex", gap:5 }}>
                                  <Btn onClick={()=>returnToTriage(d.id)} variant="amber" size="sm">↩ Confirm</Btn>
                                  <Btn onClick={()=>setReturnId(null)} variant="ghost" size="sm">Cancel</Btn>
                                </div>
                              </td>
                            </>
                          ) : (
                            <>
                              <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:12 }}>{d.serial}</td>
                              <td style={{ padding:"9px 10px", fontSize:12 }}>{d.model||"—"}</td>
                              <td style={{ padding:"9px 10px", whiteSpace:"nowrap" }}>{TYPE_ICON[d.type]} {d.type}</td>
                              <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:11, color:C.slate3 }}>{d.mac||"—"}</td>
                              <td style={{ padding:"9px 10px", color:C.slate3, whiteSpace:"nowrap" }}>{d.received}</td>
                              <td style={{ padding:"9px 10px", color:C.slate3, maxWidth:130, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{d.notes||"—"}</td>
                              <td style={{ padding:"9px 10px" }}>
                                <div style={{ display:"flex", gap:5 }}>
                                  <Btn onClick={()=>setReturnId(d.id)} variant="amber" size="sm">↩ Return to Triage</Btn>
                                  <Btn onClick={()=>{ setEditingId(d.id); setEditBuf({serial:d.serial,type:d.type,mac:d.mac||"",model:d.model||"",notes:d.notes||""}); }} variant="ghost" size="sm">✏️</Btn>
                                  <Btn onClick={()=>setDeleteId(d.id)} variant="ghost" size="sm">🗑️</Btn>
                                </div>
                              </td>
                            </>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )
        }
      </div>
      <DevSection title="QC Check" items={inQC} accent="#3B82F6"
        renderActions={d=>[
          <Btn key="w" onClick={()=>setOutcome(d.id,"Working")}     variant="success" size="sm">✓ Working</Btn>,
          <Btn key="n" onClick={()=>setOutcome(d.id,"Not Working")} variant="danger"  size="sm">✗ Not Working</Btn>,
        ]}/>
      <DevSection title="Escalated" items={escalated} accent={C.purple}
        renderActions={d=>[
          <Btn key="r" onClick={()=>requeue(d.id)}        variant="purple" size="sm">↻ Re-queue</Btn>,
          <Btn key="s" onClick={()=>escalateToScrap(d.id)} variant="danger" size="sm">→ Scrap</Btn>,
        ]}/>
    </div>
  );
}


// ─────────────────────────────────────────────
// TRANSIT RECEIPT TAB
// Admin + Stock Management confirm physical
// receipt: Working → Ready, Not Working → Scrap
// ─────────────────────────────────────────────
function TransitReceiptTab({ devices, setDevices, isMobile }) {
  const [selected, setSelected]     = useState(new Set());
  const [showConfirm, setShowConfirm] = useState(false);
  const [confirmed, setConfirmed]   = useState(false);

  const transit = devices.filter(d => d.stage === "In Transit");
  const allSelected  = transit.length > 0 && selected.size === transit.length;
  const someSelected = selected.size > 0;
  const selWorking   = transit.filter(d => selected.has(d.id) && d.outcome === "Working").length;
  const selNotWork   = transit.filter(d => selected.has(d.id) && d.outcome !== "Working").length;

  function toggleAll() {
    setSelected(allSelected ? new Set() : new Set(transit.map(d => d.id)));
  }
  function toggleOne(id) {
    setSelected(prev => {
      const n = new Set(prev);
      n.has(id) ? n.delete(id) : n.add(id);
      return n;
    });
  }

  function confirmReceipt() {
    setDevices(p => p.map(d => {
      if (!selected.has(d.id)) return d;
      const toReady = d.outcome === "Working";
      const updated = {
        ...d,
        stage:         toReady ? "Stock" : "Scrap",
        outcome:       toReady ? "Working" : "Scrap",
        sentToPartner: false,
      };
      updateDevice(d.id, {
        stage:          updated.stage,
        outcome:        updated.outcome,
        sent_to_partner: false,
      }).catch(e => console.warn("DB receipt:", e));
      return updated;
    }));
    setSelected(new Set());
    setShowConfirm(false);
    setConfirmed(true);
    setTimeout(() => setConfirmed(false), 3000);
  }

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
      <Card>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
          marginBottom:14, flexWrap:"wrap", gap:8 }}>
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <SectionTitle>🚚 Stock In Transit from Partner ({transit.length})</SectionTitle>
            {someSelected && (
              <span style={{ background:"#EFF6FF", color:"#1D4ED8", fontSize:11, fontWeight:700,
                padding:"2px 9px", borderRadius:20 }}>
                {selected.size} selected
              </span>
            )}
          </div>
          <Btn onClick={()=>setShowConfirm(true)} disabled={!someSelected}
            style={{ background:someSelected?"#16A34A":"#D1FAE5",
              color:someSelected?"#fff":"#6EE7B7", border:"none", fontWeight:800,
              boxShadow:someSelected?"0 2px 8px rgba(22,163,74,.35)":"none", transition:"all .2s" }}>
            ✓ Confirm Receipt ({selected.size})
          </Btn>
        </div>

        {confirmed && (
          <Alert type="success">
            ✅ Receipt confirmed — Working devices moved to Ready, Not Working to Scrap
          </Alert>
        )}

        <p style={{ margin:"0 0 14px", fontSize:13, color:C.slate3 }}>
          Select the devices that have been physically received from the partner.
          <strong> Working</strong> devices will move to <strong>Ready stock</strong>.
          <strong> Not Working</strong> will move to <strong>Scrap</strong>.
        </p>

        {transit.length === 0
          ? <p style={{ color:C.slate4, fontSize:13 }}>No devices currently in transit.</p>
          : isMobile
            ? (
              <>
                <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:10,
                  padding:"8px 12px", background:C.slate8, borderRadius:8 }}>
                  <input type="checkbox" checked={allSelected} onChange={toggleAll}
                    style={{ width:16, height:16, cursor:"pointer" }}/>
                  <span style={{ fontSize:13, fontWeight:600, color:C.slate3 }}>Select All</span>
                </div>
                {transit.map(d => (
                  <div key={d.id} onClick={()=>toggleOne(d.id)}
                    style={{ border:`1.5px solid ${selected.has(d.id)?"#3B82F6":C.slate6}`,
                      borderRadius:12, padding:12, marginBottom:8, cursor:"pointer",
                      background:selected.has(d.id)?"#EFF6FF":C.white, transition:"all .15s" }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
                      <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                        <input type="checkbox" checked={selected.has(d.id)}
                          onChange={()=>toggleOne(d.id)}
                          onClick={e=>e.stopPropagation()}
                          style={{ width:16, height:16, cursor:"pointer" }}/>
                        <div>
                          <div style={{ fontFamily:"monospace", fontSize:12, fontWeight:700 }}>{d.serial}</div>
                          <div style={{ fontFamily:"monospace", fontSize:10, color:C.slate4 }}>{d.mac||"—"}</div>
                        </div>
                      </div>
                      <div>
                        <span style={{ fontSize:11, fontWeight:700, padding:"2px 8px", borderRadius:20,
                          background:d.outcome==="Working"?C.greenLight:C.redLight,
                          color:d.outcome==="Working"?C.greenDark:C.redDark }}>
                          {d.outcome==="Working"?"✓ Working":"✗ Not Working"}
                        </span>
                        <div style={{ fontSize:10, color:C.slate4, textAlign:"right", marginTop:3 }}>
                          → {d.outcome==="Working"?"Ready":"Scrap"}
                        </div>
                      </div>
                    </div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:4,
                      marginTop:8, fontSize:11, color:C.slate4 }}>
                      <span>{TYPE_ICON[d.type]} {d.type}</span>
                      <span>{d.model||"—"}</span>
                    </div>
                  </div>
                ))}
              </>
            )
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead>
                    <tr style={{ borderBottom:`1.5px solid ${C.slate6}`, background:C.slate8 }}>
                      <th style={{ padding:"8px 10px", width:40 }}>
                        <input type="checkbox" checked={allSelected} onChange={toggleAll}
                          style={{ width:15, height:15, cursor:"pointer" }}/>
                      </th>
                      {["Serial Number","MAC Address","Model","Type","Received","Outcome","On Receipt"].map(h => (
                        <th key={h} style={{ padding:"8px 10px", textAlign:"left", fontSize:11,
                          fontWeight:700, color:C.slate4, letterSpacing:".05em", whiteSpace:"nowrap" }}>
                          {h.toUpperCase()}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {transit.map(d => (
                      <tr key={d.id} onClick={()=>toggleOne(d.id)}
                        style={{ borderBottom:`1px solid ${C.slate7}`, cursor:"pointer",
                          background:selected.has(d.id)?"#EFF6FF":"transparent",
                          transition:"background .15s" }}>
                        <td style={{ padding:"10px" }}>
                          <input type="checkbox" checked={selected.has(d.id)}
                            onChange={()=>toggleOne(d.id)}
                            onClick={e=>e.stopPropagation()}
                            style={{ width:15, height:15, cursor:"pointer" }}/>
                        </td>
                        <td style={{ padding:"10px", fontFamily:"monospace", fontSize:12,
                          fontWeight:700, color:C.slate2, whiteSpace:"nowrap" }}>{d.serial}</td>
                        <td style={{ padding:"10px", fontFamily:"monospace", fontSize:11,
                          color:C.slate3, whiteSpace:"nowrap" }}>{d.mac||"—"}</td>
                        <td style={{ padding:"10px", fontSize:12 }}>{d.model||"—"}</td>
                        <td style={{ padding:"10px", whiteSpace:"nowrap" }}>{TYPE_ICON[d.type]} {d.type}</td>
                        <td style={{ padding:"10px", color:C.slate3, whiteSpace:"nowrap" }}>{d.received}</td>
                        <td style={{ padding:"10px" }}>
                          <span style={{ fontSize:11, fontWeight:700, padding:"3px 10px", borderRadius:20,
                            background:d.outcome==="Working"?C.greenLight:C.redLight,
                            color:d.outcome==="Working"?C.greenDark:C.redDark }}>
                            {d.outcome==="Working"?"✓ Working":"✗ Not Working"}
                          </span>
                        </td>
                        <td style={{ padding:"10px" }}>
                          <span style={{ fontSize:11, fontWeight:700, padding:"3px 10px", borderRadius:20,
                            background:d.outcome==="Working"?"#F0FDF4":"#FEF2F2",
                            color:d.outcome==="Working"?"#166534":"#991B1B" }}>
                            → {d.outcome==="Working"?"✅ Ready":"🗑 Scrap"}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
        }
      </Card>

      {/* ── RECEIPT CONFIRMATION POPUP ── */}
      {showConfirm && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.6)", zIndex:200,
          display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
          <div style={{ background:C.white, borderRadius:16, padding:28, maxWidth:520,
            width:"100%", boxShadow:"0 24px 64px rgba(0,0,0,.3)" }}>
            <div style={{ fontSize:28, marginBottom:8 }}>📬</div>
            <h3 style={{ margin:"0 0 6px", fontSize:20, fontWeight:800, color:C.slate }}>
              Confirm Physical Receipt
            </h3>
            <p style={{ margin:"0 0 18px", fontSize:13, color:C.slate3 }}>
              Confirming that the selected devices have been physically received from the partner.
              This will update their stock status immediately.
            </p>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:16 }}>
              <div style={{ background:C.greenLight, borderRadius:12, padding:"14px 16px",
                border:`1.5px solid #BBF7D0`, textAlign:"center" }}>
                <div style={{ fontSize:30, fontWeight:800, color:C.greenDark }}>{selWorking}</div>
                <div style={{ fontSize:12, fontWeight:700, color:C.greenDark }}>✓ Working</div>
                <div style={{ fontSize:11, color:C.slate4, marginTop:3 }}>→ ✅ Ready Stock</div>
              </div>
              <div style={{ background:C.redLight, borderRadius:12, padding:"14px 16px",
                border:`1.5px solid #FECACA`, textAlign:"center" }}>
                <div style={{ fontSize:30, fontWeight:800, color:C.redDark }}>{selNotWork}</div>
                <div style={{ fontSize:12, fontWeight:700, color:C.redDark }}>✗ Not Working</div>
                <div style={{ fontSize:11, color:C.slate4, marginTop:3 }}>→ 🗑 Scrap</div>
              </div>
            </div>
            {/* Device preview list */}
            <div style={{ background:C.slate8, borderRadius:10, padding:12, marginBottom:16,
              maxHeight:180, overflowY:"auto" }}>
              <div style={{ fontSize:11, fontWeight:700, color:C.slate4, letterSpacing:".06em",
                marginBottom:8, textTransform:"uppercase" }}>Devices being confirmed</div>
              {transit.filter(d => selected.has(d.id)).map(d => (
                <div key={d.id} style={{ display:"flex", justifyContent:"space-between",
                  alignItems:"center", padding:"5px 0",
                  borderBottom:`1px solid ${C.slate6}`, gap:8 }}>
                  <div style={{ display:"flex", alignItems:"center", gap:8, minWidth:0 }}>
                    <span style={{ fontFamily:"monospace", fontSize:12, fontWeight:700,
                      color:C.slate2, flexShrink:0 }}>{d.serial}</span>
                    {d.mac && <span style={{ fontFamily:"monospace", fontSize:10,
                      color:C.slate4, flexShrink:0 }}>{d.mac}</span>}
                  </div>
                  <div style={{ display:"flex", alignItems:"center", gap:6, flexShrink:0 }}>
                    <span style={{ fontSize:11, fontWeight:600, color:C.slate4 }}>→</span>
                    <span style={{ fontSize:11, fontWeight:700, padding:"2px 8px", borderRadius:20,
                      background:d.outcome==="Working"?C.greenLight:C.redLight,
                      color:d.outcome==="Working"?C.greenDark:C.redDark }}>
                      {d.outcome==="Working"?"✅ Ready":"🗑 Scrap"}
                    </span>
                  </div>
                </div>
              ))}
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
              <Btn onClick={()=>setShowConfirm(false)} variant="ghost" full size="lg">✕ Cancel</Btn>
              <Btn onClick={confirmReceipt} full size="lg"
                style={{ background:"#16A34A", color:"#fff", border:"none", fontWeight:800,
                  boxShadow:"0 2px 10px rgba(22,163,74,.4)" }}>
                📬 Confirm Receipt
              </Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}


// ─────────────────────────────────────────────
// STOCK DEVICE LIST  (sub-component — no IIFE)
// ─────────────────────────────────────────────
function StockDeviceList({ tab, pool, shown, meta, typeFilter, setTypeFilter,
  editingId, setEditingId, changeEcusStatus, exportTab, activePool, isMobile }) {

  const STATUS_OPTIONS = ["Refurbishment","Scrap","ECUS"];

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
      {/* Sub-header with filters + download */}
      <Card>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:10 }}>
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <SectionTitle>
              {tab==="ready" ? "✅ Ready for Deployment" : tab==="scrap" ? "🗑 Scrapped Devices" : "⏸ ECUS — Held Pending"}
            </SectionTitle>
            <span style={{ background:meta.light, color:meta.dark, fontSize:12, fontWeight:700,
              padding:"2px 10px", borderRadius:20, border:`1px solid ${meta.borderColor}` }}>
              {shown.length} device{shown.length!==1?"s":""}
            </span>
          </div>
          <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center" }}>
            {["All","Router/Modem","Set-top Box","ONT/OLT"].map(f => (
              <button key={f} onClick={()=>setTypeFilter(f)}
                style={{ padding:"5px 12px", borderRadius:20, fontSize:12, fontWeight:600,
                  cursor:"pointer", border:"none",
                  background: typeFilter===f ? meta.color : "#F1F5F9",
                  color:      typeFilter===f ? "#fff"     : "#475569",
                  transition:"all .15s" }}>
                {f==="All" ? "All" : `${{ "Router/Modem":"⬡","Set-top Box":"▦","ONT/OLT":"◈" }[f]} ${f.split("/")[0]}`}
              </button>
            ))}
            <Btn onClick={()=>exportTab(activePool,
              tab==="ready" ? "Ready (Stock)" : tab==="scrap" ? "Scrap" : "ECUS",
              `stock_${tab}`)}
              style={{ background:meta.color, color:"#fff", border:"none", fontWeight:700 }} size="sm">
              ↓ Excel
            </Btn>
          </div>
        </div>
      </Card>

      {/* Device table / cards */}
      <Card>
        {shown.length === 0
          ? <p style={{ color:"#94A3B8", fontSize:13, textAlign:"center", padding:"20px 0", margin:0 }}>
              No devices in this category{typeFilter!=="All" ? ` for ${typeFilter}` : ""}.
            </p>
          : isMobile
            ? shown.map(d => (
                <div key={d.id} style={{ border:`1px solid #E2E8F0`, borderRadius:12, padding:14,
                  marginBottom:10, borderLeft:`3px solid ${meta.color}` }}>
                  <div style={{ display:"flex", justifyContent:"space-between", marginBottom:8 }}>
                    <div>
                      <span style={{ fontWeight:800, color:"#6366F1", fontSize:12 }}>{d.id}</span>
                      <span style={{ fontFamily:"monospace", fontSize:12, marginLeft:8 }}>{d.serial}</span>
                    </div>
                    <span style={{ background:meta.light, color:meta.dark, fontSize:10,
                      fontWeight:700, padding:"2px 8px", borderRadius:20 }}>{meta.badge}</span>
                  </div>
                  <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6, marginBottom: tab==="ecus"?10:0 }}>
                    <div><span style={{ fontSize:10, fontWeight:700, color:"#94A3B8", textTransform:"uppercase" }}>Type</span>
                      <div style={{ fontSize:12 }}>{{ "Router/Modem":"⬡","Set-top Box":"▦","ONT/OLT":"◈" }[d.type]} {d.type}</div></div>
                    <div><span style={{ fontSize:10, fontWeight:700, color:"#94A3B8", textTransform:"uppercase" }}>Model</span>
                      <div style={{ fontSize:12 }}>{d.model||"—"}</div></div>
                    <div><span style={{ fontSize:10, fontWeight:700, color:"#94A3B8", textTransform:"uppercase" }}>MAC</span>
                      <div style={{ fontSize:11, fontFamily:"monospace", color:"#475569" }}>{d.mac||"—"}</div></div>
                    <div><span style={{ fontSize:10, fontWeight:700, color:"#94A3B8", textTransform:"uppercase" }}>Received</span>
                      <div style={{ fontSize:12, color:"#64748B" }}>{d.received}</div></div>
                  </div>
                  {tab==="ecus" && (
                    editingId===d.id
                      ? (
                        <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                          {STATUS_OPTIONS.map(s => (
                            <button key={s} onClick={()=>changeEcusStatus(d.id, s)}
                              style={{ flex:1, padding:"6px 0", borderRadius:8, fontSize:11, fontWeight:700,
                                cursor:"pointer", border:"1.5px solid",
                                borderColor: s==="Scrap"?"#991B1B":s==="ECUS"?"#92600A":"#4338CA",
                                background:  s==="Scrap"?"#FEF2F2":s==="ECUS"?"#FFF9EB":"#EEF2FF",
                                color:       s==="Scrap"?"#991B1B":s==="ECUS"?"#92600A":"#4338CA" }}>
                              {s==="Scrap"?"🗑 Scrap":s==="ECUS"?"⏸ ECUS":"🔧 Refurb"}
                            </button>
                          ))}
                          <Btn onClick={()=>setEditingId(null)} variant="ghost" size="sm" full>✕ Cancel</Btn>
                        </div>
                      )
                      : <Btn onClick={()=>setEditingId(d.id)} variant="ghost" size="sm">✏️ Change Status</Btn>
                  )}
                </div>
              ))
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead>
                    <tr style={{ borderBottom:"1.5px solid #E2E8F0", background:"#F8FAFC" }}>
                      {["ID","Serial","MAC","Model","Type","Received","Notes",
                        ...(tab==="ecus" ? ["Change Status"] : ["Status"])].map(h=>(
                        <th key={h} style={{ padding:"8px 10px", textAlign:"left", fontSize:11,
                          fontWeight:700, color:"#94A3B8", letterSpacing:".05em", whiteSpace:"nowrap" }}>
                          {h.toUpperCase()}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {shown.map(d => (
                      <tr key={d.id} style={{ borderBottom:"1px solid #F1F5F9",
                        background: editingId===d.id ? "#FFFBEB" : "transparent" }}>
                        <td style={{ padding:"9px 10px", fontWeight:700, color:"#6366F1" }}>{d.id}</td>
                        <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:12 }}>{d.serial}</td>
                        <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:11, color:"#475569" }}>{d.mac||"—"}</td>
                        <td style={{ padding:"9px 10px", fontSize:12 }}>{d.model||"—"}</td>
                        <td style={{ padding:"9px 10px", whiteSpace:"nowrap" }}>
                          {{ "Router/Modem":"⬡","Set-top Box":"▦","ONT/OLT":"◈" }[d.type]} {d.type}
                        </td>
                        <td style={{ padding:"9px 10px", color:"#64748B", whiteSpace:"nowrap" }}>{d.received}</td>
                        <td style={{ padding:"9px 10px", color:"#64748B", maxWidth:130,
                          overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{d.notes||"—"}</td>
                        <td style={{ padding:"9px 10px" }}>
                          {tab==="ecus"
                            ? editingId===d.id
                              ? (
                                <div style={{ display:"flex", gap:5 }}>
                                  {STATUS_OPTIONS.map(s => (
                                    <button key={s} onClick={()=>changeEcusStatus(d.id, s)}
                                      style={{ padding:"4px 9px", borderRadius:7, fontSize:11, fontWeight:700,
                                        cursor:"pointer", whiteSpace:"nowrap", border:"1.5px solid",
                                        borderColor: s==="Scrap"?"#991B1B":s==="ECUS"?"#92600A":"#4338CA",
                                        background:  s==="Scrap"?"#FEF2F2":s==="ECUS"?"#FFF9EB":"#EEF2FF",
                                        color:       s==="Scrap"?"#991B1B":s==="ECUS"?"#92600A":"#4338CA" }}>
                                      {s==="Scrap"?"🗑":s==="ECUS"?"⏸":"🔧"} {s==="Refurbishment"?"Refurb":s}
                                    </button>
                                  ))}
                                  <Btn onClick={()=>setEditingId(null)} variant="ghost" size="sm">✕</Btn>
                                </div>
                              )
                              : <Btn onClick={()=>setEditingId(d.id)} variant="ghost" size="sm">✏️ Change</Btn>
                            : <span style={{ background:meta.light, color:meta.dark, fontSize:11,
                                fontWeight:700, padding:"3px 10px", borderRadius:20 }}>{meta.badge}</span>
                          }
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
        }
      </Card>
    </div>
  );
}


// ─────────────────────────────────────────────
// UNIFIED STOCK VIEW  (Dashboard + Ready + Scrap + ECUS)
// ─────────────────────────────────────────────
function StockView({ devices, setDevices, isMobile, canConfirmTransit=false }) {
  const [tab, setTab]           = useState("dashboard"); // dashboard | ready | scrap | ecus
  const [typeFilter, setTypeFilter] = useState("All");
  const [editingId,  setEditingId]  = useState(null);

  // Data slices
  const ready   = devices.filter(d => d.stage === "Stock");
  const scrap   = devices.filter(d => d.stage === "Scrap");
  const ecus    = devices.filter(d => d.stage === "ECUS");
  const transit = devices.filter(d => d.stage === "In Transit");
  const total   = ready.length + scrap.length + ecus.length + transit.length;

  // Per-tab shown (with type filter)
  const poolMap = { ready, scrap, ecus, transit };
  const activePool = tab === "dashboard" ? [] : (poolMap[tab] || []);
  const shown = typeFilter === "All" ? activePool : activePool.filter(d => d.type === typeFilter);

  // ── Excel exports ──
  function toRows(arr) {
    return arr.map(d => ({
      "Device ID":     d.id,
      "Serial Number": d.serial,
      "MAC Address":   d.mac    || "",
      "Model":         d.model  || "",
      "Device Type":   d.type,
      "Stage":         d.stage,
      "Outcome":       d.outcome || "",
      "Received Date": d.received,
      "Notes":         d.notes  || "",
    }));
  }
  function exportTab(arr, sheetName, filename) {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(toRows(arr)), sheetName);
    XLSX.writeFile(wb, `${filename}_${today()}.xlsx`);
  }
  function exportAll() {
    const wb = XLSX.utils.book_new();
    XLSX.utils.XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(toRows(ready)),   "Ready (Stock)");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(toRows(scrap)),   "Scrap");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(toRows(ecus)),    "ECUS");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(toRows(transit)), "In Transit");
    XLSX.writeFile(wb, `stock_full_report_${today()}.xlsx`);
  }

  // ── ECUS status change ──
  function changeEcusStatus(id, newStatus) {
    const isScrap = newStatus === "Scrap";
    const isEcus  = newStatus === "ECUS";
    setDevices(p => p.map(d => d.id !== id ? d : {
      ...d,
      stage:         isScrap ? "Scrap" : isEcus ? "ECUS" : "Refurbishment",
      outcome:       isScrap ? "Scrap" : null,
      sentToPartner: !isScrap && !isEcus,
      pendingAction: null,
    }));
    updateDevice(id, {
      stage:          isScrap ? "Scrap" : isEcus ? "ECUS" : "Refurbishment",
      outcome:        isScrap ? "Scrap" : null,
      sent_to_partner: !isScrap && !isEcus,
    }).catch(e => console.warn("DB ECUS status:", e));
    setEditingId(null);
  }

  // ── Shared inline bar chart ──
  function MiniBar({ label, value, max, color }) {
    const pct = max > 0 ? Math.round(value/max*100) : 0;
    return (
      <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:6 }}>
        <span style={{ width:isMobile?70:90, fontSize:11, fontWeight:600, color:C.slate3, textAlign:"right", flexShrink:0 }}>{label}</span>
        <div style={{ flex:1, background:C.slate7, borderRadius:4, height:16, overflow:"hidden" }}>
          <div style={{ width:`${pct}%`, minWidth:value?22:0, height:"100%", background:color,
            borderRadius:4, display:"flex", alignItems:"center", paddingLeft:5, transition:"width .4s" }}>
            {value>0 && <span style={{ fontSize:10, fontWeight:800, color:"#fff" }}>{value}</span>}
          </div>
        </div>
        {value===0 && <span style={{ fontSize:10, color:C.slate5 }}>0</span>}
      </div>
    );
  }

  // ── Donut chart (SVG) ──
  function DonutChart({ segments, size=120 }) {
    const r = 42, cx = 60, cy = 60, stroke = 18;
    const circ = 2 * Math.PI * r;
    const tot = segments.reduce((a,s) => a+s.value, 0);
    let offset = 0;
    const arcs = segments.map(s => {
      const dash = tot > 0 ? (s.value/tot)*circ : 0;
      const arc = { ...s, dash, gap: circ-dash, offset };
      offset += dash;
      return arc;
    });
    return (
      <svg width={size} height={size} viewBox="0 0 120 120">
        <circle cx={cx} cy={cy} r={r} fill="none" stroke={C.slate7} strokeWidth={stroke}/>
        {tot===0
          ? <circle cx={cx} cy={cy} r={r} fill="none" stroke={C.slate6} strokeWidth={stroke}/>
          : arcs.map((a,i) => a.value>0 && (
              <circle key={i} cx={cx} cy={cy} r={r} fill="none"
                stroke={a.color} strokeWidth={stroke}
                strokeDasharray={`${a.dash} ${a.gap}`}
                strokeDashoffset={-a.offset}
                style={{ transform:"rotate(-90deg)", transformOrigin:"60px 60px", transition:"all .5s" }}/>
            ))
        }
        <text x={cx} y={cy-6}  textAnchor="middle" fontSize="18" fontWeight="800" fill={C.slate}>{tot}</text>
        <text x={cx} y={cy+10} textAnchor="middle" fontSize="9"  fontWeight="600" fill={C.slate4}>TOTAL</text>
      </svg>
    );
  }

  // Tab config
  const TABS = [
    { id:"dashboard", label:"📊 Overview",         color:C.indigo  },
    { id:"ready",     label:"✅ Ready",              color:C.green   },
    { id:"scrap",     label:"🗑 Scrap",              color:C.red     },
    { id:"ecus",      label:"⏸ ECUS",               color:"#F59E0B" },
    { id:"transit",   label:"🚚 In Transit",         color:"#3B82F6" },
  ];
  const TAB_META = {
    ready:   { color:C.green,   light:C.greenLight, dark:C.greenDark, badge:"✓ Ready",       borderColor:"#BBF7D0" },
    scrap:   { color:C.red,     light:C.redLight,   dark:C.redDark,   badge:"🗑 Scrapped",    borderColor:"#FECACA" },
    ecus:    { color:"#F59E0B", light:"#FFF9EB",    dark:"#92600A",   badge:"⏸ ECUS",        borderColor:"#FCD34D" },
    transit: { color:"#3B82F6", light:"#EFF6FF",    dark:"#1D4ED8",   badge:"🚚 In Transit",  borderColor:"#BFDBFE" },
  };

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>

      {/* ── PAGE HEADER ── */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:10 }}>
        <div>
          <h2 style={{ margin:"0 0 2px", fontSize:isMobile?18:22, fontWeight:800, color:C.slate }}>Stock Management</h2>
          <p style={{ margin:0, color:C.slate4, fontSize:13 }}>Ready · Scrap · ECUS — full inventory overview</p>
        </div>
        <Btn onClick={exportAll} variant="success">↓ Export All (Excel)</Btn>
      </div>

      <DeviceSearchBar devices={devices} isMobile={isMobile}/>

      {/* ── TAB TREE ── */}
      <div style={{ background:C.white, border:`1.5px solid ${C.slate6}`, borderRadius:14, padding:"6px 8px",
        display:"flex", gap:4, flexWrap:"wrap" }}>
        {TABS.map(t => (
          <button key={t.id} onClick={()=>{ setTab(t.id); setTypeFilter("All"); setEditingId(null); }}
            style={{ padding:"8px 16px", borderRadius:10, border:"none", fontSize:13, fontWeight:700,
              cursor:"pointer", transition:"all .15s",
              background: tab===t.id ? t.color : "transparent",
              color:       tab===t.id ? "#fff"   : C.slate3,
              boxShadow:   tab===t.id ? `0 2px 8px ${t.color}55` : "none" }}>
            {t.label}
            <span style={{ marginLeft:6, fontSize:11, opacity:.85 }}>
              ({t.id==="dashboard" ? total : t.id==="ready" ? ready.length : t.id==="scrap" ? scrap.length : t.id==="ecus" ? ecus.length : transit.length})
            </span>
          </button>
        ))}
      </div>

      {/* ══════════════════════════════════════════
          DASHBOARD TAB
      ══════════════════════════════════════════ */}
      {tab==="dashboard" && (
        <div style={{ display:"flex", flexDirection:"column", gap:16 }}>

          {/* Top stat cards */}
          <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(5,1fr)", gap:12 }}>
            {[
              { label:"Total Stock",  value:total,           sub:"All statuses",        accent:C.indigo  },
              { label:"✅ Ready",     value:ready.length,    sub:"Available",           accent:C.green   },
              { label:"🗑 Scrap",     value:scrap.length,    sub:"Disposed",            accent:C.red     },
              { label:"⏸ ECUS",      value:ecus.length,     sub:"Held pending",        accent:"#F59E0B" },
              { label:"🚚 In Transit",value:transit.length,  sub:"En route from partner",accent:"#3B82F6"},
            ].map(c => (
              <Card key={c.label} style={{ borderLeft:`4px solid ${c.accent}`, cursor:"pointer" }}
                onClick={()=>c.label!=="Total Stock" && setTab(c.label.includes("Ready")?"ready":c.label.includes("Scrap")?"scrap":"ecus")}>
                <div style={{ fontSize:11, fontWeight:700, color:C.slate4, textTransform:"uppercase", letterSpacing:".07em", marginBottom:4 }}>{c.label}</div>
                <div style={{ fontSize:30, fontWeight:800, color:C.slate, lineHeight:1 }}>{c.value}</div>
                <div style={{ fontSize:12, color:C.slate4, marginTop:4 }}>{c.sub}</div>
              </Card>
            ))}
          </div>

          {/* Charts row */}
          <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr":"1fr 1fr 1fr", gap:16 }}>

            {/* Donut — status split */}
            <Card>
              <SectionTitle>Status split</SectionTitle>
              <div style={{ display:"flex", alignItems:"center", gap:20, flexWrap:"wrap" }}>
                <DonutChart segments={[
                  { value:ready.length, color:C.green  },
                  { value:scrap.length, color:C.red    },
                  { value:ecus.length,  color:"#F59E0B"},
                ]}/>
                <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
                  {[
                    ["✅ Ready", ready.length, C.green],
                    ["🗑 Scrap", scrap.length, C.red],
                    ["⏸ ECUS",  ecus.length,  "#F59E0B"],
                  ].map(([label, val, color]) => (
                    <div key={label} style={{ display:"flex", alignItems:"center", gap:8 }}>
                      <span style={{ width:10, height:10, borderRadius:"50%", background:color, flexShrink:0 }}/>
                      <span style={{ fontSize:13, color:C.slate3, fontWeight:600 }}>{label}</span>
                      <span style={{ fontSize:14, fontWeight:800, color:C.slate, marginLeft:"auto", paddingLeft:12 }}>{val}</span>
                    </div>
                  ))}
                </div>
              </div>
            </Card>

            {/* Bar — by device type */}
            <Card>
              <SectionTitle>By device type</SectionTitle>
              {DEVICE_TYPES.map(t => {
                const tTotal = devices.filter(d=>d.type===t&&["Stock","Scrap","ECUS"].includes(d.stage)).length;
                return (
                  <div key={t} style={{ marginBottom:14 }}>
                    <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
                      <span style={{ fontSize:12, fontWeight:700, color:C.slate }}>{TYPE_ICON[t]} {t}</span>
                      <span style={{ fontSize:12, fontWeight:800, color:C.slate }}>{tTotal}</span>
                    </div>
                    <div style={{ display:"flex", height:14, borderRadius:6, overflow:"hidden", gap:1 }}>
                      {[
                        [devices.filter(d=>d.type===t&&d.stage==="Stock").length, C.green],
                        [devices.filter(d=>d.type===t&&d.stage==="Scrap").length, C.red],
                        [devices.filter(d=>d.type===t&&d.stage==="ECUS").length,  "#F59E0B"],
                      ].map(([val, color], i) => tTotal>0 && val>0 && (
                        <div key={i} style={{ width:`${val/tTotal*100}%`, background:color,
                          display:"flex", alignItems:"center", justifyContent:"center" }}>
                          <span style={{ fontSize:9, fontWeight:800, color:"#fff" }}>{val}</span>
                        </div>
                      ))}
                      {tTotal===0 && <div style={{ flex:1, background:C.slate7, borderRadius:6 }}/>}
                    </div>
                    <div style={{ display:"flex", gap:10, marginTop:3 }}>
                      {[["Ready",devices.filter(d=>d.type===t&&d.stage==="Stock").length,C.green],
                        ["Scrap",devices.filter(d=>d.type===t&&d.stage==="Scrap").length,C.red],
                        ["ECUS", devices.filter(d=>d.type===t&&d.stage==="ECUS").length,"#F59E0B"]
                      ].map(([l,v,c])=>(
                        <span key={l} style={{ fontSize:10, color:c, fontWeight:600 }}>{l}: {v}</span>
                      ))}
                    </div>
                  </div>
                );
              })}
            </Card>

            {/* Bar — status per type detailed */}
            <Card>
              <SectionTitle>Status breakdown</SectionTitle>
              {[
                ["✅ Ready", ready, C.green],
                ["🗑 Scrap",  scrap,  C.red],
                ["⏸ ECUS",   ecus,   "#F59E0B"],
              ].map(([label, pool, color]) => {
                const max = Math.max(ready.length, scrap.length, ecus.length, 1);
                return (
                  <div key={label} style={{ marginBottom:16 }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:6 }}>
                      <span style={{ fontSize:12, fontWeight:700, color }}>{label}</span>
                      <span style={{ fontSize:12, fontWeight:800, color }}>{pool.length}</span>
                    </div>
                    <MiniBar label="Router/Modem" value={pool.filter(d=>d.type==="Router/Modem").length} max={max} color={color}/>
                    <MiniBar label="Set-top Box"  value={pool.filter(d=>d.type==="Set-top Box").length}  max={max} color={color}/>
                    <MiniBar label="ONT/OLT"      value={pool.filter(d=>d.type==="ONT/OLT").length}      max={max} color={color}/>
                  </div>
                );
              })}
            </Card>
          </div>

          {/* Quick-access download row */}
          <Card>
            <SectionTitle>Download reports</SectionTitle>
            <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr":"repeat(4,1fr)", gap:10 }}>
              {[
                ["↓ All (4 sheets)", exportAll,                    C.green,   C.greenLight, "#BBF7D0"],
                ["↓ Ready only",     ()=>exportTab(ready,"Ready (Stock)","stock_ready"),  C.green,   C.greenLight, "#BBF7D0"],
                ["↓ Scrap only",     ()=>exportTab(scrap,"Scrap","stock_scrap"),           C.red,     C.redLight,   "#FECACA"],
                ["↓ ECUS only",      ()=>exportTab(ecus,"ECUS","stock_ecus"),              "#92600A", "#FFF9EB",    "#FCD34D"],
              ["↓ In Transit",     ()=>exportTab(transit,"In Transit","stock_transit"),    "#1D4ED8", "#EFF6FF",    "#BFDBFE"],
              ].map(([label, fn, color, bg, border]) => (
                <button key={label} onClick={fn}
                  style={{ padding:"12px 10px", borderRadius:10, border:`1.5px solid ${border}`,
                    background:bg, color, fontSize:13, fontWeight:700, cursor:"pointer",
                    textAlign:"center", transition:"all .15s" }}>
                  {label}
                </button>
              ))}
            </div>
          </Card>
        </div>
      )}

      {/* ══════════════════════════════════════════
          DEVICE LIST TABS (ready / scrap / ecus)
      ══════════════════════════════════════════ */}
      {tab !== "dashboard" && tab !== "transit" && (
        <StockDeviceList
          tab={tab} pool={activePool} shown={shown} meta={TAB_META[tab]}
          typeFilter={typeFilter} setTypeFilter={setTypeFilter}
          editingId={editingId} setEditingId={setEditingId}
          changeEcusStatus={changeEcusStatus}
          exportTab={exportTab} activePool={activePool}
          isMobile={isMobile}/>
      )}
      {tab === "transit" && (
        canConfirmTransit
          ? <TransitReceiptTab devices={devices} setDevices={setDevices} isMobile={isMobile}/>
          : <Card><p style={{ color:C.slate4, fontSize:13, textAlign:"center", padding:"20px 0", margin:0 }}>
              ⛔ You don't have permission to confirm receipt. Contact an Admin or Stock Manager.
            </p></Card>
      )}
    </div>
  );
}



// ─────────────────────────────────────────────
// DEVICE SEARCH BAR  (inline, used on every page)
// ─────────────────────────────────────────────
function DeviceSearchBar({ devices, isMobile, scope="all", placeholder="Search serial, MAC, model, ID…" }) {
  const [query,   setQuery]   = useState("");
  const [open,    setOpen]    = useState(false);
  const [stageF,  setStageF]  = useState("All");
  const [typeF,   setTypeF]   = useState("All");

  const pool = scope === "partner" ? devices.filter(d => d.sentToPartner) : devices;

  const results = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q && stageF === "All" && typeF === "All") return [];
    return pool.filter(d => {
      const ms = stageF === "All" || d.stage === stageF;
      const mt = typeF  === "All" || d.type  === typeF;
      const mq = !q || [d.id, d.serial, d.mac, d.model, d.notes, d.type]
        .some(v => v && v.toLowerCase().includes(q));
      return ms && mt && mq;
    }).slice(0, 50);
  }, [query, stageF, typeF, pool]);

  const hasSearch = query.trim() !== "" || stageF !== "All" || typeF !== "All";

  function clear() { setQuery(""); setStageF("All"); setTypeF("All"); }

  const stageOpts = scope === "partner"
    ? ["All","Refurbishment","Confirmed","In Transit","Stock","Escalated","Scrap"]
    : ["All",...STAGES];

  return (
    <div style={{ position:"relative", marginBottom:16 }}>
      {/* Collapsed bar */}
      {!open && (
        <button onClick={()=>setOpen(true)}
          style={{ width:"100%", display:"flex", alignItems:"center", gap:10, padding:"10px 14px",
            background:C.white, border:`1.5px solid ${C.slate6}`, borderRadius:10, cursor:"pointer",
            fontSize:13, color:C.slate4, textAlign:"left", transition:"border-color .15s" }}>
          <span style={{ fontSize:16 }}>🔍</span>
          <span style={{ flex:1 }}>{placeholder}</span>
          <span style={{ fontSize:11, background:C.slate7, padding:"2px 8px",
            borderRadius:20, fontWeight:600, color:C.slate4 }}>Search</span>
        </button>
      )}

      {/* Expanded search panel */}
      {open && (
        <div style={{ background:C.white, border:`1.5px solid ${C.indigo}`,
          borderRadius:12, padding:14, boxShadow:"0 4px 20px rgba(0,0,0,.10)" }}>
          {/* Search input row */}
          <div style={{ display:"flex", gap:8, marginBottom:10, alignItems:"center" }}>
            <span style={{ fontSize:16, flexShrink:0 }}>🔍</span>
            <input value={query} onChange={e=>setQuery(e.target.value)}
              placeholder={placeholder}
              style={{ ...iStyle(), flex:1, fontSize:14, borderColor:C.indigo }}
              autoFocus/>
            <button onClick={()=>{ clear(); setOpen(false); }}
              style={{ background:"none", border:"none", cursor:"pointer",
                color:C.slate4, fontSize:20, lineHeight:1, padding:"0 4px", flexShrink:0 }}>✕</button>
          </div>

          {/* Filter chips */}
          <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginBottom:hasSearch?10:0 }}>
            <select value={stageF} onChange={e=>setStageF(e.target.value)}
              style={{ ...iStyle(), width:"auto", fontSize:11, padding:"4px 8px", height:28 }}>
              {stageOpts.map(s=><option key={s} value={s}>{s==="All"?"All stages":s}</option>)}
            </select>
            <select value={typeF} onChange={e=>setTypeF(e.target.value)}
              style={{ ...iStyle(), width:"auto", fontSize:11, padding:"4px 8px", height:28 }}>
              <option value="All">All types</option>
              {DEVICE_TYPES.map(t=><option key={t}>{t}</option>)}
            </select>
            {hasSearch && (
              <button onClick={clear}
                style={{ background:"none", border:"none", cursor:"pointer",
                  color:C.indigo, fontSize:12, fontWeight:600, padding:"4px 6px" }}>
                Clear
              </button>
            )}
          </div>

          {/* Results */}
          {hasSearch && (
            results.length === 0
              ? <p style={{ fontSize:13, color:C.slate4, margin:0, textAlign:"center", padding:"12px 0" }}>
                  No devices match
                </p>
              : (
                <div style={{ maxHeight:320, overflowY:"auto", borderTop:`1px solid ${C.slate7}`, marginTop:4 }}>
                  <div style={{ fontSize:11, fontWeight:700, color:C.slate4, padding:"8px 4px 4px",
                    letterSpacing:".05em", textTransform:"uppercase" }}>
                    {results.length} result{results.length!==1?"s":""}
                    {results.length===50?" (showing first 50)":""}
                  </div>
                  {results.map(d => (
                    <div key={d.id} style={{ padding:"8px 4px", borderBottom:`1px solid ${C.slate7}`,
                      display:"flex", justifyContent:"space-between", alignItems:"center", gap:8 }}>
                      <div style={{ minWidth:0, flex:1 }}>
                        <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:3, flexWrap:"wrap" }}>
                          <span style={{ fontWeight:800, color:C.indigo, fontSize:12,
                            flexShrink:0 }}>{d.id}</span>
                          <span style={{ fontFamily:"monospace", fontSize:12,
                            fontWeight:700 }}>{d.serial}</span>
                          {d.mac && <span style={{ fontFamily:"monospace", fontSize:10,
                            color:C.slate4 }}>{d.mac}</span>}
                        </div>
                        <div style={{ display:"flex", gap:8, fontSize:11, color:C.slate4,
                          flexWrap:"wrap" }}>
                          <span>{TYPE_ICON[d.type]} {d.type}</span>
                          {d.model && <span>{d.model}</span>}
                          <span>{d.received}</span>
                          {d.outcome && <span style={{ fontWeight:700,
                            color:d.outcome==="Working"?C.greenDark:d.outcome==="Scrap"?"#92600A":C.redDark }}>
                            {d.outcome}
                          </span>}
                        </div>
                        {d.notes && <div style={{ fontSize:11, color:C.slate4,
                          fontStyle:"italic", marginTop:2 }}>{d.notes}</div>}
                      </div>
                      <Badge stage={d.stage}/>
                    </div>
                  ))}
                </div>
              )
          )}

          {!hasSearch && (
            <p style={{ fontSize:12, color:C.slate4, margin:0, textAlign:"center", padding:"8px 0" }}>
              Type to search across all device fields
            </p>
          )}
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// ALL DEVICES  (full registry view)
// ─────────────────────────────────────────────
function AllDevices({ devices, isMobile }) {
  const [search,      setSearch]      = useState("");
  const [stageFilter, setStageFilter] = useState("All");
  const [typeFilter,  setTypeFilter]  = useState("All");
  const [showReport,  setShowReport]  = useState(false);
  const [reportStage, setReportStage] = useState("All");
  const [reportType,  setReportType]  = useState("All");
  const [reportOutcome, setReportOutcome] = useState("All");
  const [exported,    setExported]    = useState(false);

  const filtered = useMemo(() => devices.filter(d => {
    const ms = stageFilter === "All" || d.stage === stageFilter;
    const mt = typeFilter  === "All" || d.type  === typeFilter;
    const q  = search.toLowerCase();
    return ms && mt && (!q || [d.id,d.serial,d.mac,d.model,d.notes,d.type]
      .some(v => v && v.toLowerCase().includes(q)));
  }), [devices, search, stageFilter, typeFilter]);

  const reportData = useMemo(() => devices.filter(d => {
    const ms = reportStage   === "All" || d.stage   === reportStage;
    const mt = reportType    === "All" || d.type    === reportType;
    const mo = reportOutcome === "All" || (d.outcome||"") === reportOutcome;
    return ms && mt && mo;
  }), [devices, reportStage, reportType, reportOutcome]);

  function exportReport() {
    const rows = reportData.map(d => ({
      "Device ID":     d.id,
      "Serial Number": d.serial,
      "MAC Address":   d.mac    || "",
      "Model":         d.model  || "",
      "Device Type":   d.type,
      "Stage":         d.stage,
      "Outcome":       d.outcome || "",
      "Partner":       d.sentToPartner ? "Yes" : "No",
      "Received Date": d.received,
      "Notes":         d.notes  || "",
    }));
    const summary = STAGES.map(s => ({
      "Stage": s,
      "Total": reportData.filter(d=>d.stage===s).length,
      "Router/Modem": reportData.filter(d=>d.stage===s&&d.type==="Router/Modem").length,
      "Set-top Box":  reportData.filter(d=>d.stage===s&&d.type==="Set-top Box").length,
      "ONT/OLT":      reportData.filter(d=>d.stage===s&&d.type==="ONT/OLT").length,
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows),    "Devices");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "Stage Summary");
    const label = [reportStage!=="All"?reportStage:"",reportType!=="All"?reportType:"",
      reportOutcome!=="All"?reportOutcome:""].filter(Boolean).join("_") || "all";
    XLSX.writeFile(wb, `cpe_report_${label}_${today()}.xlsx`);
    setExported(true); setTimeout(()=>setExported(false), 2500);
  }

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start",
        flexWrap:"wrap", gap:10 }}>
        <div>
          <h2 style={{ margin:"0 0 2px", fontSize:isMobile?18:22, fontWeight:800, color:C.slate }}>All Devices</h2>
          <p style={{ margin:0, color:C.slate4, fontSize:13 }}>Full registry · filter, search and export</p>
        </div>
        <Btn onClick={()=>setShowReport(r=>!r)} variant={showReport?"primary":"default"}>
          📊 {showReport?"Hide":"Generate"} Report
        </Btn>
      </div>

      {/* Report builder */}
      {showReport && (
        <Card style={{ border:`1.5px solid ${C.indigo}44`, background:C.indigoLight }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
            marginBottom:14, flexWrap:"wrap", gap:8 }}>
            <SectionTitle>📊 Report Builder</SectionTitle>
            <div style={{ display:"flex", gap:8, alignItems:"center" }}>
              {exported && <span style={{ fontSize:12, color:C.greenDark, fontWeight:700 }}>✓ Downloaded!</span>}
              <Btn onClick={exportReport} variant="success">↓ Export Excel ({reportData.length})</Btn>
            </div>
          </div>
          <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(3,1fr)", gap:10, marginBottom:12 }}>
            <div><Label>FILTER BY STAGE</Label>
              <select value={reportStage} onChange={e=>setReportStage(e.target.value)} style={iStyle()}>
                <option value="All">All stages</option>
                {STAGES.map(s=><option key={s}>{s}</option>)}
              </select>
            </div>
            <div><Label>FILTER BY TYPE</Label>
              <select value={reportType} onChange={e=>setReportType(e.target.value)} style={iStyle()}>
                <option value="All">All types</option>
                {DEVICE_TYPES.map(t=><option key={t}>{t}</option>)}
              </select>
            </div>
            <div><Label>FILTER BY OUTCOME</Label>
              <select value={reportOutcome} onChange={e=>setReportOutcome(e.target.value)} style={iStyle()}>
                <option value="All">All outcomes</option>
                <option value="Working">Working</option>
                <option value="Not Working">Not Working</option>
                <option value="Scrap">Scrap</option>
              </select>
            </div>
          </div>
          <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(4,1fr)", gap:10, marginBottom:10 }}>
            {[["Total",reportData.length,C.indigo],["Working",reportData.filter(d=>d.outcome==="Working").length,C.green],
              ["Not Working",reportData.filter(d=>d.outcome==="Not Working").length,C.red],
              ["Scrap",reportData.filter(d=>d.outcome==="Scrap").length,"#92400E"]
            ].map(([l,v,c])=>(
              <div key={l} style={{ background:C.white, borderRadius:10, padding:"10px 14px", borderLeft:`3px solid ${c}` }}>
                <div style={{ fontSize:11, fontWeight:700, color:C.slate4, textTransform:"uppercase", marginBottom:2 }}>{l}</div>
                <div style={{ fontSize:22, fontWeight:800, color:c }}>{v}</div>
              </div>
            ))}
          </div>
          <p style={{ margin:0, fontSize:11, color:C.slate4 }}>
            Export generates a 2-sheet Excel: Device list · Stage summary
          </p>
        </Card>
      )}

      {/* Filters */}
      <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr":"3fr 1fr 1fr", gap:10 }}>
        <input value={search} onChange={e=>setSearch(e.target.value)}
          placeholder="Search ID, serial, MAC, model, notes…" style={iStyle()}/>
        <select value={stageFilter} onChange={e=>setStageFilter(e.target.value)} style={iStyle()}>
          <option value="All">All stages</option>
          {STAGES.map(s=><option key={s}>{s}</option>)}
        </select>
        <select value={typeFilter} onChange={e=>setTypeFilter(e.target.value)} style={iStyle()}>
          <option value="All">All types</option>
          {DEVICE_TYPES.map(t=><option key={t}>{t}</option>)}
        </select>
      </div>

      {/* Device table */}
      <Card>
        <div style={{ marginBottom:12, fontSize:13, color:C.slate4, display:"flex",
          justifyContent:"space-between", alignItems:"center" }}>
          <span>{filtered.length} device{filtered.length!==1?"s":""}</span>
          {filtered.length>0 && (
            <button onClick={()=>{ setReportStage(stageFilter); setReportType(typeFilter);
              setReportOutcome("All"); setShowReport(true); }}
              style={{ background:"none", border:"none", color:C.indigo,
                fontSize:12, fontWeight:600, cursor:"pointer" }}>
              Export this view →
            </button>
          )}
        </div>
        {filtered.length===0
          ? <p style={{ color:C.slate4, fontSize:13 }}>No devices match.</p>
          : isMobile
            ? filtered.map(d=>(
                <div key={d.id} style={{ border:`1px solid ${C.slate6}`, borderRadius:10,
                  padding:12, marginBottom:8 }}>
                  <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
                    <div>
                      <span style={{ fontWeight:800, color:C.indigo, fontSize:12 }}>{d.id}</span>
                      <span style={{ fontFamily:"monospace", fontSize:12, marginLeft:8 }}>{d.serial}</span>
                    </div>
                    <Badge stage={d.stage}/>
                  </div>
                  <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:4, fontSize:12, color:C.slate3 }}>
                    <span>{TYPE_ICON[d.type]} {d.type}</span>
                    <span>{d.model||"—"}</span>
                    <span style={{ fontFamily:"monospace", fontSize:11 }}>{d.mac||"—"}</span>
                    <span>{d.received}</span>
                  </div>
                </div>
              ))
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead><tr style={{ borderBottom:`1.5px solid ${C.slate6}`, background:C.slate8 }}>
                    {["ID","Serial","MAC","Model","Type","Stage","Outcome","Partner","Received","Notes"].map(h=>(
                      <th key={h} style={{ padding:"8px 10px", textAlign:"left", fontSize:11,
                        fontWeight:700, color:C.slate4, letterSpacing:".05em", whiteSpace:"nowrap" }}>
                        {h.toUpperCase()}
                      </th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {filtered.map(d=>(
                      <tr key={d.id} style={{ borderBottom:`1px solid ${C.slate7}` }}>
                        <td style={{ padding:"9px 10px", fontWeight:700, color:C.indigo, whiteSpace:"nowrap" }}>{d.id}</td>
                        <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:12 }}>{d.serial}</td>
                        <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:11, color:C.slate3 }}>{d.mac||"—"}</td>
                        <td style={{ padding:"9px 10px", fontSize:12 }}>{d.model||"—"}</td>
                        <td style={{ padding:"9px 10px", whiteSpace:"nowrap" }}>{TYPE_ICON[d.type]} {d.type}</td>
                        <td style={{ padding:"9px 10px" }}><Badge stage={d.stage}/></td>
                        <td style={{ padding:"9px 10px", fontWeight:600, fontSize:12,
                          color:d.outcome==="Working"?C.greenDark:d.outcome?C.redDark:C.slate4 }}>
                          {d.outcome||"—"}
                        </td>
                        <td style={{ padding:"9px 10px" }}>
                          {d.sentToPartner
                            ? <span style={{ background:C.purpleLight, color:C.purple, fontSize:11,
                                fontWeight:600, padding:"2px 8px", borderRadius:20 }}>Partner</span>
                            : <span style={{ color:C.slate5 }}>—</span>}
                        </td>
                        <td style={{ padding:"9px 10px", color:C.slate3, whiteSpace:"nowrap" }}>{d.received}</td>
                        <td style={{ padding:"9px 10px", color:C.slate4, maxWidth:130,
                          overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{d.notes||"—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
        }
      </Card>
    </div>
  );
}


// ─────────────────────────────────────────────
// PARTNER REPORTS  (proper component — hooks at top level)
// ─────────────────────────────────────────────
function PartnerReports({ devices, isMobile }) {
  const [dateFrom,    setDateFrom]    = useState("");
  const [dateTo,      setDateTo]      = useState("");
  const [repOutcome,  setRepOutcome]  = useState("All");
  const [repType,     setRepType]     = useState("All");
  const [repExported, setRepExported] = useState(false);

  const allProcessed = devices.filter(d =>
    d.sentToPartner && d.outcome && ["Stock","Escalated","Scrap"].includes(d.stage)
  );

  const reportData = allProcessed.filter(d => {
    const afterFrom = !dateFrom || d.received >= dateFrom;
    const beforeTo  = !dateTo   || d.received <= dateTo;
    const matchOut  = repOutcome === "All" || d.outcome === repOutcome;
    const matchType = repType    === "All" || d.type    === repType;
    return afterFrom && beforeTo && matchOut && matchType;
  });

  const working    = reportData.filter(d => d.outcome === "Working").length;
  const notWorking = reportData.filter(d => d.outcome === "Not Working").length;
  const rate       = reportData.length ? Math.round(working / reportData.length * 100) : 0;

  function clearFilters() {
    setDateFrom(""); setDateTo(""); setRepOutcome("All"); setRepType("All");
  }

  function exportReport() {
    const rows = reportData.map(d => ({
      "Device ID":     d.id,
      "Serial Number": d.serial,
      "MAC Address":   d.mac    || "",
      "Model":         d.model  || "",
      "Device Type":   d.type,
      "Outcome":       d.outcome || "",
      "Received Date": d.received,
      "Notes":         d.notes  || "",
      "Partner Notes": d.partnerNotes || "",
    }));
    const summaryRows = [
      { "Metric": "Period From",    "Value": dateFrom    || "All time" },
      { "Metric": "Period To",      "Value": dateTo      || "All time" },
      { "Metric": "Device Type",    "Value": repType                   },
      { "Metric": "Outcome Filter", "Value": repOutcome                },
      { "Metric": "Total Devices",  "Value": reportData.length         },
      { "Metric": "Working",        "Value": working                   },
      { "Metric": "Not Working",    "Value": notWorking                },
      { "Metric": "Success Rate %", "Value": rate                      },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows),        "Devices");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), "Summary");
    const safeType    = repType    !== "All" ? repType.replace(/[\/\s]+/g, "_") : "";
    const safeOutcome = repOutcome !== "All" ? repOutcome.replace(/\s+/g, "_")  : "";
    const label = [dateFrom || "start", dateTo || "end", safeType, safeOutcome]
                    .filter(Boolean).join("_");
    XLSX.writeFile(wb, `partner_report_${label}.xlsx`);
    setRepExported(true);
    setTimeout(() => setRepExported(false), 2500);
  }

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>

      {/* ── FILTER CARD ── */}
      <Card>
        <SectionTitle>Report Builder</SectionTitle>
        <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(4,1fr)", gap:12, marginBottom:16 }}>
          <div>
            <Label>DATE FROM</Label>
            <input type="date" value={dateFrom} onChange={e=>setDateFrom(e.target.value)} style={iStyle()}/>
          </div>
          <div>
            <Label>DATE TO</Label>
            <input type="date" value={dateTo}   onChange={e=>setDateTo(e.target.value)}   style={iStyle()}/>
          </div>
          <div>
            <Label>OUTCOME</Label>
            <select value={repOutcome} onChange={e=>setRepOutcome(e.target.value)} style={iStyle()}>
              <option value="All">All outcomes</option>
              <option value="Working">Working</option>
              <option value="Not Working">Not Working</option>
            </select>
          </div>
          <div>
            <Label>DEVICE TYPE</Label>
            <select value={repType} onChange={e=>setRepType(e.target.value)} style={iStyle()}>
              <option value="All">All types</option>
              {DEVICE_TYPES.map(t => <option key={t}>{t}</option>)}
            </select>
          </div>
        </div>

        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:10 }}>
          <span style={{ fontSize:13, color: (dateFrom||dateTo) ? C.slate3 : C.slate4 }}>
            {dateFrom || dateTo
              ? <>Period: <strong>{dateFrom||"start"}</strong> → <strong>{dateTo||"today"}</strong></>
              : "No date filter — showing all time"}
          </span>
          <div style={{ display:"flex", gap:8, alignItems:"center" }}>
            {repExported && <span style={{ fontSize:12, color:C.greenDark, fontWeight:700 }}>✓ Downloaded!</span>}
            <Btn onClick={clearFilters} variant="ghost" size="sm">Clear filters</Btn>
            <Btn onClick={exportReport} variant="success">
              ↓ Export Excel ({reportData.length})
            </Btn>
          </div>
        </div>
      </Card>

      {/* ── STAT CARDS ── */}
      <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(3,1fr)", gap:12 }}>
        <StatCard label="Total in Period" value={reportData.length} sub="Matching filters"   accent={C.purple}/>
        <StatCard label="Working"         value={working}           sub={`${rate}% success`} accent={C.green}/>
        <StatCard label="Not Working"     value={notWorking}        sub="Failed QC"          accent={C.red}/>
      </div>

      {/* ── SUCCESS RATE ── */}
      {reportData.length > 0 && (
        <Card>
          <SectionTitle>Success Rate — Filtered Period</SectionTitle>
          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
            <span style={{ fontSize:13, color:C.slate3 }}>Working after refurbishment</span>
            <span style={{ fontSize:14, fontWeight:800, color:C.greenDark }}>{rate}%</span>
          </div>
          <div style={{ background:C.slate7, borderRadius:8, height:14, overflow:"hidden", marginBottom:16 }}>
            <div style={{ width:`${rate}%`, height:"100%",
              background:"linear-gradient(90deg,#16A34A,#22C55E)", borderRadius:8, transition:"width .5s" }}/>
          </div>

          {/* Per-type breakdown */}
          <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr":"repeat(3,1fr)", gap:10 }}>
            {DEVICE_TYPES.map(t => {
              const tData    = reportData.filter(d => d.type === t);
              const tWorking = tData.filter(d => d.outcome === "Working").length;
              const tRate    = tData.length ? Math.round(tWorking / tData.length * 100) : 0;
              return (
                <div key={t} style={{ background:C.slate8, borderRadius:10, padding:"10px 14px" }}>
                  <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
                    <span style={{ fontSize:12, fontWeight:700, color:C.slate }}>{TYPE_ICON[t]} {t}</span>
                    <span style={{ fontSize:12, fontWeight:800, color:tData.length ? C.greenDark : C.slate4 }}>
                      {tData.length ? `${tRate}%` : "—"}
                    </span>
                  </div>
                  <div style={{ background:C.slate6, borderRadius:6, height:8, overflow:"hidden", marginBottom:4 }}>
                    <div style={{ width:`${tRate}%`, height:"100%", background:C.green, borderRadius:6 }}/>
                  </div>
                  <div style={{ display:"flex", justifyContent:"space-between", fontSize:11, color:C.slate4 }}>
                    <span>✓ {tWorking} working</span>
                    <span>✗ {tData.length - tWorking} not working</span>
                  </div>
                </div>
              );
            })}
          </div>
        </Card>
      )}

      {/* ── DEVICE LIST ── */}
      {reportData.length > 0 ? (
        <Card>
          <SectionTitle>Devices in report ({reportData.length})</SectionTitle>
          {isMobile
            ? reportData.map(d => (
                <div key={d.id} style={{ border:`1px solid ${C.slate6}`, borderRadius:10, padding:12, marginBottom:8,
                  borderLeft:`3px solid ${d.outcome==="Working" ? C.green : C.red}` }}>
                  <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
                    <span style={{ fontFamily:"monospace", fontSize:12, fontWeight:700 }}>{d.serial}</span>
                    <span style={{ fontSize:11, fontWeight:700, padding:"2px 8px", borderRadius:20,
                      background: d.outcome==="Working" ? C.greenLight : C.redLight,
                      color:      d.outcome==="Working" ? C.greenDark  : C.redDark }}>
                      {d.outcome==="Working" ? "✓ Working" : "✗ Not Working"}
                    </span>
                  </div>
                  <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:4, fontSize:11, color:C.slate3 }}>
                    <span>{TYPE_ICON[d.type]} {d.type}</span>
                    <span>{d.model || "—"}</span>
                    <span>Received: {d.received}</span>
                    {(d.partnerNotes||d.notes) && <span>Note: {d.partnerNotes||d.notes}</span>}
                  </div>
                </div>
              ))
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead>
                    <tr style={{ borderBottom:`1.5px solid ${C.slate6}`, background:C.slate8 }}>
                      {["Serial","Model","Type","Outcome","Received","Notes"].map(h => (
                        <th key={h} style={{ padding:"8px 10px", textAlign:"left", fontSize:11,
                          fontWeight:700, color:C.slate4, letterSpacing:".05em", whiteSpace:"nowrap" }}>
                          {h.toUpperCase()}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {reportData.map(d => (
                      <tr key={d.id} style={{ borderBottom:`1px solid ${C.slate7}` }}>
                        <td style={{ padding:"9px 10px", fontFamily:"monospace", fontSize:12, fontWeight:600 }}>{d.serial}</td>
                        <td style={{ padding:"9px 10px", fontSize:12 }}>{d.model || "—"}</td>
                        <td style={{ padding:"9px 10px", whiteSpace:"nowrap" }}>{TYPE_ICON[d.type]} {d.type}</td>
                        <td style={{ padding:"9px 10px" }}>
                          <span style={{ fontSize:11, fontWeight:700, padding:"3px 10px", borderRadius:20,
                            background: d.outcome==="Working" ? C.greenLight : C.redLight,
                            color:      d.outcome==="Working" ? C.greenDark  : C.redDark }}>
                            {d.outcome==="Working" ? "✓ Working" : "✗ Not Working"}
                          </span>
                        </td>
                        <td style={{ padding:"9px 10px", color:C.slate3, whiteSpace:"nowrap" }}>{d.received}</td>
                        <td style={{ padding:"9px 10px", color:C.slate3, maxWidth:160,
                          overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                          {d.partnerNotes || d.notes || "—"}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
          }
        </Card>
      ) : (
        <Card>
          <p style={{ color:C.slate4, fontSize:13, textAlign:"center", padding:"20px 0", margin:0 }}>
            No processed devices match the selected filters.
          </p>
        </Card>
      )}
    </div>
  );
}


// ─────────────────────────────────────────────
// PARTNER RETURN TAB
// Devices confirmed by internal → ready to be
// physically returned to partner premises
// ─────────────────────────────────────────────
function PartnerReturnTab({ devices, setDevices, isMobile }) {
  const [selected, setSelected] = useState(new Set());
  const [showConfirm, setShowConfirm] = useState(false);
  const [released, setReleased] = useState(false);

  // Devices that have been internally confirmed (Working or Not Working)
  // and are currently in Stock or Escalated stage, sent via partner
  const returnPool = devices.filter(d =>
    d.sentToPartner &&
    d.outcome &&
    ["Working","Not Working"].includes(d.outcome) &&
    d.stage === "Confirmed"
  );

  const allSelected   = returnPool.length > 0 && selected.size === returnPool.length;
  const someSelected  = selected.size > 0;
  const selWorking    = returnPool.filter(d => selected.has(d.id) && d.outcome==="Working").length;
  const selNotWorking = returnPool.filter(d => selected.has(d.id) && d.outcome==="Not Working").length;

  function toggleAll() {
    setSelected(allSelected ? new Set() : new Set(returnPool.map(d=>d.id)));
  }
  function toggleOne(id) {
    setSelected(prev => {
      const n = new Set(prev);
      n.has(id) ? n.delete(id) : n.add(id);
      return n;
    });
  }

  function releaseToTransit() {
    setDevices(p => p.map(d => {
      if (!selected.has(d.id)) return d;
      return { ...d, stage: "In Transit" };
    }));
    // DB persist
    [...selected].forEach(id =>
      updateDevice(id, { stage: "In Transit" }).catch(e => console.warn("DB transit:", e))
    );
    setSelected(new Set());
    setShowConfirm(false);
    setReleased(true);
    setTimeout(() => setReleased(false), 3000);
  }

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
      <Card>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
          marginBottom:14, flexWrap:"wrap", gap:8 }}>
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <SectionTitle>Ready to Return ({returnPool.length})</SectionTitle>
            {someSelected && (
              <span style={{ background:"#EFF6FF", color:"#1D4ED8", fontSize:11, fontWeight:700,
                padding:"2px 9px", borderRadius:20 }}>
                {selected.size} selected
              </span>
            )}
          </div>
          <Btn onClick={()=>setShowConfirm(true)} disabled={!someSelected}
            style={{ background:someSelected?"#2563EB":"#BFDBFE",
              color:someSelected?"#fff":"#93C5FD", border:"none", fontWeight:800,
              boxShadow:someSelected?"0 2px 8px rgba(37,99,235,.35)":"none", transition:"all .2s" }}>
            🚚 Release to Transit ({selected.size})
          </Btn>
        </div>

        {released && <Alert type="success">✅ Devices released — now showing as In Transit in Stock view</Alert>}

        <p style={{ margin:"0 0 14px", fontSize:13, color:C.slate3 }}>
          These devices have been confirmed by the internal team. Select the ones being physically
          returned to partner premises and release them to <strong>In Transit</strong> status.
        </p>

        {returnPool.length === 0
          ? <p style={{ color:C.slate4, fontSize:13 }}>No confirmed devices ready for return.</p>
          : isMobile
            ? (
              <>
                <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:10,
                  padding:"8px 12px", background:C.slate8, borderRadius:8 }}>
                  <input type="checkbox" checked={allSelected} onChange={toggleAll}
                    style={{ width:16, height:16, cursor:"pointer" }}/>
                  <span style={{ fontSize:13, fontWeight:600, color:C.slate3 }}>Select All</span>
                </div>
                {returnPool.map(d => (
                  <div key={d.id} onClick={()=>toggleOne(d.id)}
                    style={{ border:`1.5px solid ${selected.has(d.id)?"#3B82F6":C.slate6}`,
                      borderRadius:12, padding:12, marginBottom:8, cursor:"pointer",
                      background:selected.has(d.id)?"#EFF6FF":C.white, transition:"all .15s" }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
                      <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                        <input type="checkbox" checked={selected.has(d.id)}
                          onChange={()=>toggleOne(d.id)}
                          onClick={e=>e.stopPropagation()}
                          style={{ width:16, height:16, cursor:"pointer" }}/>
                        <div>
                          <div style={{ fontFamily:"monospace", fontSize:12, fontWeight:700 }}>{d.serial}</div>
                          <div style={{ fontFamily:"monospace", fontSize:10, color:C.slate4 }}>{d.mac||"—"}</div>
                        </div>
                      </div>
                      <span style={{ fontSize:11, fontWeight:700, padding:"2px 8px", borderRadius:20,
                        background:d.outcome==="Working"?C.greenLight:C.redLight,
                        color:d.outcome==="Working"?C.greenDark:C.redDark }}>
                        {d.outcome==="Working"?"✓ Working":"✗ Not Working"}
                      </span>
                    </div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:4,
                      marginTop:8, fontSize:11, color:C.slate4 }}>
                      <span>{TYPE_ICON[d.type]} {d.type}</span>
                      <span>{d.model||"—"}</span>
                    </div>
                  </div>
                ))}
              </>
            )
            : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead>
                    <tr style={{ borderBottom:`1.5px solid ${C.slate6}`, background:C.slate8 }}>
                      <th style={{ padding:"8px 10px", width:40 }}>
                        <input type="checkbox" checked={allSelected} onChange={toggleAll}
                          style={{ width:15, height:15, cursor:"pointer" }}/>
                      </th>
                      {["Serial Number","MAC Address","Model","Type","Received","Outcome","Notes"].map(h => (
                        <th key={h} style={{ padding:"8px 10px", textAlign:"left", fontSize:11,
                          fontWeight:700, color:C.slate4, letterSpacing:".05em", whiteSpace:"nowrap" }}>
                          {h.toUpperCase()}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {returnPool.map(d => (
                      <tr key={d.id} onClick={()=>toggleOne(d.id)}
                        style={{ borderBottom:`1px solid ${C.slate7}`, cursor:"pointer",
                          background:selected.has(d.id)?"#EFF6FF":"transparent",
                          transition:"background .15s" }}>
                        <td style={{ padding:"10px" }}>
                          <input type="checkbox" checked={selected.has(d.id)}
                            onChange={()=>toggleOne(d.id)}
                            onClick={e=>e.stopPropagation()}
                            style={{ width:15, height:15, cursor:"pointer" }}/>
                        </td>
                        <td style={{ padding:"10px", fontFamily:"monospace", fontSize:12,
                          fontWeight:700, color:C.slate2, whiteSpace:"nowrap" }}>{d.serial}</td>
                        <td style={{ padding:"10px", fontFamily:"monospace", fontSize:11,
                          color:C.slate3, whiteSpace:"nowrap" }}>{d.mac||"—"}</td>
                        <td style={{ padding:"10px", fontSize:12 }}>{d.model||"—"}</td>
                        <td style={{ padding:"10px", whiteSpace:"nowrap" }}>{TYPE_ICON[d.type]} {d.type}</td>
                        <td style={{ padding:"10px", color:C.slate3, whiteSpace:"nowrap" }}>{d.received}</td>
                        <td style={{ padding:"10px" }}>
                          <span style={{ fontSize:11, fontWeight:700, padding:"3px 10px", borderRadius:20,
                            background:d.outcome==="Working"?C.greenLight:C.redLight,
                            color:d.outcome==="Working"?C.greenDark:C.redDark }}>
                            {d.outcome==="Working"?"✓ Working":"✗ Not Working"}
                          </span>
                        </td>
                        <td style={{ padding:"10px", color:C.slate4, maxWidth:130,
                          overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                          {d.notes||"—"}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
        }
      </Card>

      {/* ── RELEASE CONFIRMATION POPUP ── */}
      {showConfirm && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.6)", zIndex:200,
          display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
          <div style={{ background:C.white, borderRadius:16, padding:28, maxWidth:500,
            width:"100%", boxShadow:"0 24px 64px rgba(0,0,0,.3)" }}>
            <div style={{ fontSize:28, marginBottom:8 }}>🚚</div>
            <h3 style={{ margin:"0 0 6px", fontSize:20, fontWeight:800, color:C.slate }}>
              Release to In Transit
            </h3>
            <p style={{ margin:"0 0 18px", fontSize:13, color:C.slate3 }}>
              The selected devices will be marked as <strong>In Transit</strong> — physically
              en route back to partner premises. Internal team will confirm receipt.
            </p>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:16 }}>
              <div style={{ background:C.greenLight, borderRadius:12, padding:"14px 16px",
                border:`1.5px solid #BBF7D0`, textAlign:"center" }}>
                <div style={{ fontSize:30, fontWeight:800, color:C.greenDark }}>{selWorking}</div>
                <div style={{ fontSize:12, fontWeight:700, color:C.greenDark }}>✓ Working</div>
                <div style={{ fontSize:11, color:C.slate4, marginTop:3 }}>Will become Ready on receipt</div>
              </div>
              <div style={{ background:C.redLight, borderRadius:12, padding:"14px 16px",
                border:`1.5px solid #FECACA`, textAlign:"center" }}>
                <div style={{ fontSize:30, fontWeight:800, color:C.redDark }}>{selNotWorking}</div>
                <div style={{ fontSize:12, fontWeight:700, color:C.redDark }}>✗ Not Working</div>
                <div style={{ fontSize:11, color:C.slate4, marginTop:3 }}>Will become Scrap on receipt</div>
              </div>
            </div>
            {/* Scrollable device list */}
            <div style={{ background:C.slate8, borderRadius:10, padding:12, marginBottom:16,
              maxHeight:180, overflowY:"auto" }}>
              <div style={{ fontSize:11, fontWeight:700, color:C.slate4, letterSpacing:".06em",
                marginBottom:8, textTransform:"uppercase" }}>Devices being released</div>
              {returnPool.filter(d=>selected.has(d.id)).map(d => (
                <div key={d.id} style={{ display:"flex", justifyContent:"space-between",
                  alignItems:"center", padding:"5px 0", borderBottom:`1px solid ${C.slate6}`, gap:8 }}>
                  <div style={{ display:"flex", alignItems:"center", gap:8, minWidth:0 }}>
                    <span style={{ fontFamily:"monospace", fontSize:12, fontWeight:700, color:C.slate2, flexShrink:0 }}>{d.serial}</span>
                    {d.mac && <span style={{ fontFamily:"monospace", fontSize:10, color:C.slate4, flexShrink:0 }}>{d.mac}</span>}
                    {d.model && <span style={{ fontSize:11, color:C.slate4, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{d.model}</span>}
                  </div>
                  <span style={{ fontSize:11, fontWeight:700, padding:"2px 9px", borderRadius:20, flexShrink:0,
                    background:d.outcome==="Working"?C.greenLight:C.redLight,
                    color:d.outcome==="Working"?C.greenDark:C.redDark }}>
                    {d.outcome==="Working"?"✓ Working":"✗ Not Working"}
                  </span>
                </div>
              ))}
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
              <Btn onClick={()=>setShowConfirm(false)} variant="ghost" full size="lg">✕ Cancel</Btn>
              <Btn onClick={releaseToTransit} full size="lg"
                style={{ background:"#2563EB", color:"#fff", border:"none", fontWeight:800,
                  boxShadow:"0 2px 10px rgba(37,99,235,.4)" }}>
                🚚 Confirm Release
              </Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}


// ─────────────────────────────────────────────
// PARTNER HISTORY
// ─────────────────────────────────────────────
function PartnerHistory({ devices, isMobile }) {
  const allProcessed = devices.filter(d =>
    d.sentToPartner && d.outcome && ["Stock","Escalated","Scrap"].includes(d.stage)
  );
  const working    = allProcessed.filter(d => d.outcome === "Working");
  const notWorking = allProcessed.filter(d => d.outcome === "Not Working");
  const byType     = ["Router/Modem","Set-top Box","ONT/OLT"].map(t => ({
    type:t, total:allProcessed.filter(d=>d.type===t).length,
    working:allProcessed.filter(d=>d.type===t&&d.outcome==="Working").length,
    notWorking:allProcessed.filter(d=>d.type===t&&d.outcome==="Not Working").length,
  }));
  const rate = allProcessed.length ? Math.round(working.length/allProcessed.length*100) : 0;
  const ICONS = { "Router/Modem":"\u2B21","Set-top Box":"\u25A6","ONT/OLT":"\u25C8" };

  function DevTable({ items, color, bg }) {
    if (!items.length) return null;
    return isMobile
      ? items.map(d=>(
          <div key={d.id} style={{ border:`1px solid #E2E8F0`,borderRadius:10,padding:12,marginBottom:8,borderLeft:`3px solid ${color}` }}>
            <div style={{ display:"flex",justifyContent:"space-between",marginBottom:6 }}>
              <span style={{ fontFamily:"monospace",fontSize:12,fontWeight:700 }}>{d.serial}</span>
              <span style={{ fontSize:11,color:"#64748B" }}>{d.received}</span>
            </div>
            <div style={{ fontSize:12,color:"#475569" }}>{ICONS[d.type]} {d.type} · {d.model||"—"}</div>
          </div>
        ))
      : (
        <div style={{ overflowX:"auto",marginBottom:16 }}>
          <table style={{ width:"100%",borderCollapse:"collapse",fontSize:13 }}>
            <thead><tr style={{ borderBottom:"1.5px solid #E2E8F0",background:bg }}>
              {["Serial","Model","Type","MAC","Received","Notes"].map(h=>(
                <th key={h} style={{ padding:"7px 10px",textAlign:"left",fontSize:11,fontWeight:700,color,letterSpacing:".05em",whiteSpace:"nowrap" }}>{h.toUpperCase()}</th>
              ))}
            </tr></thead>
            <tbody>
              {items.map(d=>(
                <tr key={d.id} style={{ borderBottom:"1px solid #F1F5F9" }}>
                  <td style={{ padding:"9px 10px",fontFamily:"monospace",fontSize:12,fontWeight:600 }}>{d.serial}</td>
                  <td style={{ padding:"9px 10px",fontSize:12 }}>{d.model||"\u2014"}</td>
                  <td style={{ padding:"9px 10px",whiteSpace:"nowrap" }}>{ICONS[d.type]} {d.type}</td>
                  <td style={{ padding:"9px 10px",fontFamily:"monospace",fontSize:11,color:"#475569" }}>{d.mac||"\u2014"}</td>
                  <td style={{ padding:"9px 10px",color:"#64748B",whiteSpace:"nowrap" }}>{d.received}</td>
                  <td style={{ padding:"9px 10px",color:"#64748B",maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{d.notes||"\u2014"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
  }

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      <div style={{ display:"grid",gridTemplateColumns:isMobile?"1fr 1fr":"repeat(3,1fr)",gap:12 }}>
        <StatCard label="Total Processed" value={allProcessed.length} sub="All confirmed" accent="#7C3AED"/>
        <StatCard label="Working"     value={working.length}    sub={`${rate}% success`} accent="#22C55E"/>
        <StatCard label="Not Working" value={notWorking.length} sub="Failed QC"          accent="#EF4444"/>
      </div>
      <Card>
        <SectionTitle>Success Rate</SectionTitle>
        <div style={{ display:"flex",justifyContent:"space-between",marginBottom:6 }}>
          <span style={{ fontSize:13,color:"#475569" }}>Working after refurbishment</span>
          <span style={{ fontSize:14,fontWeight:800,color:"#166534" }}>{rate}%</span>
        </div>
        <div style={{ background:"#F1F5F9",borderRadius:8,height:14,overflow:"hidden",marginBottom:14 }}>
          <div style={{ width:`${rate}%`,height:"100%",background:"linear-gradient(90deg,#16A34A,#22C55E)",borderRadius:8,transition:"width .5s" }}/>
        </div>
        <div style={{ display:"grid",gridTemplateColumns:isMobile?"1fr":"repeat(3,1fr)",gap:10 }}>
          {byType.map(r=>{
            const tRate = r.total ? Math.round(r.working/r.total*100) : 0;
            return (
              <div key={r.type} style={{ background:"#F8FAFC",borderRadius:10,padding:"10px 14px" }}>
                <div style={{ display:"flex",justifyContent:"space-between",marginBottom:6 }}>
                  <span style={{ fontSize:12,fontWeight:700,color:"#0F172A" }}>{ICONS[r.type]} {r.type}</span>
                  <span style={{ fontSize:12,fontWeight:800,color:r.total?"#166534":"#94A3B8" }}>{r.total?`${tRate}%`:"\u2014"}</span>
                </div>
                <div style={{ background:"#CBD5E1",borderRadius:6,height:8,overflow:"hidden",marginBottom:4 }}>
                  <div style={{ width:`${tRate}%`,height:"100%",background:"#22C55E",borderRadius:6 }}/>
                </div>
                <div style={{ display:"flex",justifyContent:"space-between",fontSize:11,color:"#94A3B8" }}>
                  <span>\u2713 {r.working} working</span><span>\u2717 {r.notWorking} not working</span>
                </div>
              </div>
            );
          })}
        </div>
      </Card>
      <Card>
        <SectionTitle>Processed Devices</SectionTitle>
        {allProcessed.length===0
          ? <p style={{ color:"#94A3B8",fontSize:13,margin:0 }}>No devices processed yet.</p>
          : <>
              {working.length>0 && (
                <div style={{ marginBottom:16 }}>
                  <div style={{ display:"flex",alignItems:"center",gap:8,marginBottom:10,paddingBottom:8,borderBottom:"2px solid #166534" }}>
                    <span style={{ fontSize:16 }}>\u2705</span>
                    <h4 style={{ margin:0,fontSize:13,fontWeight:800,color:"#166534" }}>Working</h4>
                    <span style={{ background:"#F0FDF4",color:"#166534",fontSize:11,fontWeight:700,padding:"2px 9px",borderRadius:20 }}>{working.length}</span>
                  </div>
                  <DevTable items={working} color="#166534" bg="#F0FDF4"/>
                </div>
              )}
              {notWorking.length>0 && (
                <div>
                  <div style={{ display:"flex",alignItems:"center",gap:8,marginBottom:10,paddingBottom:8,borderBottom:"2px solid #991B1B" }}>
                    <span style={{ fontSize:16 }}>\u274C</span>
                    <h4 style={{ margin:0,fontSize:13,fontWeight:800,color:"#991B1B" }}>Not Working</h4>
                    <span style={{ background:"#FEF2F2",color:"#991B1B",fontSize:11,fontWeight:700,padding:"2px 9px",borderRadius:20 }}>{notWorking.length}</span>
                  </div>
                  <DevTable items={notWorking} color="#991B1B" bg="#FEF2F2"/>
                </div>
              )}
            </>
        }
      </Card>
    </div>
  );
}

function PartnerPortal({ devices, setDevices, uploadLogs, setUploadLogs, isMobile }) {
  const [portalTab, setPortalTab]   = useState("live");
  const [edits, setEdits]           = useState({});
  const [submitting, setSubmitting] = useState(false);
  const [submitted, setSubmitted]   = useState(false);
  const [showExecConfirm, setShowExecConfirm] = useState(false);
  const [dragOver, setDragOver]     = useState(false);
  const [processing, setProcessing] = useState(false);
  const [lastResult, setLastResult] = useState(null);

  const assigned     = devices.filter(d=>d.sentToPartner && d.stage==="Refurbishment" && !d.partnerOutcome);
  const submitted2   = devices.filter(d=>d.sentToPartner && d.partnerOutcome).length;
  const done         = devices.filter(d=>d.sentToPartner && ["Stock","Escalated","Scrap"].includes(d.stage) && d.outcome).length;
  const pendingEdits = Object.values(edits).filter(v=>v.outcome).length;
  const pendingWorking    = Object.values(edits).filter(v=>v.outcome==="Working").length;
  const pendingNotWorking = Object.values(edits).filter(v=>v.outcome==="Not Working").length;

  function setEdit(id, field, value) { setEdits(p=>({ ...p, [id]:{ ...p[id], [field]:value } })); }
  function getEdit(id, field, fb="") { return edits[id]?.[field] ?? fb; }

  // Execute the queue — submit all marked outcomes
  function executeQueue() {
    setSubmitting(true);
    setTimeout(() => {
      setDevices(p=>p.map(d=>{ const e=edits[d.id]; return e?.outcome ? { ...d,partnerOutcome:e.outcome,partnerNotes:e.notes||"" } : d; }));
      setEdits({}); setSubmitting(false); setSubmitted(true);
      setShowExecConfirm(false);
      setTimeout(()=>setSubmitted(false), 3000);
    }, 400);
  }

  function normalizeRow(row) { const o={}; Object.keys(row).forEach(k=>{ o[k.trim().toLowerCase().replace(/\s+/g,"_")]=String(row[k]||"").trim(); }); return o; }
  function applyRows(rows) {
    const norm=rows.map(normalizeRow).filter(r=>r.serial_number||r.serial);
    const matched=[], unmatched=[], invalid=[], updates={};
    norm.forEach(row=>{
      const serial=(row.serial_number||row.serial||"").trim();
      const rawO=row.outcome||"";
      const outcome=rawO.toLowerCase().includes("not")?"Not Working":rawO.toLowerCase().includes("work")?"Working":null;
      const dev=devices.find(d=>d.serial.toLowerCase()===serial.toLowerCase()&&d.sentToPartner&&d.stage==="Refurbishment");
      if (!dev)    { unmatched.push(serial); return; }
      if (!outcome){ invalid.push(serial);   return; }
      matched.push(serial); updates[dev.id]={ partnerOutcome:outcome, partnerNotes:row.notes||"" };
    });
    setDevices(p=>p.map(d=>updates[d.id]?{...d,...updates[d.id]}:d));
    const log={ id:Date.now(), timestamp:new Date().toLocaleString(), total:norm.length,
      matched:matched.length, unmatched:unmatched.length, invalid:invalid.length,
      unmatchedSerials:unmatched, invalidSerials:invalid };
    setUploadLogs(p=>[log,...p]); setLastResult(log);
  }
  function parseFile(file) {
    setProcessing(true); setLastResult(null);
    const ext=file.name.split(".").pop().toLowerCase();
    if (ext==="csv") { Papa.parse(file,{ header:true, skipEmptyLines:true, complete:(r)=>{ applyRows(r.data); setProcessing(false); }, error:()=>setProcessing(false) }); }
    else if (ext==="xlsx"||ext==="xls") { const rd=new FileReader(); rd.onload=(e)=>{ const wb=XLSX.read(e.target.result,{type:"array"}); applyRows(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""})); setProcessing(false); }; rd.readAsArrayBuffer(file); }
    else setProcessing(false);
  }
  function downloadTemplate() {
    const rows=[["serial_number","outcome","notes","return_date"],["SN-EXAMPLE1","Working","Firmware updated","2025-05-20"],["SN-EXAMPLE2","Not Working","Power failure","2025-05-20"]];
    const a=document.createElement("a"); a.href=URL.createObjectURL(new Blob([rows.map(r=>r.join(",")).join("\n")],{type:"text/csv"})); a.download="partner_results_template.csv"; a.click();
  }

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      <div style={{ background:"linear-gradient(135deg,#4F46E5,#7C3AED)", borderRadius:14, padding:"20px 20px 16px" }}>
        <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:6 }}>
          <span style={{ fontSize:18 }}>🔒</span>
          <h2 style={{ margin:0, fontSize:isMobile?16:20, fontWeight:800, color:C.white }}>Partner Portal</h2>
          <span style={{ background:"rgba(255,255,255,.2)", fontSize:10, fontWeight:700, padding:"2px 8px", borderRadius:20, color:C.white }}>RESTRICTED</span>
        </div>
        <p style={{ margin:"0 0 14px", opacity:.8, fontSize:12, color:C.white }}>Submit QC outcomes — internal team will confirm before stock is updated</p>
        <div style={{ display:"flex", background:"rgba(0,0,0,.2)", borderRadius:10, padding:3, gap:2 }}>
          {[{id:"live",label:"📋 Refurb View"},{id:"upload",label:"📂 Bulk Upload"},{id:"return",label:"📦 Ready to Return"},{id:"history",label:"📊 History"},{id:"reports",label:"📑 Reports"}].map(t=>(
            <button key={t.id} onClick={()=>setPortalTab(t.id)}
              style={{ flex:1, padding:"6px 10px", border:"none", borderRadius:8, fontSize:12, fontWeight:700, cursor:"pointer",
                background:portalTab===t.id?"rgba(255,255,255,.2)":"transparent",
                color:C.white, transition:"all .15s" }}>
              {t.label}
            </button>
          ))}
        </div>
      </div>

      <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(3,1fr)", gap:10 }}>
        <StatCard label="Assigned" value={assigned.length} sub="Awaiting your results" accent={C.purple}/>
        <StatCard label="Submitted" value={submitted2} sub="Pending confirmation" accent={C.amber}/>
        <StatCard label="Completed" value={done} sub="Confirmed & closed" accent={C.green}/>
      </div>

      {portalTab !== "search" && (
        <DeviceSearchBar devices={devices} isMobile={isMobile} scope="partner"
          placeholder="Quick search partner devices…"/>
      )}

      {/* LIVE INPUT — with queue + execute */}
      {portalTab==="live" && (
        <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
          <Card>
            {/* ── Header row ── */}
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
              marginBottom:14, flexWrap:"wrap", gap:8 }}>
              <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                <SectionTitle>Assigned devices ({assigned.length})</SectionTitle>
                {pendingEdits>0 && (
                  <span style={{ background:C.indigoLight, color:C.indigoDark, fontSize:11,
                    fontWeight:700, padding:"2px 9px", borderRadius:20 }}>
                    {pendingEdits} in queue
                  </span>
                )}
              </div>
              <Btn onClick={()=>setShowExecConfirm(true)} disabled={pendingEdits===0}
                style={{ background:pendingEdits>0?"#16A34A":"#D1FAE5",
                  color:pendingEdits>0?"#fff":"#6EE7B7", border:"none", fontWeight:800,
                  boxShadow:pendingEdits>0?"0 2px 8px rgba(22,163,74,.35)":"none", transition:"all .2s" }}>
                ▶ Execute Queue {pendingEdits>0?`(${pendingEdits})`:""}
              </Btn>
            </div>

            {submitted && <Alert type="success">✓ Batch submitted — awaiting internal confirmation</Alert>}

            {assigned.length===0
              ? <p style={{ color:C.slate4, fontSize:13 }}>No devices currently assigned to your facility.</p>
              : isMobile
                /* ── MOBILE: stacked cards ── */
                ? (
                  <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
                    {assigned.map(d => {
                      const outcome = getEdit(d.id,"outcome","");
                      const notes   = getEdit(d.id,"notes","");
                      return (
                        <div key={d.id} style={{
                          border:`1.5px solid ${outcome==="Working"?"#BBF7D0":outcome==="Not Working"?"#FECACA":C.slate6}`,
                          borderRadius:12, padding:14,
                          background:outcome==="Working"?"#F0FFF4":outcome==="Not Working"?"#FFF5F5":C.white,
                          borderLeft:`3px solid ${outcome==="Working"?C.green:outcome==="Not Working"?C.red:C.slate5}` }}>
                          <div style={{ marginBottom:8 }}>
                            <span style={{ fontFamily:"monospace", fontSize:13, fontWeight:700, color:C.slate2 }}>{d.serial}</span>
                            {d.mac && <span style={{ fontFamily:"monospace", fontSize:11, color:C.slate4, marginLeft:8 }}>{d.mac}</span>}
                          </div>
                          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:4, fontSize:12, marginBottom:10 }}>
                            <span style={{ color:C.slate3 }}>{TYPE_ICON[d.type]} {d.type}</span>
                            <span style={{ color:C.slate3 }}>{d.model||"—"}</span>
                            <span style={{ color:C.slate4 }}>{d.received}</span>
                            {d.notes && <span style={{ color:C.slate4, fontStyle:"italic" }}>{d.notes}</span>}
                          </div>
                          <div style={{ display:"flex", gap:6, marginBottom:8 }}>
                            <button onClick={()=>setEdit(d.id,"outcome",outcome==="Working"?"":"Working")}
                              style={{ flex:1, padding:"7px 0", borderRadius:8, fontSize:12, fontWeight:700, cursor:"pointer",
                                border:`2px solid ${outcome==="Working"?"#16A34A":C.slate6}`,
                                background:outcome==="Working"?C.greenLight:C.white,
                                color:outcome==="Working"?C.greenDark:C.slate3 }}>
                              ✓ Working
                            </button>
                            <button onClick={()=>setEdit(d.id,"outcome",outcome==="Not Working"?"":"Not Working")}
                              style={{ flex:1, padding:"7px 0", borderRadius:8, fontSize:12, fontWeight:700, cursor:"pointer",
                                border:`2px solid ${outcome==="Not Working"?C.redDark:C.slate6}`,
                                background:outcome==="Not Working"?C.redLight:C.white,
                                color:outcome==="Not Working"?C.redDark:C.slate3 }}>
                              ✗ Not Working
                            </button>
                          </div>
                          <input value={notes} onChange={e=>setEdit(d.id,"notes",e.target.value)}
                            placeholder="Repair notes…" style={{ ...iStyle(), width:"100%" }}/>
                          {outcome && (
                            <div style={{ marginTop:8, display:"flex", alignItems:"center", gap:6 }}>
                              <span style={{ fontSize:11, fontWeight:700, padding:"2px 9px", borderRadius:20,
                                background:outcome==="Working"?C.greenLight:C.redLight,
                                color:outcome==="Working"?C.greenDark:C.redDark }}>
                                {outcome==="Working"?"✓ Working":"✗ Not Working"}
                              </span>
                              <button onClick={()=>setEdit(d.id,"outcome","")}
                                style={{ background:"none", border:"none", cursor:"pointer",
                                  color:C.slate4, fontSize:11 }}>✕ clear</button>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )
                /* ── DESKTOP: full inline table ── */
                : (
                  <div style={{ overflowX:"auto" }}>
                    <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                      <thead>
                        <tr style={{ borderBottom:`1.5px solid ${C.slate6}`, background:C.slate8 }}>
                          {["Serial Number","MAC Address","Model","Type","Received","Internal Notes",
                            "QC Outcome","Repair Notes"].map(h => (
                            <th key={h} style={{ padding:"8px 10px", textAlign:"left", fontSize:11,
                              fontWeight:700, color:C.slate4, letterSpacing:".05em", whiteSpace:"nowrap" }}>
                              {h.toUpperCase()}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {assigned.map(d => {
                          const outcome = getEdit(d.id,"outcome","");
                          const notes   = getEdit(d.id,"notes","");
                          return (
                            <tr key={d.id} style={{ borderBottom:`1px solid ${C.slate7}`,
                              background: outcome==="Working"?"#F0FFF4"
                                :outcome==="Not Working"?"#FFF5F5":"transparent",
                              transition:"background .2s" }}>

                              {/* Serial */}
                              <td style={{ padding:"10px", fontFamily:"monospace", fontSize:12,
                                fontWeight:700, color:C.slate2, whiteSpace:"nowrap" }}>
                                {d.serial}
                              </td>

                              {/* MAC */}
                              <td style={{ padding:"10px", fontFamily:"monospace", fontSize:11,
                                color:C.slate3, whiteSpace:"nowrap" }}>
                                {d.mac||"—"}
                              </td>

                              {/* Model */}
                              <td style={{ padding:"10px", fontSize:12, maxWidth:140,
                                overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                                {d.model||"—"}
                              </td>

                              {/* Type */}
                              <td style={{ padding:"10px", whiteSpace:"nowrap" }}>
                                {TYPE_ICON[d.type]} {d.type}
                              </td>

                              {/* Received */}
                              <td style={{ padding:"10px", color:C.slate3, whiteSpace:"nowrap" }}>
                                {d.received}
                              </td>

                              {/* Internal Notes */}
                              <td style={{ padding:"10px", color:C.slate4, maxWidth:130,
                                overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap",
                                fontStyle:"italic", fontSize:12 }}>
                                {d.notes||"—"}
                              </td>

                              {/* QC Outcome — inline toggle */}
                              <td style={{ padding:"8px 10px", whiteSpace:"nowrap" }}>
                                <div style={{ display:"flex", gap:5 }}>
                                  <button onClick={()=>setEdit(d.id,"outcome",outcome==="Working"?"":"Working")}
                                    style={{ padding:"5px 11px", borderRadius:7, fontSize:11, fontWeight:700,
                                      cursor:"pointer", whiteSpace:"nowrap", transition:"all .15s",
                                      border:`2px solid ${outcome==="Working"?"#16A34A":C.slate5}`,
                                      background:outcome==="Working"?C.greenLight:C.white,
                                      color:outcome==="Working"?C.greenDark:C.slate3 }}>
                                    ✓ Working
                                  </button>
                                  <button onClick={()=>setEdit(d.id,"outcome",outcome==="Not Working"?"":"Not Working")}
                                    style={{ padding:"5px 11px", borderRadius:7, fontSize:11, fontWeight:700,
                                      cursor:"pointer", whiteSpace:"nowrap", transition:"all .15s",
                                      border:`2px solid ${outcome==="Not Working"?C.redDark:C.slate5}`,
                                      background:outcome==="Not Working"?C.redLight:C.white,
                                      color:outcome==="Not Working"?C.redDark:C.slate3 }}>
                                    ✗ Not Working
                                  </button>
                                  {outcome && (
                                    <button onClick={()=>setEdit(d.id,"outcome","")}
                                      style={{ padding:"5px 7px", borderRadius:7, fontSize:11, fontWeight:700,
                                        cursor:"pointer", border:`1px solid ${C.slate5}`,
                                        background:C.white, color:C.slate4 }}>
                                      ✕
                                    </button>
                                  )}
                                </div>
                              </td>

                              {/* Repair Notes */}
                              <td style={{ padding:"8px 10px" }}>
                                <input value={notes} onChange={e=>setEdit(d.id,"notes",e.target.value)}
                                  placeholder="Repair notes…"
                                  style={{ ...iStyle(), width:160, padding:"5px 10px", fontSize:12 }}/>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )
            }

            {/* ── Queue summary bar ── */}
            {pendingEdits>0 && (
              <div style={{ marginTop:14, paddingTop:12, borderTop:`1px solid ${C.slate7}`,
                display:"flex", justifyContent:"space-between", alignItems:"center",
                flexWrap:"wrap", gap:10 }}>
                <div style={{ fontSize:13, color:C.slate3 }}>
                  <strong style={{ color:C.greenDark }}>{pendingWorking}</strong> Working ·{" "}
                  <strong style={{ color:C.redDark }}>{pendingNotWorking}</strong> Not Working queued
                </div>
                <Btn onClick={()=>setShowExecConfirm(true)} disabled={pendingEdits===0}
                  style={{ background:"#16A34A", color:"#fff", border:"none", fontWeight:800,
                    boxShadow:"0 2px 8px rgba(22,163,74,.35)" }}>
                  ▶ Execute Queue ({pendingEdits})
                </Btn>
              </div>
            )}
          </Card>

          {/* ── EXECUTE CONFIRMATION POPUP ── */}
          {showExecConfirm && (
            <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.6)", zIndex:200,
              display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
              <div style={{ background:C.white, borderRadius:16, padding:28, maxWidth:520,
                width:"100%", boxShadow:"0 24px 64px rgba(0,0,0,.3)" }}>

                <div style={{ fontSize:28, marginBottom:8 }}>📤</div>
                <h3 style={{ margin:"0 0 6px", fontSize:20, fontWeight:800, color:C.slate }}>
                  Confirm Queue Execution
                </h3>
                <p style={{ margin:"0 0 20px", fontSize:13, color:C.slate3 }}>
                  The following outcomes will be submitted to the internal team for final confirmation.
                  <strong> This action cannot be undone.</strong>
                </p>

                {/* Summary cards */}
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:16 }}>
                  <div style={{ background:C.greenLight, borderRadius:12, padding:"16px",
                    border:`1.5px solid #BBF7D0`, textAlign:"center" }}>
                    <div style={{ fontSize:32, fontWeight:800, color:C.greenDark }}>{pendingWorking}</div>
                    <div style={{ fontSize:13, fontWeight:700, color:C.greenDark, marginTop:2 }}>✓ Working</div>
                    <div style={{ fontSize:11, color:C.slate4, marginTop:4 }}>→ Sent for stock confirmation</div>
                  </div>
                  <div style={{ background:C.redLight, borderRadius:12, padding:"16px",
                    border:`1.5px solid #FECACA`, textAlign:"center" }}>
                    <div style={{ fontSize:32, fontWeight:800, color:C.redDark }}>{pendingNotWorking}</div>
                    <div style={{ fontSize:13, fontWeight:700, color:C.redDark, marginTop:2 }}>✗ Not Working</div>
                    <div style={{ fontSize:11, color:C.slate4, marginTop:4 }}>→ Sent for escalation</div>
                  </div>
                </div>

                {/* Device list preview in popup */}
                {pendingEdits > 0 && (
                  <div style={{ background:C.slate8, borderRadius:10, padding:12, marginBottom:16,
                    maxHeight:200, overflowY:"auto" }}>
                    <div style={{ fontSize:11, fontWeight:700, color:C.slate4, letterSpacing:".06em",
                      marginBottom:8, textTransform:"uppercase" }}>Devices being submitted</div>
                    {assigned
                      .filter(d => getEdit(d.id,"outcome",""))
                      .map(d => {
                        const outcome = getEdit(d.id,"outcome","");
                        return (
                          <div key={d.id} style={{ display:"flex", justifyContent:"space-between",
                            alignItems:"center", padding:"6px 0",
                            borderBottom:`1px solid ${C.slate6}`, gap:8 }}>
                            <div style={{ display:"flex", alignItems:"center", gap:8, minWidth:0 }}>
                              <span style={{ fontFamily:"monospace", fontSize:12, fontWeight:700,
                                color:C.slate2, flexShrink:0 }}>{d.serial}</span>
                              {d.mac && <span style={{ fontFamily:"monospace", fontSize:10,
                                color:C.slate4, flexShrink:0 }}>{d.mac}</span>}
                              {d.model && <span style={{ fontSize:11, color:C.slate4,
                                overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                                {d.model}
                              </span>}
                            </div>
                            <span style={{ fontSize:11, fontWeight:700, padding:"2px 9px",
                              borderRadius:20, flexShrink:0,
                              background:outcome==="Working"?C.greenLight:C.redLight,
                              color:outcome==="Working"?C.greenDark:C.redDark }}>
                              {outcome==="Working"?"✓ Working":"✗ Not Working"}
                            </span>
                          </div>
                        );
                      })
                    }
                  </div>
                )}

                {/* Unsubmitted warning */}
                {assigned.length - pendingEdits > 0 && (
                  <Alert type="warning">
                    ⚠ <strong>{assigned.length - pendingEdits}</strong> device{assigned.length-pendingEdits>1?"s":""} without an outcome will remain in the queue.
                  </Alert>
                )}

                {/* Action buttons */}
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginTop:16 }}>
                  <Btn onClick={()=>setShowExecConfirm(false)} variant="ghost" full size="lg">
                    ✕ Cancel
                  </Btn>
                  <Btn onClick={executeQueue} full size="lg" disabled={submitting}
                    style={{ background:"#16A34A", color:"#fff", border:"none", fontWeight:800,
                      boxShadow:"0 2px 10px rgba(22,163,74,.4)" }}>
                    {submitting?"Submitting…":"✓ Confirm & Execute"}
                  </Btn>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* ── READY TO RETURN TAB ── */}
      {portalTab==="return" && (
        <PartnerReturnTab devices={devices} setDevices={setDevices} isMobile={isMobile}/>
      )}

      {/* BULK UPLOAD */}
      {portalTab==="upload" && (
        <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr":"1fr 1fr", gap:16 }}>
          <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
            <Card>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
                <SectionTitle>Upload QC Results</SectionTitle>
                <Btn onClick={downloadTemplate} variant="purple" size="sm">↓ Template</Btn>
              </div>
              <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)}
                onDrop={e=>{e.preventDefault();setDragOver(false);const f=e.dataTransfer.files[0];if(f)parseFile(f);}}
                onClick={()=>document.getElementById("partner-file").click()}
                style={{ border:`2px dashed ${dragOver?C.purple:C.slate5}`, borderRadius:12, padding:"28px 16px",
                  textAlign:"center", background:dragOver?C.purpleLight:C.slate8, transition:"all .2s", cursor:"pointer" }}>
                <input id="partner-file" type="file" accept=".csv,.xlsx,.xls"
                  onChange={e=>{const f=e.target.files[0];if(f)parseFile(f);e.target.value="";}} style={{ display:"none" }}/>
                <div style={{ fontSize:28, marginBottom:6 }}>{processing?"⏳":"📂"}</div>
                <div style={{ fontSize:13, fontWeight:700, color:C.slate2, marginBottom:4 }}>{processing?"Processing…":"Drop or tap to browse"}</div>
                <div style={{ fontSize:11, color:C.slate4 }}>CSV or Excel · serial_number · outcome · notes · return_date</div>
              </div>
              <div style={{ marginTop:12, background:C.slate8, borderRadius:8, padding:"10px 14px", fontSize:12, color:C.slate3 }}>
                ℹ Bulk uploads also go through internal confirmation before stock is updated.
              </div>
            </Card>
            {lastResult && (
              <Card>
                <SectionTitle>Upload Result</SectionTitle>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:8 }}>
                  {[[lastResult.matched,"Matched",C.greenDark,C.greenLight],[lastResult.unmatched,"Unmatched",C.redDark,C.redLight],[lastResult.invalid,"Invalid","#92400E",C.amberLight]].map(([n,l,c,bg])=>(
                    <div key={l} style={{ background:bg, borderRadius:8, padding:10, textAlign:"center" }}>
                      <div style={{ fontSize:22, fontWeight:800, color:c }}>{n}</div>
                      <div style={{ fontSize:11, color:c, fontWeight:700 }}>{l}</div>
                    </div>
                  ))}
                </div>
                {lastResult.unmatchedSerials.length>0 && <div style={{ marginTop:10 }}><Alert type="danger">Unmatched: {lastResult.unmatchedSerials.join(", ")}</Alert></div>}
              </Card>
            )}
          </div>
          <Card>
            <SectionTitle>Awaiting results ({assigned.length})</SectionTitle>
            {assigned.length===0
              ? <p style={{ color:C.slate4, fontSize:13 }}>None pending.</p>
              : assigned.map(d=>(
                  <div key={d.id} style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
                    padding:"10px 0", borderBottom:`1px solid ${C.slate7}`, gap:8 }}>
                    <div>
                      <span style={{ fontFamily:"monospace", fontSize:12, fontWeight:700 }}>{d.serial}</span>
                      <span style={{ fontSize:11, color:C.slate4, marginLeft:6 }}>{d.model||d.type}</span>
                    </div>
                    <span style={{ fontSize:11, color:C.slate4 }}>{d.received}</span>
                  </div>
                ))
            }
            {uploadLogs.length>0 && (
              <div style={{ marginTop:16, borderTop:`1px solid ${C.slate7}`, paddingTop:14 }}>
                <div style={{ fontSize:11, fontWeight:700, color:C.slate4, marginBottom:8, letterSpacing:".06em" }}>UPLOAD HISTORY</div>
                {uploadLogs.slice(0,5).map(log=>(
                  <div key={log.id} style={{ display:"flex", justifyContent:"space-between", fontSize:12, marginBottom:6, color:C.slate3 }}>
                    <span>{log.timestamp}</span>
                    <span style={{ color:C.greenDark, fontWeight:700 }}>{log.matched} matched</span>
                  </div>
                ))}
              </div>
            )}
          </Card>
        </div>
      )}
      {/* ── HISTORY TAB ── */}
      {portalTab==="history" && (
        <PartnerHistory devices={devices} isMobile={isMobile}/>
      )}

    </div>
  );
}

// ─────────────────────────────────────────────
// APP SHELL
// ─────────────────────────────────────────────

// ─────────────────────────────────────────────
// ROLE CONFIGURATION
// ─────────────────────────────────────────────
const ROLES = {
  admin: {
    label:   "Administrator",
    icon:    "🛡️",
    color:   "#6366F1",
    bg:      "#EEF2FF",
    views:   ["dashboard","intake","refurb","stock","all","users","partnerportal"],
    canSwitchToPartner: true,
    canManageUsers:     true,
  },
  stock: {
    label:   "Stock Management",
    icon:    "📦",
    color:   "#16A34A",
    bg:      "#F0FDF4",
    views:   ["dashboard","intake","stock","all"],
    canSwitchToPartner: false,
    canManageUsers:     false,
  },
  partner: {
    label:   "Refurbishment Partner",
    icon:    "🔧",
    color:   "#7C3AED",
    bg:      "#FDF4FF",
    views:   [],                // partner has its own portal, no internal views
    canSwitchToPartner: true,
    canManageUsers:     false,
    defaultMode:        "partner",
  },
};

// ─────────────────────────────────────────────
// DEMO USERS  (prototype — no real auth)
// ─────────────────────────────────────────────
const DEMO_USERS_INIT = [
  { id:1, name:"Alice Admin",   username:"admin",   password:"admin123",   role:"admin"   },
  { id:2, name:"Sam Stock",     username:"stock",   password:"stock123",   role:"stock"   },
  { id:3, name:"Paula Partner", username:"partner", password:"partner123", role:"partner" },
];

// ─────────────────────────────────────────────
// LOGIN SCREEN
// ─────────────────────────────────────────────
function LoginScreen({ users, onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError]       = useState("");
  const [showPw, setShowPw]     = useState(false);
  const [loading, setLoading]   = useState(false);

  function handleLogin() {
    if (!username || !password) { setError("Please enter both username and password."); return; }
    setLoading(true);
    setTimeout(() => {
      const user = users.find(u => u.username === username && u.password === password);
      if (user) {
        onLogin(user.username, user.password);
      } else {
        setError("Invalid username or password.");
        setLoading(false);
      }
    }, 600);
  }

  return (
    <div style={{ minHeight:"100vh", background:`linear-gradient(135deg, ${C.slate} 0%, #1E293B 100%)`,
      display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
      <div style={{ width:"100%", maxWidth:400 }}>
        {/* Logo */}
        <div style={{ textAlign:"center", marginBottom:32 }}>
          <div style={{ fontSize:40, marginBottom:10 }}>⚙️</div>
          <h1 style={{ margin:"0 0 4px", fontSize:24, fontWeight:800, color:C.white }}>CPE Refurb Manager</h1>
          <p style={{ margin:0, color:C.slate4, fontSize:13 }}>Sign in to continue</p>
        </div>

        {/* Card */}
        <div style={{ background:C.white, borderRadius:16, padding:28,
          boxShadow:"0 20px 60px rgba(0,0,0,.4)" }}>
          <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
            <div>
              <Label>USERNAME</Label>
              <input value={username} onChange={e=>{ setError(""); setUsername(e.target.value); }}
                onKeyDown={e=>e.key==="Enter"&&handleLogin()}
                placeholder="Enter your username" autoFocus
                style={iStyle({ fontSize:14 })}/>
            </div>
            <div>
              <Label>PASSWORD</Label>
              <div style={{ position:"relative" }}>
                <input value={password} onChange={e=>{ setError(""); setPassword(e.target.value); }}
                  onKeyDown={e=>e.key==="Enter"&&handleLogin()}
                  type={showPw?"text":"password"} placeholder="Enter your password"
                  style={iStyle({ paddingRight:40, fontSize:14 })}/>
                <button onClick={()=>setShowPw(p=>!p)}
                  style={{ position:"absolute", right:10, top:"50%", transform:"translateY(-50%)",
                    background:"none", border:"none", cursor:"pointer", color:C.slate4, fontSize:16, padding:0 }}>
                  {showPw?"🙈":"👁"}
                </button>
              </div>
            </div>

            {error && <Alert type="danger">⚠ {error}</Alert>}

            <button onClick={handleLogin} disabled={loading}
              style={{ marginTop:4, padding:"12px 0", background:loading?"#4F46E5":C.indigo, color:C.white,
                border:"none", borderRadius:10, fontSize:14, fontWeight:800, cursor:loading?"not-allowed":"pointer",
                boxShadow:"0 2px 10px rgba(99,102,241,.4)", transition:"all .2s", opacity:loading?.8:1 }}>
              {loading ? "Signing in…" : "Sign In →"}
            </button>
          </div>

          {/* Demo credentials hint */}
          <div style={{ marginTop:20, padding:"12px 14px", background:C.slate8,
            borderRadius:10, border:`1px solid ${C.slate6}` }}>
            <div style={{ fontSize:11, fontWeight:700, color:C.slate4, letterSpacing:".06em", marginBottom:8 }}>
              DEMO CREDENTIALS
            </div>
            {[
              ["admin","admin123","🛡️ Administrator"],
              ["stock","stock123","📦 Stock Management"],
              ["partner","partner123","🔧 Refurbishment Partner"],
            ].map(([u,p,label])=>(
              <div key={u} onClick={()=>{ setUsername(u); setPassword(p); setError(""); }}
                style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
                  padding:"5px 0", borderBottom:`1px solid ${C.slate7}`, cursor:"pointer" }}>
                <span style={{ fontSize:12, color:C.slate3 }}>{label}</span>
                <span style={{ fontFamily:"monospace", fontSize:11, color:C.indigo, fontWeight:600 }}>
                  {u} / {p}
                </span>
              </div>
            ))}
            <p style={{ margin:"8px 0 0", fontSize:11, color:C.slate4 }}>Click a row to auto-fill</p>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// USER MANAGEMENT  (Admin only)
// ─────────────────────────────────────────────
function UserManagement({ users, setUsers, currentUser, isMobile }) {
  const [showForm, setShowForm]   = useState(false);
  const [editUser, setEditUser]   = useState(null);
  const [deleteId, setDeleteId]   = useState(null);
  const [form, setForm]           = useState({ name:"", username:"", password:"", role:"stock" });
  const [formError, setFormError] = useState("");

  function openAdd()  { setForm({ name:"", username:"", password:"", role:"stock" }); setEditUser(null); setFormError(""); setShowForm(true); }
  function openEdit(u){ setForm({ name:u.name, username:u.username, password:u.password, role:u.role }); setEditUser(u); setFormError(""); setShowForm(true); }

  async function saveUser() {
    if (!form.name.trim() || !form.username.trim() || !form.password.trim()) {
      setFormError("All fields are required."); return;
    }
    const duplicate = users.find(u => u.username === form.username && (!editUser || u.id !== editUser.id));
    if (duplicate) { setFormError("Username already exists."); return; }

    if (editUser) {
      setUsers(p => p.map(u => u.id === editUser.id ? { ...u, ...form } : u));
    } else {
      setUsers(p => [...p, { id: Date.now(), ...form }]);
    }
    setShowForm(false);
  }

  function deleteUser(id) {
    if (id === currentUser.id) return;
    setUsers(p => p.filter(u => u.id !== id));
    deleteUser_db(id).catch(e => console.warn("DB delete user:", e));
    setDeleteId(null);
  }

  const roleConfig = (role) => ROLES[role] || { label:role, icon:"👤", color:C.slate3, bg:C.slate7 };

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:10 }}>
        <div>
          <h2 style={{ margin:"0 0 2px", fontSize:isMobile?18:22, fontWeight:800, color:C.slate }}>User Management</h2>
          <p style={{ margin:0, color:C.slate4, fontSize:13 }}>Manage user accounts and role assignments</p>
        </div>
        <div style={{ display:"flex", gap:8 }}>
          <Btn onClick={()=>{ if(window.confirm("Reset ALL data to defaults? This clears users, devices, and logs.")) { localStorage.clear(); window.location.reload(); } }} variant="ghost" size="sm">↺ Reset data</Btn>
          <Btn onClick={openAdd} style={{ background:"#16A34A", color:"#fff", border:"none", fontWeight:800, boxShadow:"0 2px 8px rgba(22,163,74,.3)" }}>+ Add User</Btn>
        </div>
      </div>

      {/* Role summary cards */}
      <div style={{ display:"grid", gridTemplateColumns:isMobile?"1fr 1fr":"repeat(3,1fr)", gap:10 }}>
        {Object.entries(ROLES).map(([key, r]) => (
          <Card key={key} accent={r.color} style={{ display:"flex", alignItems:"center", gap:12 }}>
            <span style={{ fontSize:24 }}>{r.icon}</span>
            <div>
              <div style={{ fontSize:22, fontWeight:800, color:C.slate }}>
                {users.filter(u => u.role === key).length}
              </div>
              <div style={{ fontSize:12, color:C.slate4 }}>{r.label}</div>
            </div>
          </Card>
        ))}
      </div>

      {/* User list */}
      <Card>
        <SectionTitle>All Users ({users.length})</SectionTitle>
        {isMobile
          ? users.map(u => {
              const rc = roleConfig(u.role);
              const isDeleting = deleteId === u.id;
              if (isDeleting) return (
                <div key={u.id} style={{ border:`2px solid ${C.red}`, borderRadius:12, padding:14, marginBottom:10, background:C.redLight }}>
                  <p style={{ margin:"0 0 10px", fontSize:13, fontWeight:600, color:C.redDark }}>
                    Delete user <strong>{u.name}</strong>? This cannot be undone.
                  </p>
                  <div style={{ display:"flex", gap:8 }}>
                    <Btn onClick={()=>deleteUser(u.id)} variant="danger" full>Delete</Btn>
                    <Btn onClick={()=>setDeleteId(null)} variant="ghost" full>Cancel</Btn>
                  </div>
                </div>
              );
              return (
                <div key={u.id} style={{ border:`1px solid ${C.slate6}`, borderRadius:12, padding:14, marginBottom:10,
                  borderLeft:`3px solid ${rc.color}` }}>
                  <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}>
                    <div>
                      <span style={{ fontWeight:700, fontSize:13, color:C.slate }}>{u.name}</span>
                      {u.id === currentUser.id && <span style={{ marginLeft:6, fontSize:10, fontWeight:700, background:C.indigoLight, color:C.indigo, padding:"1px 7px", borderRadius:20 }}>YOU</span>}
                    </div>
                    <span style={{ background:rc.bg, color:rc.color, fontSize:11, fontWeight:700, padding:"2px 9px", borderRadius:20 }}>
                      {rc.icon} {rc.label}
                    </span>
                  </div>
                  <div style={{ fontSize:12, color:C.slate4, fontFamily:"monospace", marginBottom:10 }}>@{u.username}</div>
                  <div style={{ display:"flex", gap:6 }}>
                    <Btn onClick={()=>openEdit(u)} variant="ghost" size="sm" full>✏️ Edit</Btn>
                    <Btn onClick={()=>setDeleteId(u.id)} variant="ghost" size="sm" full disabled={u.id===currentUser.id}>🗑️ Delete</Btn>
                  </div>
                </div>
              );
            })
          : (
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                <thead>
                  <tr style={{ borderBottom:`1.5px solid ${C.slate6}`, background:C.slate8 }}>
                    {["Name","Username","Role","Actions"].map(h=>(
                      <th key={h} style={{ padding:"8px 12px", textAlign:"left", fontSize:11, fontWeight:700, color:C.slate4, letterSpacing:".05em" }}>{h.toUpperCase()}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {users.map(u => {
                    const rc = roleConfig(u.role);
                    const isDeleting = deleteId === u.id;
                    return (
                      <tr key={u.id} style={{ borderBottom:`1px solid ${C.slate7}`,
                        background: isDeleting ? C.redLight : "transparent" }}>
                        <td style={{ padding:"10px 12px" }}>
                          <span style={{ fontWeight:600, color:C.slate }}>{u.name}</span>
                          {u.id === currentUser.id && <span style={{ marginLeft:6, fontSize:10, fontWeight:700, background:C.indigoLight, color:C.indigo, padding:"1px 7px", borderRadius:20 }}>YOU</span>}
                        </td>
                        <td style={{ padding:"10px 12px", fontFamily:"monospace", fontSize:12, color:C.slate3 }}>@{u.username}</td>
                        <td style={{ padding:"10px 12px" }}>
                          <span style={{ background:rc.bg, color:rc.color, fontSize:12, fontWeight:700, padding:"3px 10px", borderRadius:20 }}>
                            {rc.icon} {rc.label}
                          </span>
                        </td>
                        <td style={{ padding:"10px 12px" }}>
                          {isDeleting
                            ? <div style={{ display:"flex", gap:6 }}>
                                <span style={{ fontSize:12, color:C.redDark, fontWeight:600, marginRight:4 }}>Confirm delete?</span>
                                <Btn onClick={()=>deleteUser(u.id)} variant="danger" size="sm">Delete</Btn>
                                <Btn onClick={()=>setDeleteId(null)} variant="ghost" size="sm">Cancel</Btn>
                              </div>
                            : <div style={{ display:"flex", gap:6 }}>
                                <Btn onClick={()=>openEdit(u)} variant="ghost" size="sm">✏️ Edit</Btn>
                                <Btn onClick={()=>setDeleteId(u.id)} variant="ghost" size="sm" disabled={u.id===currentUser.id}>🗑️ Delete</Btn>
                              </div>
                          }
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )
        }
      </Card>

      {/* Add / Edit modal */}
      {showForm && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.5)", zIndex:100,
          display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
          <div style={{ background:C.white, borderRadius:16, padding:28, maxWidth:420, width:"100%",
            boxShadow:"0 20px 60px rgba(0,0,0,.25)" }}>
            <h3 style={{ margin:"0 0 18px", fontSize:17, fontWeight:800, color:C.slate }}>
              {editUser ? "Edit User" : "Add New User"}
            </h3>
            <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
              <div><Label>FULL NAME</Label>
                <input value={form.name} onChange={e=>{ setFormError(""); setForm(f=>({...f,name:e.target.value})); }}
                  placeholder="e.g. Alice Admin" style={iStyle()}/>
              </div>
              <div><Label>USERNAME</Label>
                <input value={form.username} onChange={e=>{ setFormError(""); setForm(f=>({...f,username:e.target.value})); }}
                  placeholder="e.g. alice" style={iStyle({ fontFamily:"monospace" })}/>
              </div>
              <div><Label>PASSWORD</Label>
                <input value={form.password} onChange={e=>setForm(f=>({...f,password:e.target.value}))}
                  type="text" placeholder="Set a password" style={iStyle()}/>
              </div>
              <div><Label>ROLE</Label>
                <select value={form.role} onChange={e=>setForm(f=>({...f,role:e.target.value}))} style={iStyle()}>
                  {Object.entries(ROLES).map(([key, r])=>(
                    <option key={key} value={key}>{r.icon} {r.label}</option>
                  ))}
                </select>
              </div>
              {/* Role description */}
              <div style={{ background:ROLES[form.role]?.bg||C.slate8, borderRadius:8, padding:"10px 14px",
                border:`1px solid ${ROLES[form.role]?.color||C.slate6}22` }}>
                <div style={{ fontSize:11, fontWeight:700, color:ROLES[form.role]?.color||C.slate4, marginBottom:4 }}>
                  {ROLES[form.role]?.icon} {ROLES[form.role]?.label} — Access
                </div>
                <div style={{ fontSize:12, color:C.slate3 }}>
                  {form.role==="admin"   && "Full access to all views, user management, and partner portal."}
                  {form.role==="stock"   && "Dashboard, Intake & Triage, Stock & Scrap, and All Devices."}
                  {form.role==="partner" && "Partner portal only — submit QC outcomes and view assigned devices."}
                </div>
              </div>
              {formError && <Alert type="danger">⚠ {formError}</Alert>}
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginTop:20 }}>
              <Btn onClick={()=>setShowForm(false)} variant="ghost" full size="lg">Cancel</Btn>
              <Btn onClick={saveUser} full size="lg" style={{ background:"#16A34A", color:"#fff", border:"none", fontWeight:800, boxShadow:"0 2px 8px rgba(22,163,74,.3)" }}>{editUser?"Save Changes":"Add User"}</Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// NAV VIEWS CONFIG (all possible views)
// ─────────────────────────────────────────────
const ALL_VIEWS = [
  { id:"dashboard",     label:"Dashboard",        icon:"📊" },
  { id:"intake",        label:"Intake & Triage",  icon:"📥" },
  { id:"refurb",        label:"Refurbishment",    icon:"🔧" },
  { id:"stock",         label:"Stock",            icon:"📦" },
  { id:"all",           label:"All Devices",      icon:"🗂️" },
  { id:"users",         label:"User Management",  icon:"👥" },
  { id:"partnerportal", label:"Partner Portal",   icon:"🔒" },
];

// ─────────────────────────────────────────────
// APP
// ─────────────────────────────────────────────
// ── DB row mappers (snake_case → camelCase) ──────────────────────────────────
function mapDevice(d) {
  return {
    id:             d.id,
    serial:         d.serial,
    mac:            d.mac            || "",
    model:          d.model          || "",
    type:           d.type,
    stage:          d.stage,
    outcome:        d.outcome        || null,
    received:       d.received_date  || d.received || today(),
    notes:          d.notes          || "",
    sentToPartner:  d.sent_to_partner || false,
    partnerOutcome: d.partner_outcome || null,
    partnerNotes:   d.partner_notes   || "",
    pendingAction:  d.pending_action  || null,
  };
}
function mapUser(u) {
  return { id:u.id, name:u.name, username:u.username, password:u.password, role:u.role };
}
function deviceToRow(d) {
  return {
    id:              d.id,
    serial:          d.serial,
    mac:             d.mac          || "",
    model:           d.model        || "",
    type:            d.type,
    stage:           d.stage,
    outcome:         d.outcome      || null,
    received_date:   d.received     || today(),
    notes:           d.notes        || "",
    sent_to_partner: d.sentToPartner || false,
    partner_outcome: d.partnerOutcome || null,
    partner_notes:   d.partnerNotes   || "",
    pending_action:  d.pendingAction  || null,
  };
}

export default function App() {
  const isMobile = useIsMobile();

  // Auth — keep session in localStorage only (not in DB)
  const [currentUser, setCurrentUser] = useLocalStorage("cpe_current_user", null);

  // DB-backed state
  const [users,       setUsers]       = useState(DEMO_USERS_INIT);
  const [devices,     setDevices]     = useState(SEED);
  const [uploadLogs,  setUploadLogs]  = useState([]);
  const [dbReady,     setDbReady]     = useState(false);
  const [dbError,     setDbError]     = useState(null);

  // UI state
  const [view, setView]               = useState("dashboard");
  const [sidebarOpen, setSidebarOpen] = useState(false);

  // ── Preview mode: load seed data instantly ──
  useEffect(() => {
    setDbReady(true);
  }, []);

  // Derive role config and mode from current user
  const roleConfig    = currentUser ? (ROLES[currentUser.role] || ROLES.admin) : null;
  const isPartnerMode = currentUser?.role === "partner";
  const allowedViews  = roleConfig ? ALL_VIEWS.filter(v => roleConfig.views.includes(v.id)) : [];

  const pipelineCount  = devices.filter(d=>!["Stock","Scrap"].includes(d.stage)).length;
  const partnerPending = devices.filter(d=>d.sentToPartner && d.stage==="Refurbishment" && !d.partnerOutcome).length;
  const pendingConf    = devices.filter(d=>d.partnerOutcome && d.stage==="Refurbishment").length;

  async function login(username, password) {
    let found = null;
    try {
      const { data, error } = await findUserByCredentials(username, password);
      if (!error && data) found = mapUser(data);
    } catch (_) {}
    if (!found) {
      found = users.find(u => u.username === username && u.password === password) || null;
    }
    if (!found) return;
    setCurrentUser(found);
    setView("dashboard");
    setSidebarOpen(false);
  }

  function logout() {
    setCurrentUser(null);
    setView("dashboard");
    setSidebarOpen(false);
  }

  // ── Not logged in ──
  // Loading state while Supabase initialises
  // dbReady is always true in preview mode

  if (!currentUser) {
    return <LoginScreen users={users} onLogin={login}/>;
  }

  // ── Partner role → go straight to partner portal ──
  if (isPartnerMode) {
    return (
      <div style={{ minHeight:"100vh", background:C.slate8, fontFamily:"system-ui,-apple-system,sans-serif", display:"flex", flexDirection:"column" }}>
        <header style={{ background:C.slate, padding:`0 ${isMobile?14:24}px`, height:52,
          display:"flex", alignItems:"center", justifyContent:"space-between", flexShrink:0, position:"sticky", top:0, zIndex:50 }}>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <span style={{ color:C.white, fontWeight:800, fontSize:isMobile?13:15 }}>⚙️ CPE Refurb</span>
            <span style={{ background:C.purple, color:C.white, fontSize:9, fontWeight:700, padding:"2px 7px", borderRadius:20 }}>PARTNER</span>
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <span style={{ color:"rgba(255,255,255,.6)", fontSize:12 }}>
              {roleConfig.icon} {currentUser.name}
            </span>
            <button onClick={logout}
              style={{ padding:"4px 12px", background:"rgba(255,255,255,.1)", color:"rgba(255,255,255,.8)",
                border:"1px solid rgba(255,255,255,.2)", borderRadius:8, fontSize:12, fontWeight:600, cursor:"pointer" }}>
              Sign Out
            </button>
          </div>
        </header>
        <main style={{ flex:1, padding:isMobile?"14px":"28px", overflowY:"auto" }}>
          <PartnerPortal devices={devices} setDevices={setDevices}
            uploadLogs={uploadLogs} setUploadLogs={setUploadLogs} isMobile={isMobile}/>
        </main>
      </div>
    );
  }

  const props = { devices, setDevices, isMobile };

  // ── Internal roles (admin / stock) ──
  return (
    <div style={{ minHeight:"100vh", background:C.slate8, fontFamily:"system-ui,-apple-system,sans-serif", display:"flex", flexDirection:"column" }}>

      {/* ── TOP BAR ── */}
      <header style={{ background:C.slate, padding:`0 ${isMobile?14:24}px`, height:52,
        display:"flex", alignItems:"center", justifyContent:"space-between", flexShrink:0, position:"sticky", top:0, zIndex:50 }}>
        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
          {isMobile && (
            <button onClick={()=>setSidebarOpen(o=>!o)}
              style={{ background:"none", border:"none", color:C.white, fontSize:20, cursor:"pointer", padding:"0 4px", lineHeight:1 }}>☰</button>
          )}
          <span style={{ color:C.white, fontWeight:800, fontSize:isMobile?13:15 }}>⚙️ CPE Refurb</span>
          <span style={{ background:C.indigo, color:C.white, fontSize:9, fontWeight:700, padding:"2px 7px", borderRadius:20 }}>PROTO</span>
        </div>

        {/* Role badge + user info */}
        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
          <span style={{ background:roleConfig.bg, color:roleConfig.color, fontSize:11, fontWeight:700,
            padding:"3px 10px", borderRadius:20, display:isMobile?"none":"inline-flex", alignItems:"center", gap:4 }}>
            {roleConfig.icon} {roleConfig.label}
          </span>
          {!isMobile && <span style={{ color:"rgba(255,255,255,.6)", fontSize:12 }}>{currentUser.name}</span>}
          <button onClick={logout}
            style={{ padding:"4px 12px", background:"rgba(255,255,255,.1)", color:"rgba(255,255,255,.8)",
              border:"1px solid rgba(255,255,255,.2)", borderRadius:8, fontSize:12, fontWeight:600, cursor:"pointer" }}>
            {isMobile?"↩":"Sign Out"}
          </button>
        </div>
      </header>

      <div style={{ display:"flex", flex:1, overflow:"hidden", position:"relative" }}>

        {/* ── SIDEBAR / DRAWER ── */}
        <>
          {isMobile && sidebarOpen && (
            <div onClick={()=>setSidebarOpen(false)}
              style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.4)", zIndex:40 }}/>
          )}
          <nav style={{
            width:210, background:C.white, borderRight:`1.5px solid ${C.slate6}`,
            padding:"16px 0", flexShrink:0, display:"flex", flexDirection:"column", gap:2,
            ...(isMobile ? {
              position:"fixed", top:52, left:sidebarOpen?0:-230, bottom:0,
              zIndex:45, width:230, transition:"left .25s", boxShadow:sidebarOpen?"4px 0 20px rgba(0,0,0,.15)":"none"
            } : {})
          }}>
            {/* User info in sidebar */}
            <div style={{ padding:"0 16px 14px", borderBottom:`1px solid ${C.slate7}`, marginBottom:6 }}>
              <div style={{ fontSize:12, fontWeight:700, color:C.slate }}>{currentUser.name}</div>
              <span style={{ background:roleConfig.bg, color:roleConfig.color, fontSize:10, fontWeight:700,
                padding:"2px 8px", borderRadius:20, display:"inline-block", marginTop:3 }}>
                {roleConfig.icon} {roleConfig.label}
              </span>
            </div>

            {allowedViews.map(v=>(
              <button key={v.id} onClick={()=>{ setView(v.id); setSidebarOpen(false); }}
                style={{ width:"100%", padding:"11px 18px", border:"none", textAlign:"left",
                  background:view===v.id?C.indigoLight:"transparent",
                  color:view===v.id?C.indigoDark:C.slate3,
                  fontWeight:view===v.id?700:500, fontSize:13, cursor:"pointer",
                  borderLeft:view===v.id?`3px solid ${C.indigo}`:"3px solid transparent",
                  display:"flex", alignItems:"center", gap:8, transition:"all .15s" }}>
                <span>{v.icon}</span>
                {v.label}
                {v.id==="refurb" && pendingConf>0 && (
                  <span style={{ marginLeft:"auto", background:C.amber, color:C.white, fontSize:9, fontWeight:800,
                    minWidth:16, height:16, borderRadius:20, display:"flex", alignItems:"center", justifyContent:"center", padding:"0 3px" }}>{pendingConf}</span>
                )}
              </button>
            ))}

            {/* Pipeline summary — admin only */}
            {!isMobile && currentUser.role === "admin" && (
              <div style={{ marginTop:"auto", padding:"14px 18px", borderTop:`1px solid ${C.slate7}` }}>
                <div style={{ fontSize:10, color:C.slate4, fontWeight:700, letterSpacing:".07em", marginBottom:8 }}>PIPELINE</div>
                {[
                  ["Triage",       d=>["Intake","Triage"].includes(d.stage),              "#F97316"],
                  ["Refurb/QC",    d=>["Refurbishment","QC Check"].includes(d.stage),     C.amber],
                  ["Escalated",    d=>d.stage==="Escalated",                              C.purple],
                  ["At Partner",   d=>d.sentToPartner&&d.stage==="Refurbishment",         C.purple],
                  ["Pend.Confirm", d=>d.partnerOutcome&&d.stage==="Refurbishment",        C.red],
                  ["ECUS",         d=>d.stage==="ECUS",                                    "#F59E0B"],
                  ["Confirmed",    d=>d.stage==="Confirmed",                               "#22C55E"],
                  ["In Transit",   d=>d.stage==="In Transit",                              "#3B82F6"],
                ].map(([label,fn,color])=>(
                  <div key={label} style={{ display:"flex", justifyContent:"space-between", fontSize:12, color:C.slate3, marginBottom:4 }}>
                    <span style={{ display:"flex", alignItems:"center", gap:5 }}>
                      <span style={{ width:6, height:6, borderRadius:"50%", background:color, display:"inline-block" }}/>
                      {label}
                    </span>
                    <span style={{ fontWeight:700 }}>{devices.filter(fn).length}</span>
                  </div>
                ))}
              </div>
            )}
          </nav>
        </>

        {/* ── MAIN CONTENT ── */}
        <main style={{ flex:1, padding:isMobile?"14px":"28px", overflowY:"auto", paddingBottom:isMobile?"80px":"28px" }}>
          {view==="dashboard" && <Dashboard {...props}/>}
          {view==="intake"    && allowedViews.find(v=>v.id==="intake")   && <IntakeTriage {...props}/>}
          {view==="refurb"    && allowedViews.find(v=>v.id==="refurb")   && <RefurbTracking {...props}/>}
          {view==="stock"     && allowedViews.find(v=>v.id==="stock")    && <StockView {...props} canConfirmTransit={["admin","stock"].includes(currentUser?.role)}/>}
          {view==="all"       && allowedViews.find(v=>v.id==="all")      && <AllDevices {...props}/>}
          {view==="users"     && allowedViews.find(v=>v.id==="users")    && (
            <UserManagement users={users} setUsers={setUsers} currentUser={currentUser} isMobile={isMobile}/>
          )}
          {view==="partnerportal" && allowedViews.find(v=>v.id==="partnerportal") && (
            <PartnerPortal devices={devices} setDevices={setDevices}
              uploadLogs={uploadLogs} setUploadLogs={setUploadLogs} isMobile={isMobile}/>
          )}
        </main>
      </div>

      {/* ── MOBILE BOTTOM NAV ── */}
      {isMobile && (
        <nav style={{ position:"fixed", bottom:0, left:0, right:0, background:C.white,
          borderTop:`1px solid ${C.slate6}`, display:"flex", zIndex:40, paddingBottom:"env(safe-area-inset-bottom,0)" }}>
          {allowedViews.slice(0, 5).map(v=>(
            <button key={v.id} onClick={()=>{ setView(v.id); setSidebarOpen(false); }}
              style={{ flex:1, padding:"10px 4px 8px", border:"none", background:"transparent",
                display:"flex", flexDirection:"column", alignItems:"center", gap:2,
                color:view===v.id?C.indigo:C.slate4, cursor:"pointer" }}>
              <span style={{ fontSize:16, position:"relative" }}>
                {v.icon}
                {v.id==="refurb" && pendingConf>0 && (
                  <span style={{ position:"absolute", top:-3, right:-6, background:C.amber, color:C.white, fontSize:8, fontWeight:800,
                    minWidth:14, height:14, borderRadius:20, display:"flex", alignItems:"center", justifyContent:"center", padding:"0 2px" }}>{pendingConf}</span>
                )}
              </span>
              <span style={{ fontSize:9, fontWeight:view===v.id?700:500 }}>{v.label.split(" ")[0]}</span>
            </button>
          ))}
        </nav>
      )}
    </div>
  );
}
