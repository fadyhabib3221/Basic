import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// Perla Di Mare logo, embedded as a data URI so the app works as a single self-contained file.
const LOGO_DATA_URL = "data:image/png;base64,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";
import {
  Plane, Search, Trash2, Pencil, X, Check, TrendingUp, Ticket, Wallet,
  Calendar, Download, Upload, Building2, Factory, Lock, LogOut, UserPlus, Users, Eye, EyeOff,
  ShieldCheck, Wifi, User, Cloud, Globe2, List, Car, FileText, ArrowLeft,
  MapPin, Compass, Luggage, Anchor, Sparkles, Plus,
} from "lucide-react";

const MONTHS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

const monthKey = (dateStr) => (dateStr ? dateStr.slice(0, 7) : "No date");
// Storage stays in the native YYYY-MM-DD format (required by <input type="date">),
// but everywhere we display the date to the user we show it as DD-MMM-YYYY, with the
// month written as its first three letters, capitalized (e.g. "03-AUG-2026").
const formatDisplayDate = (dateStr) => {
  if (!dateStr) return "";
  const [y, m, d] = dateStr.split("-");
  if (!y || !m || !d) return dateStr;
  const monthAbbr = (MONTHS[parseInt(m, 10) - 1] || m).slice(0, 3).toUpperCase();
  return `${d}-${monthAbbr}-${y}`;
};
const monthLabel = (key) => {
  if (key === "No date") return key;
  const [y, m] = key.split("-");
  const idx = parseInt(m, 10) - 1;
  return `${MONTHS[idx] || m} ${y}`;
};

// Formats an ISO timestamp as DD-MMM-YYYY HH:MM for showing when a note edit happened.
const formatDateTime = (iso) => {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return "";
  const dd = String(d.getDate()).padStart(2, "0");
  const monthAbbr = MONTHS[d.getMonth()].slice(0, 3).toUpperCase();
  const yyyy = d.getFullYear();
  const hh = String(d.getHours()).padStart(2, "0");
  const min = String(d.getMinutes()).padStart(2, "0");
  return `${dd}-${monthAbbr}-${yyyy} ${hh}:${min}`;
};

const emptyCustomerRow = () => ({ name: "", ticketNumber: "" });

// Ticket supplier / booking source options.
const SUPPLIERS = ["Amadeus", "Sabre", "NDC", "Lowcost"];

// Saved companies were originally plain strings; this reads the name whether an entry
// is still a legacy string or the newer { name, taxNumber, commercialReg, phones } record.
const companyName = (c) => (typeof c === "string" ? c : (c && c.name) || "");

const emptyCompanyDraft = { name: "", taxNumber: "", commercialReg: "", phones: "" };

// Local YYYY-MM-DD for today, matching the native <input type="date"> format.
const todayDateStr = () => {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
};

// A function (not a static object) so every new/reset ticket picks up TODAY'S date
// at the moment it's created, rather than whatever date happened to be "today" when
// the app first loaded. The user can still change it manually afterward.
const getEmptyForm = () => ({
  id: null,
  employee: "",
  company: "",
  supplier: "",
  customersCount: 1,
  customers: [emptyCustomerRow()],
  from: "",
  to: "",
  // Multi-destination (multi-city) route support: when multiDestination is on, the
  // route is described as an ordered list of stops (e.g. ["CAI","DXB","BKK"]) instead
  // of a single from/to pair. "from"/"to" are still kept in sync (first/last stop) so
  // every place that reads a plain origin/destination keeps working unchanged.
  multiDestination: false,
  destinations: ["", ""],
  airline: "",
  date: todayDateStr(),
  netPrice: "",
  soldPrice: "",
  notes: "",
  // Reissue tracking: when isReissued is on, oldTicketNumber is looked up against
  // existing tickets to auto-fill oldTicketIssueDate and every other field below
  // (company, supplier, route, airline, prices, customer names) from that old ticket.
  isReissued: false,
  oldTicketNumber: "",
  oldTicketIssueDate: "",
});

// Renders a ticket's route as a single "A → B" (or "A → B → C → ..." for a
// multi-destination/multi-city booking) string for lists, detail views, and exports.
const routeLabel = (t) => {
  const stops = Array.isArray(t.destinations) ? t.destinations.map((d) => (d || "").trim()).filter(Boolean) : [];
  if (t.multiDestination && stops.length >= 2) return stops.join(" → ");
  return `${t.from || "-"} → ${t.to || "-"}`;
};

// Room types offered on a hotel booking's room line.
const ROOM_TYPES = [
  { value: "single", label: "Single" },
  { value: "double", label: "Double" },
  { value: "triple", label: "Triple" },
];

// Meal plan offered on a hotel booking's room line.
const MEAL_PLANS = [
  { value: "ro", label: "Room Only" },
  { value: "bb", label: "Bed & Breakfast" },
  { value: "hb", label: "Half Board" },
  { value: "fb", label: "Full Board" },
  { value: "ai", label: "All Inclusive" },
];

// Max number of adult guests a room type can hold — drives how many guest-name fields
// are shown for a room line (Single -> 1, Double -> 2, Triple -> 3).
const ROOM_CAPACITY = { single: 1, double: 2, triple: 3 };

// A single adult guest staying in a room.
const emptyGuest = () => ({
  id: `G-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
  name: "",
});

// A child staying in a room, with an age (in whole years, 0–11) alongside the name.
const emptyChild = () => ({
  id: `C-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
  name: "",
  age: "",
});

// Converts Arabic-Indic (٠-٩) and Extended Arabic-Indic (۰-۹) digits to standard 0-9,
// then strips anything that isn't a digit. Using type="text" with this instead of
// type="number" avoids the age field silently rejecting keystrokes on Arabic keyboards,
// which type="number" does with non-Latin digits.
const sanitizeAgeInput = (raw) => {
  let v = raw
    .replace(/[\u0660-\u0669]/g, (d) => String(d.charCodeAt(0) - 0x0660))
    .replace(/[\u06F0-\u06F9]/g, (d) => String(d.charCodeAt(0) - 0x06F0));
  v = v.replace(/[^0-9]/g, "");
  if (v !== "" && parseInt(v, 10) > 11) v = "11";
  return v;
};

// Resizes a room line's guest list to match its room type's capacity, keeping any
// names already entered and padding/truncating as needed.
const guestsForCapacity = (guests, capacity) => {
  const list = (Array.isArray(guests) ? guests : []).slice(0, capacity).map((g) => ({ ...g }));
  while (list.length < capacity) list.push(emptyGuest());
  return list;
};

const HOTEL_CURRENCIES = [
  { value: "EGP", label: "EGP" },
  { value: "USD", label: "USD" },
];

// A single room line within a hotel booking: a room type + meal plan combination, its own
// currency, count, and net/sold price per room per night — e.g. "1x Single, Half Board,
// EGP" and "2x Double, All Inclusive, USD" can both live inside the same booking.
const emptyRoomLine = () => ({
  id: `RL-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
  roomType: "single",
  mealPlan: "bb",
  currency: "EGP",
  count: 1,
  netPrice: "",
  soldPrice: "",
  // Each room now carries its own stay dates, since different rooms on the same
  // booking can check in/out on different days.
  checkIn: todayDateStr(),
  checkOut: todayDateStr(),
  // Adult guest names — sized to the default room type's capacity (single -> 1).
  guests: guestsForCapacity([], ROOM_CAPACITY.single),
  // Children staying in this room, each with a name and age (0–11 years).
  children: [],
});

// A function (not a static object) so every new/reset hotel booking picks up TODAY'S
// date at the moment it's created, same rationale as getEmptyForm() above.
const getEmptyHotelForm = () => ({
  id: null,
  employee: "",
  customer: "",
  hotel: "",
  supplier: "",
  roomLines: [emptyRoomLine()],
  // The date the reservation itself was made — separate from each room's own
  // check-in/check-out dates below.
  bookingDate: todayDateStr(),
  notes: "",
});

// Given a ticket number like "077-1234567890", returns the same prefix with the numeric
// part increased by one, keeping the same digit width (e.g. "077-1234567891").
// Returns "" if the ticket number doesn't match the expected PREFIX-DIGITS shape.
// Auto-sequencing only ever advances the LAST THREE digits of the serial number (wrapping
// 999 back to 000); everything before them — including the rest of the serial — stays fixed,
// since that part identifies the batch/booking rather than the individual ticket.
const nextTicketNumber = (ticketNumber) => {
  if (!ticketNumber) return "";
  const match = ticketNumber.match(/^([A-Z0-9]{3})-(\d+)$/);
  if (!match) return "";
  const [, prefix, digits] = match;
  if (digits.length <= 3) {
    const wrapped = ((parseInt(digits, 10) + 1) % (10 ** digits.length)).toString().padStart(digits.length, "0");
    return `${prefix}-${wrapped}`;
  }
  const head = digits.slice(0, -3);
  const tail = digits.slice(-3);
  const nextTail = ((parseInt(tail, 10) + 1) % 1000).toString().padStart(3, "0");
  return `${prefix}-${head}${nextTail}`;
};

// Fills/trims the customers array to match the requested count, keeping existing entries
const resizeCustomers = (customers, count) => {
  const n = Math.max(1, Math.min(50, parseInt(count, 10) || 1));
  const next = [...customers];
  while (next.length < n) next.push(emptyCustomerRow());
  next.length = n;
  return next;
};

// Job grades shown to the main account when creating/editing an employee. Picking a
// grade fills in a sensible starting set of permission toggles below (see
// ROLE_PRESETS), but every toggle can still be switched on or off by hand afterwards —
// the grade is a starting point/label, not a lock. Grade is purely descriptive; access
// is always driven by the individual toggles stored on the employee record.
const EMPLOYEE_ROLES = [
  { value: "manager", label: "Manager" },
  { value: "supervisor", label: "Supervisor" },
  { value: "employee", label: "Employee" },
  { value: "accountant", label: "Accountant" },
];

// Starting toggle values applied when a grade is picked. All six are then freely
// editable by hand, independent of which grade is selected.
const ROLE_PRESETS = {
  manager: { canViewAll: true, canAdd: true, canEdit: true, canDelete: true, isAccounting: false, canManageCompanies: true },
  supervisor: { canViewAll: true, canAdd: true, canEdit: true, canDelete: false, isAccounting: false, canManageCompanies: false },
  employee: { canViewAll: false, canAdd: true, canEdit: false, canDelete: false, isAccounting: false, canManageCompanies: false },
  accountant: { canViewAll: true, canAdd: false, canEdit: false, canDelete: false, isAccounting: true, canManageCompanies: false },
};

const roleLabel = (value) => (EMPLOYEE_ROLES.find((r) => r.value === value) || {}).label || "Employee";

// Applies the coherence rules that keep the six permission toggles consistent with
// each other, no matter which one was just changed by hand:
// - Editing, deleting, or accounting access all require view access first.
// - Accounting mode is a fixed bundle (view-only + notes) that overrides add/edit/delete.
const reconcilePermissions = (perm) => {
  if (perm.isAccounting) {
    return { ...perm, canViewAll: true, canAdd: false, canEdit: false, canDelete: false };
  }
  return { ...perm, canViewAll: perm.canViewAll || perm.canEdit || perm.canDelete };
};

const emptyNewEmployee = {
  name: "",
  username: "",
  password: "",
  role: "employee",
  // Default permissions for a newly created employee: can only see and add
  // their own tickets, cannot edit or delete anything, and is not an accounting account.
  canViewAll: false,
  canAdd: true,
  canEdit: false,
  canDelete: false,
  isAccounting: false,
  canManageCompanies: false,
};

// Per-employee permissions button used in the employee table: shows a quick summary
// of the current access and opens the full-screen EmployeePermissionsModal below on
// click. Kept as a plain button (not a popover) because the table it sits in has
// overflow-hidden on its wrapper, which would otherwise clip a dropdown open below it.
const PermissionsCell = ({ emp, onOpen }) => {
  const summary = [
    emp.canViewAll && "View all",
    emp.canEdit && "Edit",
    emp.canDelete && "Delete",
    emp.isAccounting && "Notes only",
    emp.canManageCompanies && "Companies",
  ]
    .filter(Boolean)
    .join(" · ") || "Own tickets only";

  return (
    <button
      type="button"
      onClick={onOpen}
      className="border border-stone-300 rounded-lg px-2 py-1.5 text-xs text-stone-700 hover:bg-teal-50 hover:border-teal-300 flex items-center gap-1.5 max-w-[220px]"
    >
      <span className="truncate">{summary}</span>
      <Pencil size={11} className="text-stone-400 shrink-0" />
    </button>
  );
};

// Full-screen modal for editing one employee's grade and detailed permissions. Centered
// over the whole page (not nested inside the scrollable/clipped table), so it's always
// fully visible and easy to use — this is the one place permissions for an existing
// employee are changed. Closes itself if the employee record disappears (e.g. deleted
// from another tab) or is promoted to a main account (which no longer uses these toggles).
const EmployeePermissionsModal = ({ emp, onClose, onSetRole, onSetPermission }) => {
  if (!emp) return null;
  return (
    <div className="fixed inset-0 bg-stone-900/40 flex items-center justify-center p-4 z-50" onClick={onClose}>
      <div
        className="bg-white rounded-2xl border border-stone-200 p-5 w-full max-w-sm max-h-[90vh] overflow-y-auto"
        onClick={(ev) => ev.stopPropagation()}
      >
        <div className="flex items-center justify-between mb-1">
          <h3 className="font-semibold text-stone-900">Permissions</h3>
          <button onClick={onClose} className="text-stone-400 hover:text-stone-700 p-1">
            <X size={16} />
          </button>
        </div>
        <p className="text-xs text-stone-400 mb-4">{emp.name} · {emp.username}</p>

        <label className="text-xs text-stone-500 block mb-1.5">Grade</label>
        <div className="grid grid-cols-4 gap-1.5 mb-4">
          {EMPLOYEE_ROLES.map((r) => (
            <button
              key={r.value}
              type="button"
              onClick={() => onSetRole(r.value)}
              className={`text-xs font-semibold rounded-xl px-2 py-2 border transition-colors ${
                (emp.role || "employee") === r.value
                  ? "bg-teal-800 text-white border-teal-800"
                  : "bg-white text-stone-600 border-stone-300 hover:bg-stone-50"
              }`}
            >
              {r.label}
            </button>
          ))}
        </div>

        <p className="text-xs text-stone-500 mb-1">Individual permissions</p>
        <div className="border border-stone-200 rounded-xl px-3 divide-y divide-stone-100">
          <ToggleSwitch
            label="View all tickets"
            description="See every employee's tickets, not just their own"
            checked={emp.canViewAll || emp.canEdit || emp.canDelete}
            disabled={emp.isAccounting || emp.canEdit || emp.canDelete}
            onChange={(v) => onSetPermission("canViewAll", v)}
          />
          <ToggleSwitch
            label="Edit tickets"
            description="Edit any ticket they can see"
            checked={emp.canEdit}
            disabled={emp.isAccounting}
            onChange={(v) => onSetPermission("canEdit", v)}
          />
          <ToggleSwitch
            label="Delete tickets"
            description="Permanently remove any ticket they can see"
            checked={emp.canDelete}
            disabled={emp.isAccounting}
            onChange={(v) => onSetPermission("canDelete", v)}
          />
          <ToggleSwitch
            label="Accounting mode"
            description="View all tickets, but the only edit allowed is the Notes field"
            checked={emp.isAccounting}
            onChange={(v) => onSetPermission("isAccounting", v)}
          />
          <ToggleSwitch
            label="Manage companies"
            description="Add, edit, or remove saved company records"
            checked={emp.canManageCompanies}
            onChange={(v) => onSetPermission("canManageCompanies", v)}
          />
        </div>

        <button
          onClick={onClose}
          className="mt-4 w-full bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 shadow-sm shadow-teal-800/30 transition-colors"
        >
          Done
        </button>
      </div>
    </div>
  );
};

// A small reusable on/off switch used throughout the permissions UI.
const ToggleSwitch = ({ checked, onChange, disabled, label, description }) => (
  <label className={`flex items-start justify-between gap-3 py-1.5 ${disabled ? "opacity-50" : ""}`}>
    <span>
      <span className="text-sm text-stone-700 font-medium block">{label}</span>
      {description && <span className="text-[11px] text-stone-400 block">{description}</span>}
    </span>
    <button
      type="button"
      disabled={disabled}
      onClick={() => !disabled && onChange(!checked)}
      className={`shrink-0 mt-0.5 relative inline-flex h-5 w-9 items-center rounded-full transition-colors ${
        checked ? "bg-teal-700" : "bg-stone-300"
      } ${disabled ? "cursor-not-allowed" : "cursor-pointer"}`}
    >
      <span
        className="inline-block h-3.5 w-3.5 transform rounded-full bg-white transition-transform"
        style={{ transform: checked ? "translateX(18px)" : "translateX(2px)" }}
      />
    </button>
  </label>
);

// IATA 3-digit airline accounting/ticketing prefix codes — the first 3 digits of a
// standard e-ticket number identify the issuing airline. Used to link the ticket
// number prefix with the Airline field automatically in both directions.
const AIRLINE_CODES = [
  { code: "001", iata: "AA", name: "American Airlines" }, { code: "006", iata: "DL", name: "Delta Air Lines" },
  { code: "014", iata: "AC", name: "Air Canada" }, { code: "016", iata: "UA", name: "United Airlines" },
  { code: "020", iata: "SU", name: "Aeroflot" }, { code: "022", iata: "DE", name: "Condor" },
  { code: "027", iata: "AS", name: "Alaska Airlines" }, { code: "030", iata: "VY", name: "Vueling" },
  { code: "044", iata: "AR", name: "Aerolineas Argentinas" }, { code: "045", iata: "LA", name: "LATAM Airlines" },
  { code: "050", iata: "OA", name: "Olympic Air" }, { code: "053", iata: "EI", name: "Aer Lingus" },
  { code: "055", iata: "AZ", name: "ITA Airways" }, { code: "057", iata: "AF", name: "Air France" },
  { code: "065", iata: "SV", name: "Saudia" }, { code: "071", iata: "ET", name: "Ethiopian Airlines" },
  { code: "072", iata: "GF", name: "Gulf Air" }, { code: "074", iata: "KL", name: "KLM Royal Dutch Airlines" },
  { code: "075", iata: "IB", name: "Iberia" }, { code: "076", iata: "ME", name: "Middle East Airlines" },
  { code: "077", iata: "MS", name: "EgyptAir" }, { code: "079", iata: "PR", name: "Philippine Airlines" },
  { code: "080", iata: "LO", name: "LOT Polish Airlines" }, { code: "081", iata: "QF", name: "Qantas" },
  { code: "082", iata: "SN", name: "Brussels Airlines" }, { code: "085", iata: "4Y", name: "Discover Airlines" },
  { code: "086", iata: "NZ", name: "Air New Zealand" }, { code: "087", iata: "DT", name: "TAAG Angola Airlines" },
  { code: "098", iata: "AI", name: "Air India" }, { code: "101", iata: "EN", name: "Air Dolomiti" },
  { code: "104", iata: "EW", name: "Eurowings" }, { code: "105", iata: "AY", name: "Finnair" },
  { code: "108", iata: "FI", name: "Icelandair" }, { code: "114", iata: "LY", name: "El Al" },
  { code: "115", iata: "JU", name: "Air Serbia" }, { code: "117", iata: "SK", name: "Scandinavian Airlines" },
  { code: "124", iata: "AH", name: "Air Algerie" }, { code: "125", iata: "BA", name: "British Airways" },
  { code: "126", iata: "GA", name: "Garuda Indonesia" }, { code: "127", iata: "G3", name: "Gol Transportes Aereos" },
  { code: "131", iata: "JL", name: "Japan Airlines" }, { code: "134", iata: "AV", name: "Avianca" },
  { code: "139", iata: "AM", name: "Aeromexico" }, { code: "147", iata: "AT", name: "Royal Air Maroc" },
  { code: "157", iata: "QR", name: "Qatar Airways" }, { code: "160", iata: "CX", name: "Cathay Pacific" },
  { code: "176", iata: "EK", name: "Emirates" }, { code: "180", iata: "KE", name: "Korean Air" },
  { code: "205", iata: "NH", name: "All Nippon Airways" }, { code: "217", iata: "TG", name: "Thai Airways International" },
  { code: "220", iata: "LH", name: "Lufthansa" }, { code: "230", iata: "CM", name: "Copa Airlines" },
  { code: "232", iata: "MH", name: "Malaysia Airlines" }, { code: "235", iata: "TK", name: "Turkish Airlines" },
  { code: "257", iata: "OS", name: "Austrian Airlines" }, { code: "279", iata: "B6", name: "JetBlue Airways" },
  { code: "281", iata: "RO", name: "TAROM" }, { code: "282", iata: "TP", name: "TAP Air Portugal" },
  { code: "297", iata: "CI", name: "China Airlines" }, { code: "312", iata: "6E", name: "IndiGo" },
  { code: "324", iata: "SC", name: "Shandong Airlines" }, { code: "328", iata: "DY", name: "Norwegian Air Shuttle" },
  { code: "390", iata: "A3", name: "Aegean Airlines" }, { code: "427", iata: "TX", name: "Air Caraibes" },
  { code: "465", iata: "KC", name: "Air Astana" }, { code: "479", iata: "ZH", name: "Shenzhen Airlines" },
  { code: "512", iata: "RJ", name: "Royal Jordanian" }, { code: "514", iata: "G9", name: "Air Arabia" },
  { code: "605", iata: "H2", name: "Sky Airline" }, { code: "607", iata: "EY", name: "Etihad Airways" },
  { code: "618", iata: "SQ", name: "Singapore Airlines" }, { code: "623", iata: "FB", name: "Bulgaria Air" },
  { code: "643", iata: "KM", name: "Air Malta" }, { code: "649", iata: "TS", name: "Air Transat" },
  { code: "657", iata: "BT", name: "Air Baltic" }, { code: "668", iata: "TR", name: "Scoot" },
  { code: "695", iata: "BR", name: "EVA Air" }, { code: "706", iata: "KQ", name: "Kenya Airways" },
  { code: "724", iata: "LX", name: "Swiss International Air Lines" }, { code: "731", iata: "MF", name: "Xiamen Airlines" },
  { code: "738", iata: "VN", name: "Vietnam Airlines" }, { code: "755", iata: "UX", name: "Air Europa" },
  { code: "774", iata: "FM", name: "Shanghai Airlines" }, { code: "781", iata: "MU", name: "China Eastern Airlines" },
  { code: "784", iata: "CZ", name: "China Southern Airlines" }, { code: "795", iata: "VA", name: "Virgin Australia" },
  { code: "821", iata: "NO", name: "Neos" }, { code: "831", iata: "OU", name: "Croatia Airlines" },
  { code: "838", iata: "WS", name: "WestJet" }, { code: "847", iata: "RX", name: "Riyadh Air" },
  { code: "876", iata: "3U", name: "Sichuan Airlines" }, { code: "880", iata: "HU", name: "Hainan Airlines" },
  { code: "900", iata: "F3", name: "flyadeal" }, { code: "932", iata: "VS", name: "Virgin Atlantic" },
  { code: "978", iata: "VJ", name: "VietJet Air" }, { code: "999", iata: "CA", name: "Air China" },
];
const getAirlineCode = (name) => {
  const n = (name || "").trim().toUpperCase();
  if (!n) return null;
  const match = AIRLINE_CODES.find((a) => a.name.toUpperCase() === n);
  return match ? match.code : null;
};
const getAirlineByCode = (code) => {
  const match = AIRLINE_CODES.find((a) => a.code === code);
  return match ? match.iata : null;
};
// 2-letter IATA airline designator (e.g. "MS" for EgyptAir) — this is what gets
// typed/selected into the Airline field and stored on the ticket.
const getAirlineIata = (name) => {
  const n = (name || "").trim().toUpperCase();
  if (!n) return null;
  const match = AIRLINE_CODES.find((a) => a.name.toUpperCase() === n);
  return match ? match.iata : null;
};
// Reverse lookups from the 2-letter code: the 3-digit accounting/ticketing prefix
// (used to auto-fill the ticket number) and the full airline name (shown as a hint).
const getAirlineCodeByIata = (iata) => {
  const n = (iata || "").trim().toUpperCase();
  if (!n) return null;
  const match = AIRLINE_CODES.find((a) => a.iata === n);
  return match ? match.code : null;
};
const getAirlineNameByIata = (iata) => {
  const n = (iata || "").trim().toUpperCase();
  if (!n) return null;
  const match = AIRLINE_CODES.find((a) => a.iata === n);
  return match ? match.name : null;
};

// A reference list of major world airports (IATA code + city/country), offered as
// autocomplete suggestions on the From/To fields alongside previously typed values.
const AIRPORTS = [
  ["CAI", "Cairo, Egypt"], ["HRG", "Hurghada, Egypt"], ["SSH", "Sharm El Sheikh, Egypt"],
  ["LXR", "Luxor, Egypt"], ["ASW", "Aswan, Egypt"], ["HBE", "Alexandria, Egypt"],
  ["DXB", "Dubai, UAE"], ["AUH", "Abu Dhabi, UAE"], ["SHJ", "Sharjah, UAE"],
  ["DOH", "Doha, Qatar"], ["KWI", "Kuwait City, Kuwait"], ["RUH", "Riyadh, Saudi Arabia"],
  ["JED", "Jeddah, Saudi Arabia"], ["DMM", "Dammam, Saudi Arabia"], ["MED", "Medina, Saudi Arabia"],
  ["BAH", "Manama, Bahrain"], ["MCT", "Muscat, Oman"], ["AMM", "Amman, Jordan"],
  ["BEY", "Beirut, Lebanon"], ["DAM", "Damascus, Syria"], ["BGW", "Baghdad, Iraq"],
  ["BSR", "Basra, Iraq"], ["EBL", "Erbil, Iraq"], ["TLV", "Tel Aviv, Israel"],
  ["CMN", "Casablanca, Morocco"], ["RAK", "Marrakesh, Morocco"], ["ALG", "Algiers, Algeria"],
  ["TUN", "Tunis, Tunisia"], ["TIP", "Tripoli, Libya"], ["KRT", "Khartoum, Sudan"],
  ["ADD", "Addis Ababa, Ethiopia"], ["NBO", "Nairobi, Kenya"], ["JNB", "Johannesburg, South Africa"],
  ["CPT", "Cape Town, South Africa"], ["LOS", "Lagos, Nigeria"], ["ACC", "Accra, Ghana"],
  ["DKR", "Dakar, Senegal"], ["ABJ", "Abidjan, Ivory Coast"],
  ["LHR", "London Heathrow, UK"], ["LGW", "London Gatwick, UK"], ["MAN", "Manchester, UK"],
  ["CDG", "Paris Charles de Gaulle, France"], ["ORY", "Paris Orly, France"],
  ["AMS", "Amsterdam, Netherlands"], ["FRA", "Frankfurt, Germany"], ["MUC", "Munich, Germany"],
  ["BER", "Berlin, Germany"], ["MAD", "Madrid, Spain"], ["BCN", "Barcelona, Spain"],
  ["FCO", "Rome, Italy"], ["MXP", "Milan, Italy"], ["IST", "Istanbul, Turkey"],
  ["SAW", "Istanbul Sabiha, Turkey"], ["ATH", "Athens, Greece"], ["ZRH", "Zurich, Switzerland"],
  ["GVA", "Geneva, Switzerland"], ["VIE", "Vienna, Austria"], ["BRU", "Brussels, Belgium"],
  ["CPH", "Copenhagen, Denmark"], ["OSL", "Oslo, Norway"], ["ARN", "Stockholm, Sweden"],
  ["HEL", "Helsinki, Finland"], ["DUB", "Dublin, Ireland"], ["LIS", "Lisbon, Portugal"],
  ["WAW", "Warsaw, Poland"], ["PRG", "Prague, Czech Republic"], ["BUD", "Budapest, Hungary"],
  ["OTP", "Bucharest, Romania"], ["SOF", "Sofia, Bulgaria"], ["BEG", "Belgrade, Serbia"],
  ["KEF", "Reykjavik, Iceland"], ["SVO", "Moscow, Russia"], ["LED", "St Petersburg, Russia"],
  ["DEL", "Delhi, India"], ["BOM", "Mumbai, India"], ["BLR", "Bangalore, India"],
  ["MAA", "Chennai, India"], ["HYD", "Hyderabad, India"], ["CCU", "Kolkata, India"],
  ["COK", "Kochi, India"], ["KHI", "Karachi, Pakistan"], ["LHE", "Lahore, Pakistan"],
  ["ISB", "Islamabad, Pakistan"], ["DAC", "Dhaka, Bangladesh"], ["CMB", "Colombo, Sri Lanka"],
  ["KTM", "Kathmandu, Nepal"], ["BKK", "Bangkok, Thailand"], ["HKT", "Phuket, Thailand"],
  ["SIN", "Singapore"], ["KUL", "Kuala Lumpur, Malaysia"], ["CGK", "Jakarta, Indonesia"],
  ["DPS", "Bali, Indonesia"], ["MNL", "Manila, Philippines"], ["HAN", "Hanoi, Vietnam"],
  ["SGN", "Ho Chi Minh City, Vietnam"], ["PNH", "Phnom Penh, Cambodia"], ["RGN", "Yangon, Myanmar"],
  ["HKG", "Hong Kong"], ["TPE", "Taipei, Taiwan"], ["ICN", "Seoul, South Korea"],
  ["NRT", "Tokyo Narita, Japan"], ["HND", "Tokyo Haneda, Japan"], ["KIX", "Osaka, Japan"],
  ["PEK", "Beijing, China"], ["PVG", "Shanghai, China"], ["CAN", "Guangzhou, China"],
  ["SZX", "Shenzhen, China"], ["ALA", "Almaty, Kazakhstan"], ["TAS", "Tashkent, Uzbekistan"],
  ["GYD", "Baku, Azerbaijan"], ["TBS", "Tbilisi, Georgia"], ["EVN", "Yerevan, Armenia"],
  ["JFK", "New York JFK, USA"], ["EWR", "Newark, USA"], ["LGA", "New York LaGuardia, USA"],
  ["LAX", "Los Angeles, USA"], ["ORD", "Chicago, USA"], ["MIA", "Miami, USA"],
  ["ATL", "Atlanta, USA"], ["DFW", "Dallas, USA"], ["SFO", "San Francisco, USA"],
  ["IAD", "Washington DC, USA"], ["BOS", "Boston, USA"], ["YYZ", "Toronto, Canada"],
  ["YVR", "Vancouver, Canada"], ["YUL", "Montreal, Canada"], ["MEX", "Mexico City, Mexico"],
  ["GRU", "Sao Paulo, Brazil"], ["GIG", "Rio de Janeiro, Brazil"], ["EZE", "Buenos Aires, Argentina"],
  ["SCL", "Santiago, Chile"], ["BOG", "Bogota, Colombia"], ["LIM", "Lima, Peru"],
  ["SYD", "Sydney, Australia"], ["MEL", "Melbourne, Australia"], ["BNE", "Brisbane, Australia"],
  ["PER", "Perth, Australia"], ["AKL", "Auckland, New Zealand"],

  // --- Additional Africa ---
  ["DAR", "Dar es Salaam, Tanzania"], ["ZNZ", "Zanzibar, Tanzania"], ["EBB", "Entebbe/Kampala, Uganda"],
  ["KGL", "Kigali, Rwanda"], ["BJM", "Bujumbura, Burundi"], ["LUN", "Lusaka, Zambia"],
  ["HRE", "Harare, Zimbabwe"], ["MPM", "Maputo, Mozambique"], ["WDH", "Windhoek, Namibia"],
  ["GBE", "Gaborone, Botswana"], ["TNR", "Antananarivo, Madagascar"], ["MRU", "Port Louis, Mauritius"],
  ["SEZ", "Mahe Island, Seychelles"], ["LAD", "Luanda, Angola"], ["FIH", "Kinshasa, DR Congo"],
  ["BZV", "Brazzaville, Republic of Congo"], ["LBV", "Libreville, Gabon"], ["DLA", "Douala, Cameroon"],
  ["NSI", "Yaounde, Cameroon"], ["NDJ", "N'Djamena, Chad"], ["NIM", "Niamey, Niger"],
  ["OUA", "Ouagadougou, Burkina Faso"], ["BKO", "Bamako, Mali"], ["COO", "Cotonou, Benin"],
  ["LFW", "Lome, Togo"], ["FNA", "Freetown, Sierra Leone"], ["ROB", "Monrovia, Liberia"],
  ["CKY", "Conakry, Guinea"], ["BJL", "Banjul, Gambia"], ["NKC", "Nouakchott, Mauritania"],
  ["PHC", "Port Harcourt, Nigeria"], ["ABV", "Abuja, Nigeria"], ["KAN", "Kano, Nigeria"],
  ["ASM", "Asmara, Eritrea"], ["JIB", "Djibouti City, Djibouti"], ["MGQ", "Mogadishu, Somalia"],
  ["HGA", "Hargeisa, Somaliland"], ["JUB", "Juba, South Sudan"],

  // --- Additional Middle East ---
  ["AAN", "Al Ain, UAE"], ["RKT", "Ras Al Khaimah, UAE"], ["NJF", "Najaf, Iraq"],

  // --- Additional Europe ---
  ["EDI", "Edinburgh, UK"], ["GLA", "Glasgow, UK"], ["BHX", "Birmingham, UK"], ["BRS", "Bristol, UK"],
  ["NCE", "Nice, France"], ["LYS", "Lyon, France"], ["MRS", "Marseille, France"], ["TLS", "Toulouse, France"],
  ["HAM", "Hamburg, Germany"], ["DUS", "Dusseldorf, Germany"], ["STR", "Stuttgart, Germany"], ["CGN", "Cologne, Germany"],
  ["NAP", "Naples, Italy"], ["VCE", "Venice, Italy"], ["BLQ", "Bologna, Italy"], ["TRN", "Turin, Italy"],
  ["PMI", "Palma de Mallorca, Spain"], ["AGP", "Malaga, Spain"], ["SVQ", "Seville, Spain"], ["VLC", "Valencia, Spain"],
  ["BIO", "Bilbao, Spain"], ["OPO", "Porto, Portugal"], ["FAO", "Faro, Portugal"], ["LUX", "Luxembourg City, Luxembourg"],
  ["KRK", "Krakow, Poland"], ["GDN", "Gdansk, Poland"], ["BTS", "Bratislava, Slovakia"], ["LJU", "Ljubljana, Slovenia"],
  ["ZAG", "Zagreb, Croatia"], ["SPU", "Split, Croatia"], ["DBV", "Dubrovnik, Croatia"], ["SJJ", "Sarajevo, Bosnia and Herzegovina"],
  ["SKP", "Skopje, North Macedonia"], ["TIA", "Tirana, Albania"], ["PRN", "Pristina, Kosovo"],
  ["HER", "Heraklion, Greece"], ["RHO", "Rhodes, Greece"], ["CFU", "Corfu, Greece"], ["JTR", "Santorini, Greece"],
  ["MLA", "Valletta, Malta"], ["LCA", "Larnaca, Cyprus"], ["PFO", "Paphos, Cyprus"],
  ["RIX", "Riga, Latvia"], ["VNO", "Vilnius, Lithuania"], ["TLL", "Tallinn, Estonia"], ["MSQ", "Minsk, Belarus"],
  ["KBP", "Kyiv, Ukraine"], ["ODS", "Odesa, Ukraine"], ["LWO", "Lviv, Ukraine"], ["KIV", "Chisinau, Moldova"],
  ["GOT", "Gothenburg, Sweden"], ["BGO", "Bergen, Norway"], ["TRD", "Trondheim, Norway"], ["AAL", "Aalborg, Denmark"],

  // --- Additional Asia ---
  ["PKX", "Beijing Daxing, China"], ["CTU", "Chengdu, China"], ["XIY", "Xi'an, China"], ["KMG", "Kunming, China"],
  ["WUH", "Wuhan, China"], ["NKG", "Nanjing, China"], ["TSN", "Tianjin, China"], ["HGH", "Hangzhou, China"],
  ["CSX", "Changsha, China"], ["URC", "Urumqi, China"], ["HAK", "Haikou, China"], ["SYX", "Sanya, China"],
  ["MFM", "Macau"], ["KHH", "Kaohsiung, Taiwan"], ["OKA", "Okinawa, Japan"], ["FUK", "Fukuoka, Japan"],
  ["CTS", "Sapporo, Japan"], ["NGO", "Nagoya, Japan"], ["GMP", "Seoul Gimpo, South Korea"], ["PUS", "Busan, South Korea"],
  ["CJU", "Jeju, South Korea"], ["UBN", "Ulaanbaatar, Mongolia"], ["VTE", "Vientiane, Laos"], ["LPQ", "Luang Prabang, Laos"],
  ["REP", "Siem Reap, Cambodia"], ["MDL", "Mandalay, Myanmar"], ["BWN", "Bandar Seri Begawan, Brunei"],
  ["CEB", "Cebu, Philippines"], ["DVO", "Davao, Philippines"], ["SUB", "Surabaya, Indonesia"], ["KNO", "Medan, Indonesia"],
  ["UPG", "Makassar, Indonesia"], ["PNQ", "Pune, India"], ["AMD", "Ahmedabad, India"], ["GOI", "Goa, India"],
  ["JAI", "Jaipur, India"], ["LKO", "Lucknow, India"], ["PAT", "Patna, India"], ["IXC", "Chandigarh, India"],
  ["TRV", "Thiruvananthapuram, India"], ["MLE", "Male, Maldives"],
  ["NQZ", "Astana, Kazakhstan"], ["FRU", "Bishkek, Kyrgyzstan"], ["DYU", "Dushanbe, Tajikistan"], ["ASB", "Ashgabat, Turkmenistan"],

  // --- Additional North America ---
  ["PHX", "Phoenix, USA"], ["DEN", "Denver, USA"], ["SEA", "Seattle, USA"], ["LAS", "Las Vegas, USA"],
  ["MSP", "Minneapolis, USA"], ["DTW", "Detroit, USA"], ["PHL", "Philadelphia, USA"], ["CLT", "Charlotte, USA"],
  ["HOU", "Houston Hobby, USA"], ["IAH", "Houston, USA"], ["SAN", "San Diego, USA"], ["TPA", "Tampa, USA"],
  ["MCO", "Orlando, USA"], ["FLL", "Fort Lauderdale, USA"], ["HNL", "Honolulu, USA"], ["ANC", "Anchorage, USA"],
  ["PDX", "Portland, USA"], ["AUS", "Austin, USA"], ["SLC", "Salt Lake City, USA"], ["STL", "St Louis, USA"],
  ["BWI", "Baltimore, USA"], ["DCA", "Washington Reagan, USA"], ["MSY", "New Orleans, USA"], ["OAK", "Oakland, USA"],
  ["YYC", "Calgary, Canada"], ["YEG", "Edmonton, Canada"], ["YOW", "Ottawa, Canada"], ["YHZ", "Halifax, Canada"],
  ["YWG", "Winnipeg, Canada"], ["GDL", "Guadalajara, Mexico"], ["MTY", "Monterrey, Mexico"], ["CUN", "Cancun, Mexico"],
  ["SJD", "Los Cabos, Mexico"], ["PVR", "Puerto Vallarta, Mexico"], ["TIJ", "Tijuana, Mexico"],

  // --- Central America & Caribbean ---
  ["GUA", "Guatemala City, Guatemala"], ["SAL", "San Salvador, El Salvador"], ["TGU", "Tegucigalpa, Honduras"],
  ["MGA", "Managua, Nicaragua"], ["SJO", "San Jose, Costa Rica"], ["PTY", "Panama City, Panama"],
  ["HAV", "Havana, Cuba"], ["SDQ", "Santo Domingo, Dominican Republic"], ["PUJ", "Punta Cana, Dominican Republic"],
  ["PAP", "Port-au-Prince, Haiti"], ["SJU", "San Juan, Puerto Rico"], ["MBJ", "Montego Bay, Jamaica"],
  ["KIN", "Kingston, Jamaica"], ["NAS", "Nassau, Bahamas"], ["BGI", "Bridgetown, Barbados"],
  ["POS", "Port of Spain, Trinidad and Tobago"], ["ANU", "St John's, Antigua"], ["BZE", "Belize City, Belize"],
  ["CUR", "Willemstad, Curacao"], ["AUA", "Oranjestad, Aruba"],

  // --- Additional South America ---
  ["BSB", "Brasilia, Brazil"], ["CNF", "Belo Horizonte, Brazil"], ["SSA", "Salvador, Brazil"],
  ["REC", "Recife, Brazil"], ["FOR", "Fortaleza, Brazil"], ["MAO", "Manaus, Brazil"],
  ["POA", "Porto Alegre, Brazil"], ["CWB", "Curitiba, Brazil"], ["MDZ", "Mendoza, Argentina"],
  ["COR", "Cordoba, Argentina"], ["USH", "Ushuaia, Argentina"], ["MVD", "Montevideo, Uruguay"],
  ["ASU", "Asuncion, Paraguay"], ["VVI", "Santa Cruz, Bolivia"], ["LPB", "La Paz, Bolivia"],
  ["UIO", "Quito, Ecuador"], ["GYE", "Guayaquil, Ecuador"], ["CTG", "Cartagena, Colombia"],
  ["MDE", "Medellin, Colombia"], ["CLO", "Cali, Colombia"], ["CCS", "Caracas, Venezuela"],
  ["GEO", "Georgetown, Guyana"], ["PBM", "Paramaribo, Suriname"],

  // --- Additional Oceania ---
  ["ADL", "Adelaide, Australia"], ["CNS", "Cairns, Australia"], ["OOL", "Gold Coast, Australia"],
  ["DRW", "Darwin, Australia"], ["HBA", "Hobart, Australia"], ["CHC", "Christchurch, New Zealand"],
  ["ZQN", "Queenstown, New Zealand"], ["WLG", "Wellington, New Zealand"], ["NAN", "Nadi, Fiji"],
  ["POM", "Port Moresby, Papua New Guinea"], ["NOU", "Noumea, New Caledonia"], ["PPT", "Papeete, Tahiti"],
  ["APW", "Apia, Samoa"], ["TBU", "Nuku'alofa, Tonga"], ["GUM", "Guam"], ["SPN", "Saipan"],
].map(([code, place]) => `${code} - ${place}`.toUpperCase());

export default function TicketsApp({ onChangeServer, currentServerUrl } = {}) {
  const [tickets, setTickets] = useState([]);
  const [employees, setEmployees] = useState(null); // null = not loaded yet
  const [currentUser, setCurrentUser] = useState(null); // { username, name, isAdmin }
  const [loading, setLoading] = useState(true);

  // Presence: which employees are currently connected (main account only)
  const [presenceMap, setPresenceMap] = useState({}); // username -> last-seen timestamp
  const [showOnlineList, setShowOnlineList] = useState(false);
  const [restoreError, setRestoreError] = useState("");
  const [restoreSuccess, setRestoreSuccess] = useState("");
  const fileInputRef = useRef(null);
  // Timestamp of when the current session started (login or restored on page load).
  // A remote force-sign-out (see handleForceSignOut) is only honored if it happened
  // AFTER this moment, so it can't retroactively sign someone out of a brand new session.
  const sessionStartedAtRef = useRef(0);
  // window.confirm doesn't work in this sandboxed preview, so confirmations use this
  // in-app dialog instead: { message, onConfirm } while open, null while hidden.
  const [confirmDialog, setConfirmDialog] = useState(null);
  const requestConfirm = (message, onConfirm) => setConfirmDialog({ message, onConfirm });

  const [loginUsername, setLoginUsername] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [showPassword, setShowPassword] = useState(false);

  const [setupName, setSetupName] = useState("");
  const [setupUsername, setSetupUsername] = useState("");
  const [setupPassword, setSetupPassword] = useState("");

  const [showManage, setShowManage] = useState(false);
  const [newEmployee, setNewEmployee] = useState(emptyNewEmployee);
  const [showNewEmployeePerms, setShowNewEmployeePerms] = useState(false);
  const [openPermissionsFor, setOpenPermissionsFor] = useState(null); // username, or null if closed
  const [manageError, setManageError] = useState("");
  const [visiblePasswords, setVisiblePasswords] = useState({});
  const [editingUsername, setEditingUsername] = useState(null);
  const [editDraft, setEditDraft] = useState({ name: "", username: "", password: "" });
  const [editShowPassword, setEditShowPassword] = useState(false);

  const [showManageCompanies, setShowManageCompanies] = useState(false);
  const [showCompaniesList, setShowCompaniesList] = useState(false);
  const [newCompanyDraft, setNewCompanyDraft] = useState(emptyCompanyDraft);
  const [editingCompanyName, setEditingCompanyName] = useState(null);
  const [companyError, setCompanyError] = useState("");

  const [showChangePassword, setShowChangePassword] = useState(false);
  const [currentPasswordInput, setCurrentPasswordInput] = useState("");
  const [newPasswordInput, setNewPasswordInput] = useState("");
  const [confirmPasswordInput, setConfirmPasswordInput] = useState("");
  const [passwordError, setPasswordError] = useState("");
  const [passwordSuccess, setPasswordSuccess] = useState("");

  const [form, setForm] = useState(getEmptyForm);
  // Whether the Supplier field is in "type your own name" mode (chosen via the Other option).
  const [supplierOther, setSupplierOther] = useState(false);

  // ---------- Hotels ----------
  const [hotelBookings, setHotelBookings] = useState([]);
  const [hotelForm, setHotelForm] = useState(getEmptyHotelForm);
  const [hotelError, setHotelError] = useState("");
  const [hotelEditingId, setHotelEditingId] = useState(null);
  // The hotel booking currently shown in the read-only details modal (null = closed).
  const [viewingHotelBooking, setViewingHotelBooking] = useState(null);
  // Whether the "Add supplier" / "Add hotel name" panels at the top of the Hotels
  // page are currently open, plus the text typed into each panel's input.
  const [showAddSupplierPanel, setShowAddSupplierPanel] = useState(false);
  const [showAddHotelNamePanel, setShowAddHotelNamePanel] = useState(false);
  const [newSupplierDraft, setNewSupplierDraft] = useState("");
  const [newHotelNameDraft, setNewHotelNameDraft] = useState("");
  // Whether the Hotel name / Supplier fields on the booking form are in "type your
  // own name" mode, same pattern as supplierOther for flight tickets above.
  const [hotelSupplierOther, setHotelSupplierOther] = useState(false);
  const [hotelNameOther, setHotelNameOther] = useState(false);
  // USD -> EGP exchange rate, used to also show a USD booking's value in EGP.
  // Entered by hand (no CBE API is publicly reachable from the browser), and saved so
  // everyone signed in sees today's rate without re-typing it.
  const [usdToEgpRate, setUsdToEgpRate] = useState(null);
  const [usdToEgpRateDate, setUsdToEgpRateDate] = useState("");
  const [query, setQuery] = useState("");
  const [error, setError] = useState("");
  // Clicking a ticket row opens a full detail view of that ticket (id stored here).
  const [viewingTicketId, setViewingTicketId] = useState(null);
  const [notesDraft, setNotesDraft] = useState("");
  const [notesSaved, setNotesSaved] = useState(false);
  // Refund tracking on the ticket detail view: two amounts (refunded by the airline,
  // refunded to the customer) plus which customer (when a booking has more than one)
  // the refund belongs to, recorded/edited per ticket.
  const [showRefundForm, setShowRefundForm] = useState(false);
  const [refundDraft, setRefundDraft] = useState({ airlineAmount: "", customerAmount: "", customerIndex: 0 });
  const [refundSaved, setRefundSaved] = useState(false);
  const [selectedMonth, setSelectedMonth] = useState("");
  const [selectedYear, setSelectedYear] = useState("");
  const [selectedCompany, setSelectedCompany] = useState("");
  const [selectedEmployee, setSelectedEmployee] = useState("");
  const [selectedSupplier, setSelectedSupplier] = useState("");

  // Every value ever entered (companies, customers, airlines, cities) is kept here so it
  // can be offered as an autocomplete suggestion later, even if the original ticket is deleted.
  const [suggestions, setSuggestions] = useState({ companies: [], customers: [], airlines: [], cities: [], suppliers: [], hotelNames: [] });

  // Tracks whether the one-time "create the main account" step has ever been completed.
  // Once true, the first-run setup screen must never be shown again — even if the employee
  // list later becomes empty (e.g. accounts deleted, a bad restore) — so no one can
  // create a fresh, unauthenticated admin account after the app has already been set up.
  const [setupComplete, setSetupComplete] = useState(null); // null = not loaded yet

  // Top-level section switcher: "flights" holds all existing ticket functionality;
  // "hotels" and "cars" are placeholders for future sections.
  const [activeSection, setActiveSection] = useState("flights");

  useEffect(() => {
    (async () => {
      try {
        const [ticketsRes, hotelsRes, employeesRes, sessionRes, suggestionsRes, setupRes] = await Promise.all([
          window.storage.get("tickets:list", true).catch(() => null),
          window.storage.get("tickets:hotels", true).catch(() => null),
          window.storage.get("tickets:employees", true).catch(() => null),
          window.storage.get("session:user", false).catch(() => null),
          window.storage.get("tickets:suggestions", true).catch(() => null),
          window.storage.get("tickets:setupComplete", true).catch(() => null),
        ]);
        const ticketsData = ticketsRes && ticketsRes.value ? JSON.parse(ticketsRes.value) : [];
        const hotelsData = hotelsRes && hotelsRes.value ? JSON.parse(hotelsRes.value) : [];
        const employeesData = employeesRes && employeesRes.value ? JSON.parse(employeesRes.value) : [];
        setTickets(ticketsData);
        setHotelBookings(hotelsData);
        setEmployees(employeesData);
        // If accounts already exist, the setup step has clearly already happened even if the
        // flag itself is missing (e.g. app used before this flag existed).
        setSetupComplete(!!(setupRes && setupRes.value === "true") || employeesData.length > 0);
        if (suggestionsRes && suggestionsRes.value) {
          try {
            const parsed = JSON.parse(suggestionsRes.value);
            setSuggestions({
              companies: parsed.companies || [],
              customers: [], // never restore saved customer names — this field must have no autocomplete history
              airlines: parsed.airlines || [],
              cities: parsed.cities || [],
              suppliers: parsed.suppliers || [],
              hotelNames: parsed.hotelNames || [],
            });
          } catch (e) {
            // ignore malformed suggestions data
          }
        }

        if (sessionRes && sessionRes.value) {
          const savedUsername = sessionRes.value;
          const match = employeesData.find((e) => e.username === savedUsername);
          if (match) {
            sessionStartedAtRef.current = Date.now();
            setCurrentUser({ username: match.username, name: match.name, isAdmin: !!match.isAdmin });
          }
        }
      } catch (e) {
        setEmployees([]);
        setSetupComplete(false);
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  const LIVE_REFRESH_INTERVAL_MS = 5 * 1000;

  // Keeps tickets, employee accounts, and saved suggestions (companies/customers/
  // airlines/cities) in sync across everyone who's signed in, by periodically re-reading
  // the shared storage keys. window.storage has no push/subscribe API, so short polling
  // is the only way to reflect other users' changes without a manual page refresh.
  useEffect(() => {
    if (!currentUser) return;
    let cancelled = false;
    const loadCoreData = async () => {
      try {
        const [ticketsRes, hotelsRes, employeesRes, suggestionsRes] = await Promise.all([
          window.storage.get("tickets:list", true).catch(() => null),
          window.storage.get("tickets:hotels", true).catch(() => null),
          window.storage.get("tickets:employees", true).catch(() => null),
          window.storage.get("tickets:suggestions", true).catch(() => null),
        ]);
        if (cancelled) return;
        if (ticketsRes && ticketsRes.value) {
          try {
            setTickets(JSON.parse(ticketsRes.value));
          } catch (e) {
            // ignore malformed data for this cycle, try again next poll
          }
        }
        if (hotelsRes && hotelsRes.value) {
          try {
            setHotelBookings(JSON.parse(hotelsRes.value));
          } catch (e) {
            // ignore malformed data for this cycle, try again next poll
          }
        }
        if (employeesRes && employeesRes.value) {
          try {
            setEmployees(JSON.parse(employeesRes.value));
          } catch (e) {
            // ignore malformed data for this cycle, try again next poll
          }
        }
        if (suggestionsRes && suggestionsRes.value) {
          try {
            const parsed = JSON.parse(suggestionsRes.value);
            setSuggestions({
              companies: parsed.companies || [],
              customers: [], // never restore saved customer names — this field must have no autocomplete history
              airlines: parsed.airlines || [],
              cities: parsed.cities || [],
              suppliers: parsed.suppliers || [],
              hotelNames: parsed.hotelNames || [],
            });
          } catch (e) {
            // ignore malformed data for this cycle, try again next poll
          }
        }
      } catch (e) {
        // Live refresh is best-effort; a failed poll just tries again next interval
      }
    };
    const interval = setInterval(loadCoreData, LIVE_REFRESH_INTERVAL_MS);
    return () => {
      cancelled = true;
      clearInterval(interval);
    };
  }, [currentUser]);


  const ONLINE_THRESHOLD_MS = 15 * 1000; // considered "connected" if seen in the last 15s
  const HEARTBEAT_INTERVAL_MS = 5 * 1000;

  // A short, human-readable description of what this signed-in account appears to be
  // doing right now, broadcast alongside the presence heartbeat below so the main
  // account's "online now" list can show it next to each employee.
  const myActivity = (() => {
    if (showManage) return "Managing employees";
    if (showManageCompanies) return "Managing companies";
    if (activeSection === "hotels") {
      if (viewingHotelBooking) return "Viewing a hotel booking";
      if (hotelEditingId) return "Editing a hotel booking";
      return "Hotels";
    }
    if (activeSection === "cars") return "Cars";
    if (activeSection === "files") return "Files";
    // Flights (the default section)
    if (viewingTicketId) return "Viewing a ticket";
    if (form.id) return "Editing a ticket";
    return "Flights";
  })();
  // Kept in a ref (rather than read directly) so the heartbeat interval below — which
  // only re-subscribes when currentUser changes — always sends the latest activity
  // instead of the value captured when the interval was first created.
  const myActivityRef = useRef(myActivity);
  useEffect(() => {
    myActivityRef.current = myActivity;
  }, [myActivity]);

  // While signed in, periodically mark this account as "connected" so the main account can see it
  useEffect(() => {
    if (!currentUser) return;
    let cancelled = false;
    const beat = async () => {
      try {
        await window.storage.set(
          `tickets:presence:${currentUser.username}`,
          JSON.stringify({ name: currentUser.name, ts: Date.now(), activity: myActivityRef.current }),
          true
        );
      } catch (e) {
        // Presence is a convenience feature; failures here are silent
      }
    };
    beat();
    const interval = setInterval(() => {
      if (!cancelled) beat();
    }, HEARTBEAT_INTERVAL_MS);
    return () => {
      cancelled = true;
      clearInterval(interval);
    };
  }, [currentUser]);

  // The main account polls who else is currently connected
  useEffect(() => {
    if (!currentUser || !currentUser.isAdmin) return;
    let cancelled = false;
    const loadPresence = async () => {
      try {
        const listRes = await window.storage.list("tickets:presence:", true);
        const keys = (listRes && listRes.keys) || [];
        const entries = await Promise.all(
          keys.map(async (k) => {
            try {
              const r = await window.storage.get(k, true);
              if (!r || !r.value) return null;
              const parsed = JSON.parse(r.value);
              const username = k.replace("tickets:presence:", "");
              return [username, { ts: parsed.ts, activity: parsed.activity || "" }];
            } catch (e) {
              return null;
            }
          })
        );
        if (cancelled) return;
        const map = {};
        entries.forEach((entry) => {
          if (entry) map[entry[0]] = entry[1];
        });
        setPresenceMap(map);
      } catch (e) {
        // ignore presence load failures
      }
    };
    loadPresence();
    const interval = setInterval(loadPresence, LIVE_REFRESH_INTERVAL_MS);
    return () => {
      cancelled = true;
      clearInterval(interval);
    };
  }, [currentUser]);

  const isOnline = (username) => {
    const entry = presenceMap[username];
    return !!entry && Date.now() - entry.ts < ONLINE_THRESHOLD_MS;
  };
  const onlineUsernames = Object.keys(presenceMap).filter((u) => isOnline(u));

  // Detects a remote "force sign-out": when the main account signs someone out from the
  // "online now" panel, a shared flag is written with a timestamp (see handleForceSignOut
  // below). Every signed-in client — including this one — checks its own flag on each
  // heartbeat and signs itself out automatically if the flag is newer than when this
  // particular session started (so it can never retroactively kill a brand-new login).
  useEffect(() => {
    if (!currentUser) return;
    let cancelled = false;
    const checkForceLogout = async () => {
      try {
        const res = await window.storage.get(`tickets:forceLogout:${currentUser.username}`, true).catch(() => null);
        if (cancelled || !res || !res.value) return;
        const ts = parseInt(res.value, 10);
        if (ts && ts > sessionStartedAtRef.current) {
          await handleLogout();
        }
      } catch (e) {
        // Best-effort; a missed check just retries on the next heartbeat
      }
    };
    const interval = setInterval(() => {
      if (!cancelled) checkForceLogout();
    }, HEARTBEAT_INTERVAL_MS);
    return () => {
      cancelled = true;
      clearInterval(interval);
    };
  }, [currentUser]);

  // Clears the "online now" presence flag when the page/tab is closed or navigated away
  // from, so this employee stops showing as online right away. Deliberately does NOT
  // touch the saved session here — "beforeunload"/"pagehide" also fire on a normal page
  // refresh, and clearing the session there was signing people out just from reloading
  // the page. Signing out now only happens via the explicit Sign out button, a remote
  // force-sign-out, or the inactivity timeout below.
  useEffect(() => {
    if (!currentUser) return;
    const username = currentUser.username;
    const handleUnload = () => {
      try { window.storage.delete(`tickets:presence:${username}`, true); } catch (e) {}
    };
    window.addEventListener("pagehide", handleUnload);
    window.addEventListener("beforeunload", handleUnload);
    return () => {
      window.removeEventListener("pagehide", handleUnload);
      window.removeEventListener("beforeunload", handleUnload);
    };
  }, [currentUser]);

  // Auto sign-out after 30 minutes of inactivity. Any mouse, keyboard, scroll, or touch
  // activity resets the timer; if it ever fires, the session is ended the same way the
  // Sign out button does it.
  const INACTIVITY_TIMEOUT_MS = 30 * 60 * 1000;
  useEffect(() => {
    if (!currentUser) return;
    let timeoutId;
    const resetTimer = () => {
      clearTimeout(timeoutId);
      timeoutId = setTimeout(() => {
        handleLogout();
      }, INACTIVITY_TIMEOUT_MS);
    };
    const activityEvents = ["mousedown", "mousemove", "keydown", "wheel", "touchstart", "scroll"];
    activityEvents.forEach((evt) => window.addEventListener(evt, resetTimer, { passive: true }));
    resetTimer();
    return () => {
      clearTimeout(timeoutId);
      activityEvents.forEach((evt) => window.removeEventListener(evt, resetTimer));
    };
  }, [currentUser]);

  const persistTickets = async (next) => {
    setTickets(next);
    try {
      await window.storage.set("tickets:list", JSON.stringify(next), true);
    } catch (e) {
      setError("Could not save data, please try again");
    }
  };

  // The USD -> EGP rate is entered by hand (e.g. from the CBE's published rate each
  // morning) and saved to shared storage, so every signed-in employee sees the same
  // rate without each of them having to type it in separately.
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("tickets:usdRate", true).catch(() => null);
        if (res && res.value) {
          const parsed = JSON.parse(res.value);
          setUsdToEgpRate(parsed.rate ?? null);
          setUsdToEgpRateDate(parsed.date || "");
        }
      } catch (e) {
        // no saved rate yet
      }
    })();
  }, []);

  const persistUsdRate = async (rate) => {
    const date = todayDateStr();
    setUsdToEgpRate(rate);
    setUsdToEgpRateDate(date);
    try {
      await window.storage.set("tickets:usdRate", JSON.stringify({ rate, date }), true);
    } catch (e) {
      // Saving the rate is best-effort; the typed value still applies locally either way
    }
  };

  const persistHotelBookings = async (next) => {
    setHotelBookings(next);
    try {
      await window.storage.set("tickets:hotels", JSON.stringify(next), true);
    } catch (e) {
      setHotelError("Could not save data, please try again");
    }
  };

  const persistEmployees = async (next) => {
    setEmployees(next);
    try {
      await window.storage.set("tickets:employees", JSON.stringify(next), true);
    } catch (e) {
      setManageError("Could not save the employee list, please try again");
    }
  };

  const persistSuggestions = async (next) => {
    setSuggestions(next);
    try {
      await window.storage.set("tickets:suggestions", JSON.stringify(next), true);
    } catch (e) {
      // Suggestions are a convenience feature, so failures here are silent
    }
  };

  // Remembers values entered on a ticket (airline, cities) so they keep showing up as
  // autocomplete options later, even if this ticket gets deleted. Companies are
  // intentionally excluded — a new company can only be registered via the
  // "Manage companies" button, never auto-added just by typing a new name on a ticket.
  const rememberSuggestionsFromRecord = (record) => {
    const addUnique = (list, value) => {
      const v = (value || "").trim();
      if (!v) return list;
      return list.some((existing) => existing.toLowerCase() === v.toLowerCase()) ? list : [...list, v];
    };
    let next = {
      companies: [...suggestions.companies],
      // Customer names are intentionally never remembered here — the customer field
      // must never offer autocomplete/history of previously typed names.
      customers: [],
      airlines: [...suggestions.airlines],
      cities: [...suggestions.cities],
      suppliers: [...(suggestions.suppliers || [])],
      hotelNames: [...(suggestions.hotelNames || [])],
    };
    next.airlines = addUnique(next.airlines, record.airline);
    next.cities = addUnique(next.cities, record.from);
    next.cities = addUnique(next.cities, record.to);
    if (Array.isArray(record.destinations)) {
      record.destinations.forEach((d) => { next.cities = addUnique(next.cities, d); });
    }
    persistSuggestions(next);
  };


  // Lets an admin (or an employee granted the Manage companies permission) register a
  // company's full details — name, tax number, commercial registration number, and phone
  // numbers — so they're always available to pick from the Company field and filter, even
  // before any ticket has been entered for them. If editingCompanyName is set, this saves
  // changes to that existing record instead of adding a new one.
  const handleAddCompany = () => {
    if (!canManageCompanies) return;
    const name = newCompanyDraft.name.trim();
    if (!name) return;
    const duplicate = suggestions.companies.some(
      (c) =>
        companyName(c).toLowerCase() === name.toLowerCase() &&
        companyName(c).toLowerCase() !== (editingCompanyName || "").toLowerCase()
    );
    if (duplicate) {
      setCompanyError("A company with that name already exists");
      return;
    }
    const record = {
      name,
      taxNumber: newCompanyDraft.taxNumber.trim(),
      commercialReg: newCompanyDraft.commercialReg.trim(),
      phones: newCompanyDraft.phones
        .split(/[,\n]/)
        .map((p) => p.trim())
        .filter(Boolean),
    };
    const companies = editingCompanyName
      ? suggestions.companies.map((c) => (companyName(c) === editingCompanyName ? record : c))
      : [...suggestions.companies, record];
    persistSuggestions({ ...suggestions, companies });
    setNewCompanyDraft(emptyCompanyDraft);
    setEditingCompanyName(null);
    setCompanyError("");
  };

  // Loads an existing company's saved details back into the form so they can be edited.
  const handleEditCompanyClick = (c) => {
    setEditingCompanyName(companyName(c));
    setNewCompanyDraft({
      name: companyName(c),
      taxNumber: typeof c === "object" ? c.taxNumber || "" : "",
      commercialReg: typeof c === "object" ? c.commercialReg || "" : "",
      phones: typeof c === "object" && Array.isArray(c.phones) ? c.phones.join(", ") : "",
    });
  };

  const cancelEditCompany = () => {
    setEditingCompanyName(null);
    setNewCompanyDraft(emptyCompanyDraft);
  };

  // Removes a company from the saved suggestions list. Existing tickets already
  // recorded under that company name are untouched — this only affects the picker.
  const handleDeleteCompany = (name) => {
    if (!canManageCompanies) return;
    persistSuggestions({
      ...suggestions,
      companies: suggestions.companies.filter((c) => companyName(c) !== name),
    });
    if (editingCompanyName === name) cancelEditCompany();
  };

  const profit = (net, sold) => {
    const n = parseFloat(net) || 0;
    const s = parseFloat(sold) || 0;
    return s - n;
  };

  // ---------- Hotels ----------
  const resetHotelForm = () => {
    setHotelForm(getEmptyHotelForm());
    setHotelEditingId(null);
    setHotelError("");
    setHotelSupplierOther(false);
    setHotelNameOther(false);
  };

  const addHotelRoomLine = () => {
    setHotelForm({ ...hotelForm, roomLines: [...hotelForm.roomLines, emptyRoomLine()] });
  };

  const removeHotelRoomLine = (lineId) => {
    if (hotelForm.roomLines.length <= 1) return; // always keep at least one line
    setHotelForm({ ...hotelForm, roomLines: hotelForm.roomLines.filter((l) => l.id !== lineId) });
  };

  const updateHotelRoomLine = (lineId, patch) => {
    setHotelForm({
      ...hotelForm,
      roomLines: hotelForm.roomLines.map((l) => (l.id === lineId ? { ...l, ...patch } : l)),
    });
  };

  // Updates one adult guest's name within a room line, by that guest's position.
  const updateRoomGuest = (lineId, guestIndex, name) => {
    const line = hotelForm.roomLines.find((l) => l.id === lineId);
    if (!line) return;
    const guests = (line.guests || []).map((g, i) => (i === guestIndex ? { ...g, name } : g));
    updateHotelRoomLine(lineId, { guests });
  };

  const addRoomChild = (lineId) => {
    const line = hotelForm.roomLines.find((l) => l.id === lineId);
    if (!line) return;
    updateHotelRoomLine(lineId, { children: [...(line.children || []), emptyChild()] });
  };

  const updateRoomChild = (lineId, childId, patch) => {
    const line = hotelForm.roomLines.find((l) => l.id === lineId);
    if (!line) return;
    updateHotelRoomLine(lineId, {
      children: (line.children || []).map((c) => (c.id === childId ? { ...c, ...patch } : c)),
    });
  };

  const removeRoomChild = (lineId, childId) => {
    const line = hotelForm.roomLines.find((l) => l.id === lineId);
    if (!line) return;
    updateHotelRoomLine(lineId, { children: (line.children || []).filter((c) => c.id !== childId) });
  };

  const handleSaveHotel = async () => {
    setHotelError("");
    // Company name is optional — a blank company means an Individual booking, so it's
    // no longer part of the required-fields check below.
    if (!hotelForm.hotel.trim() || !hotelForm.supplier.trim()) {
      setHotelError("Please fill in the hotel and supplier fields");
      return;
    }
    const lines = hotelForm.roomLines || [];
    if (lines.length === 0) {
      setHotelError("Please add at least one room line");
      return;
    }
    for (const l of lines) {
      if ((parseInt(l.count, 10) || 0) < 1) {
        setHotelError("Each room line needs at least 1 room");
        return;
      }
      if (l.netPrice === "" || l.soldPrice === "") {
        setHotelError("Please fill in the net and sold price for every room line");
        return;
      }
      if (!l.checkIn || !l.checkOut) {
        setHotelError("Please fill in the check-in and check-out dates for every room");
        return;
      }
      if (new Date(l.checkOut) < new Date(l.checkIn)) {
        setHotelError("Check-out date can't be before check-in date for a room");
        return;
      }
      // Only the first guest in each room is required — the rest are optional.
      if (!l.guests || !l.guests[0] || !l.guests[0].name.trim()) {
        setHotelError("Please enter at least the first guest's name for every room");
        return;
      }
    }

    if (hotelEditingId) {
      const next = hotelBookings.map((h) =>
        h.id === hotelEditingId ? { ...h, ...hotelForm, id: hotelEditingId } : h
      );
      await persistHotelBookings(next);
    } else {
      const record = {
        ...hotelForm,
        id: `H-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        employee: currentUser.name,
        employeeUsername: currentUser.username,
      };
      await persistHotelBookings([record, ...hotelBookings]);
    }
    resetHotelForm();
  };

  const handleEditHotelClick = (h) => {
    setHotelEditingId(h.id);
    setHotelForm({
      id: h.id,
      employee: h.employee || "",
      customer: h.customer || "",
      hotel: h.hotel || "",
      supplier: h.supplier || "",
      roomLines:
        Array.isArray(h.roomLines) && h.roomLines.length > 0
          ? h.roomLines.map((l) => ({
              ...l,
              id: l.id || emptyRoomLine().id,
              currency: l.currency || h.currency || "EGP",
              // Legacy bookings kept dates on the booking itself rather than per room —
              // fall back to those so older records still show something sensible.
              checkIn: l.checkIn || h.checkIn || todayDateStr(),
              checkOut: l.checkOut || h.checkOut || todayDateStr(),
              // Legacy bookings had no guest names — pad an empty list to match capacity.
              guests: guestsForCapacity(l.guests, ROOM_CAPACITY[l.roomType] || 1),
              children: Array.isArray(l.children) ? l.children : [],
            }))
          : [emptyRoomLine()],
      bookingDate: h.bookingDate || todayDateStr(),
      notes: h.notes || "",
    });
    setHotelSupplierOther(!!h.supplier && !suggestions.suppliers.includes(h.supplier));
    setHotelNameOther(!!h.hotel && !suggestions.hotelNames.includes(h.hotel));
    setHotelError("");
  };

  const handleDeleteHotel = (id) => {
    requestConfirm("Delete this hotel booking? This cannot be undone.", async () => {
      await persistHotelBookings(hotelBookings.filter((h) => h.id !== id));
      if (hotelEditingId === id) resetHotelForm();
      setConfirmDialog(null);
    });
  };

  // Registers a new supplier name so it's always available to pick from the Hotels
  // page's Supplier field, via the "+ Add supplier" button at the top of the page.
  const handleAddSupplierName = () => {
    const name = newSupplierDraft.trim();
    if (!name) return;
    const duplicate = (suggestions.suppliers || []).some((s) => s.toLowerCase() === name.toLowerCase());
    if (duplicate) {
      setHotelError("This supplier already exists");
      return;
    }
    persistSuggestions({ ...suggestions, suppliers: [...(suggestions.suppliers || []), name] });
    setNewSupplierDraft("");
    setHotelError("");
  };

  const handleDeleteSupplierName = (name) => {
    persistSuggestions({ ...suggestions, suppliers: (suggestions.suppliers || []).filter((s) => s !== name) });
  };

  // Registers a new hotel name so it's always available to pick from the Hotels
  // page's Hotel name field, via the "+ Add hotel name" button at the top of the page.
  const handleAddHotelName = () => {
    const name = newHotelNameDraft.trim();
    if (!name) return;
    const duplicate = (suggestions.hotelNames || []).some((h) => h.toLowerCase() === name.toLowerCase());
    if (duplicate) {
      setHotelError("This hotel already exists");
      return;
    }
    persistSuggestions({ ...suggestions, hotelNames: [...(suggestions.hotelNames || []), name] });
    setNewHotelNameDraft("");
    setHotelError("");
  };

  const handleDeleteHotelName = (name) => {
    persistSuggestions({ ...suggestions, hotelNames: (suggestions.hotelNames || []).filter((h) => h !== name) });
  };

  // ---------- Auth ----------
  const handleCreateFirstAdmin = async () => {
    setLoginError("");
    if (!setupName.trim() || !setupUsername.trim() || !setupPassword) {
      setLoginError("Please fill in all fields");
      return;
    }
    // The first account created becomes the main/admin account.
    // Only this account (or another account it later promotes) can manage employees.
    const admin = {
      name: setupName.trim(),
      username: setupUsername.trim(),
      password: setupPassword,
      isAdmin: true,
    };
    await persistEmployees([admin]);
    await window.storage.set("tickets:setupComplete", "true", true).catch(() => {});
    setSetupComplete(true);
    await window.storage.set("session:user", admin.username, false);
    sessionStartedAtRef.current = Date.now();
    setCurrentUser({ username: admin.username, name: admin.name, isAdmin: true });
    setSetupName(""); setSetupUsername(""); setSetupPassword("");
  };

  const handleLogin = async () => {
    setLoginError("");
    const match = (employees || []).find(
      (e) => e.username === loginUsername.trim() && e.password === loginPassword
    );
    if (!match) {
      setLoginError("Incorrect username or password");
      return;
    }
    await window.storage.set("session:user", match.username, false);
    sessionStartedAtRef.current = Date.now();
    setCurrentUser({ username: match.username, name: match.name, isAdmin: !!match.isAdmin });
    setLoginUsername(""); setLoginPassword("");
  };

  const handleLogout = async () => {
    await window.storage.delete("session:user", false).catch(() => {});
    setCurrentUser(null);
    setShowManage(false);
    setEditingUsername(null);
    setVisiblePasswords({});
  };

  // Lets the main account remotely sign out any currently-online employee (or itself)
  // from the "online now" panel. This account has no way to reach into another browser's
  // own local session storage, so instead it writes a shared timestamped flag; that
  // employee's own client picks it up on its next heartbeat (every few seconds) and signs
  // itself out. Their presence is cleared immediately here so they show as offline right away.
  const handleForceSignOut = async (username) => {
    try {
      await window.storage.set(`tickets:forceLogout:${username}`, String(Date.now()), true);
      await window.storage.delete(`tickets:presence:${username}`, true).catch(() => {});
      setPresenceMap((prev) => {
        const next = { ...prev };
        delete next[username];
        return next;
      });
    } catch (e) {
      // Best-effort; the admin can just try again
    }
  };

  const handleAddEmployee = async () => {
    setManageError("");
    if (!currentUser.isAdmin) {
      setManageError("Only the main account can add employees");
      return;
    }
    if (!newEmployee.name.trim() || !newEmployee.username.trim() || !newEmployee.password) {
      setManageError("Please fill in all fields");
      return;
    }
    if ((employees || []).some((e) => e.username === newEmployee.username.trim())) {
      setManageError("That username already exists");
      return;
    }
    const next = [
      ...(employees || []),
      {
        ...newEmployee,
        username: newEmployee.username.trim(),
        isAdmin: false,
        ...reconcilePermissions(newEmployee),
      },
    ];
    await persistEmployees(next);
    setNewEmployee(emptyNewEmployee);
    setShowNewEmployeePerms(false);
  };

  // Applies a grade's preset permissions to an employee. The grade itself is stored
  // (for the badge/label), and every toggle it sets can still be flipped individually
  // afterwards via handleTogglePermission — the preset is just a fast starting point.
  const handleRoleChange = async (username, role) => {
    if (!currentUser.isAdmin) {
      setManageError("Only the main account can change employee permissions");
      return;
    }
    const preset = ROLE_PRESETS[role] || ROLE_PRESETS.employee;
    const next = (employees || []).map((e) =>
      e.username === username ? { ...e, role, ...preset } : e
    );
    await persistEmployees(next);
  };

  // Single generic handler for every individual permission toggle (view all tickets,
  // add tickets, edit tickets, delete tickets, accounting/notes-only mode, manage
  // companies). Each toggle is independently switchable by the main account; coherence
  // between them (edit/delete requiring view, accounting overriding add/edit/delete) is
  // enforced afterwards by reconcilePermissions so the stored record never contradicts itself.
  const handleTogglePermission = async (username, field, checked) => {
    if (!currentUser.isAdmin) {
      setManageError("Only the main account can change employee permissions");
      return;
    }
    const next = (employees || []).map((e) =>
      e.username === username ? { ...e, ...reconcilePermissions({ ...e, [field]: checked }) } : e
    );
    await persistEmployees(next);
  };

  // Promotes an employee to a main/admin account. Any main account can promote another one.
  const handlePromoteToAdmin = async (username) => {
    if (!currentUser.isAdmin) {
      setManageError("Only the main account can grant main-account access");
      return;
    }
    const target = (employees || []).find((e) => e.username === username);
    if (!target) return;
    requestConfirm(
      `Make "${target.name}" a main account? They will be able to manage all employees, permissions, backups, and see every ticket.`,
      async () => {
        const next = (employees || []).map((e) =>
          e.username === username ? { ...e, isAdmin: true } : e
        );
        await persistEmployees(next);
        setConfirmDialog(null);
      }
    );
  };

  // Demotes a main account back to a regular employee. Blocked if it would leave zero main accounts.
  const handleDemoteAdmin = async (username) => {
    if (!currentUser.isAdmin) {
      setManageError("Only the main account can remove main-account access");
      return;
    }
    const admins = (employees || []).filter((e) => e.isAdmin);
    if (admins.length <= 1) {
      setManageError("There must always be at least one main account");
      return;
    }
    const target = (employees || []).find((e) => e.username === username);
    if (!target) return;
    requestConfirm(`Remove main-account access from "${target.name}"?`, async () => {
      const next = (employees || []).map((e) =>
        e.username === username ? { ...e, isAdmin: false } : e
      );
      await persistEmployees(next);
      // If the admin demoted themselves, drop their manage-panel view since they're no longer main
      if (username === currentUser.username) {
        setCurrentUser({ ...currentUser, isAdmin: false });
        setShowManage(false);
      }
      setConfirmDialog(null);
    });
  };

  const handleDeleteEmployee = async (username) => {
    if (!currentUser.isAdmin) {
      setManageError("Only the main account can remove employees");
      return;
    }
    if (username === currentUser.username) {
      setManageError("You can't delete the account you're logged in with");
      return;
    }
    await persistEmployees((employees || []).filter((e) => e.username !== username));
  };

  const togglePasswordVisible = (username) => {
    setVisiblePasswords((prev) => ({ ...prev, [username]: !prev[username] }));
  };

  const startEditEmployee = (emp) => {
    setManageError("");
    setEditShowPassword(false);
    setEditingUsername(emp.username);
    setEditDraft({ name: emp.name, username: emp.username, password: emp.password });
  };

  const cancelEditEmployee = () => {
    setEditingUsername(null);
    setEditDraft({ name: "", username: "", password: "" });
  };

  const saveEditEmployee = async () => {
    if (!currentUser.isAdmin) {
      setManageError("Only the main account can edit employee accounts");
      return;
    }
    setManageError("");
    const trimmedName = editDraft.name.trim();
    const trimmedUsername = editDraft.username.trim();
    if (!trimmedName || !trimmedUsername || !editDraft.password) {
      setManageError("Please fill in all fields");
      return;
    }
    const clash = (employees || []).some(
      (e) => e.username !== editingUsername && e.username === trimmedUsername
    );
    if (clash) {
      setManageError("That username is already taken by another account");
      return;
    }
    const next = (employees || []).map((e) =>
      e.username === editingUsername
        ? { ...e, name: trimmedName, username: trimmedUsername, password: editDraft.password }
        : e
    );
    await persistEmployees(next);

    // If the main account edited its own account, keep the current session in sync
    if (editingUsername === currentUser.username) {
      await window.storage.set("session:user", trimmedUsername, false);
      setCurrentUser({ ...currentUser, name: trimmedName, username: trimmedUsername });
    }
    cancelEditEmployee();
  };

  const handleChangePassword = async () => {
    setPasswordError("");
    setPasswordSuccess("");
    if (!currentPasswordInput || !newPasswordInput || !confirmPasswordInput) {
      setPasswordError("Please fill in all fields");
      return;
    }
    const me = (employees || []).find((e) => e.username === currentUser.username);
    if (!me || me.password !== currentPasswordInput) {
      setPasswordError("Current password is incorrect");
      return;
    }
    if (newPasswordInput !== confirmPasswordInput) {
      setPasswordError("New password and confirmation do not match");
      return;
    }
    if (newPasswordInput.length < 4) {
      setPasswordError("New password should be at least 4 characters");
      return;
    }
    const next = (employees || []).map((e) =>
      e.username === currentUser.username ? { ...e, password: newPasswordInput } : e
    );
    await persistEmployees(next);
    setPasswordSuccess("Password updated successfully");
    setCurrentPasswordInput("");
    setNewPasswordInput("");
    setConfirmPasswordInput("");
  };

  // ---------- Backup / restore (main account only) ----------
  const handleBackup = () => {
    if (!currentUser.isAdmin) return;
    const payload = {
      backupFormat: "flight-tickets-v1",
      exportedAt: new Date().toISOString(),
      exportedBy: currentUser.name,
      tickets,
      employees,
      suggestions,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `flight_tickets_backup_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const triggerRestore = () => {
    if (!currentUser.isAdmin) return;
    setRestoreError("");
    setRestoreSuccess("");
    fileInputRef.current && fileInputRef.current.click();
  };

  const handleRestoreFile = async (e) => {
    setRestoreError("");
    setRestoreSuccess("");
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      if (!parsed || !Array.isArray(parsed.tickets) || !Array.isArray(parsed.employees)) {
        setRestoreError("This file doesn't look like a valid backup");
        return;
      }
      // Normalize suggestions defensively so nothing from the backup is silently dropped,
      // even if the file is from an older/partial export.
      const s = parsed.suggestions || {};
      const normalizedSuggestions = {
        companies: Array.isArray(s.companies) ? s.companies : [],
        // Never restore saved customer names — this field must have no autocomplete history.
        customers: [],
        airlines: Array.isArray(s.airlines) ? s.airlines : [],
        cities: Array.isArray(s.cities) ? s.cities : [],
      };
      const suggestionsCount =
        normalizedSuggestions.companies.length +
        normalizedSuggestions.customers.length +
        normalizedSuggestions.airlines.length +
        normalizedSuggestions.cities.length;
      requestConfirm(
        "This will replace all current tickets and employee accounts with the data in this backup file. This cannot be undone. Continue?",
        async () => {
          await persistTickets(parsed.tickets);
          await persistEmployees(parsed.employees);
          await persistSuggestions(normalizedSuggestions);
          setRestoreSuccess(
            `Backup restored successfully: ${parsed.tickets.length} tickets, ${parsed.employees.length} employee accounts, and ${suggestionsCount} saved suggestions.`
          );
          setConfirmDialog(null);
        }
      );
    } catch (err) {
      setRestoreError("Could not read this backup file");
    } finally {
      e.target.value = "";
    }
  };

  // ---------- Tickets ----------
  // Builds a plain-language list of what changed between the ticket's previous version and
  // the edited one (e.g. "From: CAI → JED"), used to log every ticket edit — not just notes —
  // into the same edit-history trail, along with who made the change.
  const describeTicketChanges = (before, after) => {
    const changes = [];
    const fieldLabels = {
      company: "Company",
      supplier: "Supplier",
      from: "From",
      to: "To",
      airline: "Airline",
      date: "Date",
      netPrice: "Net price",
      soldPrice: "Sold price",
    };
    Object.keys(fieldLabels).forEach((key) => {
      const beforeVal = before[key] ?? "";
      const afterVal = after[key] ?? "";
      if (String(beforeVal) !== String(afterVal)) {
        changes.push(`${fieldLabels[key]}: ${beforeVal || "—"} → ${afterVal || "—"}`);
      }
    });

    const beforeCustomers = Array.isArray(before.customers) ? before.customers : [];
    const afterCustomers = Array.isArray(after.customers) ? after.customers : [];
    if (beforeCustomers.length !== afterCustomers.length) {
      changes.push(`Customers: ${beforeCustomers.length} → ${afterCustomers.length}`);
    }
    const maxLen = Math.max(beforeCustomers.length, afterCustomers.length);
    for (let i = 0; i < maxLen; i++) {
      const b = beforeCustomers[i] || { name: "", ticketNumber: "" };
      const a = afterCustomers[i] || { name: "", ticketNumber: "" };
      if ((b.name || "") !== (a.name || "")) {
        changes.push(`Customer ${i + 1} name: ${b.name || "—"} → ${a.name || "—"}`);
      }
      if ((b.ticketNumber || "") !== (a.ticketNumber || "")) {
        changes.push(`Customer ${i + 1} ticket number: ${b.ticketNumber || "—"} → ${a.ticketNumber || "—"}`);
      }
    }
    return changes;
  };

  const handleSubmit = () => {
    setError("");
    const customers = form.customers || [];
    const customersValid =
      customers.length > 0 && customers.every((c) => c.name.trim() && c.ticketNumber.trim());
    // A multi-destination route needs at least two filled-in stops; a regular route
    // needs both From and To.
    const cleanDestinations = (form.destinations || []).map((d) => (d || "").trim()).filter(Boolean);
    const routeValid = form.multiDestination
      ? cleanDestinations.length >= 2
      : form.from.trim() && form.to.trim();
    if (!customersValid || !routeValid || form.netPrice === "" || form.soldPrice === "") {
      setError("Please enter at least the customer name(s), ticket number(s), destinations, and prices");
      return;
    }
    // Keep the original owner when editing an existing ticket (so an admin editing someone
    // else's ticket doesn't reassign it to themselves); new tickets belong to whoever adds them.
    const isEditingExisting = !!(form.id && form.employeeUsername);
    const original = form.id ? tickets.find((t) => t.id === form.id) : null;
    let record = {
      ...form,
      customers,
      customersCount: customers.length,
      // For a multi-destination route, from/to are kept in sync as the first/last stop so
      // every place that reads a plain origin/destination (search, exports, older code)
      // keeps working; a regular route just keeps its own from/to untouched.
      destinations: form.multiDestination ? cleanDestinations : [],
      from: form.multiDestination ? cleanDestinations[0] || "" : form.from,
      to: form.multiDestination ? cleanDestinations[cleanDestinations.length - 1] || "" : form.to,
      employee: isEditingExisting ? form.employee : currentUser.name,
      employeeUsername: isEditingExisting ? form.employeeUsername : currentUser.username,
      id: form.id || Date.now().toString(),
    };
    // Every edit to an existing ticket — any field, not just notes — gets logged into the
    // same edit-history trail shown under Notes, recording what changed and who changed it.
    if (original) {
      const changes = describeTicketChanges(original, record);
      if (changes.length > 0) {
        const history = Array.isArray(original.notesHistory) ? original.notesHistory : [];
        record = {
          ...record,
          notesHistory: [
            ...history,
            { type: "edit", changes, by: currentUser.name, at: new Date().toISOString() },
          ],
        };
      }
    }
    let next;
    if (form.id) {
      next = tickets.map((t) => (t.id === form.id ? record : t));
    } else {
      next = [record, ...tickets];
    }
    persistTickets(next);
    rememberSuggestionsFromRecord(record);
    setForm(getEmptyForm());
    setSupplierOther(false);
  };

  // The main account can always edit tickets; an employee can too, but only if they've
  // been granted the "edit tickets" permission. Deleting stays main-account only either way.
  const handleEdit = (t) => {
    if (!currentUser.isAdmin && !canEditTickets) return;
    // Backward compatibility: older records stored a single customer/ticketNumber pair
    const customers =
      Array.isArray(t.customers) && t.customers.length > 0
        ? t.customers
        : [{ name: t.customer || "", ticketNumber: t.ticketNumber || "" }];
    // Backward compatibility: older records have no multiDestination/destinations fields.
    const destinations =
      Array.isArray(t.destinations) && t.destinations.length >= 2 ? t.destinations : [t.from || "", t.to || ""];
    setForm({ ...t, customers, customersCount: customers.length, multiDestination: !!t.multiDestination, destinations });
    setSupplierOther(!!t.supplier && !SUPPLIERS.includes(t.supplier));
  };
  const handleDelete = (id) => {
    if (!currentUser.isAdmin && !canDeleteTickets) {
      setError("You don't have permission to delete tickets");
      return;
    }
    if (form.id === id) { setForm(getEmptyForm()); setSupplierOther(false); }
    persistTickets(tickets.filter((t) => t.id !== id));
  };
  const handleCancel = () => { setForm(getEmptyForm()); setSupplierOther(false); };

  // Opens the full-detail view ("page") for a ticket, showing every field including notes.
  const openTicketDetail = (t) => {
    setViewingTicketId(t.id);
    setNotesDraft(t.notes || "");
    setNotesSaved(false);
    setShowRefundForm(false);
    setRefundDraft({
      airlineAmount: t.refund ? t.refund.airlineAmount : "",
      customerAmount: t.refund ? t.refund.customerAmount : "",
      customerIndex: t.refund && t.refund.customerIndex != null ? t.refund.customerIndex : 0,
    });
    setRefundSaved(false);
  };
  const closeTicketDetail = () => {
    setViewingTicketId(null);
    setNotesDraft("");
    setNotesSaved(false);
    setShowRefundForm(false);
    setRefundDraft({ airlineAmount: "", customerAmount: "", customerIndex: 0 });
    setRefundSaved(false);
  };
  // Saves an edit to just the notes field of a ticket, without touching anything else.
  // Every save appends an entry to notesHistory recording who made the change and when,
  // so the full edit trail (including accounting-account edits) stays visible.
  const saveTicketNotes = (id) => {
    const now = new Date().toISOString();
    const nextNotes = notesDraft.toUpperCase();
    const next = tickets.map((t) => {
      if (t.id !== id) return t;
      const history = Array.isArray(t.notesHistory) ? t.notesHistory : [];
      return {
        ...t,
        notes: nextNotes,
        notesHistory: [...history, { value: nextNotes, by: currentUser.name, at: now }],
      };
    });
    persistTickets(next);
    setNotesSaved(true);
  };

  // True once a refund (either side) has actually been recorded for a ticket.
  const hasRefund = (t) => !!(t && t.refund && (t.refund.airlineAmount !== "" || t.refund.customerAmount !== ""));

  // Accounting-adjusted figures for a ticket: a recorded refund is deducted from both
  // sides — what the airline paid back reduces our cost (net price), and what we paid
  // back to the customer reduces our revenue (sold price) — so sales/profit totals
  // everywhere (ticket rows, summary cards, monthly/company breakdowns, exports)
  // reflect the refund rather than the original pre-refund booking amounts.
  const netAfterRefund = (t) =>
    (parseFloat(t.netPrice) || 0) - (hasRefund(t) ? parseFloat(t.refund.airlineAmount) || 0 : 0);
  const soldAfterRefund = (t) =>
    (parseFloat(t.soldPrice) || 0) - (hasRefund(t) ? parseFloat(t.refund.customerAmount) || 0 : 0);
  const profitAfterRefund = (t) => soldAfterRefund(t) - netAfterRefund(t);

  // Saves the two refund amounts (from the airline, to the customer) onto a ticket. Kept
  // as its own record — separate from editing the ticket itself — with its own history
  // trail, and shows up as its own row directly under the original ticket in exports.
  const saveTicketRefund = (id) => {
    const now = new Date().toISOString();
    const refund = {
      airlineAmount: refundDraft.airlineAmount,
      customerAmount: refundDraft.customerAmount,
      customerIndex: refundDraft.customerIndex || 0,
      date: todayDateStr(),
    };
    const next = tickets.map((t) => {
      if (t.id !== id) return t;
      const history = Array.isArray(t.refundHistory) ? t.refundHistory : [];
      return {
        ...t,
        refund,
        refundHistory: [...history, { ...refund, by: currentUser.name, at: now }],
      };
    });
    persistTickets(next);
    setRefundSaved(true);
    setShowRefundForm(false);
  };

  // Removes a previously-recorded refund from a ticket (e.g. entered by mistake), keeping
  // a "cleared" entry in the history trail for the audit log.
  const clearTicketRefund = (id) => {
    const now = new Date().toISOString();
    const next = tickets.map((t) => {
      if (t.id !== id) return t;
      const history = Array.isArray(t.refundHistory) ? t.refundHistory : [];
      return {
        ...t,
        refund: null,
        refundHistory: [...history, { cleared: true, by: currentUser.name, at: now }],
      };
    });
    persistTickets(next);
    setRefundDraft({ airlineAmount: "", customerAmount: "", customerIndex: 0 });
    setRefundSaved(false);
    setShowRefundForm(false);
  };

  const handleCustomersCountChange = (value) => {
    const count = value === "" ? "" : value;
    const customers = resizeCustomers(form.customers, value);
    // When more customer rows are added, auto-sequence their ticket numbers by
    // increasing the previous customer's number by one (only if it was filled in).
    for (let i = form.customers.length; i < customers.length; i++) {
      const generated = nextTicketNumber(customers[i - 1] && customers[i - 1].ticketNumber);
      if (generated) customers[i] = { ...customers[i], ticketNumber: generated };
    }
    setForm({ ...form, customersCount: count, customers });
  };

  // From/To suggestions are shown as "CODE - City, Country" for easy searching, but only
  // the 3-letter IATA code should end up stored in the field/cell. If the typed or picked
  // value matches that "CODE - ..." shape, keep just the code; otherwise keep it as typed.
  const handleCityChange = (field, value) => {
    const raw = (value || "").toUpperCase();
    const match = raw.match(/^([A-Z]{3})\s*-\s*.+$/);
    setForm({ ...form, [field]: match ? match[1] : raw });
  };

  // Same "CODE - City, Country" → CODE cleanup as handleCityChange, but for one stop
  // in a multi-destination (multi-city) route.
  const handleDestinationChange = (index, value) => {
    const raw = (value || "").toUpperCase();
    const match = raw.match(/^([A-Z]{3})\s*-\s*.+$/);
    const clean = match ? match[1] : raw;
    const destinations = form.destinations.map((d, i) => (i === index ? clean : d));
    setForm({ ...form, destinations });
  };

  const addDestinationStop = () => {
    setForm({ ...form, destinations: [...form.destinations, ""] });
  };

  // Always keeps at least two stops (a route needs a start and an end).
  const removeDestinationStop = (index) => {
    const destinations = form.destinations.filter((_, i) => i !== index);
    setForm({ ...form, destinations: destinations.length >= 2 ? destinations : ["", ""] });
  };

  const handleAirlineChange = (value) => {
    const airline = value.toUpperCase();
    const code = getAirlineCodeByIata(airline);
    // If we recognize the airline code, pre-fill its 3-digit prefix into any customer's
    // ticket number that hasn't been typed into yet (never overwrites manual entries).
    const customers = code
      ? form.customers.map((c) => (c.ticketNumber ? c : { ...c, ticketNumber: `${code}-` }))
      : form.customers;
    setForm({ ...form, airline, customers });
  };

  const handleCustomerFieldChange = (index, field, value) => {
    let nextValue = (value || "").toUpperCase();
    if (field === "ticketNumber") {
      // Keep only letters and digits, then auto-insert a hyphen after the first 3 characters
      const clean = nextValue.replace(/[^A-Z0-9]/g, "").slice(0, 13);
      nextValue = clean.length > 3 ? `${clean.slice(0, 3)}-${clean.slice(3)}` : clean;
    }
    const customers = form.customers.map((c, i) => (i === index ? { ...c, [field]: nextValue } : c));
    let airline = form.airline;
    if (field === "ticketNumber") {
      // Auto-detect the airline from the ticket number's 3-digit prefix (only if the
      // airline field hasn't been filled in yet, so it never overrides a manual choice)
      if (!airline) {
        const match = nextValue.match(/^([A-Z0-9]{3})-/);
        if (match) {
          const detected = getAirlineByCode(match[1]);
          if (detected) airline = detected;
        }
      }
    }
    setForm({ ...form, customers, airline });
  };

  // Runs once the person leaves the ticket number field (not on every keystroke), using
  // whatever they finished typing, and auto-fills any following ticket numbers that are
  // still empty — each one increasing the previous by one. Stops at the first one someone
  // has already typed something into, so manual entries are never overwritten.
  const handleTicketNumberBlur = (index) => {
    const customers = form.customers.map((c) => ({ ...c }));
    let last = customers[index] && customers[index].ticketNumber;
    if (!last) return;
    for (let i = index + 1; i < customers.length; i++) {
      if (customers[i].ticketNumber) break;
      const generated = nextTicketNumber(last);
      if (!generated) break;
      customers[i] = { ...customers[i], ticketNumber: generated };
      last = generated;
    }
    setForm({ ...form, customers });
  };

  // Finds a saved ticket by ticket number, searching every customer row across all
  // saved tickets (old or current schema). Used when a reissued ticket references an
  // older one, both to auto-fill its issue date and to import the rest of its data.
  const findTicketByNumber = (ticketNumber) => {
    const target = (ticketNumber || "").trim().toUpperCase();
    if (!target) return null;
    for (const t of tickets) {
      const custs =
        Array.isArray(t.customers) && t.customers.length > 0
          ? t.customers
          : [{ name: t.customer || "", ticketNumber: t.ticketNumber || "" }];
      if (custs.some((c) => (c.ticketNumber || "").trim().toUpperCase() === target)) {
        return t;
      }
    }
    return null;
  };

  // Cleans up the old ticket number the same way regular ticket numbers are formatted.
  const handleOldTicketNumberChange = (value) => {
    const clean = (value || "").toUpperCase().replace(/[^A-Z0-9]/g, "").slice(0, 13);
    const nextValue = clean.length > 3 ? `${clean.slice(0, 3)}-${clean.slice(3)}` : clean;
    setForm({ ...form, oldTicketNumber: nextValue, oldTicketIssueDate: "" });
  };

  // Once the person finishes typing the old ticket number, look it up against saved
  // tickets and import that old ticket's data into the reissue form: issue date, company,
  // supplier, route (including a multi-destination route), airline, prices, and any
  // customer name not already typed. Anything the person already entered by hand is left
  // untouched — this only fills in fields that are still empty.
  const handleOldTicketNumberBlur = () => {
    const oldTicket = findTicketByNumber(form.oldTicketNumber);
    if (!oldTicket) {
      setForm({ ...form, oldTicketIssueDate: "" });
      return;
    }
    const oldCustomers =
      Array.isArray(oldTicket.customers) && oldTicket.customers.length > 0
        ? oldTicket.customers
        : [{ name: oldTicket.customer || "", ticketNumber: oldTicket.ticketNumber || "" }];
    // Fill in any customer row that doesn't have a name yet with the matching old
    // customer's name (by position); new ticket numbers are always left exactly as typed.
    const customers = form.customers.map((c, i) =>
      c.name.trim() ? c : { ...c, name: (oldCustomers[i] && oldCustomers[i].name) || c.name }
    );
    const hasOwnDestinations = (form.destinations || []).some((d) => (d || "").trim());
    const oldSupplier = oldTicket.supplier || "";
    if (!form.supplier && oldSupplier && !SUPPLIERS.includes(oldSupplier)) setSupplierOther(true);
    setForm({
      ...form,
      oldTicketIssueDate: oldTicket.date || "",
      company: form.company || oldTicket.company || "",
      supplier: form.supplier || oldSupplier,
      from: form.from || oldTicket.from || "",
      to: form.to || oldTicket.to || "",
      multiDestination: form.multiDestination || !!oldTicket.multiDestination,
      destinations: hasOwnDestinations
        ? form.destinations
        : Array.isArray(oldTicket.destinations) && oldTicket.destinations.length >= 2
        ? oldTicket.destinations
        : form.destinations,
      airline: form.airline || oldTicket.airline || "",
      netPrice: form.netPrice !== "" ? form.netPrice : oldTicket.netPrice ?? "",
      soldPrice: form.soldPrice !== "" ? form.soldPrice : oldTicket.soldPrice ?? "",
      customers,
    });
  };

  // The main account always sees everything; employees see only what they entered,
  // unless the main account has granted them permission to view all tickets — or granted
  // them permission to edit tickets, since editing every ticket requires seeing every ticket.
  // Guarded against currentUser being null (e.g. on the login/setup screens).
  const currentEmployeeRecord = currentUser
    ? (employees || []).find((e) => e.username === currentUser.username)
    : null;
  const canViewAllTickets =
    !!currentUser &&
    (currentUser.isAdmin ||
      !!(
        currentEmployeeRecord &&
        (currentEmployeeRecord.canViewAll ||
          currentEmployeeRecord.isAccounting ||
          currentEmployeeRecord.canEdit ||
          currentEmployeeRecord.canDelete)
      ));
  // Accounting accounts can see everything but cannot add tickets — their only allowed
  // edit anywhere in the app is the Notes field on a ticket's detail page.
  const isAccountingUser =
    !!currentUser && !currentUser.isAdmin && !!(currentEmployeeRecord && currentEmployeeRecord.isAccounting);
  // Every employee can add new tickets — this is no longer an individually
  // switchable permission. Accounting accounts are the one exception: their only
  // allowed edit anywhere in the app is the Notes field.
  const canAddTickets =
    !!currentUser &&
    (currentUser.isAdmin ||
      !!(currentEmployeeRecord && !currentEmployeeRecord.isAccounting));
  // A non-admin employee can be granted permission to edit tickets (within whatever
  // set of tickets they can already see). Accounting accounts are excluded even if the
  // flag is set — their only allowed edit is the Notes field, never the ticket itself.
  const canEditTickets =
    !!currentUser &&
    (currentUser.isAdmin ||
      !!(currentEmployeeRecord && currentEmployeeRecord.canEdit && !currentEmployeeRecord.isAccounting));
  // A separate, independently grantable permission: whether this employee can delete
  // tickets. Previously this was main-account only; now the main account can hand it
  // to specific employees (e.g. a Manager) without giving them full main-account access.
  const canDeleteTickets =
    !!currentUser &&
    (currentUser.isAdmin ||
      !!(currentEmployeeRecord && currentEmployeeRecord.canDelete && !currentEmployeeRecord.isAccounting));
  // A separate permission axis from ticket access: whether this account can add/edit/
  // remove saved company records (name, tax number, commercial register, phone numbers).
  const canManageCompanies =
    !!currentUser &&
    (currentUser.isAdmin || !!(currentEmployeeRecord && currentEmployeeRecord.canManageCompanies));
  const visibleTickets = !currentUser
    ? []
    : canViewAllTickets
    ? tickets
    : tickets.filter((t) =>
        t.employeeUsername ? t.employeeUsername === currentUser.username : t.employee === currentUser.name
      );

  // Hotels reuse the same view/add/edit/delete permission axis as flight tickets.
  const visibleHotelBookings = !currentUser
    ? []
    : canViewAllTickets
    ? hotelBookings
    : hotelBookings.filter((h) =>
        h.employeeUsername ? h.employeeUsername === currentUser.username : h.employee === currentUser.name
      );

  // Number of nights a single date range covers, from check-in to check-out (at least 1).
  const nightsBetween = (checkIn, checkOut) => {
    if (!checkIn || !checkOut) return 1;
    const inD = new Date(checkIn);
    const outD = new Date(checkOut);
    const diffDays = Math.round((outD - inD) / (1000 * 60 * 60 * 24));
    return diffDays > 0 ? diffDays : 1;
  };
  // Nights for one room line. Falls back to the booking's own (legacy) check-in/check-out
  // if the line itself doesn't have dates — older bookings saved before dates lived on
  // each room line.
  const roomLineNights = (l, h) => nightsBetween(l.checkIn || (h && h.checkIn), l.checkOut || (h && h.checkOut));
  // The overall date range shown for a booking: earliest check-in to latest check-out
  // across all its room lines.
  const hotelDateRange = (h) => {
    const lines = h.roomLines || [];
    const checkIns = lines.map((l) => l.checkIn || h.checkIn).filter(Boolean);
    const checkOuts = lines.map((l) => l.checkOut || h.checkOut).filter(Boolean);
    if (checkIns.length === 0 || checkOuts.length === 0) return { start: "", end: "" };
    return {
      start: checkIns.reduce((a, b) => (a < b ? a : b)),
      end: checkOuts.reduce((a, b) => (a > b ? a : b)),
    };
  };

  // Converts an amount from a room line's own currency into EGP, using the entered
  // USD->EGP rate. Returns the amount unchanged for EGP-priced lines.
  const hotelInEgp = (amount, currency) => (currency === "USD" ? amount * (usdToEgpRate || 0) : amount);

  // Per-booking totals: each room line's net/sold price is multiplied by its own room
  // count and its own number of nights, then summed across every line (e.g. 1 single
  // + 2 doubles, each possibly with different dates and currencies, all converted into
  // EGP to total).
  const hotelRoomCount = (h) => (h.roomLines || []).reduce((sum, l) => sum + (parseInt(l.count, 10) || 0), 0);
  // Raw (un-converted, in the line's own currency) total for one line — used when showing
  // a line's own subtotal next to its own currency in the form.
  const hotelLineNetTotal = (l, nights) => (parseFloat(l.netPrice) || 0) * (parseInt(l.count, 10) || 0) * nights;
  const hotelLineSoldTotal = (l, nights) => (parseFloat(l.soldPrice) || 0) * (parseInt(l.count, 10) || 0) * nights;
  const hotelNetTotal = (h) =>
    (h.roomLines || []).reduce((sum, l) => sum + hotelInEgp(hotelLineNetTotal(l, roomLineNights(l, h)), l.currency), 0);
  const hotelSoldTotal = (h) =>
    (h.roomLines || []).reduce((sum, l) => sum + hotelInEgp(hotelLineSoldTotal(l, roomLineNights(l, h)), l.currency), 0);
  const hotelProfitTotal = (h) => hotelSoldTotal(h) - hotelNetTotal(h);
  // A booking is Corporate when a company name was entered; otherwise it's an
  // Individual booking automatically — no separate toggle needed.
  const hotelBookingType = (h) => (h.customer && h.customer.trim() ? "Corporate" : "Individual");
  // A short readable summary of a booking's room lines, e.g. "1x Single (BB, EGP, 01-AUG-2026→05-AUG-2026), 2x Double (AI, USD, 01-AUG-2026→03-AUG-2026)".
  const hotelLinesSummary = (h) =>
    (h.roomLines || [])
      .map((l) => {
        const type = ROOM_TYPES.find((r) => r.value === l.roomType)?.label || l.roomType;
        const meal = MEAL_PLANS.find((m) => m.value === l.mealPlan)?.value.toUpperCase() || "";
        const checkIn = l.checkIn || h.checkIn;
        const checkOut = l.checkOut || h.checkOut;
        const dates = checkIn && checkOut ? `, ${formatDisplayDate(checkIn)}→${formatDisplayDate(checkOut)}` : "";
        return `${l.count}× ${type} (${meal}, ${l.currency}${dates})`;
      })
      .join(", ");

  const hotelTotals = visibleHotelBookings.reduce(
    (acc, h) => {
      acc.net += hotelNetTotal(h);
      acc.sold += hotelSoldTotal(h);
      acc.profit += hotelProfitTotal(h);
      return acc;
    },
    { net: 0, sold: 0, profit: 0 }
  );


  const getCustomers = (t) =>
    Array.isArray(t.customers) && t.customers.length > 0
      ? t.customers
      : [{ name: t.customer || "", ticketNumber: t.ticketNumber || "" }];

  const monthsAvailable = Array.from(new Set(visibleTickets.map((t) => monthKey(t.date)))).sort((a, b) =>
    b.localeCompare(a)
  );

  const yearsAvailable = Array.from(
    new Set(
      visibleTickets
        .map((t) => (t.date ? t.date.slice(0, 4) : ""))
        .filter(Boolean)
    )
  ).sort((a, b) => b.localeCompare(a));

  const companiesAvailable = Array.from(
    new Set(visibleTickets.map((t) => (t.company || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));

  const employeesAvailable = Array.from(
    new Set(visibleTickets.map((t) => (t.employee || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));

  const suppliersAvailable = Array.from(
    new Set(visibleTickets.map((t) => (t.supplier || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));

  const byMonth = selectedMonth
    ? visibleTickets.filter((t) => monthKey(t.date) === selectedMonth)
    : visibleTickets;

  const byYear = selectedYear
    ? byMonth.filter((t) => (t.date || "").slice(0, 4) === selectedYear)
    : byMonth;

  const byCompany = selectedCompany
    ? byYear.filter((t) => (t.company || "").trim() === selectedCompany)
    : byYear;

  const byEmployee = selectedEmployee
    ? byCompany.filter((t) => (t.employee || "").trim() === selectedEmployee)
    : byCompany;

  const bySupplier = selectedSupplier
    ? byEmployee.filter((t) => (t.supplier || "").trim() === selectedSupplier)
    : byEmployee;

  const filtered = bySupplier.filter((t) => {
    const q = query.trim().toLowerCase();
    if (!q) return true;
    const customers = getCustomers(t);
    return (
      (t.employee || "").toLowerCase().includes(q) ||
      (t.company || "").toLowerCase().includes(q) ||
      (t.from || "").toLowerCase().includes(q) ||
      (t.to || "").toLowerCase().includes(q) ||
      (Array.isArray(t.destinations) ? t.destinations.join(" ") : "").toLowerCase().includes(q) ||
      (t.airline || "").toLowerCase().includes(q) ||
      customers.some(
        (c) =>
          (c.name || "").toLowerCase().includes(q) ||
          (c.ticketNumber || "").toLowerCase().includes(q)
      )
    );
  });

  // Sort tickets by issue date (most recent first). Tickets with no date are pushed
  // to the end instead of being sorted arbitrarily.
  const sortedFiltered = [...filtered].sort((a, b) => {
    if (!a.date && !b.date) return 0;
    if (!a.date) return 1;
    if (!b.date) return -1;
    return b.date.localeCompare(a.date);
  });

  // The ticket currently open in the detail "page", if any.
  const viewingTicket = viewingTicketId ? visibleTickets.find((t) => t.id === viewingTicketId) : null;

  // Counts and sums per CUSTOMER rather than per ticket/booking: a booking with several
  // customers contributes its full (unsplit) total/profit once for each customer, and
  // each customer counts as one ticket. This keeps the summary cards, monthly totals,
  // and company breakdown consistent with the per-customer rows shown in the ticket table.
  // A recorded refund is a real, one-time amount for the booking, so it's deducted once
  // (not multiplied by customer count) — reducing sales by what went back to the customer
  // and adjusting profit by that same amount net of whatever the airline refunded back.
  const countAndSum = (rows) =>
    rows.reduce(
      (acc, t) => {
        const n = getCustomers(t).length || 1;
        const refundCustomerAmt = hasRefund(t) ? parseFloat(t.refund.customerAmount) || 0 : 0;
        const refundAirlineAmt = hasRefund(t) ? parseFloat(t.refund.airlineAmount) || 0 : 0;
        acc.count += n;
        acc.total += (parseFloat(t.soldPrice) || 0) * n - refundCustomerAmt;
        acc.profit += profit(t.netPrice, t.soldPrice) * n + refundAirlineAmt - refundCustomerAmt;
        return acc;
      },
      { count: 0, total: 0, profit: 0 }
    );

  const totals = countAndSum(bySupplier);

  const monthlyBreakdown = monthsAvailable.map((key) => {
    const rows = visibleTickets.filter((t) => monthKey(t.date) === key);
    return { key, ...countAndSum(rows) };
  });

  const companyBreakdown = companiesAvailable.map((name) => {
    const rows = visibleTickets.filter((t) => (t.company || "").trim() === name);
    const customers = Array.from(
      new Set(rows.flatMap((t) => getCustomers(t).map((c) => c.name)).filter(Boolean))
    ).sort((a, b) => a.localeCompare(b));
    return { name, customers, ...countAndSum(rows) };
  });

  // Ticket-level status text used in the exported "Status" column, replacing the old
  // per-customer numbering. "Reissued" applies to the whole booking (shown on the first
  // customer's row); "Refunded" applies only to the specific customer the refund was
  // recorded against, since a multi-customer booking may have just one refunded ticket.
  const ticketStatus = (t, i) => {
    const parts = [];
    if (i === 0 && t.isReissued) parts.push("Reissued");
    if (hasRefund(t) && (t.refund.customerIndex || 0) === i) parts.push("Refunded");
    return parts.join(" & ");
  };

  // Builds the per-customer row list for one ticket set, sorted by issue date
  // (earliest first; undated tickets pushed to the end). Tickets issued on the
  // SAME date are then ordered by ticket number ascending (numeric-aware, so
  // "077-1234567890" sorts before "077-1234567900" correctly).
  const ticketRows = (rows) => {
    const firstTicketNumber = (t) => (getCustomers(t)[0] && getCustomers(t)[0].ticketNumber) || "";
    const sorted = [...rows].sort((a, b) => {
      if (!a.date && !b.date) {
        return firstTicketNumber(a).localeCompare(firstTicketNumber(b), undefined, { numeric: true, sensitivity: "base" });
      }
      if (!a.date) return 1;
      if (!b.date) return -1;
      const dateCompare = a.date.localeCompare(b.date);
      if (dateCompare !== 0) return dateCompare;
      return firstTicketNumber(a).localeCompare(firstTicketNumber(b), undefined, { numeric: true, sensitivity: "base" });
    });
    return sorted.flatMap((t) => {
      const customers = getCustomers(t);
      const rows = customers.map((c, i) => ({
        "Type": "Ticket",
        "Employee": t.employee || "",
        "Company": t.company || "",
        "Supplier": t.supplier || "",
        "Status": ticketStatus(t, i),
        "Customer": c.name || "",
        "Ticket number": c.ticketNumber || "",
        "From": t.from,
        "To": t.to,
        "Route": routeLabel(t),
        "Airline": t.airline || "",
        "Issue date": t.date ? formatDisplayDate(t.date) : "",
        // Net/sold price and profit are the ORIGINAL booking amounts, shown once on the
        // first customer's row; if the ticket was refunded, the refund amounts appear in
        // their own row directly underneath (see below) rather than being blended in here.
        "Net price": i === 0 ? parseFloat(t.netPrice) || 0 : "",
        "Sold price": i === 0 ? parseFloat(t.soldPrice) || 0 : "",
        "Profit": i === 0 ? profit(t.netPrice, t.soldPrice) : "",
        "Refund (airline)": "",
        "Refund (customer)": "",
        "Notes": t.notes || "",
      }));
      // A refund, if one's been recorded, gets its own row directly under the original
      // ticket's row(s) — its own "Type", with the two refund amounts in their own columns,
      // and the specific customer/ticket number it was recorded against.
      if (hasRefund(t)) {
        const refundedCustomer = customers[t.refund.customerIndex || 0] || customers[0];
        rows.push({
          "Type": "Refund",
          "Employee": t.employee || "",
          "Company": t.company || "",
          "Supplier": "",
          "Status": "Refunded",
          "Customer": refundedCustomer ? refundedCustomer.name || "" : "",
          "Ticket number": `Refund — ${(refundedCustomer && refundedCustomer.ticketNumber) || firstTicketNumber(t) || "ticket"}`,
          "From": "",
          "To": "",
          "Route": "",
          "Airline": "",
          "Issue date": t.refund.date ? formatDisplayDate(t.refund.date) : "",
          "Net price": "",
          "Sold price": "",
          "Profit": "",
          "Refund (airline)": parseFloat(t.refund.airlineAmount) || 0,
          "Refund (customer)": parseFloat(t.refund.customerAmount) || 0,
          "Notes": "",
        });
      }
      return rows;
    });
  };

  // Sums net price / sold price / profit across a raw ticket list (once per booking,
  // matching how those columns are only populated on each booking's first row above),
  // with any recorded refund deducted so the totals row matches actual accounting.
  const sumTicketPrices = (rows) =>
    rows.reduce(
      (acc, t) => {
        acc.net += netAfterRefund(t);
        acc.sold += soldAfterRefund(t);
        acc.profit += profitAfterRefund(t);
        return acc;
      },
      { net: 0, sold: 0, profit: 0 }
    );

  // Appends a totals row (net price / sold price / profit) to the end of a sheet's rows.
  const rowsWithTotals = (rows) => {
    const sums = sumTicketPrices(rows);
    return [
      ...ticketRows(rows),
      {
        "Employee": "", "Company": "", "Supplier": "", "Status": "", "Customer": "",
        "Ticket number": "", "From": "", "To": "", "Airline": "", "Issue date": "TOTAL",
        "Net price": Math.round(sums.net * 100) / 100,
        "Sold price": Math.round(sums.sold * 100) / 100,
        "Profit": Math.round(sums.profit * 100) / 100,
        "Notes": "",
      },
    ];
  };

  const exportMonth = (key) => {
    const rows = visibleTickets.filter((t) => monthKey(t.date) === key);
    const ws = XLSX.utils.json_to_sheet(rowsWithTotals(rows));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Details");
    XLSX.writeFile(wb, `tickets_${key}.xlsx`);
  };

  const exportAllMonths = () => {
    const wb = XLSX.utils.book_new();
    const summarySheet = XLSX.utils.json_to_sheet(
      monthlyBreakdown.map((m) => ({
        "Month": monthLabel(m.key),
        "Tickets": m.count,
        "Total sales": Math.round(m.total * 100) / 100,
        "Total profit": Math.round(m.profit * 100) / 100,
      }))
    );
    XLSX.utils.book_append_sheet(wb, summarySheet, "Monthly totals");

    monthlyBreakdown.forEach((m) => {
      const rows = visibleTickets.filter((t) => monthKey(t.date) === m.key);
      const ws = XLSX.utils.json_to_sheet(rowsWithTotals(rows));
      const safeName = m.key.replace(/[:\\\/\?\*\[\]]/g, "-").slice(0, 31);
      XLSX.utils.book_append_sheet(wb, ws, safeName);
    });

    XLSX.writeFile(wb, "monthly_ticket_totals.xlsx");
  };

  // Exports exactly the tickets matching the currently selected month / year / company /
  // employee / supplier filters AND the search box (any combination) — the same set of
  // tickets currently shown on screen — sorted by issue date (same-day tickets ordered
  // by ticket number ascending), as a single sheet ending with a totals row.
  const hasActiveFilter = !!(selectedMonth || selectedYear || selectedCompany || selectedEmployee || selectedSupplier || query.trim());
  const exportFiltered = () => {
    const ws = XLSX.utils.json_to_sheet(rowsWithTotals(filtered));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Details");
    const parts = [
      selectedYear,
      selectedMonth,
      selectedCompany,
      selectedEmployee,
      selectedSupplier,
    ]
      .filter(Boolean)
      .map((p) => p.replace(/[^a-zA-Z0-9-]+/g, "_"));
    XLSX.writeFile(wb, `tickets_${parts.length ? parts.join("_") : "filtered"}.xlsx`);
  };

  const fmt = (n) => new Intl.NumberFormat("en-US", { maximumFractionDigits: 2 }).format(n);

  // ---------- Render: loading ----------
  if (loading || setupComplete === null) {
    return (
      <div className="w-full min-h-screen bg-gradient-to-br from-teal-50 via-stone-50 to-white flex items-center justify-center">
        <p className="text-teal-800/60 text-sm flex items-center gap-2">
          <Plane size={16} className="rotate-45 animate-pulse" /> Loading...
        </p>
      </div>
    );
  }

  // ---------- Render: first-run setup (only ever shown once, before any account exists) ----------
  if (employees && employees.length === 0 && !setupComplete) {
    return (
      <div className="w-full min-h-screen bg-gradient-to-br from-teal-50 via-stone-50 to-white flex items-center justify-center p-4">
        <div className="bg-white rounded-2xl border border-stone-200 p-6 w-full max-w-sm shadow-xl shadow-teal-900/5">
          <div className="flex items-center gap-2 mb-1">
            <div className="bg-teal-800/10 text-teal-800 rounded-xl p-1.5">
              <Lock size={16} />
            </div>
            <h1 className="font-bold text-stone-900">Create the admin account</h1>
          </div>
          <p className="text-xs text-stone-500 mb-4">
            No employees exist yet. Create the first account — it will be the main account, and only it will be able to add or remove other employees.
          </p>
          {loginError && <div className="bg-red-50 text-red-700 text-sm rounded-xl px-3 py-2 mb-3">{loginError}</div>}
          <div className="space-y-3">
            <div>
              <label className="text-xs text-stone-500 block mb-1">Full name</label>
              <input className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={setupName} onChange={(e) => setSetupName(e.target.value)} placeholder="e.g. Sara Ahmed" />
            </div>
            <div>
              <label className="text-xs text-stone-500 block mb-1">Username</label>
              <input className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={setupUsername} onChange={(e) => setSetupUsername(e.target.value)} placeholder="sara" />
            </div>
            <div>
              <label className="text-xs text-stone-500 block mb-1">Password</label>
              <input type="password" className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={setupPassword} onChange={(e) => setSetupPassword(e.target.value)} placeholder="••••••" />
            </div>
          </div>
          <button onClick={handleCreateFirstAdmin}
            className="w-full mt-4 bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10 transition-colors">
            Create account and continue
          </button>
          <p className="text-xs text-stone-400 mt-4">
            Note: this is a simple access gate stored with the app's data, not a secure authentication system — anyone with technical access to the app's data can read stored passwords. Don't reuse an important password here.
          </p>
        </div>
      </div>
    );
  }

  // ---------- Render: accounts missing after setup was already completed ----------
  // Setup has already happened once before, but no employee accounts exist right now
  // (e.g. all accounts were removed, or a restore emptied them). We deliberately do NOT
  // fall back to the unauthenticated first-run setup screen here, since that would let
  // anyone create a brand-new admin account without any credentials.
  if (employees && employees.length === 0 && setupComplete) {
    return (
      <div className="w-full min-h-screen bg-gradient-to-br from-teal-50 via-stone-50 to-white flex items-center justify-center p-4">
        <div className="bg-white rounded-2xl border border-stone-200 p-6 w-full max-w-sm text-center shadow-xl shadow-teal-900/5">
          <Lock size={22} className="text-stone-400 mx-auto mb-2" />
          <h1 className="font-bold text-stone-900 mb-1">No accounts available</h1>
          <p className="text-xs text-stone-500">
            This app was already set up before, but no employee accounts currently exist. Restore a backup that contains employee accounts, or contact whoever manages this app.
          </p>
        </div>
      </div>
    );
  }

  // ---------- Render: login screen ----------
  if (!currentUser) {
    return (
      <div className="w-full min-h-screen relative flex items-center justify-center p-4 overflow-hidden bg-gradient-to-br from-teal-900 via-teal-800 to-[#0d3b3e]" style={{ fontFamily: "'Inter', sans-serif" }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:ital,opsz,wght@0,9..144,500;0,9..144,600;1,9..144,500&family=Inter:wght@400;500;600;700&display=swap');`}</style>
        {/* Decorative sky + route backdrop */}
        <div className="pointer-events-none absolute inset-0 overflow-hidden">
          <div
            className="absolute inset-0 opacity-[0.15]"
            style={{ backgroundImage: "radial-gradient(circle, #ffffff 1px, transparent 1px)", backgroundSize: "26px 26px" }}
          />
          <div className="absolute -top-24 -left-24 w-72 h-72 bg-teal-400/25 rounded-full blur-3xl" />
          <div className="absolute -bottom-24 -right-16 w-80 h-80 bg-amber-500/20 rounded-full blur-3xl" />
          <Compass size={220} className="absolute -bottom-14 -right-14 text-white/[0.04] rotate-12" />
          <Anchor size={120} className="absolute top-[6%] -left-8 text-white/[0.05] -rotate-12" />
          <Cloud size={70} className="absolute top-[12%] left-[10%] text-white/20" />
          <Cloud size={46} className="absolute top-[22%] right-[14%] text-white/15" />
          <Cloud size={54} className="absolute bottom-[18%] left-[16%] text-white/10" />
          {/* Dashed flight path with a plane at the tip */}
          <svg className="absolute top-[8%] left-[8%] w-[84%] h-40 opacity-60" viewBox="0 0 600 140" fill="none">
            <path d="M10 120 C 160 20, 380 20, 560 60" stroke="#C9973B" strokeWidth="2" strokeDasharray="6 8" strokeLinecap="round" />
            <circle cx="10" cy="120" r="4" fill="#C9973B" />
          </svg>
          <Plane size={26} className="absolute top-[15%] right-[10%] text-white/70 rotate-45 animate-pulse" />
        </div>

        <div className="relative w-full max-w-sm">
          {/* Eyebrow route strip */}
          <div className="flex items-center justify-center gap-2 mb-4 text-amber-300/90 text-[11px] font-semibold tracking-[0.2em] uppercase">
            <Sparkles size={12} />
            Perla Di Mare Travel
            <Sparkles size={12} />
          </div>

          {/* Boarding-pass card */}
          <div className="relative bg-white rounded-3xl shadow-2xl shadow-black/30 overflow-hidden">
            {/* Branded stub */}
            <div className="relative bg-gradient-to-r from-teal-800 to-teal-900 px-6 pt-9 pb-8 text-center overflow-hidden">
              <Plane size={90} className="absolute -bottom-4 -left-6 text-white/10 rotate-12" />
              <MapPin size={54} className="absolute top-3 right-3 text-white/10" />
              <div className="relative w-full mx-auto rounded-2xl bg-white shadow-lg flex items-center justify-center mb-3 p-4">
                <img src={LOGO_DATA_URL} alt="Perla Di Mare" className="w-full h-auto object-contain" />
              </div>
              <h1 className="relative text-white font-semibold text-lg tracking-tight" style={{ fontFamily: "'Fraunces', serif" }}>Flight Ticket Manager</h1>
              <p className="relative text-teal-200/70 text-[11px] mt-0.5">By Fady Habib</p>
              <p className="relative text-teal-50/90 text-xs mt-1">Sign in to manage tickets, sales &amp; bookings</p>

              {/* Route code, like a boarding pass stub */}
              <div className="relative mt-4 flex items-center justify-center gap-3 text-white/80">
                <span className="text-sm font-bold tracking-widest">CAI</span>
                <span className="flex-1 max-w-[70px] h-px bg-white/30 relative">
                  <Plane size={12} className="absolute -top-1.5 left-1/2 -translate-x-1/2 rotate-90 text-amber-300" />
                </span>
                <span className="text-sm font-bold tracking-widest">ANY</span>
              </div>
            </div>

            {/* Perforated tear line between stub and form */}
            <div className="relative h-0">
              <div className="absolute -left-2.5 -top-2.5 w-5 h-5 rounded-full bg-teal-900" />
              <div className="absolute -right-2.5 -top-2.5 w-5 h-5 rounded-full bg-teal-900" />
              <div className="absolute left-4 right-4 top-0 border-t-2 border-dashed border-stone-200" />
            </div>

            {/* Form section */}
            <div className="relative bg-white px-6 pt-7 pb-6">
              {loginError && <div className="bg-red-50 text-red-700 text-sm rounded-2xl px-3 py-2 mb-3">{loginError}</div>}
              <div className="space-y-3">
                <div>
                  <label className="text-xs text-stone-500 block mb-1">Username</label>
                  <div className="relative">
                    <User size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
                    <input className="w-full border border-stone-300 rounded-2xl pl-9 pr-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-800 focus:border-teal-800"
                      value={loginUsername} onChange={(e) => setLoginUsername(e.target.value)}
                      onKeyDown={(e) => e.key === "Enter" && handleLogin()} placeholder="Username" autoFocus />
                  </div>
                </div>
                <div>
                  <label className="text-xs text-stone-500 block mb-1">Password</label>
                  <div className="relative">
                    <Lock size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
                    <input type={showPassword ? "text" : "password"}
                      className="w-full border border-stone-300 rounded-2xl pl-9 pr-9 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-800 focus:border-teal-800"
                      value={loginPassword} onChange={(e) => setLoginPassword(e.target.value)}
                      onKeyDown={(e) => e.key === "Enter" && handleLogin()} placeholder="Password" />
                    <button type="button" onClick={() => setShowPassword(!showPassword)}
                      className="absolute right-2 top-1/2 -translate-y-1/2 text-stone-400 hover:text-stone-600">
                      {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                    </button>
                  </div>
                </div>
              </div>
              <button onClick={handleLogin}
                className="group w-full mt-5 bg-gradient-to-r from-teal-800 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-2xl px-4 py-2.5 flex items-center justify-center gap-2 shadow-lg shadow-teal-800/30 transition-all">
                Sign in
                <Plane size={15} className="rotate-45 transition-transform group-hover:translate-x-1 group-hover:-translate-y-1" />
              </button>
              <p className="text-xs text-stone-400 mt-4 text-center flex items-center justify-center gap-1">
                <ShieldCheck size={13} /> Ask your admin if you don't have an account yet.
              </p>

              {/* Barcode flourish, echoing a real boarding pass stub */}
              <div className="flex items-end gap-[2px] justify-center mt-5 h-5 opacity-25">
                {[3,1,2,4,1,3,2,1,4,2,3,1,2,4,1,3,2,4,1,2,3,1,4,2,1,3,2,4,1,2].map((h, i) => (
                  <span key={i} className="bg-stone-900 w-[2px]" style={{ height: `${h * 4}px` }} />
                ))}
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ---------- Render: main app ----------
  return (
    <div
      dir="ltr"
      className="w-full min-h-screen bg-gradient-to-b from-stone-50 via-white to-teal-50/50 text-stone-800"
      style={{ fontFamily: "'Inter', sans-serif" }}
    >
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:ital,opsz,wght@0,9..144,500;0,9..144,600;1,9..144,500&family=Inter:wght@400;500;600;700&display=swap');`}</style>
      <div className="max-w-5xl mx-auto p-4 md:p-6">
        {/* Boarding-pass style banner */}
        <div className="relative rounded-2xl bg-gradient-to-r from-teal-800 via-teal-800 to-teal-900 shadow-lg shadow-teal-900/20 overflow-hidden mb-0">
          <Plane size={140} className="pointer-events-none absolute -bottom-8 -right-6 text-white/[0.06] rotate-45" />
          <Compass size={90} className="pointer-events-none absolute -top-6 left-[38%] text-white/[0.05]" />
          <Luggage size={70} className="pointer-events-none absolute -bottom-4 left-[18%] text-white/[0.05] hidden md:block" />
          <header className="relative flex items-center justify-between flex-wrap gap-3 px-4 py-4 md:px-6">
            <div className="flex items-center gap-3">
              <div className="bg-white rounded-2xl p-2.5 shadow-sm shrink-0 hidden sm:block">
                <img src={LOGO_DATA_URL} alt="Perla Di Mare" className="w-[120px] h-auto md:w-[150px] object-contain" />
              </div>
              <div>
                <h1 className="text-lg md:text-2xl font-semibold text-white" style={{ fontFamily: "'Fraunces', serif" }}>
                  Flight Ticket Manager <span className="text-teal-200/60 font-medium text-xs md:text-base" style={{ fontFamily: "'Inter', sans-serif" }}>By Fady Habib</span>
                </h1>
                <p className="text-teal-100/80 text-sm flex items-center gap-1.5 flex-wrap mt-0.5">
                  Signed in as {currentUser.name}
                  {currentUser.isAdmin && (
                    <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-teal-900 bg-amber-300 border border-amber-400/50 rounded-full px-2 py-0.5">
                      <ShieldCheck size={11} /> Main account
                    </span>
                  )}
                  {!currentUser.isAdmin && currentEmployeeRecord && (
                    <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-white bg-white/10 border border-white/20 rounded-full px-2 py-0.5">
                      {roleLabel(currentEmployeeRecord.role)}
                    </span>
                  )}
                  {!currentUser.isAdmin && !canViewAllTickets && (
                    <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-teal-100 bg-white/10 border border-white/20 rounded-full px-2 py-0.5">
                      Your own tickets only
                    </span>
                  )}
                  {isAccountingUser && (
                    <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-amber-100 bg-amber-500/20 border border-amber-300/30 rounded-full px-2 py-0.5">
                      Accounting — view only
                    </span>
                  )}
                  {currentUser.isAdmin && (
                    <span className="relative">
                      <button
                        type="button"
                        onClick={() => setShowOnlineList(!showOnlineList)}
                        className="inline-flex items-center gap-1 text-[10px] font-semibold text-emerald-50 bg-emerald-500/20 border border-emerald-300/30 rounded-full px-2 py-0.5 hover:bg-emerald-500/30"
                      >
                        <Wifi size={11} />
                        {onlineUsernames.length} online now
                      </button>
                      {showOnlineList && (
                        <div className="absolute z-20 top-full mt-1 left-0 w-64 bg-white border border-stone-300 rounded-2xl shadow-lg p-2">
                          {onlineUsernames.length === 0 ? (
                            <p className="text-xs text-stone-400 px-1 py-1">No one online right now</p>
                          ) : (
                            <ul className="space-y-1 max-h-56 overflow-y-auto">
                              {onlineUsernames.map((u) => {
                                const emp = (employees || []).find((e) => e.username === u);
                                const activity = presenceMap[u] && presenceMap[u].activity;
                                return (
                                  <li key={u} className="flex items-center gap-1.5 text-xs text-stone-700 px-1 py-0.5">
                                    <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 shrink-0" />
                                    <span className="flex-1 truncate">
                                      {emp ? emp.name : u}
                                      {emp && emp.isAdmin && (
                                        <span className="text-[9px] text-teal-700 font-semibold"> (main)</span>
                                      )}
                                      {activity && (
                                        <span className="block text-[10px] text-stone-400 truncate">{activity}</span>
                                      )}
                                    </span>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        if (window.confirm(`Sign out ${emp ? emp.name : u} now?`)) {
                                          handleForceSignOut(u);
                                        }
                                      }}
                                      title="Sign out this employee"
                                      className="shrink-0 inline-flex items-center gap-0.5 text-[10px] font-semibold text-red-600 hover:text-red-800 border border-red-200 hover:border-red-300 bg-red-50 rounded-full px-1.5 py-0.5"
                                    >
                                      <LogOut size={10} /> Sign out
                                    </button>
                                  </li>
                                );
                              })}
                            </ul>
                          )}
                        </div>
                      )}
                    </span>
                  )}
                </p>
              </div>
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              {currentUser.isAdmin && (
                <button onClick={handleBackup}
                  className="border border-white/20 bg-white/10 hover:bg-white/20 text-white text-sm rounded-2xl px-3 py-2 flex items-center gap-1.5 transition-colors">
                  <Download size={15} /> Backup
                </button>
              )}
              {currentUser.isAdmin && (
                <button onClick={triggerRestore}
                  className="border border-white/20 bg-white/10 hover:bg-white/20 text-white text-sm rounded-2xl px-3 py-2 flex items-center gap-1.5 transition-colors">
                  <Upload size={15} /> Restore
                </button>
              )}
              {currentUser.isAdmin && (
                <input
                  type="file"
                  accept="application/json"
                  ref={fileInputRef}
                  onChange={handleRestoreFile}
                  className="hidden"
                />
              )}
              {currentUser.isAdmin && (
                <button onClick={() => setShowManage(!showManage)}
                  className="border border-white/20 bg-white/10 hover:bg-white/20 text-white text-sm rounded-2xl px-3 py-2 flex items-center gap-1.5 transition-colors">
                  <Users size={15} /> Manage employees
                </button>
              )}
              {canManageCompanies && (
                <button onClick={() => setShowManageCompanies(!showManageCompanies)}
                  className="border border-white/20 bg-white/10 hover:bg-white/20 text-white text-sm rounded-2xl px-3 py-2 flex items-center gap-1.5 transition-colors">
                  <Factory size={15} /> Manage companies
                </button>
              )}
              <button
                onClick={() => {
                  setShowChangePassword(!showChangePassword);
                  setPasswordError("");
                  setPasswordSuccess("");
                  setCurrentPasswordInput("");
                  setNewPasswordInput("");
                  setConfirmPasswordInput("");
                }}
                className="border border-white/20 bg-white/10 hover:bg-white/20 text-white text-sm rounded-2xl px-3 py-2 flex items-center gap-1.5 transition-colors">
                <Lock size={15} /> Change password
              </button>
              <button onClick={handleLogout}
                className="border border-white/20 bg-white/10 hover:bg-white/20 text-white text-sm rounded-2xl px-3 py-2 flex items-center gap-1.5 transition-colors">
                <LogOut size={15} /> Sign out
              </button>
              {onChangeServer && (
                <button
                  onClick={() => {
                    requestConfirm(
                      `Disconnect from the current server${currentServerUrl ? ` (${currentServerUrl})` : ""} and connect to a different one?`,
                      () => {
                        setConfirmDialog(null);
                        onChangeServer();
                      }
                    );
                  }}
                  title="Change data server"
                  className="border border-white/20 bg-white/10 hover:bg-white/20 text-teal-100 text-sm rounded-2xl px-3 py-2 flex items-center gap-1.5 transition-colors"
                >
                  <Wifi size={15} /> Server
                </button>
              )}
            </div>
          </header>
        </div>
        {/* Perforated tear line, like separating a boarding-pass stub from the rest */}
        <div className="relative h-6 mb-4">
          <div className="absolute -left-2.5 top-0 w-5 h-5 rounded-full bg-stone-50" />
          <div className="absolute -right-2.5 top-0 w-5 h-5 rounded-full bg-stone-50" />
          <div className="absolute left-4 right-4 top-2.5 border-t-2 border-dashed border-teal-800/20" />
        </div>

        {(showManage || showManageCompanies) && (
          <div className="mb-6">
        {showManage && currentUser.isAdmin && (
          <div className="bg-stone-50">
            <button onClick={() => setShowManage(false)}
              className="mb-4 border border-stone-300 text-stone-600 text-sm rounded-xl px-3 py-2 flex items-center gap-1.5 hover:bg-stone-100">
              <ArrowLeft size={15} /> Back
            </button>
          <div className="bg-white rounded-2xl border border-stone-200 p-4 md:p-5 mb-6">
            <h2 className="font-semibold text-stone-900 mb-1">Employee accounts</h2>
            <p className="text-xs text-stone-400 mb-4">
              As the main account, you can view and change every employee's password, edit their name or username, add or remove accounts, assign a grade (Manager, Supervisor, Employee, Accountant), and grant or remove main-account access. A grade fills in a starting set of permissions, but every permission — view all tickets, add tickets, edit tickets, delete tickets, accounting/notes-only mode, and manage companies — is an individual on/off switch you can set by hand for each employee, click the Permissions button on their row to open it. This is a basic access gate, not a secure authentication system — anyone with technical access to the app's stored data can read these passwords. Avoid reusing important passwords here.
            </p>
            {manageError && <div className="bg-red-50 text-red-700 text-sm rounded-xl px-3 py-2 mb-3">{manageError}</div>}
            <p className="text-xs text-stone-500 mb-3 flex items-center gap-1.5">
              <Wifi size={13} className="text-emerald-600" />
              {onlineUsernames.length} of {(employees || []).length} employees connected right now
            </p>
            <div className="border border-stone-200 rounded-xl overflow-hidden mb-4">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-stone-50 text-stone-500 text-xs">
                    <th className="text-left px-3 py-2 font-medium">Status</th>
                    <th className="text-left px-3 py-2 font-medium">Name</th>
                    <th className="text-left px-3 py-2 font-medium">Username</th>
                    <th className="text-left px-3 py-2 font-medium">Password</th>
                    <th className="text-left px-3 py-2 font-medium">Grade</th>
                    <th className="text-left px-3 py-2 font-medium">Permissions</th>
                    <th className="text-left px-3 py-2 font-medium"></th>
                  </tr>
                </thead>
                <tbody>
                  {(employees || []).map((e) => {
                    const isEditing = editingUsername === e.username;
                    if (isEditing) {
                      return (
                        <tr key={e.username} className="border-t border-stone-100 bg-stone-50">
                          <td className="px-3 py-2">
                            <span className={`inline-flex items-center gap-1 text-[10px] font-semibold rounded-full px-2 py-0.5 ${isOnline(e.username) ? "text-emerald-700 bg-emerald-50 border border-emerald-200" : "text-stone-400 bg-stone-100 border border-stone-200"}`}>
                              <span className={`w-1.5 h-1.5 rounded-full ${isOnline(e.username) ? "bg-emerald-500" : "bg-stone-300"}`} />
                              {isOnline(e.username) ? "Online" : "Offline"}
                            </span>
                          </td>
                          <td className="px-3 py-2">
                            <input
                              className="w-full border border-stone-300 rounded-xl px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                              value={editDraft.name}
                              onChange={(ev) => setEditDraft({ ...editDraft, name: ev.target.value })}
                            />
                          </td>
                          <td className="px-3 py-2">
                            <input
                              className="w-full border border-stone-300 rounded-xl px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                              value={editDraft.username}
                              onChange={(ev) => setEditDraft({ ...editDraft, username: ev.target.value })}
                            />
                          </td>
                          <td className="px-3 py-2">
                            <div className="relative">
                              <input
                                type={editShowPassword ? "text" : "password"}
                                className="w-full border border-stone-300 rounded-xl pl-2 pr-8 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                                value={editDraft.password}
                                onChange={(ev) => setEditDraft({ ...editDraft, password: ev.target.value })}
                              />
                              <button
                                type="button"
                                onClick={() => setEditShowPassword(!editShowPassword)}
                                className="absolute right-2 top-1/2 -translate-y-1/2 text-stone-400"
                              >
                                {editShowPassword ? <EyeOff size={14} /> : <Eye size={14} />}
                              </button>
                            </div>
                          </td>
                          <td className="px-3 py-2 text-stone-500">
                            {e.isAdmin ? (
                              <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-teal-800 bg-teal-50 border border-teal-200 rounded-full px-2 py-0.5">
                                <ShieldCheck size={11} /> Main
                              </span>
                            ) : (
                              <select
                                value={e.role || "employee"}
                                onChange={(ev) => handleRoleChange(e.username, ev.target.value)}
                                className="border border-stone-300 rounded-lg px-2 py-1.5 text-xs focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white"
                              >
                                {EMPLOYEE_ROLES.map((r) => (
                                  <option key={r.value} value={r.value}>{r.label}</option>
                                ))}
                              </select>
                            )}
                          </td>
                          <td className="px-3 py-2">
                            {e.isAdmin ? (
                              <span className="text-xs text-stone-400">Everything (main account)</span>
                            ) : (
                              <PermissionsCell emp={e} onOpen={() => setOpenPermissionsFor(e.username)} />
                            )}
                          </td>
                          <td className="px-3 py-2">
                            <div className="flex gap-1 justify-end">
                              <button onClick={saveEditEmployee} className="text-emerald-600 hover:text-emerald-800 p-1">
                                <Check size={15} />
                              </button>
                              <button onClick={cancelEditEmployee} className="text-stone-400 hover:text-red-600 p-1">
                                <X size={15} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    }
                    return (
                      <tr key={e.username} className="border-t border-stone-100">
                        <td className="px-3 py-2">
                          <span className={`inline-flex items-center gap-1 text-[10px] font-semibold rounded-full px-2 py-0.5 ${isOnline(e.username) ? "text-emerald-700 bg-emerald-50 border border-emerald-200" : "text-stone-400 bg-stone-100 border border-stone-200"}`}>
                            <span className={`w-1.5 h-1.5 rounded-full ${isOnline(e.username) ? "bg-emerald-500" : "bg-stone-300"}`} />
                            {isOnline(e.username) ? "Online" : "Offline"}
                          </span>
                        </td>
                        <td className="px-3 py-2">{e.name}</td>
                        <td className="px-3 py-2 text-stone-500">{e.username}</td>
                        <td className="px-3 py-2 text-stone-500">
                          <div className="flex items-center gap-2">
                            <span className="font-mono">
                              {visiblePasswords[e.username] ? e.password : "••••••••"}
                            </span>
                            <button onClick={() => togglePasswordVisible(e.username)} className="text-stone-400 hover:text-teal-800">
                              {visiblePasswords[e.username] ? <EyeOff size={14} /> : <Eye size={14} />}
                            </button>
                          </div>
                        </td>
                        <td className="px-3 py-2 text-stone-500">
                          {e.isAdmin ? (
                            <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-teal-800 bg-teal-50 border border-teal-200 rounded-full px-2 py-0.5">
                              <ShieldCheck size={11} /> Main
                            </span>
                          ) : (
                            <select
                              value={e.role || "employee"}
                              onChange={(ev) => handleRoleChange(e.username, ev.target.value)}
                              className="border border-stone-300 rounded-lg px-2 py-1.5 text-xs focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white"
                            >
                              {EMPLOYEE_ROLES.map((r) => (
                                <option key={r.value} value={r.value}>{r.label}</option>
                              ))}
                            </select>
                          )}
                        </td>
                        <td className="px-3 py-2">
                          {e.isAdmin ? (
                            <span className="text-xs text-stone-400">Everything (main account)</span>
                          ) : (
                            <PermissionsCell emp={e} onOpen={() => setOpenPermissionsFor(e.username)} />
                          )}
                        </td>
                        <td className="px-3 py-2 text-right">
                          <div className="flex gap-1 justify-end">
                            {e.isAdmin ? (
                              <button
                                onClick={() => handleDemoteAdmin(e.username)}
                                title="Remove main-account access"
                                className="text-stone-400 hover:text-amber-600 text-[11px] font-semibold border border-stone-200 rounded-lg px-1.5 py-1"
                              >
                                Remove main
                              </button>
                            ) : (
                              <button
                                onClick={() => handlePromoteToAdmin(e.username)}
                                title="Make this a main account"
                                className="text-stone-400 hover:text-teal-800 text-[11px] font-semibold border border-stone-200 rounded-lg px-1.5 py-1 flex items-center gap-1"
                              >
                                <ShieldCheck size={12} /> Make main
                              </button>
                            )}
                            <button onClick={() => startEditEmployee(e)} className="text-stone-400 hover:text-teal-800 p-1">
                              <Pencil size={15} />
                            </button>
                            {!e.isAdmin && (
                              <button onClick={() => handleDeleteEmployee(e.username)} className="text-stone-400 hover:text-red-600 p-1">
                                <Trash2 size={15} />
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
              <input className="border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                placeholder="Full name" value={newEmployee.name}
                onChange={(e) => setNewEmployee({ ...newEmployee, name: e.target.value })} />
              <input className="border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                placeholder="Username" value={newEmployee.username}
                onChange={(e) => setNewEmployee({ ...newEmployee, username: e.target.value })} />
              <input type="password" className="border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                placeholder="Password" value={newEmployee.password}
                onChange={(e) => setNewEmployee({ ...newEmployee, password: e.target.value })} />
            </div>

            {/* Grade: picking one fills the toggles below with a sensible starting
                point. Every toggle can still be switched by hand afterwards. */}
            <div className="mt-3 max-w-sm">
              <label className="text-xs text-stone-500 block mb-1.5">Grade</label>
              <div className="grid grid-cols-4 gap-1.5">
                {EMPLOYEE_ROLES.map((r) => (
                  <button
                    key={r.value}
                    type="button"
                    onClick={() =>
                      setNewEmployee({ ...newEmployee, role: r.value, ...ROLE_PRESETS[r.value] })
                    }
                    className={`text-xs font-semibold rounded-xl px-2 py-2 border transition-colors ${
                      newEmployee.role === r.value
                        ? "bg-teal-800 text-white border-teal-800"
                        : "bg-white text-stone-600 border-stone-300 hover:bg-stone-50"
                    }`}
                  >
                    {r.label}
                  </button>
                ))}
              </div>
            </div>

            {/* Detailed, individually switchable permissions — the grade above is only a
                starting point; every toggle here can be set by hand regardless of grade. */}
            <div className="relative mt-3 max-w-sm">
              <button
                type="button"
                onClick={() => setShowNewEmployeePerms(!showNewEmployeePerms)}
                className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm text-stone-700 hover:bg-stone-50 flex items-center justify-between gap-2"
              >
                <span className="font-medium">Permissions</span>
                <span className="text-xs text-stone-500 truncate">
                  {[
                    newEmployee.canViewAll && "View all",
                    newEmployee.canEdit && "Edit",
                    newEmployee.canDelete && "Delete",
                    newEmployee.isAccounting && "Notes only",
                    newEmployee.canManageCompanies && "Companies",
                  ]
                    .filter(Boolean)
                    .join(" · ") || "Own tickets only"}
                </span>
              </button>

              {showNewEmployeePerms && (
                <div className="absolute z-10 mt-1 w-full bg-white border border-stone-300 rounded-xl shadow-lg p-3 divide-y divide-stone-100">
                  <ToggleSwitch
                    label="View all tickets"
                    description="See every employee's tickets, not just their own"
                    checked={newEmployee.canViewAll || newEmployee.canEdit || newEmployee.canDelete}
                    disabled={newEmployee.isAccounting || newEmployee.canEdit || newEmployee.canDelete}
                    onChange={(v) => setNewEmployee(reconcilePermissions({ ...newEmployee, canViewAll: v }))}
                  />
                  <ToggleSwitch
                    label="Edit tickets"
                    description="Edit any ticket they can see (view access included automatically)"
                    checked={newEmployee.canEdit}
                    disabled={newEmployee.isAccounting}
                    onChange={(v) => setNewEmployee(reconcilePermissions({ ...newEmployee, canEdit: v }))}
                  />
                  <ToggleSwitch
                    label="Delete tickets"
                    description="Permanently remove any ticket they can see"
                    checked={newEmployee.canDelete}
                    disabled={newEmployee.isAccounting}
                    onChange={(v) => setNewEmployee(reconcilePermissions({ ...newEmployee, canDelete: v }))}
                  />
                  <ToggleSwitch
                    label="Accounting mode"
                    description="View all tickets, but the only edit allowed is the Notes field"
                    checked={newEmployee.isAccounting}
                    onChange={(v) => setNewEmployee(reconcilePermissions({ ...newEmployee, isAccounting: v }))}
                  />
                  <ToggleSwitch
                    label="Manage companies"
                    description="Add, edit, or remove saved company records"
                    checked={newEmployee.canManageCompanies}
                    onChange={(v) => setNewEmployee({ ...newEmployee, canManageCompanies: v })}
                  />
                </div>
              )}
            </div>

            <button onClick={handleAddEmployee}
              className="mt-3 bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10 transition-colors flex items-center gap-1.5">
              <UserPlus size={15} /> Add employee
            </button>
          </div>
          </div>
        )}
        {showManageCompanies && canManageCompanies && (
          <div className="bg-stone-50">
            <button onClick={() => setShowManageCompanies(false)}
              className="mb-4 border border-stone-300 text-stone-600 text-sm rounded-xl px-3 py-2 flex items-center gap-1.5 hover:bg-stone-100">
              <ArrowLeft size={15} /> Back
            </button>
          <div className="bg-white rounded-2xl border border-stone-200 p-4 md:p-5 mb-6">
            <h2 className="font-semibold text-stone-900 mb-1 flex items-center gap-2">
              <Factory size={18} className="text-stone-500" /> Companies
            </h2>
            <p className="text-xs text-stone-400 mb-4">
              Register each company's details here so they're always available to pick from the Company field and filter, even before any ticket has been entered for them.
            </p>
            {companyError && <div className="bg-red-50 text-red-700 text-sm rounded-xl px-3 py-2 mb-3">{companyError}</div>}

            <div className="grid grid-cols-1 md:grid-cols-2 gap-3 max-w-2xl mb-3">
              <input
                className="border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                placeholder="Company name"
                value={newCompanyDraft.name}
                onChange={(e) => setNewCompanyDraft({ ...newCompanyDraft, name: e.target.value })}
              />
              <input
                className="border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                placeholder="Tax number"
                value={newCompanyDraft.taxNumber}
                onChange={(e) => setNewCompanyDraft({ ...newCompanyDraft, taxNumber: e.target.value })}
              />
              <input
                className="border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                placeholder="Commercial registration number"
                value={newCompanyDraft.commercialReg}
                onChange={(e) => setNewCompanyDraft({ ...newCompanyDraft, commercialReg: e.target.value })}
              />
              <input
                className="border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                placeholder="Phone numbers (comma separated)"
                value={newCompanyDraft.phones}
                onChange={(e) => setNewCompanyDraft({ ...newCompanyDraft, phones: e.target.value })}
              />
            </div>
            <div className="flex gap-2 mb-5">
              <button
                onClick={handleAddCompany}
                className="bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10 transition-colors flex items-center gap-1.5 whitespace-nowrap"
              >
                {editingCompanyName ? <Check size={15} /> : <Factory size={15} />}
                {editingCompanyName ? "Save changes" : "Add company"}
              </button>
              {editingCompanyName && (
                <button
                  onClick={cancelEditCompany}
                  className="border border-stone-300 text-stone-600 text-sm rounded-xl px-4 py-2 flex items-center gap-1.5"
                >
                  <X size={15} /> Cancel
                </button>
              )}
            </div>

            <div className="flex items-center justify-between mb-3">
              <p className="text-xs text-stone-400">
                {suggestions.companies.length} compan{suggestions.companies.length === 1 ? "y" : "ies"} saved
              </p>
              <button
                onClick={() => setShowCompaniesList(!showCompaniesList)}
                className="text-teal-800 border border-teal-800 hover:bg-teal-50 text-xs font-semibold rounded-xl px-3 py-1.5 flex items-center gap-1.5"
              >
                <List size={14} /> {showCompaniesList ? "Hide companies list" : "View all companies"}
              </button>
            </div>

            {showCompaniesList && (
              suggestions.companies.length === 0 ? (
                <p className="text-sm text-stone-400">No companies saved yet</p>
              ) : (
                <div className="border border-stone-200 rounded-xl overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs">
                      <thead>
                        <tr className="bg-stone-50 text-stone-500 text-[11px] uppercase tracking-wide">
                          <th className="text-left px-3 py-2 font-semibold whitespace-nowrap">Company</th>
                          <th className="text-left px-3 py-2 font-semibold whitespace-nowrap">Tax number</th>
                          <th className="text-left px-3 py-2 font-semibold whitespace-nowrap">Commercial reg.</th>
                          <th className="text-left px-3 py-2 font-semibold whitespace-nowrap">Phone</th>
                          <th className="text-right px-3 py-2 font-semibold whitespace-nowrap"></th>
                        </tr>
                      </thead>
                      <tbody>
                        {[...suggestions.companies]
                          .sort((a, b) => companyName(a).localeCompare(companyName(b)))
                          .map((c) => {
                            const name = companyName(c);
                            const taxNumber = typeof c === "object" ? c.taxNumber : "";
                            const commercialReg = typeof c === "object" ? c.commercialReg : "";
                            const phones = typeof c === "object" && Array.isArray(c.phones) ? c.phones : [];
                            return (
                              <tr
                                key={name}
                                className={`border-t border-stone-100 ${editingCompanyName === name ? "bg-teal-50/40" : "hover:bg-stone-50"}`}
                              >
                                <td className="px-3 py-2 font-medium text-stone-800 whitespace-nowrap">{name}</td>
                                <td className="px-3 py-2 text-stone-600 whitespace-nowrap">{taxNumber || "-"}</td>
                                <td className="px-3 py-2 text-stone-600 whitespace-nowrap">{commercialReg || "-"}</td>
                                <td className="px-3 py-2 text-stone-600 whitespace-nowrap">{phones.length > 0 ? phones.join(", ") : "-"}</td>
                                <td className="px-3 py-2 text-right whitespace-nowrap">
                                  <div className="flex gap-1 justify-end">
                                    <button onClick={() => handleEditCompanyClick(c)} className="text-stone-400 hover:text-teal-800 p-0.5">
                                      <Pencil size={13} />
                                    </button>
                                    <button onClick={() => handleDeleteCompany(name)} className="text-stone-400 hover:text-red-600 p-0.5">
                                      <Trash2 size={13} />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            );
                          })}
                      </tbody>
                    </table>
                  </div>
                </div>
              )
            )}
          </div>
          </div>
        )}
          </div>
        )}

        {!showManage && !showManageCompanies && (
        <>
        {/* Top-level section switcher */}
        <div className="flex items-center justify-center gap-3 mb-6">
          <button
            onClick={() => setActiveSection("flights")}
            className={`flex flex-col items-center gap-1.5 px-6 py-3 rounded-2xl border text-xs font-semibold transition-colors ${
              activeSection === "flights"
                ? "bg-gradient-to-b from-teal-700 to-teal-900 text-white border-teal-800 shadow-md shadow-teal-800/30 ring-1 ring-inset ring-amber-600/50"
                : "bg-white text-stone-500 border-stone-200 hover:border-amber-600 hover:text-teal-800 hover:shadow-sm"
            }`}
          >
            <Plane size={22} className="rotate-45" />
            Flights
          </button>
          <button
            onClick={() => setActiveSection("hotels")}
            className={`flex flex-col items-center gap-1.5 px-6 py-3 rounded-2xl border text-xs font-semibold transition-colors ${
              activeSection === "hotels"
                ? "bg-gradient-to-b from-teal-700 to-teal-900 text-white border-teal-800 shadow-md shadow-teal-800/30 ring-1 ring-inset ring-amber-600/50"
                : "bg-white text-stone-500 border-stone-200 hover:border-amber-600 hover:text-teal-800 hover:shadow-sm"
            }`}
          >
            <Building2 size={22} />
            Hotels
          </button>
          <button
            onClick={() => setActiveSection("cars")}
            className={`flex flex-col items-center gap-1.5 px-6 py-3 rounded-2xl border text-xs font-semibold transition-colors ${
              activeSection === "cars"
                ? "bg-gradient-to-b from-teal-700 to-teal-900 text-white border-teal-800 shadow-md shadow-teal-800/30 ring-1 ring-inset ring-amber-600/50"
                : "bg-white text-stone-500 border-stone-200 hover:border-amber-600 hover:text-teal-800 hover:shadow-sm"
            }`}
          >
            <Car size={22} />
            Transportation
          </button>
          <button
            onClick={() => setActiveSection("files")}
            className={`flex flex-col items-center gap-1.5 px-6 py-3 rounded-2xl border text-xs font-semibold transition-colors ${
              activeSection === "files"
                ? "bg-gradient-to-b from-teal-700 to-teal-900 text-white border-teal-800 shadow-md shadow-teal-800/30 ring-1 ring-inset ring-amber-600/50"
                : "bg-white text-stone-500 border-stone-200 hover:border-amber-600 hover:text-teal-800 hover:shadow-sm"
            }`}
          >
            <FileText size={22} />
            Files
          </button>
        </div>

        {activeSection === "flights" && (
        <>
        {currentUser.isAdmin && (restoreError || restoreSuccess) && (
          <div className={`text-sm rounded-xl px-3 py-2 mb-4 ${restoreError ? "bg-red-50 text-red-700" : "bg-emerald-50 text-emerald-700"}`}>
            {restoreError || restoreSuccess}
          </div>
        )}



        {showChangePassword && (
          <div className="bg-white rounded-2xl border border-stone-200 p-4 md:p-5 mb-6 max-w-sm">
            <h2 className="font-semibold text-stone-900 mb-1 flex items-center gap-2">
              <Lock size={16} className="text-teal-800" /> Change your password
            </h2>
            <p className="text-xs text-stone-400 mb-4">
              Signed in as {currentUser.name} ({currentUser.username})
            </p>
            {passwordError && <div className="bg-red-50 text-red-700 text-sm rounded-xl px-3 py-2 mb-3">{passwordError}</div>}
            {passwordSuccess && <div className="bg-emerald-50 text-emerald-700 text-sm rounded-xl px-3 py-2 mb-3">{passwordSuccess}</div>}
            <div className="space-y-3">
              <div>
                <label className="text-xs text-stone-500 block mb-1">Current password</label>
                <input type="password"
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={currentPasswordInput} onChange={(e) => setCurrentPasswordInput(e.target.value)} />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">New password</label>
                <input type="password"
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={newPasswordInput} onChange={(e) => setNewPasswordInput(e.target.value)} />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Confirm new password</label>
                <input type="password"
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={confirmPasswordInput} onChange={(e) => setConfirmPasswordInput(e.target.value)} />
              </div>
            </div>
            <button onClick={handleChangePassword}
              className="w-full mt-4 bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10 transition-colors">
              Update password
            </button>
          </div>
        )}

        {/* Summary cards */}
        <div className="flex items-center justify-between mb-2">
          <p className="text-sm text-stone-500">
            Totals for: <span className="font-semibold text-stone-700">
              {selectedYear ? selectedYear : ""}
              {selectedMonth ? ` · ${monthLabel(selectedMonth)}` : ""}
              {selectedCompany ? ` · ${selectedCompany}` : ""}
              {selectedEmployee ? ` · ${selectedEmployee}` : ""}
              {selectedSupplier ? ` · ${selectedSupplier}` : ""}
              {!hasActiveFilter && "all months"}
            </span>
          </p>
          <button
            onClick={() => (hasActiveFilter ? exportFiltered() : exportAllMonths())}
            className="text-teal-800 border border-teal-800 hover:bg-teal-50 text-xs font-semibold rounded-xl px-3 py-1.5 flex items-center gap-1.5"
          >
            <Download size={14} /> {hasActiveFilter ? "Export filtered results to Excel" : "Export all months to Excel"}
          </button>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-6">
          <div className="bg-white rounded-2xl border border-stone-200 p-4 flex items-center gap-3">
            <div className="bg-stone-100 rounded-xl p-2 text-stone-600"><Ticket size={20} /></div>
            <div>
              <p className="text-xs text-stone-500">Tickets</p>
              <p className="text-lg font-bold">{totals.count}</p>
            </div>
          </div>
          <div className="bg-white rounded-2xl border border-stone-200 p-4 flex items-center gap-3">
            <div className="bg-teal-50 rounded-xl p-2 text-teal-900"><Wallet size={20} /></div>
            <div>
              <p className="text-xs text-stone-500">Total sales</p>
              <p className="text-lg font-bold">{fmt(totals.total)}</p>
            </div>
          </div>
          <div className="bg-white rounded-2xl border border-stone-200 p-4 flex items-center gap-3">
            <div className="bg-emerald-50 rounded-xl p-2 text-emerald-700"><TrendingUp size={20} /></div>
            <div>
              <p className="text-xs text-stone-500">Total profit</p>
              <p className="text-lg font-bold text-emerald-700">{fmt(totals.profit)}</p>
            </div>
          </div>
        </div>

        {/* Entry form: hidden for accounting accounts (view-only + notes-only), and for
            anyone with neither add nor edit permission. Shown while editing an existing
            ticket as long as the person has edit access, even if add access is off. */}
        {!isAccountingUser && (canAddTickets || (form.id && canEditTickets)) && (
        <div className="bg-white rounded-2xl border border-stone-200 p-4 md:p-5 mb-6">
          <h2 className="font-semibold text-stone-900 mb-4">{form.id ? "Edit ticket" : "Add a new ticket"}</h2>
          {error && (
            <div className="bg-red-50 text-red-700 text-sm rounded-xl px-3 py-2 mb-3">{error}</div>
          )}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
            <div>
              <label className="text-xs text-stone-500 block mb-1">Entered by</label>
              <div className="w-full border border-stone-200 bg-stone-50 rounded-xl px-3 py-2 text-sm text-stone-600">
                {currentUser.name}
              </div>
            </div>
            <div>
              <label className="text-xs text-stone-500 block mb-1">Company (optional)</label>
              <input
                className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={form.company}
                onChange={(e) => setForm({ ...form, company: e.target.value.toUpperCase() })}
                placeholder="e.g. Acme Corp"
                list="company-suggestions"
              />
            </div>
            <div>
              <label className="text-xs text-stone-500 block mb-1">Number of customers</label>
              <input
                type="number"
                min={1}
                max={50}
                className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={form.customersCount}
                onChange={(e) => handleCustomersCountChange(e.target.value)}
                onBlur={(e) => {
                  if (e.target.value === "" || parseInt(e.target.value, 10) < 1) {
                    handleCustomersCountChange(1);
                  }
                }}
                placeholder="1"
              />
            </div>
            <div>
              <label className="text-xs text-stone-500 block mb-1">Supplier</label>
              {supplierOther ? (
                <div className="flex gap-2">
                  <input
                    className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                    value={form.supplier}
                    onChange={(e) => setForm({ ...form, supplier: e.target.value })}
                    placeholder="Enter supplier name"
                    autoFocus
                  />
                  <button
                    type="button"
                    onClick={() => { setSupplierOther(false); setForm({ ...form, supplier: "" }); }}
                    className="shrink-0 text-xs text-stone-500 hover:text-teal-800 border border-stone-300 rounded-xl px-2"
                  >
                    List
                  </button>
                </div>
              ) : (
                <select
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white"
                  value={form.supplier}
                  onChange={(e) => {
                    if (e.target.value === "__other__") {
                      setSupplierOther(true);
                      setForm({ ...form, supplier: "" });
                    } else {
                      setForm({ ...form, supplier: e.target.value });
                    }
                  }}
                >
                  <option value="">Select supplier</option>
                  {SUPPLIERS.map((s) => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                  <option value="__other__">Other</option>
                </select>
              )}
            </div>
          </div>

          {/* Dynamic customer name + ticket number cells, one row per customer */}
          <div className="mt-4">
            <label className="text-xs text-stone-500 block mb-2">
              Customers ({form.customers.length})
            </label>
            <div className="space-y-2">
              {form.customers.map((c, i) => (
                <div key={i} className="grid grid-cols-2 gap-2 md:gap-3">
                  <input
                    className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                    value={c.name}
                    onChange={(e) => handleCustomerFieldChange(i, "name", e.target.value)}
                    placeholder={`Customer ${i + 1} name`}
                  />
                  <input
                    className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                    value={c.ticketNumber}
                    onChange={(e) => handleCustomerFieldChange(i, "ticketNumber", e.target.value)}
                    onBlur={() => handleTicketNumberBlur(i)}
                    placeholder={`Ticket number ${i + 1} (e.g. 077-1234567890)`}
                  />
                </div>
              ))}
            </div>
          </div>

          {/* Reissue tracking: mark the ticket being entered as a reissue of an older
              ticket, then look that old ticket number up to auto-fill its issue date. */}
          <div className="mt-4 bg-amber-50 border border-amber-200 rounded-xl p-3">
            <label className="flex items-center gap-2 text-sm font-semibold text-amber-800 cursor-pointer select-none">
              <input
                type="checkbox"
                className="w-4 h-4 accent-amber-700"
                checked={form.isReissued}
                onChange={(e) => {
                  const isReissued = e.target.checked;
                  setForm({
                    ...form,
                    isReissued,
                    oldTicketNumber: isReissued ? form.oldTicketNumber : "",
                    oldTicketIssueDate: isReissued ? form.oldTicketIssueDate : "",
                  });
                }}
              />
              This ticket is a reissue of an older ticket
            </label>
            {form.isReissued && (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mt-3">
                <div>
                  <label className="text-xs text-stone-500 block mb-1">Old ticket number</label>
                  <input
                    className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                    value={form.oldTicketNumber}
                    onChange={(e) => handleOldTicketNumberChange(e.target.value)}
                    onBlur={handleOldTicketNumberBlur}
                    placeholder="e.g. 077-1234567890"
                  />
                </div>
                <div>
                  <label className="text-xs text-stone-500 block mb-1">Old ticket issue date</label>
                  <div className="w-full border border-stone-200 bg-stone-50 rounded-xl px-3 py-2 text-sm text-stone-600">
                    {form.oldTicketIssueDate
                      ? formatDisplayDate(form.oldTicketIssueDate)
                      : form.oldTicketNumber
                      ? "Not found among saved tickets"
                      : "Enter the old ticket number above"}
                  </div>
                </div>
              </div>
            )}
          </div>

          <div className="mt-4">
            <label className="flex items-center gap-2 text-xs font-semibold text-stone-500 cursor-pointer select-none">
              <input
                type="checkbox"
                className="w-4 h-4 accent-teal-700"
                checked={!!form.multiDestination}
                onChange={(e) => {
                  const multiDestination = e.target.checked;
                  setForm({
                    ...form,
                    multiDestination,
                    // Seed the stop list from the current From/To the first time this is
                    // switched on, so nothing already typed gets lost.
                    destinations:
                      multiDestination && !(form.destinations || []).some((d) => (d || "").trim())
                        ? [form.from || "", form.to || ""]
                        : form.destinations,
                  });
                }}
              />
              Multi-destination route (multi-city)
            </label>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mt-2">
            {form.multiDestination ? (
              <div className="md:col-span-2">
                <label className="text-xs text-stone-500 block mb-1">Route stops (in order)</label>
                <div className="space-y-2">
                  {form.destinations.map((d, i) => (
                    <div key={i} className="flex items-center gap-2">
                      <span className="text-xs text-stone-400 w-14 shrink-0">
                        {i === 0 ? "From" : i === form.destinations.length - 1 ? "Final" : `Stop ${i}`}
                      </span>
                      <input
                        className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={d}
                        onChange={(e) => handleDestinationChange(i, e.target.value)}
                        placeholder={i === 0 ? "Cairo" : "Dubai"}
                        list="city-suggestions"
                      />
                      {form.destinations.length > 2 && (
                        <button
                          type="button"
                          onClick={() => removeDestinationStop(i)}
                          className="shrink-0 text-stone-400 hover:text-red-600"
                          title="Remove stop"
                        >
                          <X size={16} />
                        </button>
                      )}
                    </div>
                  ))}
                  <button
                    type="button"
                    onClick={addDestinationStop}
                    className="inline-flex items-center gap-1 text-xs font-semibold text-teal-700 hover:text-teal-900"
                  >
                    <Plus size={14} /> Add stop
                  </button>
                </div>
              </div>
            ) : (
              <>
                <div>
                  <label className="text-xs text-stone-500 block mb-1">From</label>
                  <input
                    className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                    value={form.from}
                    onChange={(e) => handleCityChange("from", e.target.value)}
                    placeholder="Cairo"
                    list="city-suggestions"
                  />
                </div>
                <div>
                  <label className="text-xs text-stone-500 block mb-1">To</label>
                  <input
                    className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                    value={form.to}
                    onChange={(e) => handleCityChange("to", e.target.value)}
                    placeholder="Dubai"
                    list="city-suggestions"
                  />
                </div>
              </>
            )}
            <div>
              <label className="text-xs text-stone-500 mb-1 flex items-center gap-1.5">
                <span>Airline</span>
                {getAirlineNameByIata(form.airline) && (
                  <span className="bg-teal-50 text-teal-700 border border-teal-200 rounded px-1.5 py-0.5 text-[10px] font-semibold">
                    {getAirlineNameByIata(form.airline)}
                  </span>
                )}
              </label>
              <input
                className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={form.airline}
                onChange={(e) => handleAirlineChange(e.target.value)}
                placeholder="MS"
                list="airline-suggestions"
              />
            </div>
            <div>
              <label className="text-xs text-stone-500 block mb-1">Ticket issue date</label>
              <input
                type="date"
                lang="en-GB"
                max={todayDateStr()}
                className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={form.date}
                onChange={(e) => {
                  const v = e.target.value;
                  // Belt-and-braces: some browsers still let a future date be typed
                  // manually even with `max` set, so clamp it back to today here too.
                  setForm({ ...form, date: v > todayDateStr() ? todayDateStr() : v });
                }}
              />
            </div>
            <div>
              <label className="text-xs text-stone-500 block mb-1">Net price</label>
              <input
                type="number"
                className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={form.netPrice}
                onChange={(e) => setForm({ ...form, netPrice: e.target.value })}
                placeholder="0"
              />
            </div>
            <div>
              <label className="text-xs text-stone-500 block mb-1">Sold price</label>
              <input
                type="number"
                className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={form.soldPrice}
                onChange={(e) => setForm({ ...form, soldPrice: e.target.value })}
                placeholder="0"
              />
            </div>
            <div>
              <label className="text-xs text-stone-500 block mb-1">Profit (auto)</label>
              <div className="w-full border border-stone-200 bg-stone-50 rounded-xl px-3 py-2 text-sm text-emerald-700 font-semibold">
                {fmt(profit(form.netPrice, form.soldPrice))}
              </div>
            </div>
            <div className="md:col-span-3">
              <label className="text-xs text-stone-500 block mb-1">Notes</label>
              <textarea
                className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 min-h-[80px]"
                value={form.notes}
                onChange={(e) => setForm({ ...form, notes: e.target.value.toUpperCase() })}
                placeholder="Optional"
              />
            </div>
          </div>

          <div className="flex gap-2 mt-4">
            <button
              onClick={handleSubmit}
              className="bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10 transition-colors flex items-center gap-1.5"
            >
              <Check size={16} /> {form.id ? "Save changes" : "Add ticket"}
            </button>
            {form.id && (
              <button
                onClick={handleCancel}
                className="border border-stone-300 text-stone-600 text-sm rounded-xl px-4 py-2 flex items-center gap-1.5"
              >
                <X size={16} /> Cancel
              </button>
            )}
          </div>
        </div>
        )}

        {/* Search and filters */}
        <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2 mb-3">
          <div className="relative flex-1">
            <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
            <input
              className="w-full border border-stone-300 rounded-xl pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
              placeholder="Search by employee, company, ticket number, customer, destination, or airline"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
            />
          </div>
          <div className="relative sm:w-40">
            <Calendar size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none" />
            <select
              className="w-full border border-stone-300 rounded-xl pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white appearance-none"
              value={selectedYear}
              onChange={(e) => setSelectedYear(e.target.value)}
            >
              <option value="">All years</option>
              {yearsAvailable.map((y) => (
                <option key={y} value={y}>{y}</option>
              ))}
            </select>
          </div>
          <div className="relative sm:w-56">
            <Calendar size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none" />
            <select
              className="w-full border border-stone-300 rounded-xl pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white appearance-none"
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(e.target.value)}
            >
              <option value="">All months</option>
              {monthsAvailable.map((key) => (
                <option key={key} value={key}>{monthLabel(key)}</option>
              ))}
            </select>
          </div>
          <div className="relative sm:w-56">
            <Building2 size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none" />
            <select
              className="w-full border border-stone-300 rounded-xl pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white appearance-none"
              value={selectedCompany}
              onChange={(e) => setSelectedCompany(e.target.value)}
            >
              <option value="">All companies</option>
              {companiesAvailable.map((name) => (
                <option key={name} value={name}>{name}</option>
              ))}
            </select>
          </div>
          <div className="relative sm:w-56">
            <User size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none" />
            <select
              className="w-full border border-stone-300 rounded-xl pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white appearance-none"
              value={selectedEmployee}
              onChange={(e) => setSelectedEmployee(e.target.value)}
            >
              <option value="">All employees</option>
              {employeesAvailable.map((name) => (
                <option key={name} value={name}>{name}</option>
              ))}
            </select>
          </div>
          <div className="relative sm:w-56">
            <Plane size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none" />
            <select
              className="w-full border border-stone-300 rounded-xl pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white appearance-none"
              value={selectedSupplier}
              onChange={(e) => setSelectedSupplier(e.target.value)}
            >
              <option value="">All suppliers</option>
              {suppliersAvailable.map((name) => (
                <option key={name} value={name}>{name}</option>
              ))}
            </select>
          </div>
        </div>

        <datalist id="company-suggestions">
          {suggestions.companies.map((c) => (
            <option key={companyName(c)} value={companyName(c)} />
          ))}
        </datalist>
        <datalist id="airline-suggestions">
          {suggestions.airlines.map((code) => (
            <option key={`u-${code}`} value={code} />
          ))}
          {AIRLINE_CODES.map((a) => (
            <option key={`a-${a.code}`} value={a.iata} label={`${a.iata} — ${a.name}`} />
          ))}
        </datalist>
        <datalist id="city-suggestions">
          {suggestions.cities.map((name) => (
            <option key={`u-${name}`} value={name} />
          ))}
          {AIRPORTS.map((entry) => (
            <option key={`p-${entry}`} value={entry} />
          ))}
        </datalist>

        {/* Ticket list */}
        <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden">
          {filtered.length === 0 ? (
            <p className="text-center text-stone-400 text-sm py-10">
              {visibleTickets.length === 0 ? "No tickets recorded yet" : "No results match your search"}
            </p>
          ) : (
            <div className="overflow-x-auto rounded-xl border border-stone-200">
              <table className="w-full text-xs border-collapse">
                <thead>
                  <tr className="bg-teal-50/60 text-teal-800 text-[11px] uppercase tracking-wide border-b-2 border-teal-200">
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Employee</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Company</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Supplier</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Ticket #</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Customer</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Route</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Airline</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Date</th>
                    <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Net price</th>
                    <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Sold price</th>
                    <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Profit</th>
                    <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap"></th>
                  </tr>
                </thead>
                <tbody>
                  {sortedFiltered.flatMap((t) => {
                    const customers = getCustomers(t);
                    const isMulti = customers.length > 1;
                    return customers.map((c, i) => (
                      <tr
                        key={`${t.id}-${i}`}
                        onClick={() => openTicketDetail(t)}
                        className={`border-t border-stone-100 leading-tight cursor-pointer ${i > 0 ? "border-t-0" : ""} ${isMulti ? "bg-amber-50 hover:bg-amber-100" : "hover:bg-teal-50/60"}`}
                      >
                        <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">{t.employee || "-"}</td>
                        <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">
                          {t.company && t.company.trim() ? (
                            t.company
                          ) : (
                            <span className="text-stone-400 italic">Individual</span>
                          )}
                        </td>
                        <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">{t.supplier || "-"}</td>
                        <td className="px-2.5 py-1 text-stone-600 font-mono whitespace-nowrap">
                          <span className="inline-flex items-center gap-1.5">
                            {c.ticketNumber || "-"}
                            {t.isReissued && (
                              <span
                                title={`Reissued from ${t.oldTicketNumber || "an older ticket"}`}
                                className="inline-flex items-center text-[10px] font-semibold text-amber-700 bg-amber-100 border border-amber-300 rounded-full px-1.5 py-0.5"
                              >
                                Reissued
                              </span>
                            )}
                            {hasRefund(t) && (t.refund.customerIndex || 0) === i && (
                              <span
                                title={`Refunded — Airline: ${fmt(t.refund.airlineAmount)} · Customer: ${fmt(t.refund.customerAmount)}`}
                                className="inline-flex items-center text-[10px] font-semibold text-sky-700 bg-sky-100 border border-sky-300 rounded-full px-1.5 py-0.5"
                              >
                                Refunded
                              </span>
                            )}
                          </span>
                        </td>
                        <td className="px-2.5 py-1 font-medium text-stone-800 whitespace-nowrap">
                          <span className="inline-flex items-center gap-1.5">
                            {c.name || "-"}
                            {isMulti && i === 0 && (
                              <span
                                title={`This booking has ${customers.length} customers / tickets`}
                                className="inline-flex items-center gap-0.5 text-[10px] font-semibold text-amber-700 bg-amber-100 border border-amber-300 rounded-full px-1.5 py-0.5"
                              >
                                <Users size={10} /> {customers.length}
                              </span>
                            )}
                          </span>
                        </td>
                        <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">{routeLabel(t)}</td>
                        <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap" title={getAirlineNameByIata(t.airline) || t.airline || ""}>
                          {t.airline ? (getAirlineIata(t.airline) || t.airline) : "-"}
                        </td>
                        <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">{t.date ? formatDisplayDate(t.date) : "-"}</td>
                        <td className="px-2.5 py-1 text-stone-600 text-right whitespace-nowrap">{fmt(netAfterRefund(t))}</td>
                        <td className="px-2.5 py-1 text-stone-600 text-right whitespace-nowrap">{fmt(soldAfterRefund(t))}</td>
                        <td className="px-2.5 py-1 font-semibold text-emerald-700 text-right whitespace-nowrap">{fmt(profitAfterRefund(t))}</td>
                        <td className="px-2.5 py-1 text-right whitespace-nowrap">
                          {(currentUser.isAdmin || canEditTickets) ? (
                            <div className="flex gap-0.5 justify-end">
                              <button onClick={(ev) => { ev.stopPropagation(); handleEdit(t); }} className="text-stone-400 hover:text-teal-800 p-0.5">
                                <Pencil size={13} />
                              </button>
                              {(currentUser.isAdmin || canDeleteTickets) && (
                                <button onClick={(ev) => { ev.stopPropagation(); handleDelete(t.id); }} className="text-stone-400 hover:text-red-600 p-0.5">
                                  <Trash2 size={13} />
                                </button>
                              )}
                            </div>
                          ) : (
                            <span className="text-stone-300 text-[11px] block text-right">—</span>
                          )}
                        </td>
                      </tr>
                    ));
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {!selectedMonth && monthlyBreakdown.length > 0 && (
          <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden mt-6">
            <div className="px-4 py-3 border-b border-stone-100">
              <h2 className="font-semibold text-stone-900 text-sm">Totals by month</h2>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-stone-50 text-stone-500 text-xs">
                    <th className="text-left px-3 py-2 font-medium">Month</th>
                    <th className="text-left px-3 py-2 font-medium">Tickets</th>
                    <th className="text-left px-3 py-2 font-medium">Total sales</th>
                    <th className="text-left px-3 py-2 font-medium">Total profit</th>
                    <th className="text-left px-3 py-2 font-medium"></th>
                  </tr>
                </thead>
                <tbody>
                  {monthlyBreakdown.map((m) => (
                    <tr key={m.key} className="border-t border-stone-100 hover:bg-stone-50">
                      <td className="px-3 py-2 font-medium text-stone-800">{monthLabel(m.key)}</td>
                      <td className="px-3 py-2 text-stone-600">{m.count}</td>
                      <td className="px-3 py-2 text-stone-600">{fmt(m.total)}</td>
                      <td className="px-3 py-2 font-semibold text-emerald-700">{fmt(m.profit)}</td>
                      <td className="px-3 py-2">
                        <div className="flex items-center gap-3 justify-end">
                          <button
                            onClick={() => exportMonth(m.key)}
                            className="text-stone-400 hover:text-teal-800 text-xs font-medium flex items-center gap-1"
                          >
                            <Download size={13} /> Export
                          </button>
                          <button
                            onClick={() => setSelectedMonth(m.key)}
                            className="text-teal-800 text-xs font-medium hover:underline"
                          >
                            View details
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {!selectedCompany && companyBreakdown.length > 0 && (
          <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden mt-6">
            <div className="px-4 py-3 border-b border-stone-100">
              <h2 className="font-semibold text-stone-900 text-sm">Companies and their customers</h2>
            </div>
            <div className="divide-y divide-stone-100">
              {companyBreakdown.map((c) => (
                <div key={c.name} className="px-4 py-3">
                  <div className="flex items-center justify-between flex-wrap gap-2">
                    <div className="flex items-center gap-2">
                      <Building2 size={16} className="text-stone-400" />
                      <button
                        onClick={() => setSelectedCompany(c.name)}
                        className="font-medium text-stone-800 hover:text-teal-800 hover:underline text-sm"
                      >
                        {c.name}
                      </button>
                      <span className="text-xs text-stone-400">({c.count} tickets)</span>
                    </div>
                    <div className="flex items-center gap-4 text-xs text-stone-500">
                      <span>Sales: <span className="font-semibold text-stone-700">{fmt(c.total)}</span></span>
                      <span>Profit: <span className="font-semibold text-emerald-700">{fmt(c.profit)}</span></span>
                    </div>
                  </div>
                  <p className="text-xs text-stone-500 mt-1.5 pl-6">
                    Customers: {c.customers.join(", ")}
                  </p>
                </div>
              ))}
            </div>
          </div>
        )}

        <p className="text-xs text-stone-400 mt-3">
          This data is shared between signed-in employees. Login is a basic access gate, not strong security — treat it accordingly.
        </p>
        </>
        )}

        {activeSection === "hotels" && (
        <>
        {/* Buttons to register new supplier names and hotel names, so they're always
            available to pick from the Supplier / Hotel name fields below. */}
        <div className="flex flex-wrap items-center gap-2 mb-4">
          <button
            onClick={() => { setShowAddSupplierPanel(!showAddSupplierPanel); setShowAddHotelNamePanel(false); }}
            className="text-xs font-semibold text-teal-800 border border-teal-700 rounded-xl px-3 py-2 hover:bg-teal-50 flex items-center gap-1.5"
          >
            <Plus size={14} /> Add supplier
          </button>
          <button
            onClick={() => { setShowAddHotelNamePanel(!showAddHotelNamePanel); setShowAddSupplierPanel(false); }}
            className="text-xs font-semibold text-teal-800 border border-teal-700 rounded-xl px-3 py-2 hover:bg-teal-50 flex items-center gap-1.5"
          >
            <Plus size={14} /> Add hotel name
          </button>
        </div>

        {showAddSupplierPanel && (
          <div className="bg-white border border-stone-200 rounded-2xl p-4 mb-4">
            <h3 className="text-sm font-bold text-stone-700 mb-3">Suppliers</h3>
            <div className="flex gap-2 mb-3">
              <input
                className="w-full max-w-xs border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={newSupplierDraft}
                onChange={(e) => setNewSupplierDraft(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleAddSupplierName()}
                placeholder="Supplier name"
              />
              <button
                onClick={handleAddSupplierName}
                className="bg-gradient-to-b from-teal-700 to-teal-900 text-white text-sm font-semibold rounded-xl px-4 py-2 hover:brightness-110"
              >
                Add
              </button>
            </div>
            {suggestions.suppliers.length === 0 ? (
              <p className="text-xs text-stone-400">No suppliers saved yet</p>
            ) : (
              <div className="flex flex-wrap gap-2">
                {suggestions.suppliers.map((s) => (
                  <span
                    key={s}
                    className="inline-flex items-center gap-1.5 bg-stone-50 border border-stone-200 rounded-lg px-2.5 py-1 text-xs text-stone-700"
                  >
                    {s}
                    <button onClick={() => handleDeleteSupplierName(s)} className="text-red-500 hover:text-red-700">
                      <Trash2 size={12} />
                    </button>
                  </span>
                ))}
              </div>
            )}
          </div>
        )}

        {showAddHotelNamePanel && (
          <div className="bg-white border border-stone-200 rounded-2xl p-4 mb-4">
            <h3 className="text-sm font-bold text-stone-700 mb-3">Hotel names</h3>
            <div className="flex gap-2 mb-3">
              <input
                className="w-full max-w-xs border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={newHotelNameDraft}
                onChange={(e) => setNewHotelNameDraft(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleAddHotelName()}
                placeholder="Hotel name"
              />
              <button
                onClick={handleAddHotelName}
                className="bg-gradient-to-b from-teal-700 to-teal-900 text-white text-sm font-semibold rounded-xl px-4 py-2 hover:brightness-110"
              >
                Add
              </button>
            </div>
            {suggestions.hotelNames.length === 0 ? (
              <p className="text-xs text-stone-400">No hotel names saved yet</p>
            ) : (
              <div className="flex flex-wrap gap-2">
                {suggestions.hotelNames.map((hn) => (
                  <span
                    key={hn}
                    className="inline-flex items-center gap-1.5 bg-stone-50 border border-stone-200 rounded-lg px-2.5 py-1 text-xs text-stone-700"
                  >
                    {hn}
                    <button onClick={() => handleDeleteHotelName(hn)} className="text-red-500 hover:text-red-700">
                      <Trash2 size={12} />
                    </button>
                  </span>
                ))}
              </div>
            )}
          </div>
        )}

        {/* USD -> EGP exchange rate bar — entered by hand each day (e.g. from the CBE's
            published rate), saved to shared storage so every employee sees the same value. */}
        <div className="bg-white border border-stone-200 rounded-2xl px-4 py-3 mb-4 flex flex-wrap items-center gap-3">
          <span className="text-xs font-semibold text-stone-500">USD → EGP rate today:</span>
          <input
            type="number"
            step="0.01"
            className="w-28 border border-stone-300 rounded-lg px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
            value={usdToEgpRate ?? ""}
            onChange={(e) => setUsdToEgpRate(e.target.value === "" ? null : parseFloat(e.target.value))}
            onBlur={() => {
              if (usdToEgpRate !== null && !Number.isNaN(usdToEgpRate)) persistUsdRate(usdToEgpRate);
            }}
            placeholder="e.g. 51.20"
          />
          <button
            onClick={() => usdToEgpRate !== null && !Number.isNaN(usdToEgpRate) && persistUsdRate(usdToEgpRate)}
            className="text-xs font-semibold text-teal-800 border border-teal-700 rounded-lg px-3 py-1.5 hover:bg-teal-50"
          >
            Save rate
          </button>
          {usdToEgpRateDate && (
            <span className="text-xs text-stone-400">Last updated: {formatDisplayDate(usdToEgpRateDate)}</span>
          )}
        </div>

        {hotelError && (
          <div className="text-sm rounded-xl px-3 py-2 mb-4 bg-red-50 text-red-700">{hotelError}</div>
        )}

        {canAddTickets && (
          <div className="bg-white border border-stone-200 rounded-2xl p-5 mb-6">
            <h3 className="text-sm font-bold text-stone-700 mb-4">
              {hotelEditingId ? "Edit hotel booking" : "New hotel booking"}
            </h3>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
              <div>
                <label className="text-xs text-stone-500 block mb-1">
                  Company name <span className="font-normal text-stone-400">(optional — leave blank for Individual)</span>
                </label>
                <input
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={hotelForm.customer}
                  onChange={(e) => setHotelForm({ ...hotelForm, customer: e.target.value })}
                  placeholder="e.g. Perla Travel Corp — leave blank for Individual"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Hotel name</label>
                {hotelNameOther ? (
                  <div className="flex gap-2">
                    <input
                      className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                      value={hotelForm.hotel}
                      onChange={(e) => setHotelForm({ ...hotelForm, hotel: e.target.value })}
                      placeholder="Enter hotel name"
                      autoFocus
                    />
                    <button
                      type="button"
                      onClick={() => { setHotelNameOther(false); setHotelForm({ ...hotelForm, hotel: "" }); }}
                      className="shrink-0 text-xs text-stone-500 hover:text-teal-800 border border-stone-300 rounded-xl px-2"
                    >
                      List
                    </button>
                  </div>
                ) : (
                  <select
                    className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white"
                    value={hotelForm.hotel}
                    onChange={(e) => {
                      if (e.target.value === "__other__") {
                        setHotelNameOther(true);
                        setHotelForm({ ...hotelForm, hotel: "" });
                      } else {
                        setHotelForm({ ...hotelForm, hotel: e.target.value });
                      }
                    }}
                  >
                    <option value="">Select hotel</option>
                    {suggestions.hotelNames.map((hn) => (
                      <option key={hn} value={hn}>{hn}</option>
                    ))}
                    <option value="__other__">Other</option>
                  </select>
                )}
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Supplier</label>
                {hotelSupplierOther ? (
                  <div className="flex gap-2">
                    <input
                      className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                      value={hotelForm.supplier}
                      onChange={(e) => setHotelForm({ ...hotelForm, supplier: e.target.value })}
                      placeholder="Enter supplier name"
                      autoFocus
                    />
                    <button
                      type="button"
                      onClick={() => { setHotelSupplierOther(false); setHotelForm({ ...hotelForm, supplier: "" }); }}
                      className="shrink-0 text-xs text-stone-500 hover:text-teal-800 border border-stone-300 rounded-xl px-2"
                    >
                      List
                    </button>
                  </div>
                ) : (
                  <select
                    className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white"
                    value={hotelForm.supplier}
                    onChange={(e) => {
                      if (e.target.value === "__other__") {
                        setHotelSupplierOther(true);
                        setHotelForm({ ...hotelForm, supplier: "" });
                      } else {
                        setHotelForm({ ...hotelForm, supplier: e.target.value });
                      }
                    }}
                  >
                    <option value="">Select supplier</option>
                    {suggestions.suppliers.map((s) => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                    <option value="__other__">Other</option>
                  </select>
                )}
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Booking date</label>
                <input
                  type="date"
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={hotelForm.bookingDate}
                  onChange={(e) => setHotelForm({ ...hotelForm, bookingDate: e.target.value })}
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Notes</label>
                <input
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={hotelForm.notes}
                  onChange={(e) => setHotelForm({ ...hotelForm, notes: e.target.value })}
                />
              </div>
            </div>

            <p className="text-xs text-stone-500 mb-3">
              Each room has its own check-in/check-out dates — price is per room, per night.
            </p>

            {/* Room lines: one booking can mix different room types, meal plans, currencies,
                prices, and stay dates — each room keeps its own check-in/check-out. */}
            <div className="space-y-3">
              <label className="text-xs text-stone-500 block">Rooms</label>
              {hotelForm.roomLines.map((line) => (
                <div key={line.id} className="bg-stone-50 border border-stone-200 rounded-xl p-3 space-y-3">
                  {/* Row 1: room type, meal plan, dates, currency. */}
                  <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-5 gap-3 items-end">
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1">Room type</label>
                      <select
                        className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={line.roomType}
                        onChange={(e) => {
                          const roomType = e.target.value;
                          const capacity = ROOM_CAPACITY[roomType] || 1;
                          updateHotelRoomLine(line.id, { roomType, guests: guestsForCapacity(line.guests, capacity) });
                        }}
                      >
                        {ROOM_TYPES.map((r) => (
                          <option key={r.value} value={r.value}>{r.label}</option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1">Meal plan</label>
                      <select
                        className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={line.mealPlan}
                        onChange={(e) => updateHotelRoomLine(line.id, { mealPlan: e.target.value })}
                      >
                        {MEAL_PLANS.map((m) => (
                          <option key={m.value} value={m.value}>{m.label}</option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1">Check-in</label>
                      <input
                        type="date"
                        className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={line.checkIn}
                        onChange={(e) => updateHotelRoomLine(line.id, { checkIn: e.target.value })}
                      />
                    </div>
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1">Check-out</label>
                      <input
                        type="date"
                        min={line.checkIn || undefined}
                        className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={line.checkOut}
                        onChange={(e) => updateHotelRoomLine(line.id, { checkOut: e.target.value })}
                      />
                    </div>
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1">Currency</label>
                      <select
                        className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={line.currency}
                        onChange={(e) => updateHotelRoomLine(line.id, { currency: e.target.value })}
                      >
                        {HOTEL_CURRENCIES.map((c) => (
                          <option key={c.value} value={c.value}>{c.label}</option>
                        ))}
                      </select>
                    </div>
                  </div>

                  {/* Row 2: # rooms, net, sold. */}
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-3 items-start">
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1"># rooms</label>
                      <input
                        type="number"
                        min="1"
                        className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={line.count}
                        onChange={(e) => updateHotelRoomLine(line.id, { count: e.target.value })}
                      />
                    </div>
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1">Net (per room/night)</label>
                      <input
                        type="number"
                        className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={line.netPrice}
                        onChange={(e) => updateHotelRoomLine(line.id, { netPrice: e.target.value })}
                        placeholder="0"
                      />
                    </div>
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1">Sold (per room/night)</label>
                      <input
                        type="number"
                        className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={line.soldPrice}
                        onChange={(e) => updateHotelRoomLine(line.id, { soldPrice: e.target.value })}
                        placeholder="0"
                      />
                      <div className="flex items-center justify-between gap-2 mt-3">
                        <div className="text-xs text-emerald-700 font-semibold">
                          {roomLineNights(line, hotelForm)}n · {fmt(hotelLineSoldTotal(line, roomLineNights(line, hotelForm)) - hotelLineNetTotal(line, roomLineNights(line, hotelForm)))} {line.currency}
                        </div>
                        <button
                          onClick={() => removeHotelRoomLine(line.id)}
                          disabled={hotelForm.roomLines.length <= 1}
                          className="text-red-500 hover:text-red-700 disabled:opacity-30"
                          title="Remove this room line"
                        >
                          <Trash2 size={16} />
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* Adult guest names — one field per bed the room type holds, placed
                      directly above the Children section. Only the first guest is
                      mandatory; the rest are optional. */}
                  <div className="space-y-2">
                    {(line.guests || []).map((g, i) => (
                      <div key={g.id} className="bg-white border border-stone-200 rounded-lg p-2">
                        <label className="text-[11px] text-stone-500 block mb-1">
                          Guest {i + 1} name
                          {i === 0 ? <span className="text-red-500"> *</span> : (
                            <span className="text-stone-400"> (optional)</span>
                          )}
                        </label>
                        <input
                          className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                          value={g.name}
                          onChange={(e) => updateRoomGuest(line.id, i, e.target.value)}
                          placeholder={i === 0 ? "Guest 1 name (required)" : `Guest ${i + 1} name`}
                        />
                      </div>
                    ))}
                  </div>

                  {/* Children in this room — name + age in years (0–11). */}
                  <div>
                    <div className="flex items-center justify-between mb-1">
                      <label className="text-[11px] text-stone-500 block">Children</label>
                      <button
                        type="button"
                        onClick={() => addRoomChild(line.id)}
                        className="text-[11px] font-semibold text-teal-800 border border-teal-700 border-dashed rounded-lg px-2 py-1 hover:bg-teal-50"
                      >
                        + Add child
                      </button>
                    </div>
                    {(line.children || []).length > 0 && (
                      <div className="space-y-2">
                        {line.children.map((c, i) => (
                          <div key={c.id} className="grid grid-cols-1 sm:grid-cols-8 gap-3 items-end bg-white border border-stone-200 rounded-lg p-3">
                            <div className="sm:col-span-6">
                              <label className="text-[11px] text-stone-500 block mb-1">Child {i + 1} name</label>
                              <input
                                className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                                value={c.name}
                                onChange={(e) => updateRoomChild(line.id, c.id, { name: e.target.value })}
                                placeholder="Child name"
                              />
                            </div>
                            <div>
                              <label className="text-[11px] text-stone-500 block mb-1">Age (0–11)</label>
                              <input
                                type="text"
                                inputMode="numeric"
                                pattern="[0-9]*"
                                className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                                value={c.age}
                                onChange={(e) => updateRoomChild(line.id, c.id, { age: sanitizeAgeInput(e.target.value) })}
                                placeholder="e.g. 4"
                              />
                            </div>
                            <div className="flex justify-end">
                              <button
                                onClick={() => removeRoomChild(line.id, c.id)}
                                className="text-red-500 hover:text-red-700"
                                title="Remove this child"
                              >
                                <Trash2 size={14} />
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              ))}
              <button
                onClick={addHotelRoomLine}
                className="text-xs font-semibold text-teal-800 border border-teal-700 border-dashed rounded-lg px-3 py-1.5 hover:bg-teal-50"
              >
                + Add another room
              </button>
            </div>

            <div className="grid grid-cols-3 gap-3 mt-4">
              <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                <p className="text-[11px] text-stone-500">Net total (EGP)</p>
                <p className="text-sm font-bold text-stone-800">{fmt(hotelNetTotal(hotelForm))}</p>
              </div>
              <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                <p className="text-[11px] text-stone-500">Sold total (EGP)</p>
                <p className="text-sm font-bold text-stone-800">{fmt(hotelSoldTotal(hotelForm))}</p>
              </div>
              <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                <p className="text-[11px] text-stone-500">Profit (auto, EGP)</p>
                <p className="text-sm font-bold text-emerald-700">{fmt(hotelProfitTotal(hotelForm))}</p>
              </div>
            </div>
            <div className="flex items-center gap-3 mt-4">
              <button
                onClick={handleSaveHotel}
                className="bg-gradient-to-b from-teal-700 to-teal-900 text-white text-sm font-semibold rounded-xl px-5 py-2.5 hover:brightness-110"
              >
                {hotelEditingId ? "Save changes" : "Add booking"}
              </button>
              {hotelEditingId && (
                <button
                  onClick={resetHotelForm}
                  className="text-sm font-semibold text-stone-500 rounded-xl px-4 py-2.5 hover:bg-stone-50"
                >
                  Cancel
                </button>
              )}
            </div>
          </div>
        )}

        {/* Totals, converted to EGP for USD-priced bookings */}
        <div className="grid grid-cols-3 gap-3 mb-4">
          <div className="bg-white border border-stone-200 rounded-2xl px-4 py-3 text-center">
            <p className="text-xs text-stone-500">Total net (EGP)</p>
            <p className="text-lg font-bold text-stone-800">{fmt(hotelTotals.net)}</p>
          </div>
          <div className="bg-white border border-stone-200 rounded-2xl px-4 py-3 text-center">
            <p className="text-xs text-stone-500">Total sold (EGP)</p>
            <p className="text-lg font-bold text-stone-800">{fmt(hotelTotals.sold)}</p>
          </div>
          <div className="bg-white border border-stone-200 rounded-2xl px-4 py-3 text-center">
            <p className="text-xs text-stone-500">Total profit (EGP)</p>
            <p className="text-lg font-bold text-emerald-700">{fmt(hotelTotals.profit)}</p>
          </div>
        </div>

        <div className="bg-white border border-stone-200 rounded-2xl overflow-x-auto">
          <table className="w-full text-xs border-collapse">
            <thead>
              <tr className="bg-stone-50 border-b border-stone-200 text-stone-500">
                <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Company</th>
                <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Hotel</th>
                <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Supplier</th>
                <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Rooms</th>
                <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap"># rooms</th>
                <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Booking date</th>
                <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Dates</th>
                <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Net total (EGP)</th>
                <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Sold total (EGP)</th>
                <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Profit (EGP)</th>
                {(canEditTickets || canDeleteTickets) && (
                  <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Actions</th>
                )}
              </tr>
            </thead>
            <tbody>
              {visibleHotelBookings.length === 0 && (
                <tr>
                  <td colSpan={11} className="text-center text-stone-400 px-2.5 py-6">
                    No hotel bookings yet.
                  </td>
                </tr>
              )}
              {visibleHotelBookings.map((h) => (
                <tr
                  key={h.id}
                  className="border-b border-stone-100 hover:bg-stone-50 cursor-pointer"
                  onClick={() => setViewingHotelBooking(h)}
                >
                  <td className="px-2.5 py-1 text-stone-700 whitespace-nowrap">
                    {h.customer && h.customer.trim() ? (
                      h.customer
                    ) : (
                      <span className="text-stone-400 italic">Individual</span>
                    )}
                  </td>
                  <td className="px-2.5 py-1 text-stone-700 whitespace-nowrap">{h.hotel}</td>
                  <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">{h.supplier}</td>
                  <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">{hotelLinesSummary(h)}</td>
                  <td className="px-2.5 py-1 text-stone-600 text-right whitespace-nowrap">{hotelRoomCount(h)}</td>
                  <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">
                    {h.bookingDate ? formatDisplayDate(h.bookingDate) : "-"}
                  </td>
                  <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">
                    {hotelDateRange(h).start && hotelDateRange(h).end
                      ? `${formatDisplayDate(hotelDateRange(h).start)} → ${formatDisplayDate(hotelDateRange(h).end)}`
                      : "-"}
                  </td>
                  <td className="px-2.5 py-1 text-stone-600 text-right whitespace-nowrap">{fmt(hotelNetTotal(h))}</td>
                  <td className="px-2.5 py-1 text-stone-600 text-right whitespace-nowrap">{fmt(hotelSoldTotal(h))}</td>
                  <td className="px-2.5 py-1 font-semibold text-emerald-700 text-right whitespace-nowrap">
                    {fmt(hotelProfitTotal(h))}
                  </td>
                  {(canEditTickets || canDeleteTickets) && (
                    <td className="px-2.5 py-1 text-right whitespace-nowrap" onClick={(e) => e.stopPropagation()}>
                      {canEditTickets && (
                        <button
                          onClick={() => handleEditHotelClick(h)}
                          className="text-teal-700 hover:text-teal-900 mr-2"
                          title="Edit"
                        >
                          <Pencil size={14} />
                        </button>
                      )}
                      {canDeleteTickets && (
                        <button
                          onClick={() => handleDeleteHotel(h.id)}
                          className="text-red-600 hover:text-red-800"
                          title="Delete"
                        >
                          <Trash2 size={14} />
                        </button>
                      )}
                    </td>
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {viewingHotelBooking && (
          <div
            className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4"
            onClick={() => setViewingHotelBooking(null)}
          >
            <div
              className="bg-white rounded-2xl max-w-3xl w-full max-h-[85vh] overflow-y-auto p-6"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="flex items-start justify-between mb-4">
                <div>
                  <h3 className="text-lg font-bold text-stone-800">{viewingHotelBooking.hotel}</h3>
                  <p className="text-sm text-stone-500">
                    {viewingHotelBooking.customer && viewingHotelBooking.customer.trim() ? (
                      <>Company: {viewingHotelBooking.customer} <span className="text-teal-700 font-semibold">(Corporate)</span></>
                    ) : (
                      <span className="italic">Individual booking</span>
                    )}
                  </p>
                </div>
                <button
                  onClick={() => setViewingHotelBooking(null)}
                  className="text-stone-400 hover:text-stone-700"
                  title="Close"
                >
                  <X size={20} />
                </button>
              </div>

              <div className="grid grid-cols-2 gap-3 mb-4 text-sm">
                <div><span className="text-stone-500">Supplier: </span>{viewingHotelBooking.supplier || "-"}</div>
                <div><span className="text-stone-500">Booking date: </span>{viewingHotelBooking.bookingDate ? formatDisplayDate(viewingHotelBooking.bookingDate) : "-"}</div>
                <div><span className="text-stone-500">Employee: </span>{viewingHotelBooking.employee || "-"}</div>
                <div><span className="text-stone-500">Notes: </span>{viewingHotelBooking.notes || "-"}</div>
              </div>

              <div className="space-y-3">
                {(viewingHotelBooking.roomLines || []).map((l, idx) => {
                  const type = ROOM_TYPES.find((r) => r.value === l.roomType)?.label || l.roomType;
                  const meal = MEAL_PLANS.find((m) => m.value === l.mealPlan)?.label || l.mealPlan;
                  const nights = roomLineNights(l, viewingHotelBooking);
                  return (
                    <div key={l.id || idx} className="border border-stone-200 rounded-xl p-3">
                      <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
                        <span className="font-semibold text-stone-700 text-sm">
                          {l.count}× {type} — {meal}
                        </span>
                        <span className="text-xs text-stone-500">
                          {l.checkIn ? formatDisplayDate(l.checkIn) : "-"} → {l.checkOut ? formatDisplayDate(l.checkOut) : "-"} ({nights}n)
                        </span>
                      </div>
                      <div className="text-xs text-stone-600 mb-2">
                        Net: {fmt(hotelLineNetTotal(l, nights))} {l.currency} · Sold:{" "}
                        {fmt(hotelLineSoldTotal(l, nights))} {l.currency}
                      </div>
                      {Array.isArray(l.guests) && l.guests.some((g) => g.name) && (
                        <div className="text-xs text-stone-700 mb-1">
                          <span className="text-stone-500">Guests: </span>
                          {l.guests.map((g) => g.name || "-").join(", ")}
                        </div>
                      )}
                      {Array.isArray(l.children) && l.children.length > 0 && (
                        <div className="text-xs text-stone-700">
                          <span className="text-stone-500">Children: </span>
                          {l.children
                            .map((c) => `${c.name || "-"} (${c.age !== "" && c.age != null ? c.age : "-"}y)`)
                            .join(", ")}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              <div className="grid grid-cols-3 gap-3 mt-4">
                <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                  <p className="text-[11px] text-stone-500">Net total (EGP)</p>
                  <p className="text-sm font-bold text-stone-800">{fmt(hotelNetTotal(viewingHotelBooking))}</p>
                </div>
                <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                  <p className="text-[11px] text-stone-500">Sold total (EGP)</p>
                  <p className="text-sm font-bold text-stone-800">{fmt(hotelSoldTotal(viewingHotelBooking))}</p>
                </div>
                <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                  <p className="text-[11px] text-stone-500">Profit (EGP)</p>
                  <p className="text-sm font-bold text-emerald-700">{fmt(hotelProfitTotal(viewingHotelBooking))}</p>
                </div>
              </div>
            </div>
          </div>
        )}
        </>
        )}

        {activeSection === "cars" && (
          <div className="bg-white border border-stone-200 rounded-2xl p-12 text-center text-stone-400">
            <Car size={40} className="mx-auto mb-3 text-stone-300" />
            <p className="text-sm">Transportation section — nothing here yet.</p>
          </div>
        )}

        {activeSection === "files" && (
          <div className="bg-white border border-stone-200 rounded-2xl p-12 text-center text-stone-400">
            <FileText size={40} className="mx-auto mb-3 text-stone-300" />
            <p className="text-sm">Files section — nothing here yet.</p>
          </div>
        )}
        </>
        )}
      </div>

      {activeSection === "flights" && viewingTicket && (
        <div className="fixed inset-0 bg-white z-40 overflow-y-auto">
          <div className="max-w-3xl mx-auto p-4 md:p-6">
            <div className="flex items-center justify-between mb-6">
              <div className="flex items-center gap-2">
                <div className="bg-gradient-to-b from-teal-700 to-teal-900 text-white rounded-xl p-2 shadow-sm shadow-teal-800/30">
                  <Ticket size={18} />
                </div>
                <h1 className="text-lg md:text-xl font-bold text-stone-900" style={{ fontFamily: "'Fraunces', serif" }}>Ticket details</h1>
              </div>
              <button
                onClick={closeTicketDetail}
                className="border border-stone-300 text-stone-600 text-sm rounded-xl px-3 py-2 flex items-center gap-1.5"
              >
                <X size={15} /> Close
              </button>
            </div>

            <div className="bg-white border border-stone-200 rounded-2xl divide-y divide-stone-100">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 p-4 md:p-5">
                <div>
                  <p className="text-xs text-stone-400 mb-1">Entered by</p>
                  <p className="text-sm font-medium text-stone-800">{viewingTicket.employee || "-"}</p>
                </div>
                <div>
                  <p className="text-xs text-stone-400 mb-1">Company</p>
                  <p className="text-sm font-medium text-stone-800">
                    {viewingTicket.company && viewingTicket.company.trim() ? (
                      <>{viewingTicket.company} <span className="text-teal-700 font-semibold">(Corporate)</span></>
                    ) : (
                      <span className="text-stone-400 italic">Individual</span>
                    )}
                  </p>
                </div>
                <div>
                  <p className="text-xs text-stone-400 mb-1">Supplier</p>
                  <p className="text-sm font-medium text-stone-800">{viewingTicket.supplier || "-"}</p>
                </div>
                <div>
                  <p className="text-xs text-stone-400 mb-1">Route</p>
                  <p className="text-sm font-medium text-stone-800">{routeLabel(viewingTicket)}</p>
                </div>
                <div>
                  <p className="text-xs text-stone-400 mb-1">Airline</p>
                  <p className="text-sm font-medium text-stone-800" title={getAirlineNameByIata(viewingTicket.airline) || viewingTicket.airline || ""}>
                    {viewingTicket.airline ? (getAirlineIata(viewingTicket.airline) || viewingTicket.airline) : "-"}
                  </p>
                </div>
                <div>
                  <p className="text-xs text-stone-400 mb-1">Ticket issue date</p>
                  <p className="text-sm font-medium text-stone-800">
                    {viewingTicket.date ? formatDisplayDate(viewingTicket.date) : "-"}
                  </p>
                </div>
                {viewingTicket.isReissued && (
                  <div className="sm:col-span-2 bg-amber-50 border border-amber-200 rounded-xl px-3 py-2">
                    <p className="text-xs font-semibold text-amber-800 mb-1">Reissued ticket</p>
                    <p className="text-sm text-amber-900">
                      Old ticket number: {viewingTicket.oldTicketNumber || "-"}
                      {" · "}
                      Old issue date:{" "}
                      {viewingTicket.oldTicketIssueDate
                        ? formatDisplayDate(viewingTicket.oldTicketIssueDate)
                        : "not found"}
                    </p>
                  </div>
                )}
              </div>

              <div className="p-4 md:p-5">
                <p className="text-xs text-stone-400 mb-2">
                  Customers ({getCustomers(viewingTicket).length})
                </p>
                <div className="border border-stone-200 rounded-xl overflow-hidden">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="bg-stone-50 text-stone-500 text-xs">
                        <th className="text-left px-3 py-2 font-medium">Customer</th>
                        <th className="text-left px-3 py-2 font-medium">Ticket number</th>
                      </tr>
                    </thead>
                    <tbody>
                      {getCustomers(viewingTicket).map((c, i) => (
                        <tr key={i} className="border-t border-stone-100">
                          <td className="px-3 py-2 text-stone-700">
                            {c.name || "-"}
                            {hasRefund(viewingTicket) && (viewingTicket.refund.customerIndex || 0) === i && (
                              <span className="ml-2 inline-block text-[10px] font-semibold text-sky-700 bg-sky-100 rounded-full px-2 py-0.5 align-middle">
                                Refunded
                              </span>
                            )}
                          </td>
                          <td className="px-3 py-2 text-stone-700 font-mono">{c.ticketNumber || "-"}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-3 gap-4 p-4 md:p-5">
                <div>
                  <p className="text-xs text-stone-400 mb-1">Net price</p>
                  <p className="text-sm font-medium text-stone-800">{fmt(netAfterRefund(viewingTicket))}</p>
                  {hasRefund(viewingTicket) && (
                    <p className="text-[11px] text-stone-400">Before refund: {fmt(viewingTicket.netPrice)}</p>
                  )}
                </div>
                <div>
                  <p className="text-xs text-stone-400 mb-1">Sold price</p>
                  <p className="text-sm font-medium text-stone-800">{fmt(soldAfterRefund(viewingTicket))}</p>
                  {hasRefund(viewingTicket) && (
                    <p className="text-[11px] text-stone-400">Before refund: {fmt(viewingTicket.soldPrice)}</p>
                  )}
                </div>
                <div>
                  <p className="text-xs text-stone-400 mb-1">Profit</p>
                  <p className="text-sm font-semibold text-emerald-700">
                    {fmt(profitAfterRefund(viewingTicket))}
                  </p>
                  {hasRefund(viewingTicket) && (
                    <p className="text-[11px] text-stone-400">
                      Before refund: {fmt(profit(viewingTicket.netPrice, viewingTicket.soldPrice))}
                    </p>
                  )}
                </div>
              </div>

              {/* Refund: two amounts (refunded by the airline, refunded to the customer),
                  recorded separately from the ticket itself and shown as its own row
                  directly under this ticket in the exported report. */}
              <div className="p-4 md:p-5">
                <div className="flex items-center justify-between mb-2">
                  <p className="text-xs text-stone-400">Refund</p>
                  {!showRefundForm && (
                    <button
                      type="button"
                      onClick={() => setShowRefundForm(true)}
                      className="inline-flex items-center gap-1 text-xs font-semibold text-teal-700 hover:text-teal-900"
                    >
                      <Wallet size={13} /> {hasRefund(viewingTicket) ? "Edit refund" : "Add refund"}
                    </button>
                  )}
                </div>

                {!showRefundForm && (
                  hasRefund(viewingTicket) ? (
                    <div className="bg-sky-50 border border-sky-200 rounded-xl px-3 py-2 flex flex-wrap gap-4 text-sm">
                      {getCustomers(viewingTicket).length > 1 && (
                        <span className="w-full">
                          <span className="text-xs text-sky-500 block">Refunded ticket</span>
                          <span className="text-sky-900 font-medium">
                            {(() => {
                              const idx = viewingTicket.refund.customerIndex || 0;
                              const c = getCustomers(viewingTicket)[idx];
                              return c ? (c.name || `Customer ${idx + 1}`) + (c.ticketNumber ? ` — ${c.ticketNumber}` : "") : `Customer ${idx + 1}`;
                            })()}
                          </span>
                        </span>
                      )}
                      <span>
                        <span className="text-xs text-sky-500 block">Refunded by airline</span>
                        <span className="text-sky-900 font-medium">{fmt(viewingTicket.refund.airlineAmount)}</span>
                      </span>
                      <span>
                        <span className="text-xs text-sky-500 block">Refunded to customer</span>
                        <span className="text-sky-900 font-medium">{fmt(viewingTicket.refund.customerAmount)}</span>
                      </span>
                      {viewingTicket.refund.date && (
                        <span>
                          <span className="text-xs text-sky-500 block">Refund date</span>
                          <span className="text-sky-900 font-medium">{formatDisplayDate(viewingTicket.refund.date)}</span>
                        </span>
                      )}
                    </div>
                  ) : (
                    <p className="text-xs text-stone-400 italic">No refund recorded</p>
                  )
                )}

                {showRefundForm && (
                  <div className="bg-sky-50 border border-sky-200 rounded-xl p-3">
                    {getCustomers(viewingTicket).length > 1 && (
                      <div className="mb-3">
                        <label className="text-xs text-stone-500 block mb-1">Refunded ticket</label>
                        <select
                          className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white"
                          value={refundDraft.customerIndex}
                          onChange={(e) => setRefundDraft({ ...refundDraft, customerIndex: Number(e.target.value) })}
                        >
                          {getCustomers(viewingTicket).map((c, i) => (
                            <option key={i} value={i}>
                              {(c.name || `Customer ${i + 1}`) + (c.ticketNumber ? ` — ${c.ticketNumber}` : "")}
                            </option>
                          ))}
                        </select>
                      </div>
                    )}
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <label className="text-xs text-stone-500 block mb-1">Refunded by airline</label>
                        <input
                          type="number"
                          className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                          value={refundDraft.airlineAmount}
                          onChange={(e) => setRefundDraft({ ...refundDraft, airlineAmount: e.target.value })}
                          placeholder="0"
                        />
                      </div>
                      <div>
                        <label className="text-xs text-stone-500 block mb-1">Refunded to customer</label>
                        <input
                          type="number"
                          className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                          value={refundDraft.customerAmount}
                          onChange={(e) => setRefundDraft({ ...refundDraft, customerAmount: e.target.value })}
                          placeholder="0"
                        />
                      </div>
                    </div>
                    <div className="flex items-center gap-2 mt-3">
                      <button
                        onClick={() => saveTicketRefund(viewingTicket.id)}
                        className="bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 flex items-center gap-1.5 shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10"
                      >
                        <Check size={15} /> Save refund
                      </button>
                      <button
                        onClick={() => {
                          setShowRefundForm(false);
                          setRefundDraft({
                            airlineAmount: viewingTicket.refund ? viewingTicket.refund.airlineAmount : "",
                            customerAmount: viewingTicket.refund ? viewingTicket.refund.customerAmount : "",
                            customerIndex: viewingTicket.refund && viewingTicket.refund.customerIndex != null ? viewingTicket.refund.customerIndex : 0,
                          });
                        }}
                        className="border border-stone-300 text-stone-600 text-sm rounded-xl px-4 py-2"
                      >
                        Cancel
                      </button>
                      {hasRefund(viewingTicket) && (
                        <button
                          onClick={() => clearTicketRefund(viewingTicket.id)}
                          className="text-xs text-red-600 hover:text-red-800 font-semibold ml-auto"
                        >
                          Remove refund
                        </button>
                      )}
                    </div>
                  </div>
                )}
                {refundSaved && !showRefundForm && (
                  <span className="text-xs text-emerald-700 font-medium block mt-2">Saved</span>
                )}
              </div>

              <div className="p-4 md:p-5">
                <p className="text-xs text-stone-400 mb-2">Notes</p>
                <textarea
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 min-h-[100px]"
                  value={notesDraft}
                  onChange={(e) => { setNotesDraft(e.target.value.toUpperCase()); setNotesSaved(false); }}
                  placeholder="No notes yet — add some here"
                />
                <div className="flex items-center gap-2 mt-3">
                  <button
                    onClick={() => saveTicketNotes(viewingTicket.id)}
                    disabled={notesDraft === (viewingTicket.notes || "")}
                    className={`text-sm font-semibold rounded-xl px-4 py-2 flex items-center gap-1.5 transition-colors ${
                      notesDraft === (viewingTicket.notes || "")
                        ? "bg-stone-200 text-stone-400 cursor-not-allowed"
                        : "bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10"
                    }`}
                  >
                    <Check size={15} /> Save notes
                  </button>
                  {notesSaved && (
                    <span className="text-xs text-emerald-700 font-medium">Saved</span>
                  )}
                </div>

                {Array.isArray(viewingTicket.notesHistory) && viewingTicket.notesHistory.length > 0 && (
                  <div className="mt-4 pt-3 border-t border-stone-100">
                    <p className="text-xs text-stone-400 mb-2">Edit history (most recent first)</p>
                    <div className="space-y-1.5 max-h-48 overflow-y-auto">
                      {[...viewingTicket.notesHistory].reverse().map((h, idx) => (
                        <div
                          key={idx}
                          className="text-xs bg-stone-50 border border-stone-100 rounded-xl px-2.5 py-1.5 flex items-start justify-between gap-3"
                        >
                          {h.type === "edit" ? (
                            <span className="text-stone-600 break-words">
                              <span className="font-semibold text-stone-700">Ticket edited: </span>
                              {(h.changes || []).join("; ")}
                            </span>
                          ) : (
                            <span className="text-stone-600 break-words">{h.value || "(cleared)"}</span>
                          )}
                          <span className="text-stone-400 whitespace-nowrap shrink-0">
                            {h.by} · {formatDateTime(h.at)}
                          </span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {confirmDialog && (
        <div className="fixed inset-0 bg-stone-900/40 flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl border border-stone-200 p-5 w-full max-w-sm">
            <p className="text-sm text-stone-700 mb-4">{confirmDialog.message}</p>
            <div className="flex justify-end gap-2">
              <button
                onClick={() => setConfirmDialog(null)}
                className="border border-stone-300 text-stone-600 text-sm rounded-xl px-3 py-2"
              >
                Cancel
              </button>
              <button
                onClick={() => confirmDialog.onConfirm()}
                className="bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-3 py-2 shadow-sm shadow-teal-800/30 transition-colors"
              >
                Confirm
              </button>
            </div>
          </div>
        </div>
      )}

      {openPermissionsFor && (
        <EmployeePermissionsModal
          emp={(employees || []).find((e) => e.username === openPermissionsFor && !e.isAdmin)}
          onClose={() => setOpenPermissionsFor(null)}
          onSetRole={(role) => handleRoleChange(openPermissionsFor, role)}
          onSetPermission={(field, value) => handleTogglePermission(openPermissionsFor, field, value)}
        />
      )}
    </div>
  );
}
