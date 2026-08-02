import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// Perla Di Mare logo, embedded as a data URI so the app works as a single self-contained file.
const LOGO_DATA_URL = "data:image/png;base64,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";
import {
  Plane, Search, Trash2, Pencil, X, Check, TrendingUp, Ticket, Wallet,
  Calendar, Download, Upload, Building2, Factory, Lock, LogOut, UserPlus, Users, Eye, EyeOff,
  ShieldCheck, Wifi, User, Cloud, Globe2, List, Car, FileText,
} from "lucide-react";

const MONTHS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

const monthKey = (dateStr) => (dateStr ? dateStr.slice(0, 7) : "No date");
// Storage stays in the native YYYY-MM-DD format (required by <input type="date">),
// but everywhere we display the date to the user we show it starting with the day.
const formatDisplayDate = (dateStr) => {
  if (!dateStr) return "";
  const [y, m, d] = dateStr.split("-");
  if (!y || !m || !d) return dateStr;
  return `${d}-${m}-${y}`;
};
const monthLabel = (key) => {
  if (key === "No date") return key;
  const [y, m] = key.split("-");
  const idx = parseInt(m, 10) - 1;
  return `${MONTHS[idx] || m} ${y}`;
};

// Formats an ISO timestamp as DD-MM-YYYY HH:MM for showing when a note edit happened.
const formatDateTime = (iso) => {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return "";
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  const hh = String(d.getHours()).padStart(2, "0");
  const min = String(d.getMinutes()).padStart(2, "0");
  return `${dd}-${mm}-${yyyy} ${hh}:${min}`;
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
  airline: "",
  date: todayDateStr(),
  netPrice: "",
  soldPrice: "",
  notes: "",
});

// Given a ticket number like "077-1234567890", returns the same prefix with the numeric
// part increased by one, keeping the same digit width (e.g. "077-1234567891").
// Returns "" if the ticket number doesn't match the expected PREFIX-DIGITS shape.
// Auto-sequencing only ever advances the LAST TWO digits of the serial number (wrapping
// 99 back to 00); everything before them — including the rest of the serial — stays fixed,
// since that part identifies the batch/booking rather than the individual ticket.
const nextTicketNumber = (ticketNumber) => {
  if (!ticketNumber) return "";
  const match = ticketNumber.match(/^([A-Z0-9]{3})-(\d+)$/);
  if (!match) return "";
  const [, prefix, digits] = match;
  if (digits.length <= 2) {
    const wrapped = ((parseInt(digits, 10) + 1) % (10 ** digits.length)).toString().padStart(digits.length, "0");
    return `${prefix}-${wrapped}`;
  }
  const head = digits.slice(0, -2);
  const tail = digits.slice(-2);
  const nextTail = ((parseInt(tail, 10) + 1) % 100).toString().padStart(2, "0");
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

const emptyNewEmployee = {
  name: "",
  username: "",
  password: "",
  // Default permissions for a newly created employee: can only see and add
  // their own tickets, cannot edit anything, and is not an accounting account.
  canViewAll: false,
  canEdit: false,
  isAccounting: false,
  canManageCompanies: false,
};

// IATA 3-digit airline accounting/ticketing prefix codes — the first 3 digits of a
// standard e-ticket number identify the issuing airline. Used to link the ticket
// number prefix with the Airline field automatically in both directions.
const AIRLINE_CODES = [
  { code: "001", name: "American Airlines" }, { code: "006", name: "Delta Air Lines" },
  { code: "014", name: "Air Canada" }, { code: "016", name: "United Airlines" },
  { code: "020", name: "Aeroflot" }, { code: "022", name: "Condor" },
  { code: "027", name: "Alaska Airlines" }, { code: "030", name: "Vueling" },
  { code: "044", name: "Aerolineas Argentinas" }, { code: "045", name: "LATAM Airlines" },
  { code: "050", name: "Olympic Air" }, { code: "053", name: "Aer Lingus" },
  { code: "055", name: "ITA Airways" }, { code: "057", name: "Air France" },
  { code: "065", name: "Saudia" }, { code: "071", name: "Ethiopian Airlines" },
  { code: "072", name: "Gulf Air" }, { code: "074", name: "KLM Royal Dutch Airlines" },
  { code: "075", name: "Iberia" }, { code: "076", name: "Middle East Airlines" },
  { code: "077", name: "EgyptAir" }, { code: "079", name: "Philippine Airlines" },
  { code: "080", name: "LOT Polish Airlines" }, { code: "081", name: "Qantas" },
  { code: "082", name: "Brussels Airlines" }, { code: "085", name: "Discover Airlines" },
  { code: "086", name: "Air New Zealand" }, { code: "087", name: "TAAG Angola Airlines" },
  { code: "098", name: "Air India" }, { code: "101", name: "Air Dolomiti" },
  { code: "104", name: "Eurowings" }, { code: "105", name: "Finnair" },
  { code: "108", name: "Icelandair" }, { code: "114", name: "El Al" },
  { code: "115", name: "Air Serbia" }, { code: "117", name: "Scandinavian Airlines" },
  { code: "124", name: "Air Algerie" }, { code: "125", name: "British Airways" },
  { code: "126", name: "Garuda Indonesia" }, { code: "127", name: "Gol Transportes Aereos" },
  { code: "131", name: "Japan Airlines" }, { code: "134", name: "Avianca" },
  { code: "139", name: "Aeromexico" }, { code: "147", name: "Royal Air Maroc" },
  { code: "157", name: "Qatar Airways" }, { code: "160", name: "Cathay Pacific" },
  { code: "176", name: "Emirates" }, { code: "180", name: "Korean Air" },
  { code: "205", name: "All Nippon Airways" }, { code: "217", name: "Thai Airways International" },
  { code: "220", name: "Lufthansa" }, { code: "230", name: "Copa Airlines" },
  { code: "232", name: "Malaysia Airlines" }, { code: "235", name: "Turkish Airlines" },
  { code: "257", name: "Austrian Airlines" }, { code: "279", name: "JetBlue Airways" },
  { code: "281", name: "TAROM" }, { code: "282", name: "TAP Air Portugal" },
  { code: "297", name: "China Airlines" }, { code: "312", name: "IndiGo" },
  { code: "324", name: "Shandong Airlines" }, { code: "328", name: "Norwegian Air Shuttle" },
  { code: "390", name: "Aegean Airlines" }, { code: "427", name: "Air Caraibes" },
  { code: "465", name: "Air Astana" }, { code: "479", name: "Shenzhen Airlines" },
  { code: "512", name: "Royal Jordanian" }, { code: "514", name: "Air Arabia" },
  { code: "605", name: "Sky Airline" }, { code: "607", name: "Etihad Airways" },
  { code: "618", name: "Singapore Airlines" }, { code: "623", name: "Bulgaria Air" },
  { code: "643", name: "Air Malta" }, { code: "649", name: "Air Transat" },
  { code: "657", name: "Air Baltic" }, { code: "668", name: "Scoot" },
  { code: "695", name: "EVA Air" }, { code: "706", name: "Kenya Airways" },
  { code: "724", name: "Swiss International Air Lines" }, { code: "731", name: "Xiamen Airlines" },
  { code: "738", name: "Vietnam Airlines" }, { code: "755", name: "Air Europa" },
  { code: "774", name: "Shanghai Airlines" }, { code: "781", name: "China Eastern Airlines" },
  { code: "784", name: "China Southern Airlines" }, { code: "795", name: "Virgin Australia" },
  { code: "821", name: "Neos" }, { code: "831", name: "Croatia Airlines" },
  { code: "838", name: "WestJet" }, { code: "847", name: "Riyadh Air" },
  { code: "876", name: "Sichuan Airlines" }, { code: "880", name: "Hainan Airlines" },
  { code: "900", name: "flyadeal" }, { code: "932", name: "Virgin Atlantic" },
  { code: "978", name: "VietJet Air" }, { code: "999", name: "Air China" },
];
const getAirlineCode = (name) => {
  const n = (name || "").trim().toUpperCase();
  if (!n) return null;
  const match = AIRLINE_CODES.find((a) => a.name.toUpperCase() === n);
  return match ? match.code : null;
};
const getAirlineByCode = (code) => {
  const match = AIRLINE_CODES.find((a) => a.code === code);
  return match ? match.name.toUpperCase() : null;
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
  const [query, setQuery] = useState("");
  const [error, setError] = useState("");
  // Clicking a ticket row opens a full detail view of that ticket (id stored here).
  const [viewingTicketId, setViewingTicketId] = useState(null);
  const [notesDraft, setNotesDraft] = useState("");
  const [notesSaved, setNotesSaved] = useState(false);
  const [selectedMonth, setSelectedMonth] = useState("");
  const [selectedYear, setSelectedYear] = useState("");
  const [selectedCompany, setSelectedCompany] = useState("");
  const [selectedEmployee, setSelectedEmployee] = useState("");
  const [selectedSupplier, setSelectedSupplier] = useState("");

  // Every value ever entered (companies, customers, airlines, cities) is kept here so it
  // can be offered as an autocomplete suggestion later, even if the original ticket is deleted.
  const [suggestions, setSuggestions] = useState({ companies: [], customers: [], airlines: [], cities: [] });

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
        const [ticketsRes, employeesRes, sessionRes, suggestionsRes, setupRes] = await Promise.all([
          window.storage.get("tickets:list", true).catch(() => null),
          window.storage.get("tickets:employees", true).catch(() => null),
          window.storage.get("session:user", false).catch(() => null),
          window.storage.get("tickets:suggestions", true).catch(() => null),
          window.storage.get("tickets:setupComplete", true).catch(() => null),
        ]);
        const ticketsData = ticketsRes && ticketsRes.value ? JSON.parse(ticketsRes.value) : [];
        const employeesData = employeesRes && employeesRes.value ? JSON.parse(employeesRes.value) : [];
        setTickets(ticketsData);
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
            });
          } catch (e) {
            // ignore malformed suggestions data
          }
        }

        if (sessionRes && sessionRes.value) {
          const savedUsername = sessionRes.value;
          const match = employeesData.find((e) => e.username === savedUsername);
          if (match) setCurrentUser({ username: match.username, name: match.name, isAdmin: !!match.isAdmin });
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
        const [ticketsRes, employeesRes, suggestionsRes] = await Promise.all([
          window.storage.get("tickets:list", true).catch(() => null),
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

  // While signed in, periodically mark this account as "connected" so the main account can see it
  useEffect(() => {
    if (!currentUser) return;
    let cancelled = false;
    const beat = async () => {
      try {
        await window.storage.set(
          `tickets:presence:${currentUser.username}`,
          JSON.stringify({ name: currentUser.name, ts: Date.now() }),
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
              return [username, parsed.ts];
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
    const ts = presenceMap[username];
    return !!ts && Date.now() - ts < ONLINE_THRESHOLD_MS;
  };
  const onlineUsernames = Object.keys(presenceMap).filter((u) => isOnline(u));

  const persistTickets = async (next) => {
    setTickets(next);
    try {
      await window.storage.set("tickets:list", JSON.stringify(next), true);
    } catch (e) {
      setError("Could not save data, please try again");
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
    };
    next.airlines = addUnique(next.airlines, record.airline);
    next.cities = addUnique(next.cities, record.from);
    next.cities = addUnique(next.cities, record.to);
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
        // Accounting is a fixed bundle: full view access, notes-only editing —
        // it overrides whatever was picked for canViewAll / canEdit. Edit access on its
        // own also implies full view access, since editing every ticket requires seeing them.
        canViewAll: newEmployee.isAccounting || newEmployee.canEdit ? true : newEmployee.canViewAll,
        canEdit: newEmployee.isAccounting ? false : newEmployee.canEdit,
        isAccounting: newEmployee.isAccounting,
      },
    ];
    await persistEmployees(next);
    setNewEmployee(emptyNewEmployee);
    setShowNewEmployeePerms(false);
  };

  // Lets the main account set an employee's access level:
  // - "own": can only see and add their own tickets (default)
  // - "view": can see every ticket, but cannot add/edit any of them
  // - "all": can see every ticket, and add/edit like a normal employee
  // - "accounting": can see every ticket but cannot add tickets — the only edit
  //   they're allowed is the Notes field on a ticket's detail page (each note edit
  //   is timestamped and attributed to them, see saveTicketNotes).
  const handleAccessChange = async (username, value) => {
    if (!currentUser.isAdmin) {
      setManageError("Only the main account can change employee permissions");
      return;
    }
    const next = (employees || []).map((e) => {
      if (e.username !== username) return e;
      if (value === "accounting") return { ...e, isAccounting: true, canViewAll: true, canEdit: false };
      if (value === "all") return { ...e, isAccounting: false, canViewAll: true, canEdit: true };
      if (value === "view") return { ...e, isAccounting: false, canViewAll: true, canEdit: false };
      return { ...e, isAccounting: false, canViewAll: false, canEdit: false };
    });
    await persistEmployees(next);
  };

  // Independent permission toggle: whether this employee can open the Manage companies
  // panel (add/edit/remove saved company records). Separate from ticket access level.
  const handleToggleManageCompanies = async (username, checked) => {
    if (!currentUser.isAdmin) {
      setManageError("Only the main account can change employee permissions");
      return;
    }
    const next = (employees || []).map((e) =>
      e.username === username ? { ...e, canManageCompanies: checked } : e
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
    if (!customersValid || !form.from.trim() || !form.to.trim() || form.netPrice === "" || form.soldPrice === "") {
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
    setForm({ ...t, customers, customersCount: customers.length });
  };
  const handleDelete = (id) => {
    if (!currentUser.isAdmin) {
      setError("Only the main account can delete tickets");
      return;
    }
    if (form.id === id) setForm(getEmptyForm());
    persistTickets(tickets.filter((t) => t.id !== id));
  };
  const handleCancel = () => setForm(getEmptyForm());

  // Opens the full-detail view ("page") for a ticket, showing every field including notes.
  const openTicketDetail = (t) => {
    setViewingTicketId(t.id);
    setNotesDraft(t.notes || "");
    setNotesSaved(false);
  };
  const closeTicketDetail = () => {
    setViewingTicketId(null);
    setNotesDraft("");
    setNotesSaved(false);
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

  const handleAirlineChange = (value) => {
    const airline = value.toUpperCase();
    const code = getAirlineCode(airline);
    // If we recognize the airline, pre-fill its 3-digit prefix into any customer's
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
        (currentEmployeeRecord.canViewAll || currentEmployeeRecord.isAccounting || currentEmployeeRecord.canEdit)
      ));
  // Accounting accounts can see everything but cannot add tickets — their only allowed
  // edit anywhere in the app is the Notes field on a ticket's detail page.
  const isAccountingUser =
    !!currentUser && !currentUser.isAdmin && !!(currentEmployeeRecord && currentEmployeeRecord.isAccounting);
  // A non-admin employee can be granted permission to edit tickets (within whatever
  // set of tickets they can already see). Accounting accounts are excluded even if the
  // flag is set — their only allowed edit is the Notes field, never the ticket itself.
  const canEditTickets =
    !!currentUser &&
    (currentUser.isAdmin ||
      !!(currentEmployeeRecord && currentEmployeeRecord.canEdit && !currentEmployeeRecord.isAccounting));
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
      t.from.toLowerCase().includes(q) ||
      t.to.toLowerCase().includes(q) ||
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
  const countAndSum = (rows) =>
    rows.reduce(
      (acc, t) => {
        const n = getCustomers(t).length || 1;
        acc.count += n;
        acc.total += (parseFloat(t.soldPrice) || 0) * n;
        acc.profit += profit(t.netPrice, t.soldPrice) * n;
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
      return customers.map((c, i) => ({
        "Employee": t.employee || "",
        "Company": t.company || "",
        "Supplier": t.supplier || "",
        "Customer #": i + 1,
        "Customer": c.name || "",
        "Ticket number": c.ticketNumber || "",
        "From": t.from,
        "To": t.to,
        "Airline": t.airline || "",
        "Issue date": t.date ? formatDisplayDate(t.date) : "",
        // Net/sold price and profit reflect the whole booking and are shown once, on the first customer's row
        "Net price": i === 0 ? parseFloat(t.netPrice) || 0 : "",
        "Sold price": i === 0 ? parseFloat(t.soldPrice) || 0 : "",
        "Profit": i === 0 ? profit(t.netPrice, t.soldPrice) : "",
        "Notes": t.notes || "",
      }));
    });
  };

  // Sums net price / sold price / profit across a raw ticket list (once per booking,
  // matching how those columns are only populated on each booking's first row above).
  const sumTicketPrices = (rows) =>
    rows.reduce(
      (acc, t) => {
        acc.net += parseFloat(t.netPrice) || 0;
        acc.sold += parseFloat(t.soldPrice) || 0;
        acc.profit += profit(t.netPrice, t.soldPrice);
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
        "Employee": "", "Company": "", "Supplier": "", "Customer #": "", "Customer": "",
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
      <div className="w-full min-h-screen bg-slate-50 flex items-center justify-center">
        <p className="text-slate-400 text-sm">Loading...</p>
      </div>
    );
  }

  // ---------- Render: first-run setup (only ever shown once, before any account exists) ----------
  if (employees && employees.length === 0 && !setupComplete) {
    return (
      <div className="w-full min-h-screen bg-slate-50 flex items-center justify-center p-4">
        <div className="bg-white rounded-xl border border-slate-200 p-6 w-full max-w-sm">
          <div className="flex items-center gap-2 mb-1">
            <Lock size={18} className="text-teal-700" />
            <h1 className="font-bold text-slate-900">Create the admin account</h1>
          </div>
          <p className="text-xs text-slate-500 mb-4">
            No employees exist yet. Create the first account — it will be the main account, and only it will be able to add or remove other employees.
          </p>
          {loginError && <div className="bg-red-50 text-red-700 text-sm rounded-lg px-3 py-2 mb-3">{loginError}</div>}
          <div className="space-y-3">
            <div>
              <label className="text-xs text-slate-500 block mb-1">Full name</label>
              <input className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                value={setupName} onChange={(e) => setSetupName(e.target.value)} placeholder="e.g. Sara Ahmed" />
            </div>
            <div>
              <label className="text-xs text-slate-500 block mb-1">Username</label>
              <input className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                value={setupUsername} onChange={(e) => setSetupUsername(e.target.value)} placeholder="sara" />
            </div>
            <div>
              <label className="text-xs text-slate-500 block mb-1">Password</label>
              <input type="password" className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                value={setupPassword} onChange={(e) => setSetupPassword(e.target.value)} placeholder="••••••" />
            </div>
          </div>
          <button onClick={handleCreateFirstAdmin}
            className="w-full mt-4 bg-teal-700 hover:bg-teal-800 text-white text-sm font-semibold rounded-lg px-4 py-2">
            Create account and continue
          </button>
          <p className="text-xs text-slate-400 mt-4">
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
      <div className="w-full min-h-screen bg-slate-50 flex items-center justify-center p-4">
        <div className="bg-white rounded-xl border border-slate-200 p-6 w-full max-w-sm text-center">
          <Lock size={22} className="text-slate-400 mx-auto mb-2" />
          <h1 className="font-bold text-slate-900 mb-1">No accounts available</h1>
          <p className="text-xs text-slate-500">
            This app was already set up before, but no employee accounts currently exist. Restore a backup that contains employee accounts, or contact whoever manages this app.
          </p>
        </div>
      </div>
    );
  }

  // ---------- Render: login screen ----------
  if (!currentUser) {
    return (
      <div className="w-full min-h-screen relative flex items-center justify-center p-4 overflow-hidden bg-gradient-to-br from-sky-600 via-blue-700 to-indigo-900">
        {/* Decorative sky backdrop */}
        <div className="pointer-events-none absolute inset-0 overflow-hidden">
          <div className="absolute -top-24 -left-24 w-72 h-72 bg-sky-400/30 rounded-full blur-3xl" />
          <div className="absolute -bottom-24 -right-16 w-80 h-80 bg-indigo-400/20 rounded-full blur-3xl" />
          <Globe2 size={260} className="absolute -bottom-16 -right-16 text-white/5" />
          <Cloud size={70} className="absolute top-[12%] left-[10%] text-white/20" />
          <Cloud size={46} className="absolute top-[22%] right-[14%] text-white/15" />
          <Cloud size={54} className="absolute bottom-[18%] left-[16%] text-white/10" />
          {/* Dashed flight path with a plane at the tip */}
          <svg className="absolute top-[8%] left-[8%] w-[84%] h-40 opacity-40" viewBox="0 0 600 140" fill="none">
            <path d="M10 120 C 160 20, 380 20, 560 60" stroke="white" strokeWidth="2" strokeDasharray="6 8" strokeLinecap="round" />
            <circle cx="10" cy="120" r="4" fill="white" />
          </svg>
          <Plane size={26} className="absolute top-[15%] right-[10%] text-white/70 rotate-45 animate-pulse" />
        </div>

        <div className="relative w-full max-w-sm">
          <div className="bg-white rounded-3xl shadow-2xl overflow-hidden">
            {/* Branded header */}
            <div className="relative bg-gradient-to-r from-sky-600 to-blue-700 px-6 pt-9 pb-12 text-center overflow-hidden">
              <Plane size={90} className="absolute -bottom-4 -left-6 text-white/10 rotate-12" />
              <div className="relative w-full mx-auto rounded-2xl bg-white shadow-lg flex items-center justify-center mb-3 p-4">
                <img src={LOGO_DATA_URL} alt="Perla Di Mare" className="w-full h-auto object-contain" />
              </div>
              <h1 className="relative text-white font-bold text-lg tracking-tight">Flight Ticket Manager</h1>
              <p className="relative text-sky-200/70 text-[11px] mt-0.5">By Fady Habib</p>
              <p className="relative text-sky-100/90 text-xs mt-1">Sign in to manage tickets, sales &amp; bookings</p>
            </div>

            {/* Form card, slightly overlapping the header for a layered feel */}
            <div className="relative -mt-6 bg-white rounded-t-3xl px-6 pt-6 pb-6">
              {loginError && <div className="bg-red-50 text-red-700 text-sm rounded-lg px-3 py-2 mb-3">{loginError}</div>}
              <div className="space-y-3">
                <div>
                  <label className="text-xs text-slate-500 block mb-1">Username</label>
                  <div className="relative">
                    <User size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input className="w-full border border-slate-300 rounded-lg pl-9 pr-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-600 focus:border-blue-600"
                      value={loginUsername} onChange={(e) => setLoginUsername(e.target.value)}
                      onKeyDown={(e) => e.key === "Enter" && handleLogin()} placeholder="Username" autoFocus />
                  </div>
                </div>
                <div>
                  <label className="text-xs text-slate-500 block mb-1">Password</label>
                  <div className="relative">
                    <Lock size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input type={showPassword ? "text" : "password"}
                      className="w-full border border-slate-300 rounded-lg pl-9 pr-9 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-600 focus:border-blue-600"
                      value={loginPassword} onChange={(e) => setLoginPassword(e.target.value)}
                      onKeyDown={(e) => e.key === "Enter" && handleLogin()} placeholder="Password" />
                    <button type="button" onClick={() => setShowPassword(!showPassword)}
                      className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600">
                      {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                    </button>
                  </div>
                </div>
              </div>
              <button onClick={handleLogin}
                className="group w-full mt-5 bg-gradient-to-r from-sky-600 to-blue-700 hover:from-sky-500 hover:to-blue-600 text-white text-sm font-semibold rounded-lg px-4 py-2.5 flex items-center justify-center gap-2 shadow-lg shadow-blue-600/30 transition-all">
                Sign in
                <Plane size={15} className="rotate-45 transition-transform group-hover:translate-x-1 group-hover:-translate-y-1" />
              </button>
              <p className="text-xs text-slate-400 mt-4 text-center flex items-center justify-center gap-1">
                <ShieldCheck size={13} /> Ask your admin if you don't have an account yet.
              </p>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ---------- Render: main app ----------
  return (
    <div dir="ltr" className="w-full min-h-screen bg-slate-50 text-slate-800">
      <div className="max-w-5xl mx-auto p-4 md:p-6">
        <header className="flex items-center justify-between mb-6 flex-wrap gap-3">
          <div className="flex items-center gap-3">
            <div>
              <h1 className="text-xl md:text-2xl font-bold text-slate-900">
                Flight Ticket Manager <span className="text-slate-400 font-medium text-sm md:text-base">By Fady Habib</span>
              </h1>
              <p className="text-slate-500 text-sm flex items-center gap-1.5 flex-wrap">
                Signed in as {currentUser.name}
                {currentUser.isAdmin && (
                  <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-teal-700 bg-teal-50 border border-teal-200 rounded-full px-2 py-0.5">
                    <ShieldCheck size={11} /> Main account
                  </span>
                )}
                {!currentUser.isAdmin && !canViewAllTickets && (
                  <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-slate-500 bg-slate-100 border border-slate-200 rounded-full px-2 py-0.5">
                    Your own tickets only
                  </span>
                )}
                {isAccountingUser && (
                  <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-indigo-700 bg-indigo-50 border border-indigo-200 rounded-full px-2 py-0.5">
                    Accounting — view only
                  </span>
                )}
                {currentUser.isAdmin && (
                  <span className="relative">
                    <button
                      type="button"
                      onClick={() => setShowOnlineList(!showOnlineList)}
                      className="inline-flex items-center gap-1 text-[10px] font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-full px-2 py-0.5 hover:bg-emerald-100"
                    >
                      <Wifi size={11} />
                      {onlineUsernames.length} online now
                    </button>
                    {showOnlineList && (
                      <div className="absolute z-20 top-full mt-1 left-0 w-52 bg-white border border-slate-300 rounded-lg shadow-lg p-2">
                        {onlineUsernames.length === 0 ? (
                          <p className="text-xs text-slate-400 px-1 py-1">No one online right now</p>
                        ) : (
                          <ul className="space-y-1 max-h-48 overflow-y-auto">
                            {onlineUsernames.map((u) => {
                              const emp = (employees || []).find((e) => e.username === u);
                              return (
                                <li key={u} className="flex items-center gap-1.5 text-xs text-slate-700 px-1 py-0.5">
                                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 shrink-0" />
                                  {emp ? emp.name : u}
                                  {emp && emp.isAdmin && (
                                    <span className="text-[9px] text-teal-600 font-semibold">(main)</span>
                                  )}
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
            <div className="bg-white border border-slate-200 rounded-xl p-3 shadow-sm shrink-0">
              <img src={LOGO_DATA_URL} alt="Perla Di Mare" className="w-[220px] h-auto md:w-[300px] md:h-auto object-contain" />
            </div>
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            {currentUser.isAdmin && (
              <button onClick={handleBackup}
                className="border border-slate-300 text-slate-600 text-sm rounded-lg px-3 py-2 flex items-center gap-1.5">
                <Download size={15} /> Backup
              </button>
            )}
            {currentUser.isAdmin && (
              <button onClick={triggerRestore}
                className="border border-slate-300 text-slate-600 text-sm rounded-lg px-3 py-2 flex items-center gap-1.5">
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
                className="border border-slate-300 text-slate-600 text-sm rounded-lg px-3 py-2 flex items-center gap-1.5">
                <Users size={15} /> Manage employees
              </button>
            )}
            {canManageCompanies && (
              <button onClick={() => setShowManageCompanies(!showManageCompanies)}
                className="border border-slate-300 text-slate-600 text-sm rounded-lg px-3 py-2 flex items-center gap-1.5">
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
              className="border border-slate-300 text-slate-600 text-sm rounded-lg px-3 py-2 flex items-center gap-1.5">
              <Lock size={15} /> Change password
            </button>
            <button onClick={handleLogout}
              className="border border-slate-300 text-slate-600 text-sm rounded-lg px-3 py-2 flex items-center gap-1.5">
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
                className="border border-slate-300 text-slate-500 text-sm rounded-lg px-3 py-2 flex items-center gap-1.5"
              >
                <Wifi size={15} /> Server
              </button>
            )}
          </div>
        </header>

        {/* Top-level section switcher */}
        <div className="flex items-center justify-center gap-3 mb-6">
          <button
            onClick={() => setActiveSection("flights")}
            className={`flex flex-col items-center gap-1.5 px-6 py-3 rounded-xl border text-xs font-semibold transition-colors ${
              activeSection === "flights"
                ? "bg-teal-700 text-white border-teal-700 shadow-sm"
                : "bg-white text-slate-500 border-slate-200 hover:border-teal-300 hover:text-teal-700"
            }`}
          >
            <Plane size={22} className="rotate-45" />
            Flights
          </button>
          <button
            onClick={() => setActiveSection("hotels")}
            className={`flex flex-col items-center gap-1.5 px-6 py-3 rounded-xl border text-xs font-semibold transition-colors ${
              activeSection === "hotels"
                ? "bg-teal-700 text-white border-teal-700 shadow-sm"
                : "bg-white text-slate-500 border-slate-200 hover:border-teal-300 hover:text-teal-700"
            }`}
          >
            <Building2 size={22} />
            Hotels
          </button>
          <button
            onClick={() => setActiveSection("cars")}
            className={`flex flex-col items-center gap-1.5 px-6 py-3 rounded-xl border text-xs font-semibold transition-colors ${
              activeSection === "cars"
                ? "bg-teal-700 text-white border-teal-700 shadow-sm"
                : "bg-white text-slate-500 border-slate-200 hover:border-teal-300 hover:text-teal-700"
            }`}
          >
            <Car size={22} />
            Transportation
          </button>
          <button
            onClick={() => setActiveSection("files")}
            className={`flex flex-col items-center gap-1.5 px-6 py-3 rounded-xl border text-xs font-semibold transition-colors ${
              activeSection === "files"
                ? "bg-teal-700 text-white border-teal-700 shadow-sm"
                : "bg-white text-slate-500 border-slate-200 hover:border-teal-300 hover:text-teal-700"
            }`}
          >
            <FileText size={22} />
            Files
          </button>
        </div>

        {activeSection === "flights" && (
        <>
        {currentUser.isAdmin && (restoreError || restoreSuccess) && (
          <div className={`text-sm rounded-lg px-3 py-2 mb-4 ${restoreError ? "bg-red-50 text-red-700" : "bg-emerald-50 text-emerald-700"}`}>
            {restoreError || restoreSuccess}
          </div>
        )}

        {showManage && currentUser.isAdmin && (
          <div className="bg-white rounded-xl border border-slate-200 p-4 md:p-5 mb-6">
            <h2 className="font-semibold text-slate-900 mb-1">Employee accounts</h2>
            <p className="text-xs text-slate-400 mb-4">
              As the main account, you can view and change every employee's password, edit their name or username, add or remove accounts, control who can see all tickets, and grant or remove main-account access. This is a basic access gate, not a secure authentication system — anyone with technical access to the app's stored data can read these passwords. Avoid reusing important passwords here.
            </p>
            {manageError && <div className="bg-red-50 text-red-700 text-sm rounded-lg px-3 py-2 mb-3">{manageError}</div>}
            <p className="text-xs text-slate-500 mb-3 flex items-center gap-1.5">
              <Wifi size={13} className="text-emerald-600" />
              {onlineUsernames.length} of {(employees || []).length} employees connected right now
            </p>
            <div className="border border-slate-200 rounded-lg overflow-hidden mb-4">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-slate-500 text-xs">
                    <th className="text-left px-3 py-2 font-medium">Status</th>
                    <th className="text-left px-3 py-2 font-medium">Name</th>
                    <th className="text-left px-3 py-2 font-medium">Username</th>
                    <th className="text-left px-3 py-2 font-medium">Password</th>
                    <th className="text-left px-3 py-2 font-medium">Role</th>
                    <th className="text-left px-3 py-2 font-medium">Access</th>
                    <th className="text-left px-3 py-2 font-medium">Companies</th>
                    <th className="text-left px-3 py-2 font-medium"></th>
                  </tr>
                </thead>
                <tbody>
                  {(employees || []).map((e) => {
                    const isEditing = editingUsername === e.username;
                    if (isEditing) {
                      return (
                        <tr key={e.username} className="border-t border-slate-100 bg-slate-50">
                          <td className="px-3 py-2">
                            <span className={`inline-flex items-center gap-1 text-[10px] font-semibold rounded-full px-2 py-0.5 ${isOnline(e.username) ? "text-emerald-700 bg-emerald-50 border border-emerald-200" : "text-slate-400 bg-slate-100 border border-slate-200"}`}>
                              <span className={`w-1.5 h-1.5 rounded-full ${isOnline(e.username) ? "bg-emerald-500" : "bg-slate-300"}`} />
                              {isOnline(e.username) ? "Online" : "Offline"}
                            </span>
                          </td>
                          <td className="px-3 py-2">
                            <input
                              className="w-full border border-slate-300 rounded-lg px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                              value={editDraft.name}
                              onChange={(ev) => setEditDraft({ ...editDraft, name: ev.target.value })}
                            />
                          </td>
                          <td className="px-3 py-2">
                            <input
                              className="w-full border border-slate-300 rounded-lg px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                              value={editDraft.username}
                              onChange={(ev) => setEditDraft({ ...editDraft, username: ev.target.value })}
                            />
                          </td>
                          <td className="px-3 py-2">
                            <div className="relative">
                              <input
                                type={editShowPassword ? "text" : "password"}
                                className="w-full border border-slate-300 rounded-lg pl-2 pr-8 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                                value={editDraft.password}
                                onChange={(ev) => setEditDraft({ ...editDraft, password: ev.target.value })}
                              />
                              <button
                                type="button"
                                onClick={() => setEditShowPassword(!editShowPassword)}
                                className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400"
                              >
                                {editShowPassword ? <EyeOff size={14} /> : <Eye size={14} />}
                              </button>
                            </div>
                          </td>
                          <td className="px-3 py-2 text-slate-500">
                            {e.isAdmin ? (
                              <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-teal-700 bg-teal-50 border border-teal-200 rounded-full px-2 py-0.5">
                                <ShieldCheck size={11} /> Main
                              </span>
                            ) : (
                              "Employee"
                            )}
                          </td>
                          <td className="px-3 py-2 text-slate-500">
                            {e.isAdmin ? (
                              <span className="text-xs">Always (Main account)</span>
                            ) : (
                              <select
                                value={e.isAccounting ? "accounting" : e.canEdit ? "all" : e.canViewAll ? "view" : "own"}
                                onChange={(ev) => handleAccessChange(e.username, ev.target.value)}
                                className="border border-slate-300 rounded-md px-2 py-1.5 text-xs focus:outline-none focus:ring-2 focus:ring-teal-600 bg-white"
                              >
                                <option value="own">Own tickets only</option>
                                <option value="view">View all tickets (no edit)</option>
                                <option value="all">All tickets (view &amp; edit)</option>
                                <option value="accounting">Accounting (view + notes only)</option>
                              </select>
                            )}
                          </td>
                          <td className="px-3 py-2">
                            {e.isAdmin ? (
                              <span className="text-xs text-slate-400">Always</span>
                            ) : (
                              <input
                                type="checkbox"
                                checked={!!e.canManageCompanies}
                                onChange={(ev) => handleToggleManageCompanies(e.username, ev.target.checked)}
                                className="w-4 h-4 accent-teal-700"
                              />
                            )}
                          </td>
                          <td className="px-3 py-2">
                            <div className="flex gap-1 justify-end">
                              <button onClick={saveEditEmployee} className="text-emerald-600 hover:text-emerald-800 p-1">
                                <Check size={15} />
                              </button>
                              <button onClick={cancelEditEmployee} className="text-slate-400 hover:text-red-600 p-1">
                                <X size={15} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    }
                    return (
                      <tr key={e.username} className="border-t border-slate-100">
                        <td className="px-3 py-2">
                          <span className={`inline-flex items-center gap-1 text-[10px] font-semibold rounded-full px-2 py-0.5 ${isOnline(e.username) ? "text-emerald-700 bg-emerald-50 border border-emerald-200" : "text-slate-400 bg-slate-100 border border-slate-200"}`}>
                            <span className={`w-1.5 h-1.5 rounded-full ${isOnline(e.username) ? "bg-emerald-500" : "bg-slate-300"}`} />
                            {isOnline(e.username) ? "Online" : "Offline"}
                          </span>
                        </td>
                        <td className="px-3 py-2">{e.name}</td>
                        <td className="px-3 py-2 text-slate-500">{e.username}</td>
                        <td className="px-3 py-2 text-slate-500">
                          <div className="flex items-center gap-2">
                            <span className="font-mono">
                              {visiblePasswords[e.username] ? e.password : "••••••••"}
                            </span>
                            <button onClick={() => togglePasswordVisible(e.username)} className="text-slate-400 hover:text-teal-700">
                              {visiblePasswords[e.username] ? <EyeOff size={14} /> : <Eye size={14} />}
                            </button>
                          </div>
                        </td>
                        <td className="px-3 py-2 text-slate-500">
                          {e.isAdmin ? (
                            <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-teal-700 bg-teal-50 border border-teal-200 rounded-full px-2 py-0.5">
                              <ShieldCheck size={11} /> Main
                            </span>
                          ) : (
                            "Employee"
                          )}
                        </td>
                        <td className="px-3 py-2 text-slate-500">
                          {e.isAdmin ? (
                            <span className="text-xs">Always (Main account)</span>
                          ) : (
                            <select
                              value={e.isAccounting ? "accounting" : e.canEdit ? "all" : e.canViewAll ? "view" : "own"}
                              onChange={(ev) => handleAccessChange(e.username, ev.target.value)}
                              className="border border-slate-300 rounded-md px-2 py-1.5 text-xs focus:outline-none focus:ring-2 focus:ring-teal-600 bg-white"
                            >
                              <option value="own">Own tickets only</option>
                              <option value="view">View all tickets (no edit)</option>
                              <option value="all">All tickets (view &amp; edit)</option>
                              <option value="accounting">Accounting (view + notes only)</option>
                            </select>
                          )}
                        </td>
                        <td className="px-3 py-2">
                          {e.isAdmin ? (
                            <span className="text-xs text-slate-400">Always</span>
                          ) : (
                            <input
                              type="checkbox"
                              checked={!!e.canManageCompanies}
                              onChange={(ev) => handleToggleManageCompanies(e.username, ev.target.checked)}
                              className="w-4 h-4 accent-teal-700"
                            />
                          )}
                        </td>
                        <td className="px-3 py-2 text-right">
                          <div className="flex gap-1 justify-end">
                            {e.isAdmin ? (
                              <button
                                onClick={() => handleDemoteAdmin(e.username)}
                                title="Remove main-account access"
                                className="text-slate-400 hover:text-amber-600 text-[11px] font-semibold border border-slate-200 rounded-md px-1.5 py-1"
                              >
                                Remove main
                              </button>
                            ) : (
                              <button
                                onClick={() => handlePromoteToAdmin(e.username)}
                                title="Make this a main account"
                                className="text-slate-400 hover:text-teal-700 text-[11px] font-semibold border border-slate-200 rounded-md px-1.5 py-1 flex items-center gap-1"
                              >
                                <ShieldCheck size={12} /> Make main
                              </button>
                            )}
                            <button onClick={() => startEditEmployee(e)} className="text-slate-400 hover:text-teal-700 p-1">
                              <Pencil size={15} />
                            </button>
                            {!e.isAdmin && (
                              <button onClick={() => handleDeleteEmployee(e.username)} className="text-slate-400 hover:text-red-600 p-1">
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
              <input className="border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                placeholder="Full name" value={newEmployee.name}
                onChange={(e) => setNewEmployee({ ...newEmployee, name: e.target.value })} />
              <input className="border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                placeholder="Username" value={newEmployee.username}
                onChange={(e) => setNewEmployee({ ...newEmployee, username: e.target.value })} />
              <input type="password" className="border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                placeholder="Password" value={newEmployee.password}
                onChange={(e) => setNewEmployee({ ...newEmployee, password: e.target.value })} />
            </div>

            {/* Multi-select permissions picker: choose any combination of view/edit
                access to grant this employee before creating the account. */}
            <div className="relative mt-3 max-w-sm">
              <button
                type="button"
                onClick={() => setShowNewEmployeePerms(!showNewEmployeePerms)}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm text-slate-700 hover:bg-slate-50 flex items-center justify-between gap-2"
              >
                <span className="font-medium">Permissions</span>
                <span className="text-xs text-slate-500 truncate">
                  {(newEmployee.isAccounting
                    ? "Accounting (view all, notes only)"
                    : newEmployee.canEdit
                    ? "View all tickets + edit"
                    : newEmployee.canViewAll
                    ? "View all tickets, no edit"
                    : "Own tickets only") + (newEmployee.canManageCompanies ? " · Companies" : "")}
                </span>
              </button>

              {showNewEmployeePerms && (
                <div className="absolute z-10 mt-1 w-full bg-white border border-slate-300 rounded-lg shadow-lg p-3 space-y-2.5">
                  <label className="flex items-center gap-2 text-sm text-slate-700">
                    <input
                      type="checkbox"
                      checked={newEmployee.canViewAll || newEmployee.canEdit}
                      disabled={newEmployee.isAccounting || newEmployee.canEdit}
                      onChange={(e) => setNewEmployee({ ...newEmployee, canViewAll: e.target.checked })}
                    />
                    View all tickets
                  </label>
                  <label className="flex items-center gap-2 text-sm text-slate-700">
                    <input
                      type="checkbox"
                      checked={newEmployee.canEdit}
                      disabled={newEmployee.isAccounting}
                      onChange={(e) =>
                        setNewEmployee({
                          ...newEmployee,
                          canEdit: e.target.checked,
                          // Editing every ticket requires seeing every ticket first.
                          canViewAll: e.target.checked ? true : newEmployee.canViewAll,
                        })
                      }
                    />
                    Edit tickets (all tickets, view access included automatically)
                  </label>
                  <label className="flex items-center gap-2 text-sm text-slate-700 border-t border-slate-100 pt-2.5">
                    <input
                      type="checkbox"
                      checked={newEmployee.isAccounting}
                      onChange={(e) =>
                        setNewEmployee({
                          ...newEmployee,
                          isAccounting: e.target.checked,
                          // Accounting always implies full view access and no ticket editing.
                          canViewAll: e.target.checked ? true : newEmployee.canViewAll,
                          canEdit: e.target.checked ? false : newEmployee.canEdit,
                        })
                      }
                    />
                    Accounting (view all, notes only)
                  </label>
                  <label className="flex items-center gap-2 text-sm text-slate-700 border-t border-slate-100 pt-2.5">
                    <input
                      type="checkbox"
                      checked={newEmployee.canManageCompanies}
                      onChange={(e) => setNewEmployee({ ...newEmployee, canManageCompanies: e.target.checked })}
                    />
                    Manage companies (add/edit/remove saved company records)
                  </label>
                  <p className="text-[11px] text-slate-400 border-t border-slate-100 pt-2">
                    Leave everything unchecked for "own tickets only" (default): the employee
                    will only see and add the tickets they personally enter.
                  </p>
                </div>
              )}
            </div>

            <button onClick={handleAddEmployee}
              className="mt-3 bg-teal-700 hover:bg-teal-800 text-white text-sm font-semibold rounded-lg px-4 py-2 flex items-center gap-1.5">
              <UserPlus size={15} /> Add employee
            </button>
          </div>
        )}

        {showManageCompanies && canManageCompanies && (
          <div className="bg-white rounded-xl border border-slate-200 p-4 md:p-5 mb-6">
            <h2 className="font-semibold text-slate-900 mb-1 flex items-center gap-2">
              <Factory size={18} className="text-slate-500" /> Companies
            </h2>
            <p className="text-xs text-slate-400 mb-4">
              Register each company's details here so they're always available to pick from the Company field and filter, even before any ticket has been entered for them.
            </p>
            {companyError && <div className="bg-red-50 text-red-700 text-sm rounded-lg px-3 py-2 mb-3">{companyError}</div>}

            <div className="grid grid-cols-1 md:grid-cols-2 gap-3 max-w-2xl mb-3">
              <input
                className="border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                placeholder="Company name"
                value={newCompanyDraft.name}
                onChange={(e) => setNewCompanyDraft({ ...newCompanyDraft, name: e.target.value })}
              />
              <input
                className="border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                placeholder="Tax number"
                value={newCompanyDraft.taxNumber}
                onChange={(e) => setNewCompanyDraft({ ...newCompanyDraft, taxNumber: e.target.value })}
              />
              <input
                className="border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                placeholder="Commercial registration number"
                value={newCompanyDraft.commercialReg}
                onChange={(e) => setNewCompanyDraft({ ...newCompanyDraft, commercialReg: e.target.value })}
              />
              <input
                className="border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                placeholder="Phone numbers (comma separated)"
                value={newCompanyDraft.phones}
                onChange={(e) => setNewCompanyDraft({ ...newCompanyDraft, phones: e.target.value })}
              />
            </div>
            <div className="flex gap-2 mb-5">
              <button
                onClick={handleAddCompany}
                className="bg-teal-700 hover:bg-teal-800 text-white text-sm font-semibold rounded-lg px-4 py-2 flex items-center gap-1.5 whitespace-nowrap"
              >
                {editingCompanyName ? <Check size={15} /> : <Factory size={15} />}
                {editingCompanyName ? "Save changes" : "Add company"}
              </button>
              {editingCompanyName && (
                <button
                  onClick={cancelEditCompany}
                  className="border border-slate-300 text-slate-600 text-sm rounded-lg px-4 py-2 flex items-center gap-1.5"
                >
                  <X size={15} /> Cancel
                </button>
              )}
            </div>

            <div className="flex items-center justify-between mb-3">
              <p className="text-xs text-slate-400">
                {suggestions.companies.length} compan{suggestions.companies.length === 1 ? "y" : "ies"} saved
              </p>
              <button
                onClick={() => setShowCompaniesList(!showCompaniesList)}
                className="text-teal-700 border border-teal-700 hover:bg-teal-50 text-xs font-semibold rounded-lg px-3 py-1.5 flex items-center gap-1.5"
              >
                <List size={14} /> {showCompaniesList ? "Hide companies list" : "View all companies"}
              </button>
            </div>

            {showCompaniesList && (
              suggestions.companies.length === 0 ? (
                <p className="text-sm text-slate-400">No companies saved yet</p>
              ) : (
                <div className="border border-slate-200 rounded-lg overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs">
                      <thead>
                        <tr className="bg-slate-50 text-slate-500 text-[11px] uppercase tracking-wide">
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
                                className={`border-t border-slate-100 ${editingCompanyName === name ? "bg-teal-50/40" : "hover:bg-slate-50"}`}
                              >
                                <td className="px-3 py-2 font-medium text-slate-800 whitespace-nowrap">{name}</td>
                                <td className="px-3 py-2 text-slate-600 whitespace-nowrap">{taxNumber || "-"}</td>
                                <td className="px-3 py-2 text-slate-600 whitespace-nowrap">{commercialReg || "-"}</td>
                                <td className="px-3 py-2 text-slate-600 whitespace-nowrap">{phones.length > 0 ? phones.join(", ") : "-"}</td>
                                <td className="px-3 py-2 text-right whitespace-nowrap">
                                  <div className="flex gap-1 justify-end">
                                    <button onClick={() => handleEditCompanyClick(c)} className="text-slate-400 hover:text-teal-700 p-0.5">
                                      <Pencil size={13} />
                                    </button>
                                    <button onClick={() => handleDeleteCompany(name)} className="text-slate-400 hover:text-red-600 p-0.5">
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
        )}

        {showChangePassword && (
          <div className="bg-white rounded-xl border border-slate-200 p-4 md:p-5 mb-6 max-w-sm">
            <h2 className="font-semibold text-slate-900 mb-1 flex items-center gap-2">
              <Lock size={16} className="text-teal-700" /> Change your password
            </h2>
            <p className="text-xs text-slate-400 mb-4">
              Signed in as {currentUser.name} ({currentUser.username})
            </p>
            {passwordError && <div className="bg-red-50 text-red-700 text-sm rounded-lg px-3 py-2 mb-3">{passwordError}</div>}
            {passwordSuccess && <div className="bg-emerald-50 text-emerald-700 text-sm rounded-lg px-3 py-2 mb-3">{passwordSuccess}</div>}
            <div className="space-y-3">
              <div>
                <label className="text-xs text-slate-500 block mb-1">Current password</label>
                <input type="password"
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                  value={currentPasswordInput} onChange={(e) => setCurrentPasswordInput(e.target.value)} />
              </div>
              <div>
                <label className="text-xs text-slate-500 block mb-1">New password</label>
                <input type="password"
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                  value={newPasswordInput} onChange={(e) => setNewPasswordInput(e.target.value)} />
              </div>
              <div>
                <label className="text-xs text-slate-500 block mb-1">Confirm new password</label>
                <input type="password"
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                  value={confirmPasswordInput} onChange={(e) => setConfirmPasswordInput(e.target.value)} />
              </div>
            </div>
            <button onClick={handleChangePassword}
              className="w-full mt-4 bg-teal-700 hover:bg-teal-800 text-white text-sm font-semibold rounded-lg px-4 py-2">
              Update password
            </button>
          </div>
        )}

        {/* Summary cards */}
        <div className="flex items-center justify-between mb-2">
          <p className="text-sm text-slate-500">
            Totals for: <span className="font-semibold text-slate-700">
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
            className="text-teal-700 border border-teal-700 hover:bg-teal-50 text-xs font-semibold rounded-lg px-3 py-1.5 flex items-center gap-1.5"
          >
            <Download size={14} /> {hasActiveFilter ? "Export filtered results to Excel" : "Export all months to Excel"}
          </button>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-6">
          <div className="bg-white rounded-xl border border-slate-200 p-4 flex items-center gap-3">
            <div className="bg-slate-100 rounded-lg p-2 text-slate-600"><Ticket size={20} /></div>
            <div>
              <p className="text-xs text-slate-500">Tickets</p>
              <p className="text-lg font-bold">{totals.count}</p>
            </div>
          </div>
          <div className="bg-white rounded-xl border border-slate-200 p-4 flex items-center gap-3">
            <div className="bg-blue-50 rounded-lg p-2 text-blue-700"><Wallet size={20} /></div>
            <div>
              <p className="text-xs text-slate-500">Total sales</p>
              <p className="text-lg font-bold">{fmt(totals.total)}</p>
            </div>
          </div>
          <div className="bg-white rounded-xl border border-slate-200 p-4 flex items-center gap-3">
            <div className="bg-emerald-50 rounded-lg p-2 text-emerald-700"><TrendingUp size={20} /></div>
            <div>
              <p className="text-xs text-slate-500">Total profit</p>
              <p className="text-lg font-bold text-emerald-700">{fmt(totals.profit)}</p>
            </div>
          </div>
        </div>

        {/* Entry form (accounting accounts are view-only + notes-only, so this is hidden for them) */}
        {!isAccountingUser && (
        <div className="bg-white rounded-xl border border-slate-200 p-4 md:p-5 mb-6">
          <h2 className="font-semibold text-slate-900 mb-4">{form.id ? "Edit ticket" : "Add a new ticket"}</h2>
          {error && (
            <div className="bg-red-50 text-red-700 text-sm rounded-lg px-3 py-2 mb-3">{error}</div>
          )}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
            <div>
              <label className="text-xs text-slate-500 block mb-1">Entered by</label>
              <div className="w-full border border-slate-200 bg-slate-50 rounded-lg px-3 py-2 text-sm text-slate-600">
                {currentUser.name}
              </div>
            </div>
            <div>
              <label className="text-xs text-slate-500 block mb-1">Company (optional)</label>
              <input
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                value={form.company}
                onChange={(e) => setForm({ ...form, company: e.target.value.toUpperCase() })}
                placeholder="e.g. Acme Corp"
                list="company-suggestions"
              />
            </div>
            <div>
              <label className="text-xs text-slate-500 block mb-1">Number of customers</label>
              <input
                type="number"
                min={1}
                max={50}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
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
              <label className="text-xs text-slate-500 block mb-1">Supplier</label>
              <select
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600 bg-white"
                value={form.supplier}
                onChange={(e) => setForm({ ...form, supplier: e.target.value })}
              >
                <option value="">Select supplier</option>
                {SUPPLIERS.map((s) => (
                  <option key={s} value={s}>{s}</option>
                ))}
              </select>
            </div>
          </div>

          {/* Dynamic customer name + ticket number cells, one row per customer */}
          <div className="mt-4">
            <label className="text-xs text-slate-500 block mb-2">
              Customers ({form.customers.length})
            </label>
            <div className="space-y-2">
              {form.customers.map((c, i) => (
                <div key={i} className="grid grid-cols-2 gap-2 md:gap-3">
                  <input
                    className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                    value={c.name}
                    onChange={(e) => handleCustomerFieldChange(i, "name", e.target.value)}
                    placeholder={`Customer ${i + 1} name`}
                  />
                  <input
                    className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                    value={c.ticketNumber}
                    onChange={(e) => handleCustomerFieldChange(i, "ticketNumber", e.target.value)}
                    onBlur={() => handleTicketNumberBlur(i)}
                    placeholder={`Ticket number ${i + 1} (e.g. 077-1234567890)`}
                  />
                </div>
              ))}
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mt-4">
            <div>
              <label className="text-xs text-slate-500 block mb-1">From</label>
              <input
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                value={form.from}
                onChange={(e) => handleCityChange("from", e.target.value)}
                placeholder="Cairo"
                list="city-suggestions"
              />
            </div>
            <div>
              <label className="text-xs text-slate-500 block mb-1">To</label>
              <input
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                value={form.to}
                onChange={(e) => handleCityChange("to", e.target.value)}
                placeholder="Dubai"
                list="city-suggestions"
              />
            </div>
            <div>
              <label className="text-xs text-slate-500 block mb-1">Airline</label>
              <input
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                value={form.airline}
                onChange={(e) => handleAirlineChange(e.target.value)}
                placeholder="EgyptAir"
                list="airline-suggestions"
              />
            </div>
            <div>
              <label className="text-xs text-slate-500 block mb-1">Ticket issue date</label>
              <input
                type="date"
                lang="en-GB"
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                value={form.date}
                onChange={(e) => setForm({ ...form, date: e.target.value })}
              />
            </div>
            <div>
              <label className="text-xs text-slate-500 block mb-1">Net price</label>
              <input
                type="number"
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                value={form.netPrice}
                onChange={(e) => setForm({ ...form, netPrice: e.target.value })}
                placeholder="0"
              />
            </div>
            <div>
              <label className="text-xs text-slate-500 block mb-1">Sold price</label>
              <input
                type="number"
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
                value={form.soldPrice}
                onChange={(e) => setForm({ ...form, soldPrice: e.target.value })}
                placeholder="0"
              />
            </div>
            <div>
              <label className="text-xs text-slate-500 block mb-1">Profit (auto)</label>
              <div className="w-full border border-slate-200 bg-slate-50 rounded-lg px-3 py-2 text-sm text-emerald-700 font-semibold">
                {fmt(profit(form.netPrice, form.soldPrice))}
              </div>
            </div>
            <div className="md:col-span-3">
              <label className="text-xs text-slate-500 block mb-1">Notes</label>
              <textarea
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600 min-h-[80px]"
                value={form.notes}
                onChange={(e) => setForm({ ...form, notes: e.target.value.toUpperCase() })}
                placeholder="Optional"
              />
            </div>
          </div>

          <div className="flex gap-2 mt-4">
            <button
              onClick={handleSubmit}
              className="bg-teal-700 hover:bg-teal-800 text-white text-sm font-semibold rounded-lg px-4 py-2 flex items-center gap-1.5"
            >
              <Check size={16} /> {form.id ? "Save changes" : "Add ticket"}
            </button>
            {form.id && (
              <button
                onClick={handleCancel}
                className="border border-slate-300 text-slate-600 text-sm rounded-lg px-4 py-2 flex items-center gap-1.5"
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
            <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              className="w-full border border-slate-300 rounded-lg pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600"
              placeholder="Search by employee, company, ticket number, customer, destination, or airline"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
            />
          </div>
          <div className="relative sm:w-40">
            <Calendar size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
            <select
              className="w-full border border-slate-300 rounded-lg pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600 bg-white appearance-none"
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
            <Calendar size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
            <select
              className="w-full border border-slate-300 rounded-lg pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600 bg-white appearance-none"
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
            <Building2 size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
            <select
              className="w-full border border-slate-300 rounded-lg pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600 bg-white appearance-none"
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
            <User size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
            <select
              className="w-full border border-slate-300 rounded-lg pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600 bg-white appearance-none"
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
            <Plane size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
            <select
              className="w-full border border-slate-300 rounded-lg pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600 bg-white appearance-none"
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
          {suggestions.airlines.map((name) => (
            <option key={`u-${name}`} value={name} />
          ))}
          {AIRLINE_CODES.map((a) => (
            <option key={`a-${a.code}`} value={a.name.toUpperCase()} />
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
        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
          {filtered.length === 0 ? (
            <p className="text-center text-slate-400 text-sm py-10">
              {visibleTickets.length === 0 ? "No tickets recorded yet" : "No results match your search"}
            </p>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-xs border-collapse">
                <thead>
                  <tr className="bg-slate-50 text-slate-500 text-[11px] uppercase tracking-wide">
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
                        className={`border-t border-slate-100 leading-tight cursor-pointer ${i > 0 ? "border-t-0" : ""} ${isMulti ? "bg-amber-50 hover:bg-amber-100" : "hover:bg-slate-50"}`}
                      >
                        <td className="px-2.5 py-1 text-slate-600 whitespace-nowrap">{t.employee || "-"}</td>
                        <td className="px-2.5 py-1 text-slate-600 whitespace-nowrap">{t.company || "-"}</td>
                        <td className="px-2.5 py-1 text-slate-600 whitespace-nowrap">{t.supplier || "-"}</td>
                        <td className="px-2.5 py-1 text-slate-600 font-mono whitespace-nowrap">
                          {c.ticketNumber || "-"}
                        </td>
                        <td className="px-2.5 py-1 font-medium text-slate-800 whitespace-nowrap">
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
                        <td className="px-2.5 py-1 text-slate-600 whitespace-nowrap">{t.from} → {t.to}</td>
                        <td className="px-2.5 py-1 text-slate-600 whitespace-nowrap">{t.airline || "-"}</td>
                        <td className="px-2.5 py-1 text-slate-600 whitespace-nowrap">{t.date ? formatDisplayDate(t.date) : "-"}</td>
                        <td className="px-2.5 py-1 text-slate-600 text-right whitespace-nowrap">{fmt(t.netPrice)}</td>
                        <td className="px-2.5 py-1 text-slate-600 text-right whitespace-nowrap">{fmt(t.soldPrice)}</td>
                        <td className="px-2.5 py-1 font-semibold text-emerald-700 text-right whitespace-nowrap">{fmt(profit(t.netPrice, t.soldPrice))}</td>
                        <td className="px-2.5 py-1 text-right whitespace-nowrap">
                          {(currentUser.isAdmin || canEditTickets) ? (
                            <div className="flex gap-0.5 justify-end">
                              <button onClick={(ev) => { ev.stopPropagation(); handleEdit(t); }} className="text-slate-400 hover:text-teal-700 p-0.5">
                                <Pencil size={13} />
                              </button>
                              {currentUser.isAdmin && (
                                <button onClick={(ev) => { ev.stopPropagation(); handleDelete(t.id); }} className="text-slate-400 hover:text-red-600 p-0.5">
                                  <Trash2 size={13} />
                                </button>
                              )}
                            </div>
                          ) : (
                            <span className="text-slate-300 text-[11px] block text-right">—</span>
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
          <div className="bg-white rounded-xl border border-slate-200 overflow-hidden mt-6">
            <div className="px-4 py-3 border-b border-slate-100">
              <h2 className="font-semibold text-slate-900 text-sm">Totals by month</h2>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-slate-500 text-xs">
                    <th className="text-left px-3 py-2 font-medium">Month</th>
                    <th className="text-left px-3 py-2 font-medium">Tickets</th>
                    <th className="text-left px-3 py-2 font-medium">Total sales</th>
                    <th className="text-left px-3 py-2 font-medium">Total profit</th>
                    <th className="text-left px-3 py-2 font-medium"></th>
                  </tr>
                </thead>
                <tbody>
                  {monthlyBreakdown.map((m) => (
                    <tr key={m.key} className="border-t border-slate-100 hover:bg-slate-50">
                      <td className="px-3 py-2 font-medium text-slate-800">{monthLabel(m.key)}</td>
                      <td className="px-3 py-2 text-slate-600">{m.count}</td>
                      <td className="px-3 py-2 text-slate-600">{fmt(m.total)}</td>
                      <td className="px-3 py-2 font-semibold text-emerald-700">{fmt(m.profit)}</td>
                      <td className="px-3 py-2">
                        <div className="flex items-center gap-3 justify-end">
                          <button
                            onClick={() => exportMonth(m.key)}
                            className="text-slate-400 hover:text-teal-700 text-xs font-medium flex items-center gap-1"
                          >
                            <Download size={13} /> Export
                          </button>
                          <button
                            onClick={() => setSelectedMonth(m.key)}
                            className="text-teal-700 text-xs font-medium hover:underline"
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
          <div className="bg-white rounded-xl border border-slate-200 overflow-hidden mt-6">
            <div className="px-4 py-3 border-b border-slate-100">
              <h2 className="font-semibold text-slate-900 text-sm">Companies and their customers</h2>
            </div>
            <div className="divide-y divide-slate-100">
              {companyBreakdown.map((c) => (
                <div key={c.name} className="px-4 py-3">
                  <div className="flex items-center justify-between flex-wrap gap-2">
                    <div className="flex items-center gap-2">
                      <Building2 size={16} className="text-slate-400" />
                      <button
                        onClick={() => setSelectedCompany(c.name)}
                        className="font-medium text-slate-800 hover:text-teal-700 hover:underline text-sm"
                      >
                        {c.name}
                      </button>
                      <span className="text-xs text-slate-400">({c.count} tickets)</span>
                    </div>
                    <div className="flex items-center gap-4 text-xs text-slate-500">
                      <span>Sales: <span className="font-semibold text-slate-700">{fmt(c.total)}</span></span>
                      <span>Profit: <span className="font-semibold text-emerald-700">{fmt(c.profit)}</span></span>
                    </div>
                  </div>
                  <p className="text-xs text-slate-500 mt-1.5 pl-6">
                    Customers: {c.customers.join(", ")}
                  </p>
                </div>
              ))}
            </div>
          </div>
        )}

        <p className="text-xs text-slate-400 mt-3">
          This data is shared between signed-in employees. Login is a basic access gate, not strong security — treat it accordingly.
        </p>
        </>
        )}

        {activeSection === "hotels" && (
          <div className="bg-white border border-slate-200 rounded-xl p-12 text-center text-slate-400">
            <Building2 size={40} className="mx-auto mb-3 text-slate-300" />
            <p className="text-sm">Hotels section — nothing here yet.</p>
          </div>
        )}

        {activeSection === "cars" && (
          <div className="bg-white border border-slate-200 rounded-xl p-12 text-center text-slate-400">
            <Car size={40} className="mx-auto mb-3 text-slate-300" />
            <p className="text-sm">Transportation section — nothing here yet.</p>
          </div>
        )}

        {activeSection === "files" && (
          <div className="bg-white border border-slate-200 rounded-xl p-12 text-center text-slate-400">
            <FileText size={40} className="mx-auto mb-3 text-slate-300" />
            <p className="text-sm">Files section — nothing here yet.</p>
          </div>
        )}
      </div>

      {activeSection === "flights" && viewingTicket && (
        <div className="fixed inset-0 bg-white z-40 overflow-y-auto">
          <div className="max-w-3xl mx-auto p-4 md:p-6">
            <div className="flex items-center justify-between mb-6">
              <div className="flex items-center gap-2">
                <div className="bg-teal-700 text-white rounded-lg p-2">
                  <Ticket size={18} />
                </div>
                <h1 className="text-lg md:text-xl font-bold text-slate-900">Ticket details</h1>
              </div>
              <button
                onClick={closeTicketDetail}
                className="border border-slate-300 text-slate-600 text-sm rounded-lg px-3 py-2 flex items-center gap-1.5"
              >
                <X size={15} /> Close
              </button>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl divide-y divide-slate-100">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 p-4 md:p-5">
                <div>
                  <p className="text-xs text-slate-400 mb-1">Entered by</p>
                  <p className="text-sm font-medium text-slate-800">{viewingTicket.employee || "-"}</p>
                </div>
                <div>
                  <p className="text-xs text-slate-400 mb-1">Company</p>
                  <p className="text-sm font-medium text-slate-800">{viewingTicket.company || "-"}</p>
                </div>
                <div>
                  <p className="text-xs text-slate-400 mb-1">Supplier</p>
                  <p className="text-sm font-medium text-slate-800">{viewingTicket.supplier || "-"}</p>
                </div>
                <div>
                  <p className="text-xs text-slate-400 mb-1">Route</p>
                  <p className="text-sm font-medium text-slate-800">{viewingTicket.from} → {viewingTicket.to}</p>
                </div>
                <div>
                  <p className="text-xs text-slate-400 mb-1">Airline</p>
                  <p className="text-sm font-medium text-slate-800">{viewingTicket.airline || "-"}</p>
                </div>
                <div>
                  <p className="text-xs text-slate-400 mb-1">Ticket issue date</p>
                  <p className="text-sm font-medium text-slate-800">
                    {viewingTicket.date ? formatDisplayDate(viewingTicket.date) : "-"}
                  </p>
                </div>
              </div>

              <div className="p-4 md:p-5">
                <p className="text-xs text-slate-400 mb-2">
                  Customers ({getCustomers(viewingTicket).length})
                </p>
                <div className="border border-slate-200 rounded-lg overflow-hidden">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="bg-slate-50 text-slate-500 text-xs">
                        <th className="text-left px-3 py-2 font-medium">Customer</th>
                        <th className="text-left px-3 py-2 font-medium">Ticket number</th>
                      </tr>
                    </thead>
                    <tbody>
                      {getCustomers(viewingTicket).map((c, i) => (
                        <tr key={i} className="border-t border-slate-100">
                          <td className="px-3 py-2 text-slate-700">{c.name || "-"}</td>
                          <td className="px-3 py-2 text-slate-700 font-mono">{c.ticketNumber || "-"}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-3 gap-4 p-4 md:p-5">
                <div>
                  <p className="text-xs text-slate-400 mb-1">Net price</p>
                  <p className="text-sm font-medium text-slate-800">{fmt(viewingTicket.netPrice)}</p>
                </div>
                <div>
                  <p className="text-xs text-slate-400 mb-1">Sold price</p>
                  <p className="text-sm font-medium text-slate-800">{fmt(viewingTicket.soldPrice)}</p>
                </div>
                <div>
                  <p className="text-xs text-slate-400 mb-1">Profit</p>
                  <p className="text-sm font-semibold text-emerald-700">
                    {fmt(profit(viewingTicket.netPrice, viewingTicket.soldPrice))}
                  </p>
                </div>
              </div>

              <div className="p-4 md:p-5">
                <p className="text-xs text-slate-400 mb-2">Notes</p>
                <textarea
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-600 min-h-[100px]"
                  value={notesDraft}
                  onChange={(e) => { setNotesDraft(e.target.value.toUpperCase()); setNotesSaved(false); }}
                  placeholder="No notes yet — add some here"
                />
                <div className="flex items-center gap-2 mt-3">
                  <button
                    onClick={() => saveTicketNotes(viewingTicket.id)}
                    disabled={notesDraft === (viewingTicket.notes || "")}
                    className="bg-teal-700 hover:bg-teal-800 disabled:opacity-40 disabled:cursor-not-allowed text-white text-sm font-semibold rounded-lg px-4 py-2 flex items-center gap-1.5"
                  >
                    <Check size={15} /> Save notes
                  </button>
                  {notesSaved && (
                    <span className="text-xs text-emerald-700 font-medium">Saved</span>
                  )}
                </div>

                {Array.isArray(viewingTicket.notesHistory) && viewingTicket.notesHistory.length > 0 && (
                  <div className="mt-4 pt-3 border-t border-slate-100">
                    <p className="text-xs text-slate-400 mb-2">Edit history (most recent first)</p>
                    <div className="space-y-1.5 max-h-48 overflow-y-auto">
                      {[...viewingTicket.notesHistory].reverse().map((h, idx) => (
                        <div
                          key={idx}
                          className="text-xs bg-slate-50 border border-slate-100 rounded-lg px-2.5 py-1.5 flex items-start justify-between gap-3"
                        >
                          {h.type === "edit" ? (
                            <span className="text-slate-600 break-words">
                              <span className="font-semibold text-slate-700">Ticket edited: </span>
                              {(h.changes || []).join("; ")}
                            </span>
                          ) : (
                            <span className="text-slate-600 break-words">{h.value || "(cleared)"}</span>
                          )}
                          <span className="text-slate-400 whitespace-nowrap shrink-0">
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
        <div className="fixed inset-0 bg-slate-900/40 flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-xl border border-slate-200 p-5 w-full max-w-sm">
            <p className="text-sm text-slate-700 mb-4">{confirmDialog.message}</p>
            <div className="flex justify-end gap-2">
              <button
                onClick={() => setConfirmDialog(null)}
                className="border border-slate-300 text-slate-600 text-sm rounded-lg px-3 py-2"
              >
                Cancel
              </button>
              <button
                onClick={() => confirmDialog.onConfirm()}
                className="bg-teal-700 hover:bg-teal-800 text-white text-sm font-semibold rounded-lg px-3 py-2"
              >
                Confirm
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
