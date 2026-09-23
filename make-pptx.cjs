const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "TravelMate Admin Dashboard";

const C = {
  darkBg:   "0F1F2E",
  lightBg:  "F0F4F8",
  white:    "FFFFFF",
  teal:     "0EA5E9",
  tealMid:  "38BDF8",
  gray:     "64748B",
  lightGray:"CBD5E1",
  cardBg:   "FFFFFF",
  text:     "1E293B",
  muted:    "64748B",
}

const makeShadow = () => ({ type:"outer", color:"000000", blur:8, offset:2, angle:45, opacity:0.10 })

// KPI card — fixed height, short label
function kpiCard(slide, x, y, icon, label, value, bg) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w:2.1, h:0.9, rectRadius:0.08,
    fill:{ color: bg || "EFF6FF" },
    shadow: makeShadow(),
    line:{ color:"E2E8F0", width:0.5 }
  })
  slide.addText(icon, { x:x+0.1, y:y+0.1, w:0.36, h:0.36, fontSize:15, align:"center", margin:0 })
  slide.addText(value, { x:x+0.5, y:y+0.08, w:1.55, h:0.3, fontSize:12, bold:true, color:C.text, align:"left", margin:0 })
  slide.addText(label, { x:x+0.5, y:y+0.42, w:1.55, h:0.42, fontSize:8.5, color:C.muted, align:"left", margin:0 })
}

// Feature card — fills more vertical space
function featureCard(slide, x, y, w, h, icon, title, desc, accent) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius:0.1,
    fill:{ color:C.cardBg }, shadow:makeShadow(), line:{ color:"E2E8F0", width:0.5 }
  })
  slide.addShape(pres.shapes.OVAL, {
    x:x+0.15, y:y+0.15, w:0.52, h:0.52,
    fill:{ color: accent, transparency:80 }, line:{ color:"E0E0E0", width:0.3 }
  })
  slide.addText(icon, { x:x+0.15, y:y+0.15, w:0.52, h:0.52, fontSize:17, align:"center", margin:0 })
  slide.addText(title, { x:x+0.78, y:y+0.18, w:w-0.92, h:0.3, fontSize:11.5, bold:true, color:C.text, align:"left", margin:0 })
  slide.addText(desc, { x:x+0.15, y:y+0.6, w:w-0.3, h:h-0.75, fontSize:9.5, color:C.muted, align:"left", margin:0 })
}

function slideTitle(slide, text, light=true) {
  slide.addText(text, {
    x:0.5, y:0.2, w:9, h:0.62,
    fontSize:26, bold:true, color: light ? C.text : C.white,
    align:"left", margin:0
  })
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 1 — Title
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide()
  s.background = { color: C.darkBg }

  // decorative circles — kept inside slide
  s.addShape(pres.shapes.OVAL, { x:7.5, y:0.2, w:3.0, h:3.0, fill:{ color:C.teal, transparency:88 }, line:{ color:"E0E0E0", width:0 } })
  s.addShape(pres.shapes.OVAL, { x:8.4, y:3.8, w:1.5, h:1.5, fill:{ color:C.teal, transparency:82 }, line:{ color:"E0E0E0", width:0 } })

  // chip
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.55, y:1.1, w:2.2, h:0.34, rectRadius:0.14, fill:{ color:"0EA5E9", transparency:80 }, line:{ color:"E0E0E0", width:0 } })
  s.addText("ADMIN DASHBOARD", { x:0.6, y:1.12, w:2.1, h:0.28, fontSize:9, bold:true, color:C.tealMid, charSpacing:2, align:"center", margin:0 })

  s.addText("TravelMate", { x:0.5, y:1.52, w:8, h:1.05, fontSize:54, bold:true, color:C.white, align:"left", margin:0 })
  s.addText("Intelligent Travel Platform Management System", { x:0.5, y:2.62, w:7.5, h:0.45, fontSize:16, color:C.tealMid, align:"left", margin:0 })

  s.addShape(pres.shapes.LINE, { x:0.5, y:3.2, w:4.5, h:0, line:{ color:C.teal, width:1 } })

  s.addText("Mohamed Kotb", { x:0.5, y:3.35, w:4.5, h:0.3, fontSize:12, bold:true, color:C.lightGray, align:"left", margin:0 })
  s.addText("June 2026  ·  React · Supabase · Tailwind CSS · AI-Powered", { x:0.5, y:3.68, w:7, h:0.28, fontSize:10, color:C.gray, align:"left", margin:0 })

  // bottom stats row
  const stats = [["6","Pages"],["24","Unique KPIs"],["4","Real-time feeds"],["35","Page Docs"]]
  stats.forEach(([num, lbl], i) => {
    const x = 0.5 + i*2.3
    s.addText(num, { x, y:4.3, w:2.1, h:0.55, fontSize:32, bold:true, color:C.teal, align:"center", margin:0 })
    s.addText(lbl, { x, y:4.85, w:2.1, h:0.3, fontSize:10, color:C.lightGray, align:"center", margin:0 })
  })

  s.addNotes("Welcome everyone. Today I will walk you through TravelMate Admin Dashboard — a full-stack web system built to manage the TravelMate travel planning platform.")
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 2 — Project Overview
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide()
  s.background = { color: C.lightBg }
  slideTitle(s, "What Is This Project?")

  const pts = [
    { icon:"🌍", bold:"TravelMate Mobile App", desc:"An AI-powered travel planning app used by travelers to discover destinations and generate personalised itineraries." },
    { icon:"🖥️", bold:"The Admin Dashboard", desc:"A separate web system built for administrators to manage, monitor, and control the entire platform — no database access required." },
    { icon:"🎯", bold:"Project Goal", desc:"Give admins full visibility and control through a professional, real-time, dark-mode-ready interface built with modern web technologies." },
  ]

  pts.forEach(({ icon, bold, desc }, i) => {
    const y = 1.0 + i * 1.3
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.5, y, w:5.7, h:1.12, rectRadius:0.1, fill:{ color:C.cardBg }, shadow:makeShadow(), line:{ color:"E2E8F0", width:0.5 } })
    s.addShape(pres.shapes.OVAL, { x:0.7, y:y+0.3, w:0.52, h:0.52, fill:{ color:"EFF6FF" }, line:{ color:"E0E0E0", width:0 } })
    s.addText(icon, { x:0.7, y:y+0.29, w:0.52, h:0.52, fontSize:18, align:"center", margin:0 })
    s.addText([
      { text:bold+"\n", options:{ bold:true, fontSize:13, color:C.text, breakLine:true } },
      { text:desc, options:{ fontSize:10.5, color:C.muted } }
    ], { x:1.35, y:y+0.1, w:4.75, h:0.92, align:"left", margin:0 })
  })

  // right — tech grid
  s.addText("Built With", { x:6.6, y:0.95, w:3.1, h:0.35, fontSize:12, bold:true, color:C.muted, align:"center", margin:0 })
  const techs = [
    ["⚛️","React 18","1C7293"],["⚡","Vite","0369A1"],
    ["🎨","Tailwind CSS","7C3AED"],["🗄️","Supabase","16A34A"],
    ["📡","Realtime","0EA5E9"],["🤖","AI API","DC2626"],
  ]
  techs.forEach(([icon, label, color], i) => {
    const col = i % 2, row = Math.floor(i / 2)
    const x = 6.55 + col * 1.65, y = 1.38 + row * 1.1
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:1.5, h:0.9, rectRadius:0.08, fill:{ color:"FFFFFF" }, shadow:makeShadow(), line:{ color:"E2E8F0", width:0.5 } })
    s.addText(icon, { x, y:y+0.08, w:1.5, h:0.38, fontSize:18, align:"center", margin:0 })
    s.addText(label, { x, y:y+0.5, w:1.5, h:0.32, fontSize:9, bold:true, color:color, align:"center", margin:0 })
  })

  s.addNotes("TravelMate has two parts: a mobile app for travelers, and this admin dashboard for platform managers. The goal was to build something production-ready.")
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 3 — Tech Stack
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide()
  s.background = { color: C.darkBg }
  slideTitle(s, "Technology Stack", false)

  const items = [
    { icon:"⚛️", name:"React 18",          desc:"Functional components, hooks, Context API for global state", color:"38BDF8" },
    { icon:"⚡", name:"Vite",              desc:"Lightning-fast dev server with hot module replacement",        color:"A78BFA" },
    { icon:"🎨", name:"Tailwind CSS",      desc:"Utility-first styling — full dark mode on every element",     color:"34D399" },
    { icon:"🔀", name:"React Router v6",   desc:"Client-side navigation with protected route guards",           color:"FB923C" },
    { icon:"🗄️", name:"Supabase",          desc:"PostgreSQL + Auth + Realtime — anon & service-role clients",   color:"4ADE80" },
    { icon:"📡", name:"Supabase Realtime", desc:"WebSocket subscriptions — live notifications & feed updates",  color:"38BDF8" },
    { icon:"📊", name:"Recharts",          desc:"ComposedChart, AreaChart, BarChart, PieChart visualisations",  color:"F472B6" },
    { icon:"🤖", name:"AI API (Railway)",  desc:"Generates full travel itineraries from structured prompts",    color:"FBBF24" },
  ]

  items.forEach(({ icon, name, desc, color }, i) => {
    const col = i % 2, row = Math.floor(i / 2)
    const x = 0.4 + col * 4.75, y = 0.98 + row * 1.12
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:4.4, h:1.0, rectRadius:0.1, fill:{ color:"162130" }, line:{ color:"1E3A5F", width:0.5 } })
    s.addShape(pres.shapes.OVAL, { x:x+0.12, y:y+0.24, w:0.5, h:0.5, fill:{ color, transparency:82 }, line:{ color:"E0E0E0", width:0 } })
    s.addText(icon, { x:x+0.12, y:y+0.23, w:0.5, h:0.5, fontSize:18, align:"center", margin:0 })
    s.addText(name, { x:x+0.72, y:y+0.12, w:3.55, h:0.3, fontSize:12, bold:true, color:color, align:"left", margin:0 })
    s.addText(desc, { x:x+0.72, y:y+0.44, w:3.55, h:0.48, fontSize:9.5, color:"94A3B8", align:"left", margin:0 })
  })

  s.addNotes("The stack was chosen for speed and integration. Supabase gives us database, auth, and real-time in one service. Two Supabase clients: anon (respects RLS) and service-role (bypasses RLS for admin operations).")
}

// ════════════════════════════════════════════════════════════════════════════
// PAGE SLIDE FACTORY
// ════════════════════════════════════════════════════════════════════════════
function pageSlide(title, kpis, features, notes, accent) {
  const s = pres.addSlide()
  s.background = { color: C.lightBg }
  slideTitle(s, title)

  s.addText("KEY METRICS", { x:0.5, y:0.92, w:3, h:0.2, fontSize:7.5, bold:true, color:C.muted, charSpacing:2, margin:0 })
  const kpiBgs = ["EFF6FF","F0FDF4","FFF7ED","FDF4FF"]
  kpis.forEach(([icon, label, value], i) => {
    kpiCard(s, 0.45 + i*2.35, 1.14, icon, label, value, kpiBgs[i])
  })

  s.addText("KEY FEATURES", { x:0.5, y:2.17, w:3, h:0.2, fontSize:7.5, bold:true, color:C.muted, charSpacing:2, margin:0 })

  // 2 cols × 2 rows of feature cards, filling remaining space
  features.forEach(([icon, ftitle, desc], i) => {
    const col = i % 2, row = Math.floor(i / 2)
    const x = 0.45 + col * 4.8, y = 2.4 + row * 1.55
    featureCard(s, x, y, 4.52, 1.42, icon, ftitle, desc, accent)
  })

  s.addNotes(notes)
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 4 — Dashboard
// ════════════════════════════════════════════════════════════════════════════
pageSlide(
  "Dashboard — Platform Overview",
  [
    ["👥","Total registered accounts", "Total Users"],
    ["🔥","Users active this week",    "Active (7 Days)"],
    ["🔁","Last month's users back",   "Retention Rate"],
    ["📅","New accounts today",        "Joined Today"],
  ],
  [
    ["📈","User Growth Chart","Smooth area (new signups) + dashed cumulative line — ComposedChart"],
    ["⚡","Live Activity Feed","Supabase Realtime subscription — new entries appear instantly"],
    ["🔧","System Health Panel","Live ping to 4 services with ms response times"],
    ["🌗","Full Dark Mode","Every card, chart, input, and tooltip has a dark variant"],
  ],
  "The Dashboard is the landing page. The Retention Rate is the most strategic card — it tells us if users come back, not just how many we have.",
  "0EA5E9"
)

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 5 — Users
// ════════════════════════════════════════════════════════════════════════════
pageSlide(
  "Users — Management & Role Control",
  [
    ["📊","Avg. fields filled per user","Avg. Profile Completion"],
    ["✈️","Users with 3+ fields filled","Profiles ≥ 50% Filled"],
    ["📈","This month vs last month",   "MoM Growth"],
    ["⚠️","Users with <3 fields",       "Incomplete Profiles"],
  ],
  [
    ["🔍","Search & Filter","Real-time search by name, country, travel style — instant results"],
    ["👑","Role Management","Promote or demote admin with one click — action logged automatically"],
    ["🗑️","Bulk Delete","Multi-select users and delete with a confirmation step"],
    ["📉","Field Fill Rate Chart","See exactly which profile fields users skip the most"],
  ],
  "The Users page is the main operational tool. Avg Profile Completion replaced a redundant Total Users card. The fill rate chart helps understand data quality — critical for the AI trip generator.",
  "10B981"
)

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 6 — Activity Logs
// ════════════════════════════════════════════════════════════════════════════
pageSlide(
  "Activity Logs — Admin Audit Trail",
  [
    ["⚡","Signups + updates last 60m", "Events This Hour"],
    ["🕐","Time since last action",     "Last Event"],
    ["🔄","Users who updated profile",  "Update Rate %"],
    ["🔥","Days in a row with signup",  "Signup Streak"],
  ],
  [
    ["📋","Permanent Record","Every admin action logged — role changes, deletions, updates"],
    ["📡","Real-Time Updates","New entries appear instantly via Supabase Realtime subscription"],
    ["🔎","Filter & Search","Filter by date range and action type simultaneously"],
    ["🔒","Tamper-Evident","Logs are never deleted through the UI — full history always preserved"],
  ],
  "The Activity Logs page creates full accountability. Every role change, deletion, or update is permanently recorded with admin ID and timestamp — critical for security audits.",
  "8B5CF6"
)

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 7 — Analytics
// ════════════════════════════════════════════════════════════════════════════
pageSlide(
  "Analytics — Deep User Insights",
  [
    ["🌍","Most popular country",      "Top Country"],
    ["🌟","Most popular style + %",    "Top Travel Style"],
    ["🎂","Mean age from birth dates", "Avg. User Age"],
    ["☀️","Preferred climate type",    "Top Weather Pref."],
  ],
  [
    ["📅","Date Range Filter","All charts and KPIs respond to the selected period simultaneously"],
    ["👨‍👩‍👧","Children × Style","Do family travelers choose differently? Full cross-tab breakdown"],
    ["🎯","Age × Style Matrix","Which age groups prefer which travel styles — heatmap view"],
    ["🌦️","Weather Distribution","Climate preferences visualised across the entire user base"],
  ],
  "Analytics is a dedicated page for demographic insights. The date filter scopes everything at once. Children x Style and Age x Style matrices give unique insights not available anywhere else.",
  "F59E0B"
)

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 8 — Trip Simulator
// ════════════════════════════════════════════════════════════════════════════
pageSlide(
  "Trip Simulator — AI Itinerary Testing",
  [
    ["🗺️","Places in the database",    "Destinations"],
    ["🎯","Configured pref. categories","Pref. Categories"],
    ["🤖","LLaMA 3.3-70B via Railway", "AI Model"],
    ["⏱️","Time to generate itinerary", "Avg. Response Time"],
  ],
  [
    ["🧠","AI Generation","Structured prompt → full markdown itinerary — sent to Railway API"],
    ["📅","Day-by-Day Output","Custom parser: Morning / Afternoon / Evening per day — handles emoji headers"],
    ["🎨","Multiple Styles","Select several travel styles — AI blends them into one itinerary"],
    ["📋","Preference Panel","Read-only reference showing style-to-preference mappings from DB"],
  ],
  "The Trip Simulator is the most impressive feature. Our custom parsing engine handles emoji prefixes like 📅 Day 1 that the AI adds. Output matches exactly what mobile app users see.",
  "EC4899"
)

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 9 — Settings
// ════════════════════════════════════════════════════════════════════════════
pageSlide(
  "Settings — Config & Admin Profile",
  [
    ["⏱️","Time logged in (live timer)", "Session Duration"],
    ["⚠️","Users without travel style",  "Unconfigured Users"],
    ["📊","Admin accounts ÷ total × 100","Admin Coverage %"],
    ["👤","Regular account count",        "Non-Admin Users"],
  ],
  [
    ["⭐","Preference Editor","Edit travel style → pref. mappings that feed directly into the AI engine"],
    ["👥","Role Management","Grant or revoke admin access for any user from this page"],
    ["🔑","Change Password","Show/hide toggle on both fields — secure Supabase Auth updateUser call"],
    ["🚨","Danger Zone","Sign Out All Devices — global JWT invalidation across every active session"],
  ],
  "The Settings page doubles as the admin profile manager. The session duration is a live hook updating every 10 seconds. Danger Zone calls Supabase global sign-out — invalidates ALL sessions everywhere.",
  "6366F1"
)

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 10 — Cross-Cutting Features
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide()
  s.background = { color: C.darkBg }
  slideTitle(s, "What Makes It Stand Out", false)

  const cards = [
    { icon:"🌗", title:"Full Dark Mode",       color:"38BDF8", desc:"Every page, card, chart, input, and tooltip has an explicit dark variant. Preference saved in localStorage and survives refresh and browser restart." },
    { icon:"📡", title:"Real-Time Updates",    color:"4ADE80", desc:"Supabase Realtime WebSocket subscriptions push live data to the notification bell, activity feed, and top bar name updates — instantly, without polling." },
    { icon:"🔒", title:"Secure by Design",     color:"F472B6", desc:"JWT auth · RLS on all tables · self-demotion blocked · global sign-out · passwords never stored or logged in plain text at any point." },
    { icon:"🃏", title:"24 Unique KPI Cards",  color:"FBBF24", desc:"All 24 KPI cards across 6 pages are unique — no metric ever repeated. Three redundancies were identified and replaced with better insights during development." },
  ]

  cards.forEach(({ icon, title, color, desc }, i) => {
    const col = i % 2, row = Math.floor(i / 2)
    const x = 0.4 + col * 4.75, y = 1.0 + row * 2.2
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:4.4, h:2.05, rectRadius:0.12, fill:{ color:"162130" }, line:{ color:"1E3A5F", width:0.5 } })
    s.addShape(pres.shapes.OVAL, { x:x+0.18, y:y+0.2, w:0.6, h:0.6, fill:{ color, transparency:80 }, line:{ color:"E0E0E0", width:0 } })
    s.addText(icon, { x:x+0.18, y:y+0.2, w:0.6, h:0.6, fontSize:22, align:"center", margin:0 })
    s.addText(title, { x:x+0.9, y:y+0.25, w:3.35, h:0.35, fontSize:14, bold:true, color:color, align:"left", margin:0 })
    s.addText(desc, { x:x+0.18, y:y+0.9, w:4.05, h:1.05, fontSize:10, color:"94A3B8", align:"left", margin:0 })
  })

  s.addNotes("These four properties distinguish this from a basic CRUD app. Dark mode is complete — not partial. Real-time is used in 4 places. Security was designed in from the start. KPI uniqueness was enforced as a hard design rule.")
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 11 — Conclusion
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide()
  s.background = { color: C.darkBg }

  // decorative circles — inside bounds
  s.addShape(pres.shapes.OVAL, { x:0.2, y:0.2, w:2.5, h:2.5, fill:{ color:C.teal, transparency:90 }, line:{ color:"E0E0E0", width:0 } })
  s.addShape(pres.shapes.OVAL, { x:8.0, y:3.6, w:1.8, h:1.8, fill:{ color:C.teal, transparency:85 }, line:{ color:"E0E0E0", width:0 } })

  s.addText("Thank You", { x:0.5, y:0.35, w:9, h:0.9, fontSize:44, bold:true, color:C.white, align:"center", margin:0 })
  s.addText("TravelMate Admin Dashboard — A Complete Platform Control System", { x:0.5, y:1.28, w:9, h:0.38, fontSize:13, color:C.tealMid, align:"center", margin:0 })

  const pts = [
    ["✅","Production-Ready","6 pages, real-time, dark mode, AI integration, role management — fully functional"],
    ["✅","Well-Documented","35-page technical document covering every design decision and bug fix"],
    ["✅","Intentional Design","24 unique KPIs · no redundant sections · every feature justified"],
  ]

  pts.forEach(([icon, bold, desc], i) => {
    const y = 1.88 + i * 0.95
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.7, y, w:8.6, h:0.82, rectRadius:0.1, fill:{ color:"162130" }, line:{ color:"1E3A5F", width:0.5 } })
    s.addText(icon, { x:0.9, y:y+0.22, w:0.38, h:0.38, fontSize:16, align:"center", margin:0 })
    s.addText([
      { text:bold+" — ", options:{ bold:true, fontSize:12, color:C.tealMid } },
      { text:desc, options:{ fontSize:11, color:"94A3B8" } }
    ], { x:1.38, y:y+0.15, w:7.8, h:0.52, align:"left", margin:0 })
  })

  s.addText('"Built not just to work — but to impress."', { x:0.5, y:4.88, w:9, h:0.4, fontSize:13, italic:true, color:C.teal, align:"center", margin:0 })

  s.addNotes("Close by inviting questions. The project has a 35-page documentation file and a live demo ready to go.")
}

// ════════════════════════════════════════════════════════════════════════════
pres.writeFile({ fileName:"TravelMate_Presentation.pptx" })
  .then(() => console.log("✅  TravelMate_Presentation.pptx written"))
  .catch(e => { console.error("❌", e); process.exit(1) })
