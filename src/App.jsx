import { useState, useEffect } from "react";
import "./premium-ui.css";
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import { Document, Packer, Paragraph, TextRun } from "docx";
import { FiLayout, FiEdit3, FiBookOpen } from "react-icons/fi";

import { createClient } from "@supabase/supabase-js";
const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;
const supabase = createClient(supabaseUrl, supabaseAnonKey);

async function fetchWithAuth(url, options = {}, token = null) {
  let accessToken = token;

  // Try provided token, then session check
  if (!accessToken) {
    try {
      const { data } = await supabase.auth.getSession();
      accessToken = data?.session?.access_token;
    } catch (e) {
      console.warn("Manual session check failed in fetchWithAuth:", e);
    }
  }

  if (!accessToken) {
    console.error("No auth token available for:", url);
    throw new Error("Please sign in to continue.");
  }
  
  console.log("Fetching:", url, "with auth token length:", accessToken.length);
  const headers = { ...options.headers, Authorization: `Bearer ${accessToken}` };
  return fetch(url, { ...options, headers });
}

async function exportElementToPdf(element, fileName) {
  if (!element) return;
  const canvas = await html2canvas(element, { scale: 2, useCORS: true, backgroundColor: "#ffffff" });
  const imgData = canvas.toDataURL("image/png");
  const pdf = new jsPDF({
    orientation: canvas.width >= canvas.height ? "landscape" : "portrait",
    unit: "px",
    format: [canvas.width, canvas.height],
  });
  pdf.addImage(imgData, "PNG", 0, 0, canvas.width, canvas.height);
  pdf.save(fileName);
}

function normalizeExportText(value) {
  const raw = String(value || "");
  if (!raw) return "";
  if (typeof document !== "undefined") {
    const temp = document.createElement("div");
    temp.innerHTML = raw
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/(p|div|li|h[1-6])>/gi, "\n");
    return (temp.textContent || temp.innerText || raw).replace(/\n{3,}/g, "\n\n").trim();
  }
  return raw.replace(/<[^>]*>/g, " ").replace(/\s{2,}/g, " ").trim();
}

async function exportTextToDocx(text, fileName, title = "Rezolt Output") {
  const cleaned = normalizeExportText(text);
  const lines = cleaned.split(/\n+/).map(l => l.trim()).filter(Boolean);

  const isSectionHeader = line => /^[A-Z][A-Z\s&\/\-]+:?\s*$/.test(line);
  const isBullet = line => line.startsWith("•") || line.startsWith("-");
  const isContactLine = line => /\|/.test(line) || /^[+\d][\d\s\-]+$/.test(line) || /@/.test(line);

  const paragraphs = lines.map(line => {
    if (isSectionHeader(line)) {
      return new Paragraph({
        children: [new TextRun({ text: line.replace(/:$/, ""), bold: true, size: 26, color: "031D40" })],
        spacing: { before: 200, after: 80 },
        border: { bottom: { style: "single", size: 4, color: "031D40", space: 4 } },
      });
    }
    if (isBullet(line)) {
      return new Paragraph({
        children: [new TextRun({ text: line, size: 22 })],
        spacing: { after: 60 },
        indent: { left: 360 },
      });
    }
    if (isContactLine(line)) {
      return new Paragraph({
        children: [new TextRun({ text: line, size: 22, color: "444444" })],
        spacing: { after: 40 },
        alignment: "center",
      });
    }
    return new Paragraph({
      children: [new TextRun({ text: line, size: 22 })],
      spacing: { after: 80 },
    });
  });

  const doc = new Document({
    sections: [{
      properties: {},
      children: paragraphs.length ? paragraphs : [new Paragraph("No content available for export yet.")],
    }],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

// ─── COLOUR SYSTEM ────────────────────────────────────────────────────────────
const N1 = "#031D40"; // primary navy
const N2 = "#08284F"; // supporting navy
const AC = "#031D40"; // requested replacement accent
const LB = "#A4B4C9"; // muted navy tint
const PB = "#E4BE47"; // premium gold highlight
const G = AC; // success uses the brand color
const ER = "#EF4444"; // error
// Legacy aliases (used throughout components)
const O = AC;
const DARK = "var(--text-primary)";
const MID = "var(--text-secondary)";
const MUTED = "var(--text-muted)";
const FAINT = "var(--text-muted)";
const BORDER = "var(--border)";
const BG = "var(--app-bg)";
const WHITE = "var(--surface)";
const RESUME_FONT = "'Calibri', 'Carlito', 'Segoe UI', Arial, sans-serif";

// ─── RAZORPAY ────────────────────────────────────────────────────────────────

const RZP_KEY = import.meta.env.VITE_RAZORPAY_KEY_ID;
const RZP_IS_TEST = RZP_KEY?.startsWith("rzp_test_");

function loadRazorpay() {
  return new Promise((resolve, reject) => {
    if (window.Razorpay) {
      console.log("✓ Razorpay already loaded");
      return resolve(true);
    }
    
    const s = document.createElement("script");
    s.src = "https://checkout.razorpay.com/v1/checkout.js";
    s.async = true;
    s.onload = () => {
      console.log("✓ Razorpay script loaded");
      resolve(true);
    };
    s.onerror = (err) => {
      console.error("✗ Razorpay load failed:", err);
      reject(new Error("Razorpay CDN unreachable"));
    };
    document.body.appendChild(s);
  });
}

async function openRazorpay({ planId, amount, name, description, prefill, onSuccess, onDismiss }) {
  if (!RZP_KEY) {
    alert("Razorpay key is missing. Please check your Vercel environment variables.");
    return false;
  }

  let loaded = false;
  try {
    loaded = await loadRazorpay();
  } catch (e) {
    console.error("Razorpay CDN load failed:", e);
    alert("Failed to load payment gateway. Please disable any ad blockers and try again.");
    return false;
  }
  if (!loaded) { alert("Failed to load payment gateway. Please try again."); return false; }

  let orderId = null;
  let orderAmountPaise = Math.round(amount * 100);
  try {
    const res = await fetchWithAuth("/api/create-order", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        planId,
        receipt: `rezolt_${Date.now()}`,
        notes: { plan: planId, description },
      }),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || "Failed to create Razorpay order");
    if (!data.orderId) throw new Error("Could not create a secure payment order.");
    orderId = data.orderId;
    if (Number.isFinite(Number(data.amount))) {
      orderAmountPaise = Math.round(Number(data.amount));
    }
  } catch (e) {
    console.warn("Order creation failed:", e);
    alert(e?.message || "Please sign in again before making a payment.");
    return false;
  }

  return new Promise((resolve) => {
    const options = {
      key: RZP_KEY,
      amount: orderAmountPaise,
      currency: "INR",
      order_id: orderId || undefined,
      name: "Rezolt",
      description,
      prefill: { name: prefill.name, email: prefill.email },
      notes: { plan: planId || name, user_id: prefill.userId || "" },
      theme: { color: "#031D40" },
      method: { upi: true, card: true, netbanking: true, wallet: true },
      handler: async (response) => {
        try {
          await onSuccess?.(response);
        } finally {
          resolve(response);
        }
      },
      modal: {
        ondismiss: async () => {
          try {
            await onDismiss?.();
          } finally {
            resolve(null);
          }
        },
      },
    };
    console.log("Razorpay opening with options:", { key: RZP_KEY, amount: orderAmountPaise, order_id: orderId });
    const rzp = new window.Razorpay(options);
    rzp.on("payment.failed", (resp) => { console.error("Payment failed:", resp); resolve(false); });
    rzp.open();
    console.log("Razorpay rzp.open() called");
  });
}


// ─── TABS ─────────────────────────────────────────────────────────────────────

const TABS = [
  { id: "resume", iconClass: "fa-solid fa-file", label: "Resume", minPlan: "Free" },
  { id: "cover", iconClass: "fa-solid fa-envelope", label: "Cover Letter", minPlan: "starter" },
  { id: "referral", iconClass: "fa-solid fa-handshake", label: "Referral", minPlan: "starter" },
  { id: "interview", iconClass: "fa-solid fa-bullseye", label: "Interview Prep", minPlan: "starter" },
  { id: "reach", iconClass: "fa-solid fa-magnifying-glass", label: "Find & Reach", minPlan: "starter" },
  { id: "negotiate", iconClass: "fa-solid fa-dollar-sign", label: "Negotiate", minPlan: "unlimited" },
];
const PLAN_ORDER = ["Free", "starter", "Pro", "unlimited"];
const normalizePlan = (plan) => {
  const map = { free: "Free", starter: "starter", pro: "Pro", unlimited: "unlimited" };
  return plan ? (map[plan.toLowerCase()] ?? null) : null;
};
const canAccess = (userPlan, minPlan) => PLAN_ORDER.indexOf(userPlan ?? "Free") >= PLAN_ORDER.indexOf(minPlan);
const ADMIN_EMAIL = "hema.manoharan13@outlook.com";
const MAX_RESUME_CHARS = 12000;
const MAX_JD_CHARS = 8000;

function prepareInputForAi(text, maxChars, label) {
  const cleaned = String(text || "")
    .replace(/\r/g, "")
    .replace(/\u0000/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  if (cleaned.length <= maxChars) {
    return { text: cleaned, trimmed: false, originalLength: cleaned.length };
  }

  const head = Math.floor(maxChars * 0.72);
  const tail = maxChars - head;
  return {
    text: `${cleaned.slice(0, head)}\n\n[Rezolt note: ${label} trimmed for safe AI processing. Focus on the most relevant experience and requirements.]\n\n${cleaned.slice(-tail)}`,
    trimmed: true,
    originalLength: cleaned.length,
  };
}

function isUsableGeneration(text) {
  return typeof text === "string" && text.trim() && !/(network error|something went wrong|session expired|please sign in)/i.test(text);
}

// ─── PROMPTS ──────────────────────────────────────────────────────────────────

const TEMPLATE_INSTRUCTIONS = {
  creative: "STYLE: Use a creative, modern layout. Lead with a bold professional summary. Use clear section dividers. Skills as grouped tags. Emphasise achievements with strong metrics.",
  modern: "STYLE: Clean minimal layout. Short punchy summary (2-3 sentences). Prioritise whitespace. No fluff. Bullet points tight and action-led.",
  bold: "STYLE: Bold, ATS-heavy format. Name prominent at top. Section headers in ALL CAPS and underlined with dashes. Dense, keyword-rich content throughout.",
  classic: "STYLE: Traditional formal format. Name centred at top. Use formal language. Objective statement instead of summary. Conservative section ordering: Summary, Experience, Education, Skills.",
  elegant: "STYLE: Sophisticated two-tone layout. Serif-style writing. Elevated vocabulary. Use full sentences for bullet points, not fragments. Professional and polished throughout.",
  compact: "STYLE: Two-column dense format. Left column: contact, skills, education. Right column: summary and experience. Maximum information density. Keep bullets to one line each.",
  minimal: "STYLE: Ultra-clean. Extreme whitespace. Only the most essential information. Maximum 3 bullets per role. Summary one sentence only. No clutter whatsoever.",
  tech: "STYLE: Developer/tech format. Lead with technical skills. GitHub/portfolio links prominent. Bullet points focus on technologies, systems built, and scale. Use technical terminology confidently.",
  warm: "STYLE: Friendly approachable tone. Use first person where natural. Highlight teamwork, collaboration, and culture fit alongside achievements. Warm professional language throughout.",
};

const PROMPTS = {
  resume: (r, jd, template = "creative") => `You are an elite resume strategist who has helped candidates land roles at Google, Amazon, Microsoft, Meta, Flipkart, Swiggy, and top Indian MNCs. Your rewrites compete against thousands of applicants for highly selective roles.

CRITICAL: Output ONLY the final rewritten resume. No thinking, no steps, no preamble. Start immediately with CONTACT INFORMATION.

COMPANY RULE: The TARGET company is named in the JOB DESCRIPTION. The RESUME contains current and past employers. Never confuse them. The summary must name the TARGET company from the JD.

${TEMPLATE_INSTRUCTIONS[template] || TEMPLATE_INSTRUCTIONS.creative}

RESUME STANDARDS:
- Lead with impact, not responsibility. What changed because of this person?
- Power verbs: Spearheaded, Orchestrated, Transformed, Scaled, Optimised, Delivered, Drove, Partnered, Streamlined, Championed
- Use real metrics from the resume where they exist (%, headcount, revenue, time saved, scale)
- If the candidate did not provide a metric for something, describe the impact clearly and specifically — do NOT invent or estimate numbers that are not in the resume
- ATS keywords must be exact-match to the JD, not synonyms
- Summary must position the candidate as a domain expert, not a generalist
- Highlight scope: team size led, geography, business units, budget ownership — only if mentioned in the resume

FORMATTING:
- Plain text only. No markdown (no **, *, ##, ---)
- No em dashes. Use commas or full stops.
- Section headers in ALL CAPS followed by a colon
- Bullets use: •
- Blank line between sections
- Every role: 4-6 bullet points. Use a metric only when one is genuinely present in the source resume.

OUTPUT FORMAT (start immediately, no preamble):

CONTACT INFORMATION:
[Full Name]
[Phone] | [Email] | [LinkedIn URL] | [Portfolio/GitHub if relevant]
[City, State]

PROFESSIONAL SUMMARY:
[5-6 sentences. Open with years of experience and domain expertise. Name the target company and role. Cover the top 4 JD requirements with specific proof from the resume. End with a statement of unique value that differentiates this candidate.]

CORE EXPERTISE:
[6-8 domain-specific areas as a comma-separated list matching JD language exactly]

WORK EXPERIENCE:

[Company Name] | [Job Title] | [Month Year] to [Month Year or Present]
• [Power verb] [specific initiative] resulting in [X% improvement / INR X revenue / X users impacted] by [specific method]
• [Power verb] [specific system/process/team] achieving [measurable outcome] across [scope]
• [Power verb] [specific deliverable] reducing [cost/time/errors] by [X%] within [timeframe]
• [Power verb] [cross-functional collaboration] enabling [business outcome] for [stakeholder group]
• [Power verb] [data/analysis/technical work] driving [decision or outcome] with [business impact]
• [Power verb] [leadership or mentorship] resulting in [team outcome or growth metric]

[Repeat for EVERY role — minimum 6 bullets each, all quantified]

KEY SKILLS:

Core Expertise: [domain skills matching JD exactly, 6-8 items]
Technical Skills: [tools, platforms, methodologies, 8-10 items]
Tools and Platforms: [software, systems, databases used]

EDUCATION:
[Degree] | [Institution] | [Year] | [CGPA or distinction if strong]

CERTIFICATIONS AND AWARDS:
[Certification name] | [Issuing body] | [Year]

RULE: Use only real details from the resume. Expand and reframe what exists. Do not invent experiences, numbers, or metrics that are not in the source resume.

RESUME:
${r}

JOB DESCRIPTION:
${jd}`,

  cover: (r, jd) => `You are an elite career coach writing a cover letter that will be read by a hiring manager at a top-tier company. This letter must be extraordinary, not standard, not generic, not template-driven.

CRITICAL COMPANY RULE: The TARGET company is named in the JOB DESCRIPTION. Address only the TARGET company. Never reference a resume company as the target.

FORMATTING: Plain text only. No markdown. No em dashes. Three paragraphs. No headers. No "I am writing to apply."

FAANG/MAANG COVER LETTER STANDARDS:
- Open with a bold, specific claim that earns immediate attention, not a compliment to the company
- Every claim backed by a real number or outcome from the resume
- Show you understand the company's actual problems or goals (infer from the JD)
- Avoid all clichés: "passionate", "results-driven", "team player", "dynamic", "excited to"
- Tone: confident, direct, intelligent, like a peer conversation not a plea
- End with a specific confident call to action

FORMAT:

Dear Hiring Manager,

[Paragraph 1: 3 sentences maximum. Open with your most impressive relevant achievement, lead with the number. Connect it directly to why you are the right person for THIS role at THIS company. Make them want to keep reading.]

[Paragraph 2: 3-4 sentences. Pick 2 quantified achievements from the resume that map directly to JD requirements. Show you understand what the company is trying to solve. Demonstrate you have done this before at measurable scale.]

[Paragraph 3: 2-3 sentences. Show genuine insight about the company or role from the JD. Express specific interest, not general excitement. Close with a confident direct call to action: not "I hope to hear from you" but "I would welcome a conversation."]

Warm regards,
[Full name from resume]

RESUME:
${r}

JOB DESCRIPTION:
${jd}`,

  referral: (r, jd) => `You are a networking expert writing referral messages for a candidate targeting a competitive role. These messages must feel human, specific, and compelling, not like templates.

CRITICAL COMPANY RULE: The TARGET company is named in the JOB DESCRIPTION. All messages must reference that company. Resume companies are past or current employers only.

FORMATTING: Plain text only. No markdown. No em dashes. Use exact labels below.

FAANG/MAANG REFERRAL STANDARDS:
- Be specific, generic messages get ignored
- Name one real impressive achievement from the resume to establish credibility fast
- Show the contact exactly why THIS candidate is worth referring, make it easy for them to say yes
- The DM must feel written by a real person, not an AI
- Include a specific ask that is easy for the contact to action

OUTPUT FORMAT:

VERSION 1: COLD OUTREACH MESSAGE
(Under 300 characters. For a new or weak contact. Specific, direct, no fluff.)

[Message: Open with their name. Name the exact role and company from JD. Drop one specific impressive stat or strength from the resume. End with a friendly low-friction ask.]

VERSION 2: WARM CONTACT MESSAGE
(For a warm contact. 8-10 lines. Human, specific, easy to action.)

[Line 1: Warm personal opener]
[Lines 2-3: Mention the specific role at the target company, show genuine excitement for THIS role]
[Lines 4-6: Two specific quantified reasons they are a strong fit, pulled directly from the resume. Make the contact feel confident referring them.]
[Lines 7-8: Specific easy ask, "would you be open to a referral or share any insights on the process?"]
[Lines 9-10: Offer to send resume directly. Thank them genuinely.]

RESUME:
${r}

JOB DESCRIPTION:
${jd}`,

  interview: (r, jd) => `You are a senior interview coach who has prepped candidates for Google, Amazon, Microsoft, Meta, McKinsey, and top Indian product and consulting companies. Your prep is specific, tactical, and battle-tested.

CRITICAL COMPANY RULE: The company they are interviewing at is named in the JOB DESCRIPTION. All questions and answers must be tailored to that company. Resume companies are past employers.

FORMATTING: Plain text only. No markdown. No em dashes. Use exact labels.

FAANG/MAANG INTERVIEW PREP STANDARDS:
- Questions must be the ones THIS company actually asks for THIS type of role
- STAR answers must use REAL details from the resume, no generic examples
- Every answer must have a metric in the Result section
- The TIE-BACK must connect explicitly to a JD requirement
- Smart questions must show strategic thinking and company knowledge
- Include one curveball or values-based question that top companies commonly ask

OUTPUT FORMAT:

SECTION 1: TOP 5 QUESTIONS THEY WILL ASK

Q1: [Most likely first question for this specific role and company]
WHY THEY ASK: [One sentence, what competency or signal are they testing]
SITUATION: [Specific scenario from this candidate's actual resume, name the company, project, or initiative]
ACTION: [Exactly what they did, specific steps, decisions, methods]
RESULT: [The outcome with a real number from the resume, if no number derive a credible proxy]
TIE-BACK: [One sentence connecting this answer to the specific JD requirement it addresses]

[Repeat format for Q2, Q3, Q4, Q5. Include at least one behavioural, one situational, one role-specific, and one values or culture question.]

SECTION 2: 2 POWER QUESTIONS TO ASK THE INTERVIEWER

Q1: [Strategic question showing deep knowledge of the company's current priorities, infer from JD]
WHY THIS WORKS: [One sentence explaining what signal this sends about the candidate]

Q2: [Forward-looking question about growth, team, or impact, shows genuine interest and ambition]
WHY THIS WORKS: [One sentence]

SECTION 3: 3 WATCH-OUTS FOR THIS INTERVIEW

1. [Specific trap or common mistake candidates make for this role or company, with how to avoid it]
2. [Second watch-out, could be about tone, format, or common interview bias for this role]
3. [Third watch-out, practical tip for standing out in the final few minutes]

RESUME:
${r}

JOB DESCRIPTION:
${jd}`,

  reach: (r, jd) => `You are a LinkedIn outreach and job search strategist helping a candidate break into a highly competitive company. Your guidance is tactical, specific, and immediately actionable.

CRITICAL COMPANY RULE: The TARGET company to reach is named in the JOB DESCRIPTION. All search queries and messages must target that company. Resume companies are past employers.

FORMATTING: Plain text only. No markdown. No em dashes. Use exact labels.

FAANG/MAANG OUTREACH STANDARDS:
- Search queries must be precise and immediately usable
- Messages must reference one specific impressive achievement from the resume
- Every message must feel human, not templated
- Include the psychology behind each tip
- Give actual Boolean strings, not descriptions of Boolean strings

OUTPUT FORMAT:

SECTION 1: FIND THE EXACT RIGHT PERSON

LinkedIn Search Query (copy and paste directly):
[Exact search string using target company name from JD and most relevant TA or hiring title for this role]

Boolean Search String:
[Full Boolean string using OR, AND, "" ready to paste into LinkedIn search bar]

Who to prioritise contacting (in order):
• [Most senior relevant TA or hiring manager title] — owns the hiring decision for this level
• [Second contact type] — likely screens candidates and knows the pipeline
• [Third contact type] — team member who can provide referral weight

SECTION 2: OUTREACH MESSAGES

Connection Request Note (under 300 characters):
[Message: Name + target role + company + one specific achievement from resume + low-friction ask]

Follow-Up DM (send within 48 hours of connection accepting):
[8-10 lines. Reference their specific background or post if possible. Drop your most relevant achievement with a number. Make a single specific ask. Offer to make it easy for them.]

Cold InMail Subject Line:
[Subject line that earns a click, specific, intriguing, not salesy]

Cold InMail Body (6-8 lines):
[Open with shared context or their company achievement. One line about yourself with your best metric. The ask. Simple and respectful.]

SECTION 3: 5 TACTICS THAT ACTUALLY WORK FOR INDIAN MNC HIRING

1. [Tactic with the psychology behind it, why it increases response rate]
2. [Tactic, specific and actionable, not generic advice]
3. [Tactic, timing, follow-up cadence, or platform-specific tip]
4. [Tactic, how to use mutual connections or alumni networks in India]
5. [Tactic, what to do if you get no response after two messages]

RESUME:
${r}

JOB DESCRIPTION:
${jd}`,
};

// ─── RESUME TEMPLATE RENDERER ────────────────────────────────────────────────

function parseResumeSections(text) {
  const clean = text.replace(/\*\*(.*?)\*\*/g, "$1").replace(/\*(.*?)\*/g, "$1").replace(/^#{1,3}\s*/gm, "").replace(/---/g, "").replace(/\u2014/g, ",");
  const lines = clean.split("\n");
  const sections = {};
  let currentSection = "header";
  let buffer = [];

  const flush = () => { if (buffer.length) { sections[currentSection] = ((sections[currentSection] || "") + buffer.join("\n")).trim(); buffer = []; } };

  lines.forEach(line => {
    const t = line.trim();
    // Match ANY all-caps line ending with colon as a section header
    const isHeader = t.length > 2 && t.length < 70 && t === t.toUpperCase() && (t.endsWith(":") || /^(CONTACT|PROFESSIONAL|WORK|KEY SKILLS|SKILLS|EDUCATION|CERTIFICATIONS|SUMMARY|EXPERIENCE)/.test(t)) && !/^[•\d\-]/.test(t);
    if (isHeader) {
      flush();
      currentSection = t.replace(/:$/, "").trim().toLowerCase().replace(/\s+/g, "_");
    } else {
      buffer.push(line);
    }
  });
  flush();

  // Normalize common section name variants
  const normalize = (obj) => {
    const map = {
      contact_information: ["contact", "personal_information", "personal_details"],
      professional_summary: ["summary", "profile", "objective", "professional_profile", "career_summary"],
      work_experience: ["experience", "employment", "employment_history", "professional_experience", "career_history"],
      key_skills: ["skills", "technical_skills", "core_competencies", "competencies", "areas_of_expertise"],
      education: ["educational_background", "academic_background", "qualifications"],
      certifications: ["certification", "courses", "training", "professional_development"],
    };
    Object.entries(map).forEach(([canonical, aliases]) => {
      if (!obj[canonical]) {
        for (const alias of aliases) {
          if (obj[alias]) { obj[canonical] = obj[alias]; break; }
        }
      }
    });
    return obj;
  };

  return normalize(sections);
}

function renderResumeWithTemplate(text, template, photoUrl, showBranding = true) {
  if (!text) return null;
  const s = parseResumeSections(text);

  // If parsing yielded nothing useful, fall back to standard renderer
  const hasContent = s.professional_summary || s.work_experience || s.contact_information;
  if (!hasContent) return <div style={{ textAlign: "left", fontFamily: RESUME_FONT }}>{renderOutput(text)}</div>;

  const headerLines = (s.contact_information || s.header || "").trim().split("\n").filter(Boolean);
  const name = headerLines[0] || "";
  const contact = headerLines[1] || "";
  const location = headerLines[2] || "";
  const summary = (s.professional_summary || "").trim();
  const experience = (s.work_experience || "").trim();
  const skills = (s.key_skills || s.skills || "").trim();
  const education = (s.education || "").trim();
  const certs = (s.certifications || "").trim();

  const tmplColors = {
    creative: "#031D40", modern: "#374151", bold: "#111827",
    classic: "#1E3A5F", elegant: "#78350F", compact: "#374151",
    minimal: "#6B7280", tech: "#031D40", warm: "#92400e",
  };
  const accentColor = tmplColors[template] || "#02457A";

  // Template name badge shown at top for the free plan only
  const badge = showBranding ? (
    <div style={{ display: "inline-flex", alignItems: "center", gap: 6, background: `${accentColor}15`, border: `1px solid ${accentColor}30`, borderRadius: 20, padding: "3px 10px", marginBottom: 16, fontSize: 10, fontWeight: 700, color: accentColor, textTransform: "uppercase", letterSpacing: ".06em" }}>
      {template} template
    </div>
  ) : null;

  const renderLines = (block, bulletColor) => block.split("\n").filter(l => l.trim()).map((l, i) => {
    const t = l.trim();
    if (t.startsWith("•") || t.startsWith("-")) return <div key={i} style={{ display: "flex", gap: 8, marginBottom: 5, alignItems: "flex-start" }}><span style={{ color: bulletColor, flexShrink: 0, marginTop: 1 }}>•</span><span style={{ fontSize: 12, color: "var(--text-mid)", lineHeight: 1.7 }}>{t.replace(/^[•\-]\s*/, "")}</span></div>;
    if (t.includes("|")) return <div key={i} style={{ marginTop: 12, marginBottom: 4 }}><span style={{ fontSize: 13, fontWeight: 700, color: "var(--text)" }}>{t.split("|")[0].trim()}</span><span style={{ fontSize: 11, color: "var(--text-muted)", marginLeft: 8 }}>{t.split("|").slice(1).join("|").trim()}</span></div>;
    if (t) return <div key={i} style={{ fontSize: 12, color: "var(--text-muted)", marginBottom: 4 }}>{t}</div>;
    return null;
  });

  // ── CREATIVE (sidebar layout, blue) ──
  if (template === "creative") return (
    <div style={{ fontFamily: RESUME_FONT }}>
      {badge}
      <div style={{ display: "flex", gap: 0, fontFamily: "inherit", minHeight: 400, borderRadius: 12, overflow: "hidden", border: "1px solid var(--border)" }}>
        <div style={{ width: 185, flexShrink: 0, background: "linear-gradient(180deg,#001B48,#02457A)", padding: "24px 16px" }}>
          <div style={{ width: 52, height: 52, borderRadius: "50%", background: photoUrl ? "transparent" : "#02457A", overflow: "hidden", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, fontWeight: 700, color: "white", marginBottom: 12 }}>
            {photoUrl ? <img src={photoUrl} alt="Photo" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : (name[0] || "?")}
          </div>
          <div style={{ fontSize: 13, fontWeight: 700, color: "white", marginBottom: 3, lineHeight: 1.3 }}>{name}</div>
          <div style={{ fontSize: 10, color: "#C9D6E4", marginBottom: 16 }}>{contact.split("|")[0]?.trim()}</div>
          <div style={{ fontSize: 9, color: "rgba(255,255,255,.45)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 6 }}>Contact</div>
          {contact.split("|").map((c, i) => <div key={i} style={{ fontSize: 9, color: "rgba(255,255,255,.7)", marginBottom: 3, wordBreak: "break-all" }}>{c.trim()}</div>)}
          {location && <div style={{ fontSize: 9, color: "rgba(255,255,255,.6)", marginBottom: 14, marginTop: 4 }}>{location}</div>}
          {skills && <><div style={{ fontSize: 9, color: "rgba(255,255,255,.45)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 7, marginTop: 12 }}>Key Skills</div><div style={{ display: "flex", flexWrap: "wrap", gap: 3 }}>{skills.split(",").slice(0, 10).map((sk, i) => <span key={i} style={{ background: "rgba(3,29,64,.28)", color: "#C9D6E4", fontSize: 8, padding: "2px 6px", borderRadius: 3 }}>{sk.trim()}</span>)}</div></>}
        </div>
        <div style={{ flex: 1, padding: "24px 20px", background: "white", minWidth: 0 }}>
          {summary && <><div style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#02457A", marginBottom: 7, borderBottom: "1.5px solid #02457A", paddingBottom: 4, display: "inline-block" }}>Professional Summary</div><p style={{ fontSize: 11, color: "#374151", lineHeight: 1.8, marginBottom: 18 }}>{summary}</p></>}
          {experience && <><div style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#02457A", marginBottom: 8, borderBottom: "1.5px solid #02457A", paddingBottom: 4, display: "inline-block" }}>Experience</div><div style={{ marginTop: 4 }}>{renderLines(experience, "#02457A")}</div></>}
          {education && <><div style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#02457A", marginBottom: 8, borderBottom: "1.5px solid #02457A", paddingBottom: 4, display: "inline-block", marginTop: 16 }}>Education</div><div style={{ marginTop: 4 }}>{renderLines(education, "#02457A")}</div></>}
        </div>
      </div>
    </div>
  );

  // ── MODERN (clean, avatar header) ──
  if (template === "modern") return (
    <div style={{ fontFamily: RESUME_FONT }}>
      {badge}
      <div style={{ border: "1px solid var(--border)", borderRadius: 12, overflow: "hidden" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14, padding: "20px 24px", borderBottom: "1px solid var(--border)", background: "var(--surface2)" }}>
          <div style={{ width: 48, height: 48, borderRadius: "50%", background: photoUrl ? "transparent" : "#374151", overflow: "hidden", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20, fontWeight: 700, color: "white", flexShrink: 0 }}>
            {photoUrl ? <img src={photoUrl} alt="Photo" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : name[0] || "?"}
          </div>
          <div><div style={{ fontSize: 18, fontWeight: 700, color: "var(--text)" }}>{name}</div><div style={{ fontSize: 12, color: "var(--text-muted)", marginTop: 2 }}>{contact}{location ? ` · ${location}` : ""}</div></div>
        </div>
        <div style={{ padding: "20px 24px", background: "var(--surface)" }}>
          {skills && <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 16 }}>{skills.split(",").slice(0, 12).map((sk, i) => <span key={i} style={{ background: "var(--surface2)", color: "var(--text-mid)", fontSize: 11, padding: "3px 10px", borderRadius: 20, border: "1px solid var(--border)" }}>{sk.trim()}</span>)}</div>}
          {summary && <p style={{ fontSize: 13, color: "var(--text-mid)", lineHeight: 1.85, marginBottom: 18 }}>{summary}</p>}
          {experience && <><div style={{ fontSize: 10, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "var(--text-muted)", marginBottom: 10 }}>Experience</div>{renderLines(experience, "#374151")}</>}
          {education && <><div style={{ fontSize: 10, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "var(--text-muted)", marginBottom: 8, marginTop: 16 }}>Education</div>{renderLines(education, "#374151")}</>}
        </div>
      </div>
    </div>
  );

  // ── BOLD (dark header bar) ──
  if (template === "bold") return (
    <div style={{ fontFamily: RESUME_FONT }}>
      {badge}
      <div style={{ border: "1px solid #111827", borderRadius: 12, overflow: "hidden" }}>
        <div style={{ background: "#111827", padding: "20px 24px" }}>
          <div style={{ fontSize: 22, fontWeight: 800, color: "white", letterSpacing: "-.01em" }}>{name}</div>
          <div style={{ fontSize: 12, color: "rgba(255,255,255,.55)", marginTop: 5 }}>{contact}{location ? ` | ${location}` : ""}</div>
        </div>
        <div style={{ padding: "20px 24px", background: "var(--surface)" }}>
          {summary && <><div style={{ fontSize: 10, fontWeight: 800, letterSpacing: ".1em", color: "#111827", borderBottom: "2px solid #111827", paddingBottom: 4, marginBottom: 8, display: "inline-block" }}>PROFESSIONAL SUMMARY</div><p style={{ fontSize: 12, color: "var(--text-mid)", lineHeight: 1.8, marginBottom: 18 }}>{summary}</p></>}
          {experience && <><div style={{ fontSize: 10, fontWeight: 800, letterSpacing: ".1em", color: "#111827", borderBottom: "2px solid #111827", paddingBottom: 4, marginBottom: 8, display: "inline-block" }}>WORK EXPERIENCE</div><div style={{ marginTop: 4 }}>{renderLines(experience, "#111827")}</div></>}
          {skills && <><div style={{ fontSize: 10, fontWeight: 800, letterSpacing: ".1em", color: "#111827", borderBottom: "2px solid #111827", paddingBottom: 4, marginTop: 16, marginBottom: 8, display: "inline-block" }}>KEY SKILLS</div><div style={{ fontSize: 12, color: "var(--text-mid)", lineHeight: 1.9, marginTop: 4 }}>{skills}</div></>}
          {education && <><div style={{ fontSize: 10, fontWeight: 800, letterSpacing: ".1em", color: "#111827", borderBottom: "2px solid #111827", paddingBottom: 4, marginTop: 16, marginBottom: 8, display: "inline-block" }}>EDUCATION</div><div style={{ marginTop: 4 }}>{renderLines(education, "#111827")}</div></>}
        </div>
      </div>
    </div>
  );

  // ── CLASSIC (centred, serif) ──
  if (template === "classic") return (
    <div style={{ fontFamily: RESUME_FONT }}>
      {badge}
      <div style={{ fontFamily: RESUME_FONT, border: "1px solid var(--border)", borderRadius: 10, padding: "28px 28px", background: "var(--surface)" }}>
        <div style={{ textAlign: "center", marginBottom: 18, paddingBottom: 16, borderBottom: "2px solid #1E3A5F" }}>
          <div style={{ fontSize: 22, fontWeight: 700, color: "#1E3A5F", letterSpacing: ".02em" }}>{name}</div>
          <div style={{ fontSize: 12, color: "#5A6D88", marginTop: 5 }}>{contact}</div>
          {location && <div style={{ fontSize: 12, color: "#5A6D88" }}>{location}</div>}
        </div>
        {summary && <><div style={{ fontSize: 10, fontWeight: 700, letterSpacing: ".12em", textTransform: "uppercase", color: "#1E3A5F", marginBottom: 8 }}>Objective</div><p style={{ fontSize: 12, color: "#374151", lineHeight: 1.9, marginBottom: 18 }}>{summary}</p><div style={{ borderBottom: "1px solid #CBD5E1", marginBottom: 16 }} /></>}
        {experience && <><div style={{ fontSize: 10, fontWeight: 700, letterSpacing: ".12em", textTransform: "uppercase", color: "#1E3A5F", marginBottom: 10 }}>Professional Experience</div>{renderLines(experience, "#1E3A5F")}<div style={{ borderBottom: "1px solid #CBD5E1", margin: "16px 0" }} /></>}
        {skills && <><div style={{ fontSize: 10, fontWeight: 700, letterSpacing: ".12em", textTransform: "uppercase", color: "#1E3A5F", marginBottom: 8 }}>Skills</div><div style={{ fontSize: 12, color: "#374151", lineHeight: 1.9, marginBottom: 16 }}>{skills}</div><div style={{ borderBottom: "1px solid #CBD5E1", marginBottom: 16 }} /></>}
        {education && <><div style={{ fontSize: 10, fontWeight: 700, letterSpacing: ".12em", textTransform: "uppercase", color: "#1E3A5F", marginBottom: 8 }}>Education</div>{renderLines(education, "#1E3A5F")}</>}
      </div>
    </div>
  );

  // ── ELEGANT (gold left bar, serif) ──
  if (template === "elegant") return (
    <div style={{ fontFamily: RESUME_FONT }}>
      {badge}
      <div style={{ display: "flex", gap: 0, fontFamily: RESUME_FONT, border: "1px solid #C9A96E", borderRadius: 10, overflow: "hidden" }}>
        <div style={{ width: 5, background: "linear-gradient(180deg,#C9A96E,#78350F)", flexShrink: 0 }} />
        <div style={{ flex: 1, padding: "24px 24px", background: "var(--surface)" }}>
          <div style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 22, fontWeight: 700, color: "#2C1810" }}>{name}</div>
            <div style={{ fontSize: 12, color: "#C9A96E", marginTop: 4 }}>{contact}{location ? ` · ${location}` : ""}</div>
          </div>
          <div style={{ height: 1, background: "linear-gradient(90deg,#C9A96E,transparent)", marginBottom: 18 }} />
          {summary && <><div style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#78350F", marginBottom: 8 }}>Profile</div><p style={{ fontSize: 12, color: "#374151", lineHeight: 1.95, marginBottom: 18, fontStyle: "italic" }}>{summary}</p></>}
          {experience && <><div style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#78350F", marginBottom: 10 }}>Experience</div>{renderLines(experience, "#C9A96E")}</>}
          {skills && <><div style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#78350F", marginBottom: 8, marginTop: 16 }}>Competencies</div><div style={{ fontSize: 12, color: "#374151", lineHeight: 1.9 }}>{skills}</div></>}
          {education && <><div style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#78350F", marginBottom: 8, marginTop: 16 }}>Education</div>{renderLines(education, "#C9A96E")}</>}
        </div>
      </div>
    </div>
  );

  // ── COMPACT (two-column) ──
  if (template === "compact") return (
    <div style={{ fontFamily: RESUME_FONT }}>
      {badge}
      <div style={{ border: "1px solid var(--border)", borderRadius: 10, overflow: "hidden" }}>
        <div style={{ background: "var(--surface2)", padding: "14px 20px", borderBottom: "1px solid var(--border)" }}>
          <div style={{ fontSize: 17, fontWeight: 700, color: "var(--text)" }}>{name}</div>
          <div style={{ fontSize: 11, color: "var(--text-muted)", marginTop: 3 }}>{contact}</div>
        </div>
        <div style={{ display: "flex", gap: 0, background: "var(--surface)" }}>
          <div style={{ flex: 2, padding: "18px 20px", borderRight: "1px solid var(--border)", minWidth: 0 }}>
            {summary && <><div style={{ fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".1em", color: "var(--text-muted)", marginBottom: 6 }}>Summary</div><p style={{ fontSize: 11, color: "var(--text-mid)", lineHeight: 1.7, marginBottom: 14 }}>{summary}</p></>}
            {experience && <><div style={{ fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".1em", color: "var(--text-muted)", marginBottom: 6 }}>Experience</div>{renderLines(experience, "var(--text-muted)")}</>}
          </div>
          <div style={{ width: 160, flexShrink: 0, padding: "18px 16px" }}>
            {location && <div style={{ fontSize: 10, color: "var(--text-muted)", marginBottom: 14 }}>{location}</div>}
            {skills && <><div style={{ fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".1em", color: "var(--text-muted)", marginBottom: 7 }}>Skills</div><div style={{ display: "flex", flexWrap: "wrap", gap: 3, marginBottom: 14 }}>{skills.split(",").map((sk, i) => <span key={i} style={{ background: "var(--surface2)", color: "var(--text-mid)", fontSize: 9, padding: "2px 6px", borderRadius: 3, border: "1px solid var(--border)" }}>{sk.trim()}</span>)}</div></>}
            {education && <><div style={{ fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".1em", color: "var(--text-muted)", marginBottom: 6 }}>Education</div>{renderLines(education, "var(--text-muted)")}</>}
            {certs && <><div style={{ fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".1em", color: "var(--text-muted)", marginBottom: 6, marginTop: 12 }}>Certifications</div>{renderLines(certs, "var(--text-muted)")}</>}
          </div>
        </div>
      </div>
    </div>
  );

  // ── MINIMAL (ultra clean whitespace) ──
  if (template === "minimal") return (
    <div style={{ fontFamily: RESUME_FONT }}>
      {badge}
      <div style={{ maxWidth: 560, padding: "8px 4px" }}>
        <div style={{ marginBottom: 32 }}>
          <div style={{ fontSize: 28, fontWeight: 300, color: "var(--text)", letterSpacing: "-.03em" }}>{name}</div>
          <div style={{ fontSize: 12, color: "var(--text-faint)", marginTop: 6, letterSpacing: ".02em" }}>{contact}{location ? ` · ${location}` : ""}</div>
        </div>
        {summary && <p style={{ fontSize: 13, color: "var(--text-mid)", lineHeight: 2, marginBottom: 32, fontWeight: 300 }}>{summary.split(".").slice(0, 2).join(".")}.</p>}
        {experience && <><div style={{ fontSize: 9, letterSpacing: ".2em", textTransform: "uppercase", color: "var(--text-faint)", marginBottom: 14 }}>Experience</div>{renderLines(experience, "var(--text-faint)")}</>}
        {skills && <><div style={{ fontSize: 9, letterSpacing: ".2em", textTransform: "uppercase", color: "var(--text-faint)", marginBottom: 10, marginTop: 28 }}>Skills</div><div style={{ fontSize: 12, color: "var(--text-mid)", lineHeight: 2, fontWeight: 300 }}>{skills}</div></>}
        {education && <><div style={{ fontSize: 9, letterSpacing: ".2em", textTransform: "uppercase", color: "var(--text-faint)", marginBottom: 10, marginTop: 28 }}>Education</div>{renderLines(education, "var(--text-faint)")}</>}
      </div>
    </div>
  );

  // ── TECH (dark, developer style) ──
  if (template === "tech") return (
    <div style={{ fontFamily: RESUME_FONT }}>
      {badge}
      <div style={{ background: "#0d1117", borderRadius: 10, padding: "20px 24px", fontFamily: RESUME_FONT, border: "1px solid #30363d" }}>
        <div style={{ borderBottom: "1px solid #21262d", paddingBottom: 14, marginBottom: 16 }}>
          <div style={{ fontSize: 18, fontWeight: 700, color: "#C9D6E4" }}>{name}</div>
          <div style={{ fontSize: 11, color: "#8b949e", marginTop: 5 }}>{contact}{location ? ` // ${location}` : ""}</div>
        </div>
        {skills && <><div style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#C9D6E4", marginBottom: 8 }}>{">"} Stack</div><div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginBottom: 18 }}>{skills.split(",").map((sk, i) => <span key={i} style={{ background: "#161b22", border: "1px solid #30363d", color: "#E2EAF3", fontSize: 10, padding: "3px 9px", borderRadius: 4 }}>{sk.trim()}</span>)}</div></>}
        {summary && <><div style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#C9D6E4", marginBottom: 8 }}>{">"} About</div><p style={{ fontSize: 11, color: "#8b949e", lineHeight: 1.8, marginBottom: 18 }}>{summary}</p></>}
        {experience && <><div style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#C9D6E4", marginBottom: 10 }}>{">"} Experience</div>{experience.split("\n").filter(l => l.trim()).map((l, i) => { const t = l.trim(); if (t.startsWith("•") || t.startsWith("-")) return <div key={i} style={{ fontSize: 11, color: "#8b949e", marginBottom: 5, paddingLeft: 14 }}>{"// "}{t.replace(/^[•\-]\s*/, "")}</div>; if (t.includes("|")) return <div key={i} style={{ fontSize: 12, color: "#e6edf3", fontWeight: 700, marginTop: 12, marginBottom: 4 }}>{t}</div>; return t ? <div key={i} style={{ fontSize: 11, color: "#C9D6E4", marginBottom: 3 }}>{t}</div> : null; })}</>}
        {education && <><div style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#C9D6E4", marginBottom: 8, marginTop: 18 }}>{">"} Education</div>{education.split("\n").filter(l => l.trim()).map((l, i) => <div key={i} style={{ fontSize: 11, color: "#8b949e", marginBottom: 4 }}>{l.trim()}</div>)}</>}
      </div>
    </div>
  );

  // ── WARM (earthy amber) ──
  if (template === "warm") return (
    <div style={{ fontFamily: RESUME_FONT }}>
      {badge}
      <div style={{ background: "#fffbf5", borderRadius: 10, border: "1px solid #fde68a", overflow: "hidden" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14, padding: "18px 22px", borderBottom: "2px solid #fde68a" }}>
          <div style={{ width: 48, height: 48, borderRadius: "50%", background: "linear-gradient(135deg,#f59e0b,#d97706)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20, fontWeight: 700, color: "white", flexShrink: 0 }}>{name[0] || "?"}</div>
          <div><div style={{ fontSize: 18, fontWeight: 700, color: "#78350f" }}>{name}</div><div style={{ fontSize: 12, color: "#d97706", marginTop: 3 }}>{contact}{location ? ` · ${location}` : ""}</div></div>
        </div>
        <div style={{ padding: "20px 22px" }}>
          {summary && <div style={{ background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 10, padding: "12px 16px", marginBottom: 18 }}><p style={{ fontSize: 12, color: "#374151", lineHeight: 1.85, margin: 0 }}>{summary}</p></div>}
          {experience && <><div style={{ fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".1em", color: "#92400e", marginBottom: 10 }}>Experience</div>{renderLines(experience, "#d97706")}</>}
          {skills && <><div style={{ fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".1em", color: "#92400e", marginBottom: 8, marginTop: 16 }}>Skills</div><div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>{skills.split(",").map((sk, i) => <span key={i} style={{ background: "#fde68a", color: "#92400e", fontSize: 11, padding: "3px 10px", borderRadius: 20 }}>{sk.trim()}</span>)}</div></>}
          {education && <><div style={{ fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".1em", color: "#92400e", marginBottom: 8, marginTop: 16 }}>Education</div>{renderLines(education, "#d97706")}</>}
        </div>
      </div>
    </div>
  );

  // ── DEFAULT fallback ──
  return <div style={{ textAlign: "left", fontFamily: RESUME_FONT }}>{renderOutput(text)}</div>;
}
// ─── RENDERER ────────────────────────────────────────────────────────────────

function renderOutput(text) {
  if (!text) return null;
  const clean = text
    .replace(/\*\*(.*?)\*\*/g, "$1")
    .replace(/\*(.*?)\*/g, "$1")
    .replace(/^#{1,3}\s*/gm, "")
    .replace(/---/g, "")
    .replace(/\u2014/g, ",");

  const lines = clean.split("\n");
  const elements = [];

  lines.forEach((line, i) => {
    const t = line.trim();
    if (!t) { elements.push(<div key={i} style={{ height: 10 }} />); return; }

    const isHeader = (t === t.toUpperCase() && t.length > 4 && t.length < 60 && t.endsWith(":") && !/^[•Q\d]/.test(t)) ||
      t.match(/^(SECTION \d|VERSION \d|PROFESSIONAL SUMMARY|WORK EXPERIENCE|SKILLS|EDUCATION|CERTIFICATIONS):?$/i);

    if (isHeader) {
      elements.push(
        <div key={i} style={{ marginTop: elements.length > 0 ? 24 : 0, marginBottom: 10 }}>
          <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.14em", color: O, borderBottom: `2px solid ${O}`, paddingBottom: 5, display: "inline-block", textTransform: "uppercase" }}>
            {t.replace(/:$/, "")}
          </div>
        </div>
      );
      return;
    }

    if (/^Q\d:/.test(t)) {
      elements.push(
        <div key={i} style={{ background: "rgba(249,115,22,0.05)", borderLeft: `3px solid ${O}`, borderRadius: "0 8px 8px 0", padding: "10px 14px", margin: "16px 0 6px" }}>
          <div style={{ fontSize: 14, fontWeight: 700, color: DARK }}>{t.replace(/^Q\d:\s*/, "")}</div>
        </div>
      );
      return;
    }

    if (/^(WHY THEY ASK|WHY THIS WORKS|TIE-BACK):/.test(t)) {
      const colon = t.indexOf(":");
      elements.push(
        <div key={i} style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "flex-start" }}>
          <span style={{ fontSize: 11, fontWeight: 700, color: MUTED, minWidth: 110, paddingTop: 3, flexShrink: 0, textTransform: "uppercase", letterSpacing: "0.06em" }}>{t.substring(0, colon)}:</span>
          <span style={{ fontSize: 14, color: MUTED, lineHeight: 1.7, fontStyle: "italic" }}>{t.substring(colon + 1).trim()}</span>
        </div>
      );
      return;
    }

    if (/^(SITUATION|ACTION|RESULT):/.test(t)) {
      const colon = t.indexOf(":");
      elements.push(
        <div key={i} style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "flex-start", paddingLeft: 8 }}>
          <span style={{ fontSize: 11, fontWeight: 700, color: O, minWidth: 80, paddingTop: 3, flexShrink: 0, textTransform: "uppercase", letterSpacing: "0.06em" }}>{t.substring(0, colon)}:</span>
          <span style={{ fontSize: 14, color: MID, lineHeight: 1.7 }}>{t.substring(colon + 1).trim()}</span>
        </div>
      );
      return;
    }

    if (/^[•\-]/.test(t)) {
      elements.push(
        <div key={i} style={{ display: "flex", gap: 10, marginBottom: 7, paddingLeft: 4 }}>
          <span style={{ color: O, flexShrink: 0, marginTop: 4, fontSize: 14, lineHeight: 1 }}>•</span>
          <span style={{ fontSize: 14, color: MID, lineHeight: 1.75 }}>{t.replace(/^[•\-]\s*/, "")}</span>
        </div>
      );
      return;
    }

    if (/^\d+\./.test(t)) {
      const num = t.match(/^(\d+)\./)[1];
      elements.push(
        <div key={i} style={{ display: "flex", gap: 10, marginBottom: 8 }}>
          <span style={{ background: O, color: "#fff", fontSize: 11, fontWeight: 700, width: 20, height: 20, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, marginTop: 3 }}>{num}</span>
          <span style={{ fontSize: 14, color: MID, lineHeight: 1.75 }}>{t.replace(/^\d+\.\s*/, "")}</span>
        </div>
      );
      return;
    }

    if (t.includes("|")) {
      const parts = t.split("|").map(p => p.trim());
      elements.push(
        <div key={i} style={{ marginTop: 14, marginBottom: 6 }}>
          <span style={{ fontSize: 14, fontWeight: 700, color: DARK }}>{parts[0]}</span>
          {parts[1] && <span style={{ fontSize: 13, color: MUTED, marginLeft: 8 }}>| {parts[1]}</span>}
          {parts[2] && <span style={{ fontSize: 13, color: FAINT, marginLeft: 8 }}>| {parts[2]}</span>}
        </div>
      );
      return;
    }

    if (/^VERSION \d:/.test(t)) {
      elements.push(
        <div key={i} style={{ marginTop: 20, marginBottom: 10 }}>
          <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.14em", color: "#fff", background: O, padding: "3px 10px", borderRadius: 4, textTransform: "uppercase" }}>
            {t.replace(/:$/, "")}
          </span>
        </div>
      );
      return;
    }

    if (t.includes(":") && t.indexOf(":") < 25 && t.indexOf(":") > 1) {
      const colon = t.indexOf(":");
      const label = t.substring(0, colon).trim();
      const val = t.substring(colon + 1).trim();
      if (val && label.length < 25) {
        elements.push(
          <div key={i} style={{ display: "flex", gap: 10, marginBottom: 7, alignItems: "flex-start" }}>
            <span style={{ fontSize: 12, fontWeight: 700, color: DARK, minWidth: 100, paddingTop: 3, flexShrink: 0 }}>{label}:</span>
            <span style={{ fontSize: 14, color: MID, lineHeight: 1.75 }}>{val}</span>
          </div>
        );
        return;
      }
    }

    elements.push(<p key={i} style={{ fontSize: 14, color: MID, lineHeight: 1.85, marginBottom: 8 }}>{t}</p>);
  });

  return elements;
}

function extractOutreachBlock(text, regex) {
  const match = String(text || "").match(regex);
  return match ? match[1].trim() : "";
}

function OutreachQuickActions({ type, text }) {
  const [variant, setVariant] = useState("warm");
  const [copied, setCopied] = useState(false);

  const referralCold = extractOutreachBlock(text, /VERSION 1:\s*(?:LINKEDIN CONNECTION REQUEST|COLD OUTREACH MESSAGE)\s*\n(?:\(.*?\)\s*\n)?([\s\S]*?)(?=VERSION 2:|$)/i);
  const referralWarm = extractOutreachBlock(text, /VERSION 2:\s*(?:DIRECT MESSAGE|WARM CONTACT MESSAGE)\s*\n(?:\(.*?\)\s*\n)?([\s\S]*?)$/i);
  const linkedInQuery = extractOutreachBlock(text, /LinkedIn Search Query \(copy and paste directly\):\s*\n([\s\S]*?)(?=Boolean Search String:|Who to prioritise contacting|SECTION 2:|$)/i);
  const connectionNote = extractOutreachBlock(text, /Connection Request Note \(under 300 characters\):\s*\n([\s\S]*?)(?=Follow-Up DM|Cold InMail Subject Line|SECTION 3:|$)/i);
  const followUpDm = extractOutreachBlock(text, /Follow-Up DM \(send within 48 hours of connection accepting\):\s*\n([\s\S]*?)(?=Cold InMail Subject Line:|Cold InMail Body|SECTION 3:|$)/i);
  const coldSubject = extractOutreachBlock(text, /Cold InMail Subject Line:\s*\n([\s\S]*?)(?=Cold InMail Body \(6-8 lines\):|SECTION 3:|$)/i);
  const coldBody = extractOutreachBlock(text, /Cold InMail Body \(6-8 lines\):\s*\n([\s\S]*?)(?=SECTION 3:|$)/i);

  const options = type === "referral"
    ? [
      { id: "warm", label: "Warm contact", body: referralWarm || followUpDm || connectionNote },
      { id: "cold", label: "Cold outreach", body: referralCold || connectionNote || coldBody },
    ]
    : [
      { id: "warm", label: "Warm follow-up", body: followUpDm || connectionNote },
      { id: "cold", label: "Cold InMail", body: [coldSubject, coldBody].filter(Boolean).join("\n\n") || connectionNote },
    ];

  const selected = options.find(item => item.id === variant) || options[0];

  const copySelected = async () => {
    const payload = selected?.body || linkedInQuery;
    if (!payload) return;
    try {
      await navigator.clipboard.writeText(payload);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1800);
    } catch { }
  };

  const launchLinkedIn = () => {
    const query = encodeURIComponent(linkedInQuery || "talent acquisition recruiter hiring manager");
    window.open(`https://www.linkedin.com/search/results/people/?keywords=${query}`, "_blank", "noopener,noreferrer");
  };

  if (!selected?.body && !linkedInQuery) return null;

  return (
    <div style={{ marginBottom: 16, background: "var(--surface2)", border: `1px solid ${BORDER}`, borderRadius: 16, padding: "14px 16px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center", flexWrap: "wrap", marginBottom: 10 }}>
        <div>
          <div style={{ fontSize: 12, fontWeight: 700, color: DARK }}>{type === "referral" ? "Warm or cold referral helper" : "LinkedIn launch helper"}</div>
          <div style={{ fontSize: 11, color: MUTED, marginTop: 3 }}>{type === "referral" ? "Pick the right tone before you copy and send." : "Switch message styles and launch your search faster."}</div>
        </div>
        {linkedInQuery && (
          <button onClick={launchLinkedIn} style={{ background: "var(--grad)", color: WHITE, border: "none", borderRadius: 999, padding: "8px 12px", fontSize: 11, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
            Launch on LinkedIn
          </button>
        )}
      </div>

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: selected?.body ? 10 : 0 }}>
        {options.map(option => (
          <button key={option.id} onClick={() => setVariant(option.id)} style={{ background: variant === option.id ? "var(--accent-soft)" : WHITE, color: variant === option.id ? O : MUTED, border: `1px solid ${variant === option.id ? O : BORDER}`, borderRadius: 999, padding: "7px 12px", fontSize: 11, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
            {option.label}
          </button>
        ))}
      </div>

      {selected?.body && (
        <div style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 12, padding: "12px 14px", fontSize: 13, color: MID, lineHeight: 1.75, whiteSpace: "pre-wrap" }}>
          {selected.body}
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center", flexWrap: "wrap", marginTop: 10 }}>
        <div style={{ fontSize: 11, color: FAINT }}>{linkedInQuery ? "LinkedIn search query is ready to use." : "Use the tone that best matches your relationship."}</div>
        <button onClick={copySelected} style={{ background: WHITE, color: O, border: `1px solid ${BORDER}`, borderRadius: 10, padding: "8px 12px", fontSize: 11, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
          {copied ? "Copied!" : `Copy ${selected?.label || "message"}`}
        </button>
      </div>
    </div>
  );
}

// ─── GLOBAL CSS ───────────────────────────────────────────────────────────────

const GLOBAL_CSS = `
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg: #F9FAFB;
    --surface: rgba(255,255,255,0.94);
    --surface-strong: #FFFFFF;
    --surface2: #F3F4F6;
    --surface3: #F9FAFB;
    --border: rgba(31,41,55,0.06);
    --text: #1F2937;
    --text-mid: #374151;
    --text-muted: #6B7280;
    --text-faint: #94A3B8;
    --accent: #031D40;
    --accent-glow: rgba(3,29,64,0.14);
    --accent-soft: rgba(3,29,64,0.06);
    --gold: #E4BE47;
    --gold-soft: rgba(228,190,71,0.14);
    --navy: #031D40;
    --navy2: #08284F;
    --pale: #E5E7EB;
    --grad: linear-gradient(135deg, #031D40 0%, #08284F 48%, #031D40 82%, #E4BE47 122%);
    --grad-subtle: linear-gradient(135deg, #FFFFFF 0%, #F9FAFB 60%, #F3F4F6 100%);
    --soft-shadow: 0 8px 30px rgba(0,0,0,0.04);
    --shadow-sm: 0 8px 30px rgba(0,0,0,0.04);
    --shadow-md: 0 16px 36px rgba(0,0,0,0.05);
    --shadow-lg: 0 22px 50px rgba(0,0,0,0.06);
    --shadow-accent: 0 14px 34px rgba(3,29,64,0.12);
    --ring: inset 0 1px 0 rgba(255,255,255,0.92);
    --radius: 20px;
    --radius-lg: 30px;
  }

  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #020B1E;
      --surface: rgba(3,28,74,0.90);
      --surface-strong: #031D40;
      --surface2: #08284F;
      --surface3: #08264F;
      --border: rgba(228,190,71,0.14);
      --text: #EEF3F9;
      --text-mid: #C7D4E5;
      --text-muted: #9BAEC7;
      --text-faint: rgba(199,212,229,0.58);
      --accent: #A4B4C9;
      --accent-glow: rgba(164,180,201,0.22);
      --accent-soft: rgba(164,180,201,0.12);
      --gold: #E4BE47;
      --gold-soft: rgba(228,190,71,0.14);
      --pale: rgba(220,229,240,0.08);
      --grad: linear-gradient(135deg, #031D40 0%, #08284F 48%, #031D40 82%, #E4BE47 122%);
      --grad-subtle: linear-gradient(135deg, #02132D 0%, #08284F 100%);
      --shadow-sm: 0 8px 20px rgba(0,0,0,0.30);
      --shadow-md: 0 18px 40px rgba(0,0,0,0.40);
      --shadow-lg: 0 30px 70px rgba(0,0,0,0.52);
      --ring: inset 0 1px 0 rgba(255,255,255,0.05);
    }
  }

  html { scroll-behavior: smooth; }

  body {
    font-family: 'Open Sans', sans-serif;
    background:
      radial-gradient(circle at top left, rgba(201,169,110,0.12), transparent 0, transparent 26%),
      radial-gradient(circle at top right, rgba(3,29,64,0.10), transparent 0, transparent 32%),
      radial-gradient(circle at bottom left, rgba(3,29,64,0.08), transparent 0, transparent 28%),
      linear-gradient(180deg, #FAFCFE 0%, #F3F7FB 48%, #FAFCFE 100%);
    color: var(--text);
    line-height: 1.6;
    -webkit-font-smoothing: antialiased;
  }

  #root {
    position: relative;
  }
  #root::before {
    content: '';
    position: fixed;
    inset: 0;
    pointer-events: none;
    background: linear-gradient(180deg, rgba(255,255,255,0.22), transparent 18%, transparent 82%, rgba(201,169,110,0.03));
    opacity: 0.55;
  }

  textarea, input, select, button { font-family: 'Open Sans', sans-serif; }
  [style*="DM Serif Display"], [style*="Cormorant Garamond"] { font-family: 'Raleway', sans-serif !important; }
  [style*="Plus Jakarta Sans"], [style*="Manrope"] { font-family: 'Open Sans', sans-serif !important; }
  .hero-title { font-family: 'Raleway', sans-serif !important; }

  input:focus, textarea:focus, select:focus {
    border-color: var(--accent) !important;
    box-shadow: 0 0 0 3px var(--accent-glow) !important;
    outline: none !important;
  }

  /* ── ANIMATIONS ─────────────────────────────────────────────────── */
  @keyframes spin         { to { transform: rotate(360deg); } }
  @keyframes fadeUp       { from { opacity:0; transform:translateY(24px); } to { opacity:1; transform:none; } }
  @keyframes fadeIn       { from { opacity:0; transform:translateY(10px); } to { opacity:1; transform:none; } }
  @keyframes slideRight   { from { opacity:0; transform:translateX(-16px); } to { opacity:1; transform:none; } }
  @keyframes scaleIn      { from { opacity:0; transform:scale(0.95); } to { opacity:1; transform:scale(1); } }
  @keyframes modalIn      { from { opacity:0; transform:translate(-50%,-50%) scale(0.95); } to { opacity:1; transform:translate(-50%,-50%) scale(1); } }
  @keyframes float        { 0%,100% { transform:translateY(0); } 50% { transform:translateY(-10px); } }
  @keyframes pulse-ring   { 0% { transform:scale(1); opacity:.8; } 100% { transform:scale(1.6); opacity:0; } }
  @keyframes gradShift    { 0%,100% { background-position:0% 50%; } 50% { background-position:100% 50%; } }
  @keyframes shimmer      { 0% { opacity:.4; } 50% { opacity:1; } 100% { opacity:.4; } }
  @keyframes dotPulse     { 0%,100%{opacity:.3;transform:scale(.8);} 50%{opacity:1;transform:scale(1);} }

  @keyframes scanline   { from { top: -10% } to { top: 110% } }
  @keyframes arrowPulse { 0%,100% { transform: translateX(0) } 50% { transform: translateX(5px) } }
  @keyframes glow       { 0%,100% { box-shadow: 0 0 12px rgba(3,29,64,.22) } 50% { box-shadow: 0 0 28px rgba(3,29,64,.45) } }

  @keyframes spin       { from { transform: rotate(0deg) } to { transform: rotate(360deg) } }
  @keyframes bounce     { 0%, 100% { transform: translateY(0) } 50% { transform: translateY(-8px) } }
  @keyframes slideIn    { from { opacity: 0; transform: translateX(-20px) } to { opacity: 1; transform: translateX(0) } }
  @keyframes slideInRight { from { opacity: 0; transform: translateX(20px) } to { opacity: 1; transform: translateX(0) } }
  @keyframes scaleIn    { from { opacity: 0; transform: scale(0.95) } to { opacity: 1; transform: scale(1) } }
  @keyframes pulse-soft { 0%, 100% { opacity: 1 } 50% { opacity: 0.7 } }
  @keyframes vibrate    { 0%, 100% { transform: translateX(0) } 25% { transform: translateX(-2px) } 75% { transform: translateX(2px) } }
  @keyframes rotate-light { from { transform: rotate(0deg) } to { transform: rotate(6deg) } }
  @keyframes ambientDrift { 0%,100% { transform: translate3d(0,0,0) scale(1); } 50% { transform: translate3d(0,-10px,0) scale(1.02); } }
  @keyframes sheenMove { 0% { transform: translateX(-130%) skewX(-18deg); } 100% { transform: translateX(130%) skewX(-18deg); } }
  @keyframes pageFade { from { opacity: 0; } to { opacity: 1; } }
  .scan-line { position: absolute; left: 0; right: 0; height: 2px; background: linear-gradient(90deg, transparent, rgba(3,29,64,.6), transparent); animation: scanline 2.5s ease-in-out infinite; pointer-events: none; }
  .arrow-pulse { animation: arrowPulse 1.2s ease-in-out infinite; }
  .after-glow  { animation: glow 2.5s ease-in-out infinite; }
  .fade-in2 { animation: fadeUp 0.4s ease 0.1s both; }
  .fade-in3 { animation: fadeUp 0.4s ease 0.2s both; }

  /* ── SCROLLBAR ──────────────────────────────────────────────────── */
  ::-webkit-scrollbar { width: 5px; }
  ::-webkit-scrollbar-track { background: var(--surface2); }
  ::-webkit-scrollbar-thumb { background: var(--accent); border-radius: 4px; }

  /* ── BUTTON HOVER STATES ────────────────────────────────────────── */
  .btn-primary {
    transition: transform 0.24s cubic-bezier(.4,0,.2,1), box-shadow 0.24s cubic-bezier(.4,0,.2,1), filter 0.24s ease, background-position 0.35s ease !important;
    position: relative;
    overflow: hidden;
    letter-spacing: 0.01em;
    border: 1px solid rgba(201,169,110,0.22);
    box-shadow: inset 0 1px 0 rgba(255,255,255,0.16), 0 10px 24px rgba(6,26,53,0.10);
  }
  .btn-primary::before {
    content: '';
    position: absolute;
    inset: 0;
    background: linear-gradient(115deg, transparent 10%, rgba(255,255,255,0.28) 48%, transparent 86%);
    transform: translateX(-120%);
    transition: transform 0.55s ease;
    pointer-events: none;
  }
  .btn-primary:hover:not(:disabled)::before {
    transform: translateX(120%);
  }
  .btn-primary:hover:not(:disabled) {
    transform: translateY(-3px) !important;
    box-shadow: 0 18px 40px rgba(6,26,53,0.16), 0 10px 28px rgba(3,29,64,0.18) !important;
    filter: brightness(1.04) saturate(1.05) !important;
  }
  .btn-primary:active:not(:disabled) {
    transform: translateY(0) !important;
  }

  .btn-ghost {
    transition: transform 0.22s cubic-bezier(.4,0,.2,1), background 0.22s ease, border-color 0.22s ease, color 0.22s ease !important;
    position: relative;
    backdrop-filter: blur(8px);
  }
  .btn-ghost:hover {
    background: rgba(255,255,255,0.18) !important;
    color: var(--gold) !important;
    border-color: rgba(201,169,110,0.40) !important;
    transform: translateY(-2px) !important;
  }

  /* ── CARD HOVER ─────────────────────────────────────────────────── */
  .card-hover {
    transition: transform 0.28s cubic-bezier(.4,0,.2,1), box-shadow 0.28s cubic-bezier(.4,0,.2,1), border-color 0.28s ease !important;
    position: relative;
    backdrop-filter: blur(14px) saturate(125%);
    -webkit-backdrop-filter: blur(14px) saturate(125%);
  }
  .card-hover::before {
    content: '';
    position: absolute;
    inset: 1px;
    border-radius: inherit;
    background: linear-gradient(180deg, rgba(255,255,255,0.46), transparent 22%, rgba(3,29,64,0.03) 100%);
    pointer-events: none;
  }
  .card-hover::after {
    content: '';
    position: absolute;
    left: 18px;
    right: 18px;
    bottom: 0;
    height: 2px;
    border-radius: 2px;
    background: linear-gradient(90deg, transparent, rgba(201,169,110,0.65), rgba(3,29,64,0.55), transparent);
    opacity: 0;
    transition: opacity 0.25s ease;
    pointer-events: none;
  }
  .card-hover:hover::after {
    opacity: 1;
  }
  .card-hover:hover {
    transform: translateY(-8px) !important;
    box-shadow: 0 24px 54px rgba(6,26,53,0.10), 0 10px 24px rgba(3,29,64,0.10) !important;
    border-color: rgba(201,169,110,0.26) !important;
  }

  /* ── NAV HOVER ──────────────────────────────────────────────────── */
  .nav-item { 
    transition: all 0.22s cubic-bezier(.4,0,.2,1) !important; 
    position: relative; 
  }
  .nav-item::after {
    content: '';
    position: absolute;
    bottom: -4px; left: 50%; right: 50%;
    height: 2px;
    background: linear-gradient(90deg, var(--accent), var(--gold));
    border-radius: 2px;
    transition: all 0.28s cubic-bezier(.4,0,.2,1);
  }
  .nav-item:hover::after { 
    left: 10px; 
    right: 10px;
  }
  .nav-item:hover { 
    color: var(--accent) !important; 
    background: rgba(3,29,64,0.07) !important;
    transform: translateY(-1px) !important;
  }

  .tab-pill { 
    transition: all 0.22s cubic-bezier(.4,0,.2,1) !important;
    position: relative;
  }
  .tab-pill:hover { 
    background: rgba(3,29,64,0.08) !important; 
    color: var(--accent) !important;
    transform: translateY(-1px) !important;
  }

  .copy-btn { 
    transition: all 0.2s cubic-bezier(.4,0,.2,1) !important;
  }
  .copy-btn:hover { 
    background: var(--accent-soft) !important;
    transform: translateY(-1px) !important;
  }

  .gen-btn { 
    transition: all 0.25s cubic-bezier(.4,0,.2,1) !important;
    position: relative;
  }
  .gen-btn:hover:not(:disabled) {
    transform: translateY(-3px) scale(1.02) !important;
    box-shadow: 0 10px 30px rgba(3,29,64,0.25), var(--shadow-accent) !important;
    filter: brightness(1.1) !important;
  }

  .hero-btn {
    background-size: 200% auto !important;
    transition: all 0.35s cubic-bezier(.4,0,.2,1) !important;
    position: relative;
    overflow: hidden;
  }
  .hero-btn::after {
    content: '';
    position: absolute;
    inset: 0;
    background: radial-gradient(circle at var(--mouse-x, 50%) var(--mouse-y, 50%), rgba(255,255,255,0.2) 0%, transparent 70%);
    opacity: 0;
    transition: opacity 0.3s ease;
  }
  .hero-btn:hover::after {
    opacity: 1;
  }
  .hero-btn:hover {
    background-position: right center !important;
    transform: translateY(-4px) scale(1.03) !important;
    box-shadow: 0 12px 40px rgba(3,29,64,0.30), 0 8px 30px rgba(3,29,64,0.38) !important;
  }

  .ghost-btn { 
    transition: all 0.25s cubic-bezier(.4,0,.2,1) !important;
    position: relative;
  }
  .ghost-btn:hover {
    background: var(--accent) !important;
    color: white !important;
    transform: translateY(-3px) scale(1.02) !important;
  }

  /* ── HERO GRADIENT ANIMATION ────────────────────────────────────── */
  .topbar-pad {
    background: linear-gradient(180deg, rgba(255,255,255,0.94), rgba(255,255,255,0.84)) !important;
    border: 1px solid rgba(6,26,53,0.08) !important;
    border-radius: 24px;
    backdrop-filter: blur(18px) saturate(160%) !important;
    -webkit-backdrop-filter: blur(18px) saturate(160%) !important;
    box-shadow: 0 16px 34px rgba(6,26,53,0.08);
    max-width: calc(100% - 28px);
    margin: 10px auto 0;
  }

  .section-pad {
    position: relative;
  }

  .hero-gradient {
    max-width: calc(100% - 32px);
    margin: 16px auto 0;
    border-radius: 32px;
    background: linear-gradient(135deg, rgba(255,255,255,0.98) 0%, rgba(248,250,252,0.98) 50%, rgba(243,246,250,0.98) 100%);
    box-shadow: 0 24px 60px rgba(6,26,53,0.08);
    border: 1px solid rgba(3,29,64,0.08);
  }
  .hero-gradient::before {
    content: '';
    position: absolute;
    inset: 0;
    background:
      radial-gradient(circle at top right, rgba(228,190,71,0.18), transparent 0, transparent 26%),
      radial-gradient(circle at bottom left, rgba(3,29,64,0.10), transparent 0, transparent 30%),
      linear-gradient(120deg, rgba(255,255,255,0.72), transparent 28%, transparent 72%, rgba(3,29,64,0.03));
    pointer-events: none;
    animation: ambientDrift 14s ease-in-out infinite;
  }
  .hero-gradient::after {
    content: '';
    position: absolute;
    inset: 16px;
    border: 1px solid rgba(3,29,64,0.08);
    border-radius: 26px;
    pointer-events: none;
    opacity: 0.9;
    box-shadow: inset 0 1px 0 rgba(255,255,255,0.72);
  }

  .three-col > div,
  .pricing-grid > div,
  .article-featured > div,
  .article-stats > div,
  .auth-card {
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
  }

  /* ── GRADIENT TEXT (cross-browser) ─────────────────────────────────── */
  .grad-text {
    background: var(--grad);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    -moz-background-clip: text;
    background-clip: text;
    color: transparent;
  }

  /* ── FEATURE CARD EQUAL HEIGHT ──────────────────────────────────── */
  .feature-grid {
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 18px;
    align-items: stretch;
    width: 100%;
    max-width: 1280px;
    margin: 0 auto;
  }
  .feature-grid > * {
    flex: 0 0 calc(25% - 13.5px);
    min-width: 0;
  }
  @media (max-width: 1200px) {
    .feature-grid > * {
      flex: 0 0 calc(33.333% - 12px);
    }
  }
  @media (max-width: 768px) {
    .feature-grid > * {
      flex: 0 0 calc(50% - 9px) !important;
    }
  }
  @media (max-width: 560px) {
    .feature-grid > * {
      flex: 0 0 100% !important;
    }
  }
  .feature-card {
    display: flex;
    flex-direction: column;
    align-items: center !important;
    justify-content: flex-start;
    text-align: center !important;
    transition: transform 0.28s cubic-bezier(.4,0,.2,1), box-shadow 0.28s cubic-bezier(.4,0,.2,1), border-color 0.28s ease;
    position: relative;
    min-height: 232px;
    padding: 34px 22px 24px !important;
    border-radius: 20px !important;
    background: linear-gradient(180deg, rgba(255,255,255,0.98), rgba(246,249,252,0.95)) !important;
    box-shadow: var(--ring), 0 12px 28px rgba(6,26,53,0.05) !important;
  }
  .feature-card::before {
    content: '';
    position: absolute;
    inset: 0;
    background: linear-gradient(180deg, rgba(255,255,255,0.54) 0%, transparent 22%, rgba(3,29,64,0.04) 100%);
    opacity: 1;
    border-radius: inherit;
    pointer-events: none;
  }
  .feature-card:hover {
    transform: translateY(-8px);
    box-shadow: 0 24px 52px rgba(6,26,53,0.10), 0 10px 22px rgba(3,29,64,0.10) !important;
    border-color: rgba(201,169,110,0.28) !important;
  }
  .feature-copy {
    width: 100%;
    text-align: center !important;
    position: relative;
    z-index: 1;
  }
  .feature-plan {
    position: relative;
    z-index: 3;
    box-shadow: 0 8px 18px rgba(6,26,53,0.08);
    backdrop-filter: blur(8px);
  }
  .feature-card .icon-wrap {
    width: 52px; height: 52px;
    border-radius: 18px;
    background: linear-gradient(180deg, rgba(255,255,255,0.18), rgba(255,255,255,0.06));
    border: 1px solid rgba(255,255,255,0.16);
    display: flex; align-items: center; justify-content: center;
    font-size: 24px;
    margin: 8px auto 16px !important;
    transition: transform 0.28s cubic-bezier(.4,0,.2,1), box-shadow 0.28s cubic-bezier(.4,0,.2,1);
    box-shadow: 0 10px 22px rgba(3,29,64,0.12), inset 0 1px 0 rgba(255,255,255,0.16);
    position: relative;
    z-index: 1;
  }
  .feature-card:hover .icon-wrap {
    transform: translateY(-2px) scale(1.04);
  }

  /* ── DECORATIVE BLOBS ───────────────────────────────────────────── */
  .blob {
    position: absolute;
    border-radius: 50%;
    filter: blur(60px);
    opacity: 0.15;
    pointer-events: none;
  }

  /* ── INPUT FIELDS & FORM ELEMENTS ────────────────────────────────── */
  input[type="text"],
  input[type="email"],
  input[type="password"],
  input[type="file"],
  textarea,
  select {
    transition: all 0.25s cubic-bezier(.4,0,.2,1) !important;
  }
  
  input[type="text"]:focus,
  input[type="email"]:focus,
  input[type="password"]:focus,
  input[type="file"]:focus,
  textarea:focus,
  select:focus {
    transform: translateY(-1px) !important;
    box-shadow: 0 0 0 2px rgba(3,29,64,0.10), 0 4px 16px rgba(3,29,64,0.15) !important;
  }

  /* ── INTERACTIVE LINKS ───────────────────────────────────────────── */
  a {
    transition: all 0.22s cubic-bezier(.4,0,.2,1);
    position: relative;
  }
  a:not([class*="btn"]):hover {
    color: var(--accent);
  }

  /* ── TOGGLE & CHECKBOX ENHANCEMENTS ──────────────────────────────── */
  input[type="checkbox"],
  input[type="radio"] {
    transition: all 0.2s ease;
    cursor: pointer;
  }
  
  input[type="checkbox"]:hover,
  input[type="radio"]:hover {
    transform: scale(1.1);
  }

  /* ── MODE TOGGLE BUTTONS ─────────────────────────────────────────── */
  .mode-toggle {
    transition: all 0.3s cubic-bezier(.4,0,.2,1);
    position: relative;
    overflow: hidden;
  }
  
  .mode-toggle::before {
    content: '';
    position: absolute;
    inset: 0;
    background: rgba(3,29,64,0.10);
    transform: scaleX(0);
    transform-origin: left;
    transition: transform 0.3s cubic-bezier(.4,0,.2,1);
  }
  
  .mode-toggle:hover::before {
    transform: scaleX(1);
  }

  /* ── PILL BUTTONS/TOGGLES ────────────────────────────────────────── */
  .pill-btn {
    transition: all 0.25s cubic-bezier(.4,0,.2,1);
    position: relative;
  }
  
  .pill-btn:hover {
    transform: scale(1.05);
  }

  /* ── MENU TOGGLE ANIMATION ───────────────────────────────────────── */
  .menu-toggle {
    transition: transform 0.3s cubic-bezier(.4,0,.2,1);
  }

  .menu-toggle:hover {
    transform: scale(1.1) rotate(5deg);
  }

  .menu-toggle.active {
    transform: rotate(90deg);
  }

  /* ── ENTRANCE ANIMATIONS ─────────────────────────────────────────── */
  .fade-in {
    animation: pageFade 0.55s cubic-bezier(.22,.61,.36,1) both;
  }
  .fade-in.delay-1 { animation: fadeUp 0.5s ease 0.1s both; }
  .fade-in.delay-2 { animation: fadeUp 0.5s ease 0.2s both; }
  .fade-in.delay-3 { animation: fadeUp 0.5s ease 0.3s both; }

  .slide-in {
    animation: slideIn 0.5s cubic-bezier(.4,0,.2,1) both;
  }
  .slide-in-right {
    animation: slideInRight 0.5s cubic-bezier(.4,0,.2,1) both;
  }

  .scale-in {
    animation: scaleIn 0.4s cubic-bezier(.4,0,.2,1) both;
  }

  /* ── STAGGER ANIMATIONS FOR LISTS ────────────────────────────────── */
  .grid-item {
    animation: fadeUp 0.5s ease forwards;
  }

  .grid-item:nth-child(1) { animation-delay: 0.05s; }
  .grid-item:nth-child(2) { animation-delay: 0.1s; }
  .grid-item:nth-child(3) { animation-delay: 0.15s; }
  .grid-item:nth-child(4) { animation-delay: 0.2s; }
  .grid-item:nth-child(5) { animation-delay: 0.25s; }
  .grid-item:nth-child(6) { animation-delay: 0.3s; }
  .grid-item:nth-child(n+7) { animation-delay: 0.35s; }

  /* ── SUBTLE CONTINUOUS ANIMATIONS ───────────────────────────────── */
  .pulse-subtle {
    animation: pulse-soft 3s ease-in-out infinite;
  }

  .float-animation {
    animation: float 3.5s ease-in-out infinite;
  }

  /* ── ICON ANIMATIONS ─────────────────────────────────────────────── */
  .icon-bounce {
    animation: bounce 1.5s ease-in-out infinite;
  }

  .icon-spin {
    animation: spin 2s linear infinite;
  }

  .icon-spin.slow {
    animation: spin 4s linear infinite;
  }

  /* ── MODAL/POPUP ANIMATIONS ──────────────────────────────────────── */
  .modal-enter {
    animation: modalIn 0.3s cubic-bezier(.4,0,.2,1) both;
  }

  /* ── HOVER LIFT EFFECT ───────────────────────────────────────────── */
  .lift-on-hover {
    transition: all 0.3s cubic-bezier(.4,0,.2,1);
    transform: translateZ(0);
  }

  .lift-on-hover:hover {
    transform: translateY(-8px);
    box-shadow: 0 16px 40px rgba(3,29,64,0.25);
  }

  .marine-hero {
    position: relative;
    display: grid;
    grid-template-columns: minmax(0, 1.2fr) minmax(260px, 0.8fr);
    gap: 20px;
    align-items: end;
    max-width: 1320px;
    margin: 0 auto 22px;
    padding: 34px 34px;
    border-radius: 30px;
    overflow: hidden;
    background: linear-gradient(135deg, rgba(255,255,255,0.99) 0%, rgba(248,250,252,0.98) 52%, rgba(243,246,250,0.98) 100%);
    border: 1px solid rgba(3,29,64,0.08);
    box-shadow: 0 28px 70px rgba(6,26,53,0.10);
  }
  .marine-hero::before {
    content: '';
    position: absolute;
    right: -70px;
    top: 50%;
    transform: translateY(-50%);
    width: 420px;
    height: 420px;
    border-radius: 50%;
    background:
      repeating-radial-gradient(circle, rgba(3,29,64,0.08) 0 2px, transparent 2px 24px),
      radial-gradient(circle at center, rgba(228,190,71,0.16), transparent 64%);
    opacity: 0.36;
    pointer-events: none;
    animation: ambientDrift 16s ease-in-out infinite;
  }
  .marine-hero::after {
    content: '';
    position: absolute;
    left: 22px;
    bottom: 18px;
    width: 96px;
    height: 40px;
    background-image: radial-gradient(rgba(3,29,64,0.18) 1px, transparent 1px);
    background-size: 10px 10px;
    opacity: 0.42;
    pointer-events: none;
  }
  .marine-kicker {
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 10px;
  }
  .marine-chip-row {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
    margin-top: 18px;
  }
  .marine-chip {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 7px 10px;
    border-radius: 999px;
    background: rgba(3,29,64,0.04);
    border: 1px solid rgba(3,29,64,0.08);
    color: var(--text-mid);
    font-size: 12px;
    font-weight: 600;
    box-shadow: inset 0 1px 0 rgba(255,255,255,0.6);
  }
  .marine-hero-side {
    position: relative;
    z-index: 1;
    display: grid;
    gap: 10px;
    justify-self: end;
    width: min(100%, 320px);
  }
  .marine-step {
    display: grid;
    grid-template-columns: 34px 1fr;
    gap: 10px;
    align-items: center;
    padding: 10px 12px;
    border-radius: 14px;
    background: linear-gradient(180deg, rgba(255,255,255,0.86), rgba(243,246,250,0.92));
    border: 1px solid rgba(3,29,64,0.08);
    backdrop-filter: blur(10px);
  }
  .marine-step strong {
    font-family: 'Raleway', sans-serif;
    font-size: 20px;
    line-height: 1;
    color: var(--accent);
  }
  .marine-step span {
    font-size: 12px;
    color: var(--text-mid);
    font-weight: 600;
  }
  .panel-shell {
    background: linear-gradient(180deg, rgba(255,255,255,0.99), rgba(249,250,251,0.98));
    border: 1px solid rgba(31,41,55,0.04);
    border-radius: 30px;
    box-shadow: var(--soft-shadow);
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }
  .panel-head {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 12px;
    padding: 18px 20px;
    border-bottom: 1px solid var(--border);
    background: linear-gradient(180deg, rgba(255,255,255,1), rgba(249,250,251,0.96));
    position: relative;
  }
  .panel-label {
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 0.04em;
    color: var(--text-mid);
  }
  .panel-body {
    padding: 20px;
    display: flex;
    flex-direction: column;
    flex: 1;
    gap: 14px;
  }
  .dock-bar {
    display: grid;
    grid-template-columns: minmax(0, 1.2fr) minmax(0, 1fr) auto;
    gap: 18px;
    align-items: center;
    background: linear-gradient(180deg, rgba(255,255,255,0.99), rgba(249,250,251,0.97));
    border: 1px solid rgba(31,41,55,0.04);
    border-radius: 30px;
    padding: 22px;
    box-shadow: var(--soft-shadow);
  }
  .dock-photo {
    display: flex;
    align-items: center;
    gap: 14px;
    min-width: 0;
  }
  .dock-photo-frame {
    width: 62px;
    height: 62px;
    border-radius: 18px;
    overflow: hidden;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
    background: linear-gradient(135deg, #ECF3F9, #DCEAF4);
    border: 1px solid rgba(201,169,110,0.18);
    box-shadow: inset 0 1px 0 rgba(255,255,255,0.82), 0 8px 18px rgba(6,26,53,0.06);
  }
  .output-badges {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
  }
  .action-stack {
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    gap: 12px;
  }

  .auth-card {
    border-radius: 30px !important;
    box-shadow: var(--soft-shadow) !important;
  }

  .paper-preview {
    max-width: 860px;
    margin: 0 auto;
    background: #FFFFFF;
    border: 1px solid rgba(31,41,55,0.05);
    border-radius: 30px;
    padding: 22px;
    box-shadow: var(--soft-shadow);
  }

  @media (max-width: 1100px) {
    .marine-hero,
    .dock-bar {
      grid-template-columns: 1fr;
    }
    .marine-hero-side {
      justify-self: stretch;
      width: 100%;
    }
    .action-stack {
      align-items: stretch;
    }
  }

  @media (prefers-reduced-motion: reduce) {
    *, *::before, *::after {
      animation: none !important;
      transition-duration: 0.01ms !important;
      scroll-behavior: auto !important;
    }
  }

  /* ── MOBILE ─────────────────────────────────────────────────────── */
  @media (max-width: 768px) {
    .desktop-only { display: none !important; }

    /* ── TYPOGRAPHY ── */
    .hero-title { font-size: clamp(30px, 8vw, 48px) !important; }

    /* ── LAYOUT OVERRIDES ── */
    .hero-pad         { padding: 44px 20px 52px !important; }
    .section-pad      { padding: 48px 20px !important; }
    .topbar-pad       { padding: 0 16px !important; }
    .page-pad         { padding: 20px 16px 80px !important; }
    .output-pad       { padding: 16px 14px 28px !important; }
    .footer-grid      { flex-direction: column !important; gap: 10px !important; text-align: center !important; }

    /* ── GRID COLLAPSES ── */
    .three-col        { grid-template-columns: 1fr !important; }
    .two-col-md       { grid-template-columns: 1fr !important; }
    .pricing-grid     { grid-template-columns: 1fr !important; }
    .feature-grid > * { flex: 0 0 calc(50% - 9px) !important; }
    .kit-grid         { grid-template-columns: 1fr !important; }

    /* ── HERO before/after ── */
    .hero-ba-grid     { grid-template-columns: 1fr !important; gap: 20px !important; }
    .hero-ba-right    { display: none !important; }

    /* ── STATS BAR ── */
    .stats-bar        { grid-template-columns: repeat(2,1fr) !important; }

    /* ── DASHBOARD ── */
    .dash-stats       { grid-template-columns: repeat(2,1fr) !important; }
    .dash-actions     { flex-direction: column !important; }
    /* ── KIT GENERATOR ── */
    .kit-inputs       { grid-template-columns: 1fr !important; }
    .tab-pills        { gap: 6px !important; }
    .tab-pill         { padding: 7px 10px !important; font-size: 12px !important; }
    .output-header    { flex-direction: column !important; align-items: flex-start !important; gap: 10px !important; }
    .template-switcher { display: none !important; }

    /* ── NEGOTIATE ── */
    .neg-grid         { grid-template-columns: 1fr !important; }
    /* ── ARTICLES ── */
    .article-featured { grid-template-columns: 1fr !important; }
    .article-stats    { display: none !important; }

    /* ── AUTH ── */
    .auth-card        { padding: 24px 20px !important; }

    /* ── GENERAL ── */
    .mobile-full      { width: 100% !important; }
    .mobile-hide      { display: none !important; }
    .mobile-small-text { font-size: 12px !important; }

    /* ── FIX OVERFLOW ── */
    body { overflow-x: hidden; }
    * { max-width: 100%; }
    .hero-gradient { overflow: hidden; }
  }

  @media (min-width: 769px) {
    .mobile-only { display: none !important; }
  }

  /* ── BOTTOM NAV PADDING ── */
  @media (max-width: 768px) {
    .page-pad { padding-bottom: 160px !important; }
    .section-pad:last-of-type { padding-bottom: 90px !important; }
  }
  @media (max-width: 380px) {
    .hero-title { font-size: 28px !important; }
    .feature-grid > * { flex: 0 0 100% !important; }
    .pricing-grid { grid-template-columns: 1fr !important; }
    .dash-stats { grid-template-columns: 1fr !important; }
    .stats-bar { grid-template-columns: repeat(2,1fr) !important; }
  }
`;


if (typeof window !== "undefined") {
  if (!window["pdfjs-dist/build/pdf"]) {
    const script = document.createElement("script");
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
    document.head.appendChild(script);
  }
  if (!window.mammoth) {
    const script = document.createElement("script");
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js";
    document.head.appendChild(script);
  }
}

const inp = {
  fontFamily: "'Open Sans', sans-serif",
  color: "var(--text)",
  background: "var(--surface-strong)",
  border: "1px solid var(--border)",
  borderRadius: 16,
  padding: "16px 18px",
  fontSize: 14,
  width: "100%",
  transition: "all 0.2s ease",
};

function BrandLogo({ height = 100, style = {}, alt = "Rezolt" }) {
  return (
    <img
      src="/rezolt-brand.svg"
      alt={alt}
      style={{ height, width: "auto", display: "block", objectFit: "contain", ...style }}
    />
  );
}

// ─── TOP BAR ─────────────────────────────────────────────────────────────────

function TopBar({ page, setPage, user, onSignOut }) {
  const [menuOpen, setMenuOpen] = useState(false);
  const [mobileNav, setMobileNav] = useState(false);
  const creditCount = user?.credits ?? 0;
  const creditLabel = !user
    ? ""
    : user?.plan === "unlimited"
      ? "Unlimited kits"
      : creditCount === 0
        ? "No credits left"
        : user?.plan === "Free"
          ? `${creditCount} free resume left`
          : `${creditCount} kit${creditCount === 1 ? "" : "s"} left`;
  const creditTone = !user ? AC : user?.plan === "unlimited" || creditCount > 0 ? AC : ER;

  return (
    <>
      <div style={{ height: 3, background: "linear-gradient(90deg, #001B48, #031D40, #08284F)" }} />
      <div style={{
        background: "var(--surface)",
        border: "1px solid var(--border)",
        backdropFilter: "blur(14px)",
        WebkitBackdropFilter: "blur(14px)",
        height: 74,
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        padding: "0 clamp(14px,2.8vw,22px)",
        position: "sticky",
        top: 10,
        zIndex: 100,
        boxShadow: "0 18px 40px rgba(6,26,53,0.08)",
      }} className="topbar-pad">

        {/* Logo */}
        <div onClick={() => setPage("landing")} style={{ display: "flex", alignItems: "center", cursor: "pointer", userSelect: "none" }}>
          <BrandLogo height={50} />
        </div>

        {/* Desktop nav */}
        {!user ? (
          <div className="desktop-only" style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <button onClick={() => setPage("landing")} className="nav-item" style={{ background: "none", border: "none", padding: "8px 14px", fontSize: 14, color: "var(--text-muted)", cursor: "pointer", borderRadius: 8, fontFamily: "inherit", fontWeight: 500 }}>Home</button>
            <button onClick={() => setPage("articles")} className="nav-item" style={{ background: "none", border: "none", padding: "8px 14px", fontSize: 14, color: page === "articles" ? AC : "var(--text-muted)", cursor: "pointer", borderRadius: 8, fontFamily: "inherit", fontWeight: page === "articles" ? 700 : 500 }}>Articles</button>
            <button onClick={() => { setPage("landing"); setTimeout(() => { document.getElementById("pricing-section")?.scrollIntoView({ behavior: "smooth" }); }, 120); }} className="nav-item" style={{ background: "none", border: "none", padding: "8px 14px", fontSize: 14, color: "var(--text-muted)", cursor: "pointer", borderRadius: 8, fontFamily: "inherit", fontWeight: 500 }}>Pricing</button>
            <button onClick={() => setPage("auth")} className="btn-primary" style={{ background: "var(--grad)", color: WHITE, border: "none", borderRadius: 10, padding: "10px 22px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", boxShadow: "var(--shadow-accent)" }}>Get Started</button>
          </div>
        ) : (
          <div className="desktop-only" style={{ display: "flex", gap: 4, alignItems: "center" }}>
            {[
              { key: "home", label: "Home", page: "landing" },
              { key: "articles", label: "Articles", page: "articles" },
              { key: "pricing", label: "Pricing", page: "landing", scroll: "pricing-section" },
              { key: "dashboard", label: "Dashboard", page: "dashboard" },
              { key: "create", label: "Create Kit", page: "generate" },
              ...(user?.email?.trim()?.toLowerCase() === ADMIN_EMAIL.trim().toLowerCase() ? [{ key: "admin", label: "Admin", page: "admin" }] : []),
            ].map(item => (
              <button key={item.key} onClick={() => { setPage(item.page); if (item.scroll) setTimeout(() => document.getElementById(item.scroll)?.scrollIntoView({ behavior: "smooth" }), 100); }} className="nav-item" style={{
                background: page === item.page && !item.scroll ? "var(--accent-soft)" : "none",
                border: "none", padding: "8px 14px", fontSize: 14,
                color: page === item.page && !item.scroll ? AC : "var(--text-muted)",
                cursor: "pointer", borderRadius: 8, fontFamily: "inherit",
                fontWeight: page === item.page && !item.scroll ? 700 : 500,
              }}>{item.label}</button>
            ))}
            <div style={{ width: 1, height: 24, background: "var(--border)", margin: "0 6px" }} />
            <div style={{ position: "relative" }}>
              <div onClick={() => setMenuOpen(p => !p)} style={{
                width: 36, height: 36, borderRadius: "50%",
                background: "var(--grad)", cursor: "pointer",
                display: "flex", alignItems: "center", justifyContent: "center",
                fontSize: 14, fontWeight: 800, color: "white",
                boxShadow: "var(--shadow-accent)", userSelect: "none",
                transition: "transform 0.2s ease",
              }}
                onMouseEnter={e => e.currentTarget.style.transform = "scale(1.08)"}
                onMouseLeave={e => e.currentTarget.style.transform = "scale(1)"}
              >
                {user.name?.[0]?.toUpperCase()}
              </div>
              {menuOpen && (
                <>
                  <div onClick={() => setMenuOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 98 }} />
                  <div style={{
                    position: "absolute", top: 46, right: 0,
                    background: "var(--surface)", border: "1px solid var(--border)",
                    borderRadius: 16, padding: "8px", minWidth: 220,
                    boxShadow: "var(--shadow-lg)", zIndex: 99,
                    animation: "scaleIn 0.15s ease",
                  }}>
                    <div style={{ padding: "12px 16px", borderBottom: "1px solid var(--border)", marginBottom: 6 }}>
                      <div style={{ fontSize: 14, fontWeight: 700, color: "var(--text)" }}>{user.name}</div>
                      <div style={{ fontSize: 12, color: "var(--text-muted)", marginTop: 2 }}>{user.email}</div>
                    </div>
                    {[
                      { key: "plans", label: (<><i className="fa-solid fa-dollar-sign" style={{ marginRight: 8 }} />Plans & Billing</>), action: () => { setPage("payment"); setMenuOpen(false); } },
                      { key: "dashboard", label: (<><FiLayout style={{ marginRight: 8 }} />Dashboard</>), action: () => { setPage("dashboard"); setMenuOpen(false); } },
                      { key: "create", label: (<><FiEdit3 style={{ marginRight: 8 }} />Create Kit</>), action: () => { setPage("generate"); setMenuOpen(false); } },
                    ].map(item => (
                      <button key={item.key} onClick={item.action} style={{ width: "100%", textAlign: "left", background: "none", border: "none", padding: "10px 16px", fontSize: 13, color: "var(--text-mid)", cursor: "pointer", borderRadius: 10, fontFamily: "inherit", display: "block", transition: "all 0.15s ease" }}
                        onMouseEnter={e => { e.currentTarget.style.background = "var(--accent-soft)"; e.currentTarget.style.color = AC; }}
                        onMouseLeave={e => { e.currentTarget.style.background = "none"; e.currentTarget.style.color = "var(--text-mid)"; }}>
                        {item.label}
                      </button>
                    ))}
                    <div style={{ borderTop: "1px solid var(--border)", marginTop: 6, paddingTop: 6 }}>
                      <button onClick={() => { onSignOut(); setMenuOpen(false); }} style={{ width: "100%", textAlign: "left", background: "none", border: "none", padding: "10px 16px", fontSize: 13, color: ER, cursor: "pointer", borderRadius: 10, fontFamily: "inherit", transition: "all 0.15s ease" }}
                        onMouseEnter={e => e.currentTarget.style.background = "rgba(239,68,68,0.08)"}
                        onMouseLeave={e => e.currentTarget.style.background = "none"}>
                        <i className="fa-solid fa-right-from-bracket" style={{ marginRight: 8 }} />Sign Out
                      </button>
                    </div>
                  </div>
                </>
              )}
            </div>
          </div>
        )}

        {/* Mobile hamburger */}
        <div className="mobile-only" style={{ display: "flex", alignItems: "center", gap: 8 }}>
          {user && (
            <div style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "var(--surface2)", border: "1px solid var(--border)", borderRadius: 999, padding: "5px 9px" }}>
              <span style={{ width: 7, height: 7, borderRadius: "50%", background: creditTone, flexShrink: 0 }} />
              <span style={{ fontSize: 11, fontWeight: 700, color: creditTone }}>{user.plan === "unlimited" ? "∞ kits" : `${creditCount} kit${creditCount === 1 ? "" : "s"}`}</span>
            </div>
          )}
          <button className="menu-toggle" onClick={() => setMobileNav(p => !p)} style={{ background: "none", border: "1px solid var(--border)", borderRadius: 8, padding: "6px 10px", cursor: "pointer", fontSize: 18, color: "var(--text)" }}>
            <i className={mobileNav ? "fa-solid fa-xmark" : "fa-solid fa-bars"} />
          </button>
        </div>
      </div>

      {/* Mobile nav drawer */}
      {mobileNav && (
        <div className="mobile-only" style={{ background: "var(--surface)", borderBottom: "1px solid var(--border)", padding: "8px 16px 16px", boxShadow: "var(--shadow-md)", zIndex: 99, position: "relative" }}>
          {!user ? (
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              <button onClick={() => { setPage("landing"); setMobileNav(false); }} style={{ background: "none", border: "none", padding: "12px 14px", fontSize: 15, color: "var(--text-mid)", cursor: "pointer", borderRadius: 10, fontFamily: "inherit", textAlign: "left", fontWeight: 600 }}><i className="fa-solid fa-house" style={{ marginRight: 8 }} />Home</button>
              <button onClick={() => { setPage("articles"); setMobileNav(false); }} style={{ background: "none", border: "none", padding: "12px 14px", fontSize: 15, color: "var(--text-mid)", cursor: "pointer", borderRadius: 10, fontFamily: "inherit", textAlign: "left", fontWeight: 600 }}><i className="fa-solid fa-newspaper" style={{ marginRight: 8 }} />Articles</button>
              <button onClick={() => { setPage("landing"); setTimeout(() => document.getElementById("pricing-section")?.scrollIntoView({ behavior: "smooth" }), 120); setMobileNav(false); }} style={{ background: "none", border: "none", padding: "12px 14px", fontSize: 15, color: "var(--text-mid)", cursor: "pointer", borderRadius: 10, fontFamily: "inherit", textAlign: "left", fontWeight: 600 }}><i className="fa-solid fa-dollar-sign" style={{ marginRight: 8 }} />Pricing</button>
              <div style={{ height: 1, background: "var(--border)", margin: "4px 0" }} />
              <button onClick={() => { setPage("auth"); setMobileNav(false); }} style={{ background: "var(--grad)", color: WHITE, border: "none", borderRadius: 12, padding: "14px", fontSize: 15, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>Get Started Free</button>
              <button onClick={() => { setPage("auth"); setMobileNav(false); }} style={{ background: "none", border: "1px solid var(--border)", borderRadius: 12, padding: "13px", fontSize: 15, fontWeight: 600, cursor: "pointer", fontFamily: "inherit", color: "var(--text-muted)" }}>Sign In</button>
            </div>
          ) : (
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "12px 14px", background: "var(--surface2)", borderRadius: 12, marginBottom: 8 }}>
                <div style={{ width: 40, height: 40, borderRadius: "50%", background: "var(--grad)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16, fontWeight: 800, color: "white", flexShrink: 0 }}>{user.name?.[0]?.toUpperCase()}</div>
                <div>
                  <div style={{ fontSize: 14, fontWeight: 700, color: "var(--text)" }}>{user.name}</div>
                  <div style={{ fontSize: 11, color: "var(--text-muted)", fontWeight: 600, textTransform: "capitalize", marginTop: 2 }}>{user.plan || "starter"} plan</div>
                </div>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                {[
                  { key: "dashboard", label: (<><FiLayout style={{ marginRight: 8 }} />Dashboard</>), page: "dashboard" },
                  { key: "create", label: (<><FiEdit3 style={{ marginRight: 8 }} />Create Kit</>), page: "generate" },
                  { key: "home", label: (<><i className="fa-solid fa-house" style={{ marginRight: 8 }} />Home</>), page: "landing" },
                  { key: "pricing", label: (<><i className="fa-solid fa-dollar-sign" style={{ marginRight: 8 }} />Pricing</>), page: "landing", scroll: "pricing-section" },
                  { key: "articles", label: (<><FiBookOpen style={{ marginRight: 8 }} />Articles</>), page: "articles" },
                  { key: "plans", label: (<><i className="fa-solid fa-dollar-sign" style={{ marginRight: 8 }} />Plans & Billing</>), page: "payment" },
                  ...(user?.email?.trim()?.toLowerCase() === ADMIN_EMAIL.trim().toLowerCase() ? [{ key: "admin", label: (<><i className="fa-solid fa-sliders" style={{ marginRight: 8 }} />Admin</>), page: "admin" }] : []),
                ].map(item => (
                  <button key={item.key} onClick={() => { setPage(item.page); if (item.scroll) setTimeout(() => document.getElementById(item.scroll)?.scrollIntoView({ behavior: "smooth" }), 100); setMobileNav(false); }} style={{ background: "none", border: "none", padding: "12px 14px", fontSize: 14, color: "var(--text-mid)", cursor: "pointer", borderRadius: 10, fontFamily: "inherit", textAlign: "left", fontWeight: 500 }}>{item.label}</button>
                ))}
                <div style={{ height: 1, background: "var(--border)", margin: "4px 0" }} />
                <button onClick={() => { onSignOut(); setMobileNav(false); }} style={{ background: "none", border: "none", padding: "12px 14px", fontSize: 14, color: ER, cursor: "pointer", borderRadius: 10, fontFamily: "inherit", textAlign: "left" }}><i className="fa-solid fa-right-from-bracket" style={{ marginRight: 8 }} />Sign Out</button>
              </div>
            </div>
          )}
        </div>
      )}
    </>
  );
}

// ─── LANDING PAGE ─────────────────────────────────────────────────────────────

function LandingPage({ setPage, user, selectedTemplate, setSelectedTemplate, setSelectedArticle }) {
  const [sampleTab, setSampleTab] = useState("resume");
  const [sampleModalOpen, setSampleModalOpen] = useState(false);

  const sampleModalContent = {
    resume: {
      title: "Sample resume kit",
      subtitle: "HRBP Manager · pharma analytics hiring",
      bullets: [
        "Role-matched summary with stronger positioning",
        "Quantified achievements recruiters can scan fast",
        "Cleaner ATS-friendly structure and sharper keywords",
      ],
      excerpt: [
        "Lead Talent Acquisition Specialist with 4.5 years driving hiring across pharma analytics and commercial roles.",
        "Reduced time-to-hire by 28%, built Power BI workforce dashboards, and supported campus partnerships that led to 12 hires in FY2024.",
      ],
    },
    cover: {
      title: "Sample cover letter",
      subtitle: "Personal, specific, and recruiter-ready",
      bullets: [
        "Connects the candidate's background to the role in the first paragraph",
        "Uses concrete proof points instead of generic claims",
        "Ends with a warm, professional call to action",
      ],
      excerpt: [
        "With 4.5 years building talent pipelines for pharma analytics functions, I bring the mix of stakeholder management and execution speed this HRBP Manager role needs.",
        "At my current organisation, I reduced time-to-hire by 28% while also building reporting dashboards used across 150+ headcount decisions.",
      ],
    },
    referral: {
      title: "Sample referral outreach",
      subtitle: "Warm note plus a direct follow-up DM",
      bullets: [
        "Short enough for LinkedIn but still specific to the role",
        "Feels human and respectful, not templated or spammy",
        "Makes it easy for a contact to reply or refer",
      ],
      excerpt: [
        "Hi Anjali — I came across the HRBP Manager opening at [Company] and it aligns closely with my 4.5 years in pharma TA and HR analytics.",
        "If you're open to it, I’d really value any guidance on the process or a referral. Happy to share my resume directly.",
      ],
    },
  };

  const templates = [
    {
      id: "creative", name: "Creative", desc: "Sidebar · Blue accents", best: "Analytics, HR, Tech",
      header: "linear-gradient(135deg,#001B48,#02457A)", headerTxt: "white",
      preview: (
        <div style={{ display: "flex", gap: 6, minHeight: 100, padding: 12, background: "#F0F7FF" }}>
          <div style={{ width: 42, background: "linear-gradient(180deg,#001B48,#02457A)", borderRadius: 6, padding: "7px 5px", flexShrink: 0 }}>
            <div style={{ width: 18, height: 18, borderRadius: "50%", background: "#02457A", margin: "0 auto 5px" }} />
            <div style={{ height: 2, background: "rgba(255,255,255,.3)", borderRadius: 1, marginBottom: 3 }} />
            <div style={{ height: 2, background: "rgba(255,255,255,.15)", borderRadius: 1, width: "80%", marginBottom: 8 }} />
            {[100, 85, 70].map((w, i) => <div key={i} style={{ height: 4, background: "rgba(3,29,64,.35)", borderRadius: 2, width: w + "%", marginBottom: 3 }} />)}
          </div>
          <div style={{ flex: 1 }}>
            <div style={{ height: 6, background: "#001B48", borderRadius: 2, marginBottom: 3, width: "70%" }} />
            <div style={{ height: 3, background: "#02457A", borderRadius: 2, marginBottom: 7, width: "40%" }} />
            {[100, 85, 70].map((w, i) => <div key={i} style={{ height: 2, background: "#CBD5E1", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
            <div style={{ height: 4, background: "#001B48", borderRadius: 2, marginBottom: 4, width: "50%", marginTop: 7 }} />
            {[100, 80].map((w, i) => <div key={i} style={{ height: 2, background: "#CBD5E1", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
          </div>
        </div>
      )
    },
    {
      id: "modern", name: "Modern", desc: "Clean · Minimal", best: "Any industry",
      header: "#F8FAFC", headerTxt: "#374151",
      preview: (
        <div style={{ minHeight: 100, padding: 12, background: "white" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 8 }}>
            <div style={{ width: 22, height: 22, borderRadius: "50%", background: "#F1F5F9", flexShrink: 0 }} />
            <div><div style={{ height: 5, background: "#111827", borderRadius: 2, width: 65, marginBottom: 3 }} /><div style={{ height: 3, background: "#CBD5E1", borderRadius: 1, width: 42 }} /></div>
          </div>
          <div style={{ height: 1, background: "#E2E8F0", marginBottom: 7 }} />
          {[100, 85, 68].map((w, i) => <div key={i} style={{ height: 2, background: "#E2E8F0", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
          <div style={{ height: 4, background: "#111827", borderRadius: 2, marginBottom: 4, width: "44%", marginTop: 7 }} />
          {[100, 80].map((w, i) => <div key={i} style={{ height: 2, background: "#E2E8F0", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
        </div>
      )
    },
    {
      id: "bold", name: "Bold", desc: "Header-first · ATS-heavy", best: "Corporate, Finance",
      header: "#F8FAFC", headerTxt: "#374151",
      preview: (
        <div style={{ minHeight: 100, padding: 12, background: "white" }}>
          <div style={{ height: 22, background: "#111827", borderRadius: 4, marginBottom: 7, display: "flex", alignItems: "center", padding: "0 7px" }}>
            <div style={{ height: 4, background: "rgba(255,255,255,.85)", borderRadius: 2, width: "55%" }} />
          </div>
          <div style={{ display: "flex", gap: 5, marginBottom: 7 }}>
            {[30, 22].map((w, i) => <div key={i} style={{ height: 3, background: i === 0 ? "#111827" : "#CBD5E1", borderRadius: 1, width: w + "%" }} />)}
          </div>
          {[100, 88, 72].map((w, i) => <div key={i} style={{ height: 2, background: "#E2E8F0", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
          <div style={{ height: 4, background: "#111827", borderRadius: 2, marginBottom: 4, width: "48%", marginTop: 7 }} />
          {[100, 83].map((w, i) => <div key={i} style={{ height: 2, background: "#E2E8F0", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
        </div>
      )
    },
    {
      id: "classic", name: "Classic", desc: "Centered header · Traditional", best: "Government, Law, Banking",
      header: "#F8FAFC", headerTxt: "#374151",
      preview: (
        <div style={{ minHeight: 100, padding: 12, background: "white" }}>
          <div style={{ textAlign: "center", marginBottom: 7 }}>
            <div style={{ height: 6, background: "#1E3A5F", borderRadius: 2, width: "65%", margin: "0 auto 3px" }} />
            <div style={{ height: 3, background: "#5A6D88", borderRadius: 2, width: "40%", margin: "0 auto 4px" }} />
            <div style={{ height: 1, background: "#CBD5E1", width: "80%", margin: "0 auto" }} />
          </div>
          <div style={{ height: 3, background: "#1E3A5F", borderRadius: 2, marginBottom: 5, width: "35%", marginTop: 8 }} />
          {[100, 88, 75].map((w, i) => <div key={i} style={{ height: 2, background: "#E2E8F0", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
          <div style={{ height: 3, background: "#1E3A5F", borderRadius: 2, marginBottom: 5, width: "35%", marginTop: 7 }} />
          {[100, 82].map((w, i) => <div key={i} style={{ height: 2, background: "#E2E8F0", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
        </div>
      )
    },
    {
      id: "elegant", name: "Elegant", desc: "Serif · Sophisticated", best: "Consulting, PR, Management",
      header: "linear-gradient(135deg,#2C1810,#78350F)", headerTxt: "#fef3c7",
      preview: (
        <div style={{ display: "flex", gap: 0, minHeight: 100, background: "white", padding: 12 }}>
          <div style={{ width: 4, background: "linear-gradient(180deg,#C9A96E,#78350F)", borderRadius: "3px 0 0 3px", flexShrink: 0 }} />
          <div style={{ flex: 1, paddingLeft: 10 }}>
            <div style={{ height: 7, background: "#2C1810", borderRadius: 2, marginBottom: 3, width: "75%" }} />
            <div style={{ height: 3, background: "#C9A96E", borderRadius: 2, marginBottom: 8, width: "45%" }} />
            <div style={{ height: 1, background: "#C9A96E", opacity: 0.4, marginBottom: 7 }} />
            <div style={{ height: 3, background: "#78350F", borderRadius: 2, marginBottom: 5, width: "42%" }} />
            {[100, 86, 72].map((w, i) => <div key={i} style={{ height: 2, background: "#D1C5B8", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
          </div>
        </div>
      )
    },
    {
      id: "compact", name: "Compact", desc: "Two-column · Dense", best: "Senior roles, info-heavy",
      header: "#F8FAFC", headerTxt: "#374151",
      preview: (
        <div style={{ minHeight: 100, padding: 12, background: "white", display: "flex", gap: 8 }}>
          <div style={{ flex: 1 }}>
            <div style={{ height: 5, background: "#111827", borderRadius: 2, marginBottom: 5 }} />
            {[100, 90, 80].map((w, i) => <div key={i} style={{ height: 2, background: "#E2E8F0", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
            <div style={{ height: 3, background: "#374151", borderRadius: 2, marginBottom: 4, width: "60%", marginTop: 7 }} />
            {[100, 85].map((w, i) => <div key={i} style={{ height: 2, background: "#E2E8F0", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
          </div>
          <div style={{ width: 1, background: "#E2E8F0" }} />
          <div style={{ width: 75, flexShrink: 0 }}>
            <div style={{ height: 3, background: "#374151", borderRadius: 2, marginBottom: 5, width: "70%" }} />
            <div style={{ display: "flex", flexWrap: "wrap", gap: 2, marginBottom: 7 }}>
              {[48, 44, 100].map((w, i) => <div key={i} style={{ height: 9, background: "#F1F5F9", borderRadius: 3, width: w + "%" }} />)}
            </div>
            <div style={{ height: 3, background: "#374151", borderRadius: 2, marginBottom: 4, width: "80%" }} />
            {[100, 90].map((w, i) => <div key={i} style={{ height: 2, background: "#E2E8F0", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
          </div>
        </div>
      )
    },
    {
      id: "minimal", name: "Minimal", desc: "Ultra clean · Whitespace", best: "Design, Creative industries",
      header: "#FAFAFA", headerTxt: "#374151",
      preview: (
        <div style={{ minHeight: 100, padding: 16, background: "white" }}>
          <div style={{ height: 7, background: "#111827", borderRadius: 1, marginBottom: 4, width: "55%" }} />
          <div style={{ height: 3, background: "#D1D5DB", borderRadius: 1, marginBottom: 20, width: "35%" }} />
          {[100, 88, 72].map((w, i) => <div key={i} style={{ height: 2, background: "#F3F4F6", borderRadius: 1, marginBottom: 3, width: w + "%" }} />)}
          <div style={{ height: 2, background: "#D1D5DB", borderRadius: 1, marginBottom: 5, width: "32%", marginTop: 12 }} />
          {[100, 80].map((w, i) => <div key={i} style={{ height: 2, background: "#F3F4F6", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
        </div>
      )
    },
    {
      id: "tech", name: "Tech", desc: "Dark · Developer style", best: "Engineering, Data Science",
      header: "#0d1117", headerTxt: "#e6edf3",
      preview: (
        <div style={{ minHeight: 100, padding: 12, background: "#0d1117" }}>
          <div style={{ height: 5, background: "#C9D6E4", borderRadius: 2, marginBottom: 4, width: "65%" }} />
          <div style={{ height: 3, background: "#30363d", borderRadius: 1, marginBottom: 2, width: "50%" }} />
          <div style={{ height: 1, background: "#21262d", marginBottom: 7 }} />
          <div style={{ display: "flex", gap: 3, marginBottom: 7, flexWrap: "wrap" }}>
            {[28, 22, 32].map((w, i) => <div key={i} style={{ height: 10, background: "#161b22", border: "1px solid #30363d", borderRadius: 4, width: w + "%" }} />)}
          </div>
          <div style={{ height: 3, background: "#C9D6E4", borderRadius: 2, marginBottom: 4, width: "40%", opacity: 0.6 }} />
          {[100, 88, 72].map((w, i) => <div key={i} style={{ height: 2, background: "#21262d", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
        </div>
      )
    },
    {
      id: "warm", name: "Warm", desc: "Earthy · Friendly", best: "Marketing, Sales, NGO",
      header: "linear-gradient(135deg,#92400e,#b45309)", headerTxt: "#fef3c7",
      preview: (
        <div style={{ minHeight: 100, padding: 12, background: "#fffbf5" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
            <div style={{ width: 22, height: 22, borderRadius: "50%", background: "linear-gradient(135deg,#f59e0b,#d97706)", flexShrink: 0 }} />
            <div>
              <div style={{ height: 5, background: "#78350f", borderRadius: 2, width: 65, marginBottom: 3 }} />
              <div style={{ height: 3, background: "#d97706", borderRadius: 1, width: 42 }} />
            </div>
          </div>
          <div style={{ height: 1, background: "#fde68a", marginBottom: 7 }} />
          {[100, 85, 70].map((w, i) => <div key={i} style={{ height: 2, background: "#fef3c7", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
          <div style={{ height: 3, background: "#92400e", borderRadius: 2, marginBottom: 5, width: "44%", marginTop: 7 }} />
          {[100, 78].map((w, i) => <div key={i} style={{ height: 2, background: "#fef3c7", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
        </div>
      )
    },
  ];

  const selTmpl = templates.find(t => t.id === selectedTemplate) || templates[0];

  return (
    <div className="fade-in">
      {/* ── HERO ── */}
      <div className="hero-gradient" style={{ padding: "48px clamp(16px,5vw,80px) 48px", position: "relative", overflow: "hidden" }}>
        <div style={{ position: "absolute", top: -80, right: -80, width: 350, height: 350, background: "rgba(3,29,64,.12)", borderRadius: "50%", filter: "blur(70px)" }} />
        <div style={{ position: "absolute", bottom: -60, left: -40, width: 250, height: 250, background: "rgba(3,29,64,.08)", borderRadius: "50%", filter: "blur(50px)" }} />

        <div className="two-col-md hero-ba-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 48, alignItems: "center", position: "relative", zIndex: 1 }}>
          {/* Left */}
          <div>
            <div style={{ display: "inline-flex", alignItems: "center", gap: 7, background: "var(--accent-soft)", border: "1px solid rgba(3,29,64,.08)", borderRadius: 30, padding: "6px 14px", marginBottom: 24, backdropFilter: "blur(8px)" }}>
              <div style={{ width: 7, height: 7, borderRadius: "50%", background: PB, position: "relative" }}>
                <div style={{ position: "absolute", inset: 0, borderRadius: "50%", background: PB, animation: "pulse-ring 1.5s ease infinite" }} />
              </div>
              <span style={{ fontSize: 12, fontWeight: 700, color: AC }}>Resume help built for modern Indian job searches</span>
            </div>
            <h1 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(34px,4vw,56px)", fontWeight: 400, lineHeight: 1.06, letterSpacing: "-.03em", color: "var(--text)", marginBottom: 18 }}>
              Build a stronger resume<br />for the <em style={{ color: PB }}>role you want.</em>
            </h1>
            <p style={{ fontSize: 15, color: "var(--text-muted)", lineHeight: 1.8, marginBottom: 24, maxWidth: 470 }}>
              Rezolt helps you turn your current resume into a cleaner, more tailored career kit with ATS-friendly wording, stronger positioning, and ready-to-use job search outputs.
            </p>
            <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 18 }}>
              {user ? (
                <>
                  <button className="hero-btn btn-primary" onClick={() => setPage("dashboard")} style={{ background: "linear-gradient(135deg,#E4BE47,#F2D46C,#E4BE47)", backgroundSize: "200% auto", color: N1, border: "none", borderRadius: 12, padding: "14px 28px", fontSize: 14, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" }}>Go to Dashboard</button>
                  <button className="ghost-btn" onClick={() => setPage("generate")} style={{ background: "rgba(255,255,255,.86)", color: "var(--text)", border: "1px solid rgba(3,29,64,.10)", borderRadius: 12, padding: "13px 22px", fontSize: 14, cursor: "pointer", fontFamily: "inherit", fontWeight: 600 }}>Create New Kit</button>
                </>
              ) : (
                <>
                  <button className="hero-btn btn-primary" onClick={() => setPage("auth")} style={{ background: "linear-gradient(135deg,#E4BE47,#F2D46C,#E4BE47)", backgroundSize: "200% auto", color: N1, border: "none", borderRadius: 12, padding: "14px 28px", fontSize: 14, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" }}>Get Started Free</button>
                  <button className="ghost-btn" onClick={() => setPage("auth")} style={{ background: "rgba(255,255,255,.86)", color: "var(--text)", border: "1px solid rgba(3,29,64,.10)", borderRadius: 12, padding: "13px 22px", fontSize: 14, cursor: "pointer", fontFamily: "inherit", fontWeight: 600 }}>Sign In</button>
                </>
              )}
            </div>
            <div style={{ display: "grid", gap: 10, marginBottom: 18, maxWidth: 520 }}>
              {[
                "Tailored resume and cover letter for the exact role",
                "Referral, outreach, and interview support in one place",
                "ATS-aware formatting built for Indian hiring flows",
              ].map(item => (
                <div key={item} style={{ display: "flex", alignItems: "center", gap: 10, fontSize: 13, color: MID }}>
                  <span style={{ width: 22, height: 22, borderRadius: "50%", background: "var(--accent-soft)", color: AC, display: "inline-flex", alignItems: "center", justifyContent: "center", fontWeight: 700, flexShrink: 0 }}>✓</span>
                  <span>{item}</span>
                </div>
              ))}
            </div>
            <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
              {[["4.8/5", " average rating from early users"], ["60s", "average turnaround"], ["3 free", "starter career kits"]].map(([v, l]) => (
                <div key={v} style={{ display: "flex", alignItems: "center", gap: 8, padding: "10px 12px", background: "rgba(255,255,255,.82)", border: "1px solid rgba(3,29,64,.06)", borderRadius: 999 }}>
                  <span style={{ fontSize: 18, fontWeight: 700, color: PB }}>{v}</span>
                  <span style={{ fontSize: 11, color: "var(--text-muted)" }}>{l}</span>
                </div>
              ))}
            </div>
          </div>

          {/* Right: Before/After — hidden on mobile */}
          <div className="hero-ba-right" style={{ position: "relative" }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr auto 1fr", gap: 8, alignItems: "start" }}>
              {/* Before */}
              <div style={{ background: "rgba(255,255,255,.92)", border: "1px solid rgba(3,29,64,.08)", borderRadius: 14, overflow: "hidden", animation: "fadeUp 0.5s ease both", boxShadow: "0 10px 24px rgba(3,29,64,.08)" }}>
                <div style={{ padding: "8px 12px", borderBottom: "1px solid rgba(3,29,64,.08)", display: "flex", justifyContent: "space-between", alignItems: "center", background: "#F8FAFC" }}>
                  <span style={{ fontSize: 10, fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: ".08em" }}>Before</span>
                  <span style={{ fontSize: 9, color: "var(--text-faint)" }}>resume_old.docx</span>
                </div>
                <div style={{ padding: 12, fontSize: 9, color: "#6B7280", lineHeight: 1.7, fontFamily: "monospace" }}>
                  <div style={{ marginBottom: 5, color: "#475569", fontWeight: 700 }}>PRIYA SHARMA</div>
                  <div style={{ marginBottom: 6 }}>HR professional with exp in recruitment. Worked at pharma company doing hiring for 4 years. Good at Excel.</div>
                  <div style={{ color: "#94A3B8", marginBottom: 3 }}>EXPERIENCE</div>
                  <div>TA Specialist - Current Company</div>
                  <div>- Did hiring</div>
                  <div>- Made reports</div>
                  <div style={{ color: "#94A3B8", marginTop: 5, marginBottom: 3 }}>SKILLS</div>
                  <div>Excel, Naukri, LinkedIn</div>
                </div>
              </div>

              {/* Arrow */}
              <div style={{ display: "flex", flexDirection: "column", alignItems: "center", paddingTop: 40, gap: 5 }}>
                <div style={{ background: "linear-gradient(135deg,#031D40,#08284F)", borderRadius: "50%", width: 30, height: 30, display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 0 16px rgba(3,29,64,.45)", animation: "arrowPulse 1.2s ease-in-out infinite" }}>
                  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="9,18 15,12 9,6" /></svg>
                </div>
                <div style={{ fontSize: 8, color: "var(--text-muted)", textAlign: "center", fontWeight: 700, lineHeight: 1.4 }}>Career<br />Partner</div>
              </div>

              {/* After */}
              <div style={{ background: "white", borderRadius: 14, overflow: "hidden", animation: "fadeUp 0.5s ease 0.2s both", boxShadow: "0 0 24px rgba(3,29,64,.28)" }}>
                <div style={{ padding: "8px 12px", background: "linear-gradient(135deg,#001B48,#02457A)", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontSize: 10, fontWeight: 700, color: "rgba(255,255,255,.7)", textTransform: "uppercase", letterSpacing: ".08em" }}>After</span>
                  <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                    <div style={{ width: 5, height: 5, borderRadius: "50%", background: "#22C55E" }} />
                    <span style={{ fontSize: 9, color: "rgba(255,255,255,.5)" }}>ATS 94/100</span>
                  </div>
                </div>
                <div style={{ display: "flex", minHeight: 150 }}>
                  <div style={{ width: 48, flexShrink: 0, background: "linear-gradient(180deg,#001B48,#02457A)", padding: "8px 6px" }}>
                    <div style={{ width: 20, height: 20, borderRadius: "50%", background: "#02457A", margin: "0 auto 6px" }} />
                    <div style={{ height: 2, background: "rgba(255,255,255,.35)", borderRadius: 1, marginBottom: 3 }} />
                    <div style={{ height: 2, background: "rgba(255,255,255,.18)", borderRadius: 1, width: "80%", marginBottom: 9 }} />
                    {[100, 85, 70].map((w, i) => <div key={i} style={{ height: 4, background: "rgba(3,29,64,.35)", borderRadius: 2, width: w + "%", marginBottom: 3 }} />)}
                  </div>
                  <div style={{ flex: 1, padding: 10 }}>
                    <div style={{ height: 6, background: "#001B48", borderRadius: 2, marginBottom: 3, width: "75%" }} />
                    <div style={{ height: 3, background: "#02457A", borderRadius: 2, marginBottom: 8, width: "45%" }} />
                    {[100, 90, 75].map((w, i) => <div key={i} style={{ height: 2, background: "#E2E8F0", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
                    <div style={{ height: 3, background: "#001B48", borderRadius: 2, marginBottom: 5, width: "50%", marginTop: 8 }} />
                    {[100, 85, 70].map((w, i) => <div key={i} style={{ height: 2, background: "#E2E8F0", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
                    <div style={{ height: 3, background: "#001B48", borderRadius: 2, marginBottom: 5, width: "55%", marginTop: 8 }} />
                    {[100, 80].map((w, i) => <div key={i} style={{ height: 2, background: "#E2E8F0", borderRadius: 1, marginBottom: 2, width: w + "%" }} />)}
                  </div>
                </div>
                <div style={{ padding: "7px 10px", background: "#F4F7FB", borderTop: "1px solid rgba(3,29,64,.1)", display: "flex", alignItems: "center", gap: 5 }}>
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#22C55E" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12" /></svg>
                  <span style={{ fontSize: 9, color: "#02457A", fontWeight: 600 }}>ATS Optimized ✓</span>
                </div>
              </div>
            </div>
            <div style={{ textAlign: "center", marginTop: 12 }}>
              <span style={{ background: "rgba(255,255,255,.92)", border: "1px solid rgba(3,29,64,.08)", borderRadius: 20, padding: "4px 14px", fontSize: 10, color: "var(--text-muted)", boxShadow: "0 8px 20px rgba(3,29,64,.06)" }}>Same experience. Clearer positioning. Better first impression.</span>
            </div>
          </div>
        </div>
      </div>

      <div className="section-pad" style={{ padding: "18px clamp(16px,5vw,80px) 42px", background: "var(--bg)" }}>
        <div className="three-col" style={{ maxWidth: 1180, margin: "0 auto", display: "grid", gridTemplateColumns: "repeat(4, minmax(0, 1fr))", gap: 14 }}>
          {[
            { value: "ATS-ready", label: "resume structure and wording" },
            { value: "7 outputs", label: "for one complete application flow" },
            { value: "India-first", label: "built around local hiring reality" },
            { value: "3 free kits", label: "to help you get started" },
          ].map(item => (
            <div key={item.value} style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 18, padding: "16px 18px", boxShadow: "var(--soft-shadow)" }}>
              <div style={{ fontSize: 15, fontWeight: 700, color: DARK, marginBottom: 4 }}>{item.value}</div>
              <div style={{ fontSize: 12, color: MUTED, lineHeight: 1.6 }}>{item.label}</div>
            </div>
          ))}
        </div>
      </div>

      {/* ── HOW IT WORKS ── */}
      <div className="section-pad" style={{ padding: "80px clamp(20px,5vw,80px)", background: "var(--bg)" }}>
        <div style={{ textAlign: "center", marginBottom: 52 }}>
          <div style={{ fontSize: 12, fontWeight: 700, letterSpacing: ".16em", textTransform: "uppercase", color: AC, marginBottom: 14 }}>How it works</div>
          <h2 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(28px,4vw,44px)", fontWeight: 400, letterSpacing: "-.02em", background: "var(--grad)", WebkitBackgroundClip: "text", MozBackgroundClip: "text", WebkitTextFillColor: "transparent", backgroundClip: "text", color: "transparent" }}>A simple path to a sharper application</h2>
          <p style={{ fontSize: 15, color: "var(--text-muted)", maxWidth: 620, margin: "12px auto 0", lineHeight: 1.75 }}>Built to feel easy, clear, and recruiter-aware from your first draft to your final send.</p>
        </div>
        <div className="three-col" style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 24 }}>
          {[
            { num: "01", kicker: "Start with your base", title: "Upload or paste your resume", desc: "Bring the version you already have. Rezolt reshapes it without losing your core story.", foot: "No rework from scratch" },
            { num: "02", kicker: "Target the role", title: "Add the role you want", desc: "Share the job description so every output feels relevant, tailored, and more intentional.", foot: "Matched to the target role" },
            { num: "03", kicker: "Review and send", title: "Review your full kit", desc: "Get resume, cover letter, outreach, and prep materials in one place — ready for a human final pass.", foot: "Built in under a minute" },
          ].map((s, i) => (
            <div key={i} className="card-hover grid-item" style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 24, padding: "28px 24px", position: "relative", overflow: "hidden", boxShadow: "var(--soft-shadow)" }}>

              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, marginBottom: 18 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10, minWidth: 0 }}>
                  <span style={{ width: 28, height: 3, borderRadius: 999, background: "var(--grad)", display: "inline-block", flexShrink: 0 }} />
                  <span style={{ fontSize: 10, fontWeight: 700, color: AC, letterSpacing: ".08em", textTransform: "uppercase" }}>{s.kicker}</span>
                </div>
                <span style={{ fontSize: 11, fontWeight: 700, color: AC, background: "var(--accent-soft)", border: "1px solid var(--border)", padding: "4px 10px", borderRadius: 999 }}>{s.num}</span>
              </div>
              <h3 style={{ fontSize: 17, fontWeight: 700, color: "var(--text)", marginBottom: 10 }}>{s.title}</h3>
              <p style={{ fontSize: 14, color: "var(--text-muted)", lineHeight: 1.75, marginBottom: 16 }}>{s.desc}</p>
              <div style={{ fontSize: 12, color: MID, background: "var(--surface2)", border: "1px solid var(--border)", borderRadius: 12, padding: "8px 10px" }}>{s.foot}</div>
            </div>
          ))}
        </div>
      </div>

      {/* ── SAMPLE OUTPUT ── */}
      <div style={{ background: "var(--grad-subtle)", borderTop: "1px solid var(--border)", borderBottom: "1px solid var(--border)" }}>
        <div className="section-pad" style={{ padding: "80px clamp(20px,5vw,80px)" }}>
          <div style={{ textAlign: "center", marginBottom: 44 }}>
            <div style={{ fontSize: 12, fontWeight: 700, letterSpacing: ".16em", textTransform: "uppercase", color: AC, marginBottom: 14 }}>See it in action</div>
            <h2 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(28px,4vw,44px)", fontWeight: 400, letterSpacing: "-.02em", background: "var(--grad)", WebkitBackgroundClip: "text", MozBackgroundClip: "text", WebkitTextFillColor: "transparent", backgroundClip: "text", color: "transparent" }}>See what a stronger application can look like</h2>
            <p style={{ fontSize: 14, color: "var(--text-muted)", maxWidth: 460, margin: "10px auto 0", lineHeight: 1.7 }}>Real sample outputs based on a genuine resume and role brief — designed to feel clearer, sharper, and more recruiter-ready.</p>
          </div>

          {/* Sample tabs */}
          <div style={{ display: "flex", gap: 8, justifyContent: "center", marginBottom: 14, flexWrap: "wrap" }}>
            {[["resume", "Resume"], ["cover", "Cover Letter"], ["referral", "Referral DM"]].map(([id, label]) => (
              <button key={id} className="tab-pill pill-btn" onClick={() => setSampleTab(id)} style={{
                padding: "9px 20px", borderRadius: 30, fontSize: 13, fontWeight: sampleTab === id ? 700 : 600, cursor: "pointer", fontFamily: "inherit", border: "none", transition: "all .18s ease",
                background: sampleTab === id ? N1 : "var(--surface)",
                color: sampleTab === id ? "white" : "var(--text-muted)",
                boxShadow: sampleTab === id ? "var(--shadow-md)" : "none",
              }}>{label}</button>
            ))}
          </div>

          <div style={{ display: "flex", justifyContent: "center", marginBottom: 24 }}>
            <button onClick={() => setSampleModalOpen(true)} style={{ background: WHITE, color: O, border: `1px solid ${BORDER}`, borderRadius: 999, padding: "10px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", boxShadow: "var(--soft-shadow)" }}>
              View sample kit in detail
            </button>
          </div>

          <div style={{ maxWidth: 700, margin: "0 auto" }}>
            {/* Resume Sample */}
            {sampleTab === "resume" && (
              <div style={{ background: "var(--surface)", borderRadius: 20, overflow: "hidden", boxShadow: "var(--shadow-lg)", border: "1px solid var(--border)", animation: "fadeUp .3s ease" }}>
                <div style={{ background: "linear-gradient(135deg,#001B48,#02457A)", padding: "14px 22px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <div style={{ width: 7, height: 7, borderRadius: "50%", background: "#22C55E" }} />
                    <span style={{ color: "rgba(255,255,255,.65)", fontSize: 12 }}>resume_rewrite.pdf</span>
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <span style={{ background: "rgba(34,197,94,.2)", color: "#86efac", fontSize: 10, fontWeight: 700, padding: "2px 10px", borderRadius: 10 }}>ATS 94/100</span>
                    <span style={{ background: "rgba(255,255,255,.1)", color: "#C9D6E4", fontSize: 10, fontWeight: 600, padding: "2px 10px", borderRadius: 10 }}>Tailored</span>
                  </div>
                </div>
                <div style={{ padding: "24px 28px" }}>
                  <div style={{ display: "flex", gap: 20 }}>
                    <div style={{ width: 155, flexShrink: 0, background: "linear-gradient(180deg,#001B48,#02457A)", borderRadius: 10, padding: "16px 12px", color: "white" }}>
                      <div style={{ width: 40, height: 40, borderRadius: "50%", background: "#02457A", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16, fontWeight: 700, marginBottom: 10 }}>P</div>
                      <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 2 }}>Priya Sharma</div>
                      <div style={{ fontSize: 9, color: "#C9D6E4", marginBottom: 10 }}>Lead TA Specialist</div>
                      <div style={{ fontSize: 8, color: "rgba(255,255,255,.4)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 4 }}>Contact</div>
                      <div style={{ fontSize: 9, color: "rgba(255,255,255,.65)", marginBottom: 2 }}>priya@email.com</div>
                      <div style={{ fontSize: 9, color: "rgba(255,255,255,.65)", marginBottom: 10 }}>Bangalore, KA</div>
                      <div style={{ fontSize: 8, color: "rgba(255,255,255,.4)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 4 }}>Skills</div>
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 3 }}>
                        {["Power BI", "Naukri", "HRBP", "SQL"].map(s => <span key={s} style={{ background: "rgba(3,29,64,.28)", color: "#C9D6E4", fontSize: 8, padding: "2px 5px", borderRadius: 3 }}>{s}</span>)}
                      </div>
                    </div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: 8, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#02457A", marginBottom: 5 }}>Professional Summary</div>
                      <p style={{ fontSize: 10, color: "var(--text-mid)", lineHeight: 1.75, marginBottom: 14 }}>Lead Talent Acquisition Specialist with 4.5 years driving end-to-end recruitment for pharma analytics and commercial intelligence roles. Targeting the HRBP Manager position at the target company, bringing expertise in Power BI dashboards, stakeholder management, and MOU partnership execution.</p>
                      <div style={{ fontSize: 8, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#02457A", marginBottom: 7 }}>Experience</div>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 2 }}>
                        <span style={{ fontSize: 11, fontWeight: 700, color: N1 }}>Lead TA Specialist</span>
                        <span style={{ fontSize: 9, color: "#C9D6E4" }}>2020 – Present</span>
                      </div>
                      <div style={{ fontSize: 10, color: "#02457A", marginBottom: 4 }}>Current Company · Bangalore</div>
                      <div style={{ fontSize: 9, color: "var(--text-mid)", lineHeight: 1.65 }}>• Reduced time-to-hire by 28% across 40+ analytics and engineering roles<br />• Built Power BI WFH dashboard tracking 150+ headcount across 4 quarters<br />• Executed MOUs with Dayananda Sagar College and MIT Pune — 12 hires in FY2024</div>
                    </div>
                  </div>
                  <div style={{ marginTop: 16, padding: "10px 14px", background: "var(--bg)", borderRadius: 8, display: "flex", alignItems: "center", gap: 8 }}>
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={AC} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10" /><polyline points="12,6 12,12 16,14" /></svg>
                    <span style={{ fontSize: 11, color: "var(--text-muted)" }}>Generated in <strong style={{ color: N1 }}>58 seconds</strong> · Tailored to the exact job description</span>
                  </div>
                </div>
              </div>
            )}

            {/* Cover Letter Sample */}
            {sampleTab === "cover" && (
              <div style={{ background: "var(--surface)", borderRadius: 20, overflow: "hidden", boxShadow: "var(--shadow-lg)", border: "1px solid var(--border)", animation: "fadeUp .3s ease" }}>
                <div style={{ background: "linear-gradient(135deg,#02457A,#02457A)", padding: "14px 22px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <div style={{ width: 7, height: 7, borderRadius: "50%", background: "#22C55E" }} />
                    <span style={{ color: "rgba(255,255,255,.65)", fontSize: 12 }}>cover_letter.pdf</span>
                  </div>
                  <span style={{ background: "rgba(255,255,255,.1)", color: "#D6E8EE", fontSize: 10, fontWeight: 600, padding: "2px 10px", borderRadius: 10 }}>Tailored</span>
                </div>
                <div style={{ padding: "32px 36px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 22 }}>
                    <div><div style={{ fontSize: 15, fontWeight: 700, color: "var(--text)" }}>Priya Sharma</div><div style={{ fontSize: 11, color: "var(--text-muted)" }}>priya@email.com · Bangalore, KA</div></div>
                    <div style={{ fontSize: 11, color: "var(--text-faint)" }}>5 April 2026</div>
                  </div>
                  <div style={{ width: 36, height: 3, background: "var(--grad)", borderRadius: 2, marginBottom: 18 }} />
                  <p style={{ fontSize: 13, color: "var(--text-mid)", lineHeight: 1.85, marginBottom: 14 }}>Dear Hiring Manager,</p>
                  <p style={{ fontSize: 13, color: "var(--text-mid)", lineHeight: 1.85, marginBottom: 14 }}>With 4.5 years building talent pipelines for pharma analytics functions, I bring exactly what the HRBP Manager role requires — domain fluency you cannot onboard, and execution speed your team can rely on from day one.</p>
                  <p style={{ fontSize: 13, color: "var(--text-mid)", lineHeight: 1.85, marginBottom: 14 }}>At my current organisation, I reduced time-to-hire by 28% across 40+ analytics and engineering roles by restructuring sourcing channels. I also built a Power BI WFH dashboard covering 150+ headcount — giving me a dual lens on both talent strategy and HR data storytelling.</p>
                  <p style={{ fontSize: 13, color: "var(--text-mid)", lineHeight: 1.85, marginBottom: 20 }}>I would welcome a conversation about how my background aligns with your team&apos;s priorities. Available for an interview at your convenience.</p>
                  <p style={{ fontSize: 13, color: "var(--text-mid)", lineHeight: 1.85 }}>Warm regards,<br /><strong style={{ color: "var(--text)" }}>Priya Sharma</strong></p>
                </div>
              </div>
            )}

            {/* Referral DM Sample */}
            {sampleTab === "referral" && (
              <div style={{ background: "var(--surface)", borderRadius: 20, overflow: "hidden", boxShadow: "var(--shadow-lg)", border: "1px solid var(--border)", animation: "fadeUp .3s ease" }}>
                <div style={{ background: "linear-gradient(135deg,#031D40,#08284F)", padding: "14px 22px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <div style={{ width: 7, height: 7, borderRadius: "50%", background: "#22C55E" }} />
                    <span style={{ color: "rgba(0,27,72,.7)", fontSize: 12 }}>referral_messages.txt</span>
                  </div>
                  <span style={{ background: "rgba(0,27,72,.12)", color: N1, fontSize: 10, fontWeight: 700, padding: "2px 10px", borderRadius: 10 }}>Outreach-ready</span>
                </div>
                <div className="output-pad" style={{ padding: "28px 32px" }}>
                  <div style={{ marginBottom: 22 }}>
                    <div style={{ display: "inline-block", background: "var(--accent-soft)", color: AC, fontSize: 9, fontWeight: 700, padding: "3px 10px", borderRadius: 4, marginBottom: 12, letterSpacing: ".06em" }}>VERSION 1 · CONNECTION REQUEST</div>
                    <div style={{ background: "var(--bg)", borderLeft: `3px solid ${AC}`, borderRadius: "0 10px 10px 0", padding: "14px 16px", fontSize: 13, color: "var(--text-mid)", lineHeight: 1.75 }}>
                      Hi Anjali, I noticed you&apos;re at [Company] — I&apos;m applying for the HRBP Manager role and have 4.5 years in pharma TA. Would love to connect!
                    </div>
                  </div>
                  <div style={{ height: 1, background: "var(--border)", marginBottom: 22 }} />
                  <div>
                    <div style={{ display: "inline-block", background: "var(--accent-soft)", color: AC, fontSize: 9, fontWeight: 700, padding: "3px 10px", borderRadius: 4, marginBottom: 12, letterSpacing: ".06em" }}>VERSION 2 · DIRECT MESSAGE</div>
                    <div style={{ background: "var(--bg)", borderLeft: `3px solid ${LB}`, borderRadius: "0 10px 10px 0", padding: "14px 16px", fontSize: 13, color: "var(--text-mid)", lineHeight: 1.85 }}>
                      Hi Anjali, hope you&apos;re doing well!<br /><br />
                      I came across the HRBP Manager opening at [Company] and I&apos;m genuinely excited — it aligns closely with what I&apos;ve been building over 4.5 years in pharma TA and HR analytics.<br /><br />
                      Would you be open to referring me or sharing tips on the process? Happy to send my resume directly. Really appreciate it!
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>

          <div style={{ textAlign: "center", marginTop: 24 }}>
            <button className="btn-primary" onClick={() => setPage(user ? "generate" : "auth")} style={{ background: "var(--grad)", color: "white", border: "none", borderRadius: 12, padding: "13px 30px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", boxShadow: "var(--shadow-accent)" }}>
              {user ? "Generate yours now" : "Get yours free, no card needed"}
            </button>
          </div>

          {sampleModalOpen && (
            <>
              <div onClick={() => setSampleModalOpen(false)} style={{ position: "fixed", inset: 0, background: "rgba(3,29,64,0.45)", zIndex: 300 }} />
              <div className="modal-enter" style={{ position: "fixed", top: "50%", left: "50%", width: "min(760px, calc(100vw - 24px))", maxHeight: "85vh", overflowY: "auto", background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 24, padding: 22, boxShadow: "var(--shadow-lg)", zIndex: 301 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, marginBottom: 18 }}>
                  <div>
                    <div style={{ fontSize: 12, fontWeight: 700, letterSpacing: ".12em", textTransform: "uppercase", color: AC, marginBottom: 8 }}>Sample career kit</div>
                    <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 28, color: DARK, marginBottom: 4 }}>{sampleModalContent[sampleTab].title}</div>
                    <div style={{ fontSize: 13, color: MUTED }}>{sampleModalContent[sampleTab].subtitle}</div>
                  </div>
                  <button onClick={() => setSampleModalOpen(false)} style={{ background: "var(--surface2)", border: `1px solid ${BORDER}`, borderRadius: 10, width: 34, height: 34, cursor: "pointer", color: MUTED, fontSize: 16 }}><i className="fa-solid fa-xmark" /></button>
                </div>
                <div className="two-col-md" style={{ display: "grid", gridTemplateColumns: "1.1fr .9fr", gap: 14 }}>
                  <div style={{ background: "var(--surface2)", border: `1px solid ${BORDER}`, borderRadius: 18, padding: "16px 18px" }}>
                    <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: AC, marginBottom: 10 }}>What it sounds like</div>
                    {sampleModalContent[sampleTab].excerpt.map((line, index) => (
                      <p key={index} style={{ fontSize: 13, color: MID, lineHeight: 1.8, marginBottom: index === sampleModalContent[sampleTab].excerpt.length - 1 ? 0 : 12 }}>{line}</p>
                    ))}
                  </div>
                  <div style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 18, padding: "16px 18px" }}>
                    <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: AC, marginBottom: 10 }}>Why it works</div>
                    <div style={{ display: "grid", gap: 8 }}>
                      {sampleModalContent[sampleTab].bullets.map((item) => (
                        <div key={item} style={{ display: "flex", gap: 8, alignItems: "flex-start", fontSize: 13, color: MID, lineHeight: 1.7 }}>
                          <span style={{ color: AC, fontWeight: 700 }}>✓</span>
                          <span>{item}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
                <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 18, flexWrap: "wrap" }}>
                  <button onClick={() => setSampleModalOpen(false)} style={{ background: WHITE, color: MUTED, border: `1px solid ${BORDER}`, borderRadius: 10, padding: "11px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>Close</button>
                  <button onClick={() => { setSampleModalOpen(false); setPage(user ? "generate" : "auth"); }} style={{ background: "var(--grad)", color: WHITE, border: "none", borderRadius: 10, padding: "11px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
                    Try this with my role
                  </button>
                </div>
              </div>
            </>
          )}
        </div>
      </div>

      {/* ── WHAT YOU GET ── */}
      <div className="section-pad" style={{ padding: "80px clamp(20px,5vw,80px)", background: "var(--bg)" }}>
        <div style={{ textAlign: "center", marginBottom: 52 }}>
          <div style={{ fontSize: 12, fontWeight: 700, letterSpacing: ".16em", textTransform: "uppercase", color: AC, marginBottom: 14 }}>What you get</div>
          <h2 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(28px,4vw,44px)", fontWeight: 400, letterSpacing: "-.02em", background: "var(--grad)", WebkitBackgroundClip: "text", MozBackgroundClip: "text", WebkitTextFillColor: "transparent", backgroundClip: "text", color: "transparent" }}>Everything you need to apply with confidence</h2>
          <p style={{ fontSize: 15, color: "var(--text-muted)", marginTop: 12, maxWidth: 560, margin: "12px auto 0" }}>Every output is tailored to the role and company you&apos;re targeting, so your application feels more complete from the start.</p>
        </div>
        <div className="feature-grid">
          {[
            { title: "Resume Rewrite", desc: "ATS-optimized, role-specific, quantified achievements. Ready to send.", plan: "Starter", planColor: LB, planBg: "rgba(3,29,64,0.08)", grad: "linear-gradient(135deg,#001B48,#031D40)", svg: <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14,2 14,8 20,8" /><line x1="16" y1="13" x2="8" y2="13" /><line x1="16" y1="17" x2="8" y2="17" /></svg> },
            { title: "Cover Letter", desc: "3 paragraphs, confident tone. Reads like you wrote it, not a template.", plan: "Starter", planColor: LB, planBg: "rgba(3,29,64,0.08)", grad: "linear-gradient(135deg,#031D40,#031D40)", svg: <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z" /><polyline points="22,6 12,13 2,6" /></svg> },
            { title: "Referral Message", desc: "LinkedIn note and DM. Makes it easy for your contact to say yes.", plan: "Starter", planColor: LB, planBg: "rgba(3,29,64,0.08)", grad: "linear-gradient(135deg,#031D40,#08284F)", svg: <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" /><circle cx="9" cy="7" r="4" /><path d="M23 21v-2a4 4 0 0 0-3-3.87" /><path d="M16 3.13a4 4 0 0 1 0 7.75" /></svg> },
            { title: "Interview Prep", desc: "Top 5 questions with STAR answers based on your actual experience.", plan: "Pro+", planColor: AC, planBg: "var(--accent-soft)", grad: "linear-gradient(135deg,#001B48,#031D40)", svg: <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10" /><circle cx="12" cy="12" r="6" /><circle cx="12" cy="12" r="2" /></svg> },
            { title: "Find & Reach", desc: "Exact LinkedIn search query and cold outreach to find the right recruiter.", plan: "Pro+", planColor: AC, planBg: "var(--accent-soft)", grad: "linear-gradient(135deg,#031D40,#08284F)", svg: <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" /></svg> },
            { title: "Salary Negotiation", desc: "Enter offer vs expectation — get a ready-to-use negotiation script.", plan: "Unlimited", planColor: N2, planBg: "rgba(3,29,64,0.08)", grad: "linear-gradient(135deg,#001B48,#031D40,#08284F)", svg: <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="1" x2="12" y2="23" /><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6" /></svg> },
            { title: "Notice Period", desc: "Verbal script, HR email, LinkedIn — all tailored to your notice period.", plan: "Unlimited", planColor: N2, planBg: "rgba(3,29,64,0.08)", grad: "linear-gradient(135deg,#031D40,#08284F)", svg: <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2" /><line x1="16" y1="2" x2="16" y2="6" /><line x1="8" y1="2" x2="8" y2="6" /><line x1="3" y1="10" x2="21" y2="10" /></svg> },
          ].map((f, i) => (
            <div key={i} className="feature-card card-hover grid-item" style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 18, padding: "28px 20px 22px", position: "relative" }}>
              <div style={{ position: "absolute", top: 14, right: 14 }}>
                <span className="feature-plan" style={{ fontSize: 10, fontWeight: 700, color: f.planColor, background: f.planBg, padding: "3px 9px", borderRadius: 20, letterSpacing: ".04em", whiteSpace: "nowrap", border: `1px solid ${f.planColor}30` }}>{f.plan}</span>
              </div>
              <div className="icon-wrap" style={{ width: 54, height: 54, borderRadius: 16, background: f.grad, display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 16, margin: "0 auto 16px", boxShadow: `0 6px 20px ${f.planColor}40` }}>{f.svg}</div>
              <div className="feature-copy">
                <div style={{ fontSize: 13, fontWeight: 700, color: "var(--text)", marginBottom: 8, textAlign: "center" }}>{f.title}</div>
                <div style={{ fontSize: 12, color: "var(--text-muted)", lineHeight: 1.7, textAlign: "center" }}>{f.desc}</div>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* ── TESTIMONIALS ── */}
      <div style={{ background: "var(--grad-subtle)", borderTop: "1px solid var(--border)", borderBottom: "1px solid var(--border)" }}>
        <div className="section-pad" style={{ padding: "72px clamp(20px,5vw,80px)" }}>
          <div style={{ textAlign: "center", marginBottom: 44 }}>
            <div style={{ fontSize: 12, fontWeight: 700, letterSpacing: ".16em", textTransform: "uppercase", color: AC, marginBottom: 14 }}>From our users</div>
            <h2 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(26px,3.5vw,40px)", fontWeight: 400, letterSpacing: "-.02em", background: "var(--grad)", WebkitBackgroundClip: "text", MozBackgroundClip: "text", WebkitTextFillColor: "transparent", backgroundClip: "text", color: "transparent" }}>What job seekers are saying</h2>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(260px,1fr))", gap: 18, maxWidth: 1080, margin: "0 auto" }}>
            {[
              { name: "Najma K.", role: "Human Resource Executive", company: "Bangalore", quote: "I had hit a wall applying to 30 roles with the same resume. Rezolt rebuilt it around the JD in under a minute. Got two calls within a week.", stars: 5 },
              { name: "Soham S.", role: "Frontend Developer", company: "Mumbai", quote: "My previous resume felt very generic. This tool didn't just summarize my work, it actually highlighted the exact React experience they asked for in the JD.", stars: 5 },
              { name: "Shreshth J.", role: "Data Analyst", company: "Chandigarh", quote: "The referral messages are the real hidden gem. I reached out to 4 contacts and 3 responded. That never happened before.", stars: 5 },
              { name: "Ehsan S.", role: "Product Manager", company: "Gurgaon", quote: "I usually spend 2 hours tweaking my resume for safety. Did it in 45 seconds here. The wording feels natural, not AI-generated.", stars: 5 },
            ].map((t, i) => (
              <div key={i} style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 20, padding: "22px 22px 18px", boxShadow: "var(--soft-shadow)", display: "flex", flexDirection: "column", gap: 12 }}>
                <div style={{ display: "flex", gap: 2 }}>
                  {Array.from({ length: t.stars }).map((_, j) => <span key={j} style={{ color: "#FBBF24", fontSize: 14 }}>★</span>)}
                </div>
                <p style={{ fontSize: 14, color: "var(--text-mid)", lineHeight: 1.75, fontStyle: "italic", flex: 1 }}>&ldquo;{t.quote}&rdquo;</p>
                <div style={{ display: "flex", alignItems: "center", gap: 10, borderTop: "1px solid var(--border)", paddingTop: 12 }}>
                  <div style={{ width: 36, height: 36, borderRadius: "50%", background: "linear-gradient(135deg,#001B48,#02457A)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, fontWeight: 700, color: "white", flexShrink: 0 }}>{t.name[0]}</div>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: "var(--text)" }}>{t.name}</div>
                    <div style={{ fontSize: 11, color: "var(--text-muted)" }}>{t.role} · {t.company}</div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* ── PRICING ── */}
      <div id="pricing-section" className="section-pad" style={{ padding: "80px clamp(20px,5vw,80px)", background: "var(--bg)" }}>
        <div style={{ textAlign: "center", marginBottom: 52 }}>
          <div style={{ fontSize: 12, fontWeight: 700, letterSpacing: ".16em", textTransform: "uppercase", color: AC, marginBottom: 14 }}>Pricing</div>
          <h2 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(28px,4vw,44px)", fontWeight: 400, letterSpacing: "-.02em", background: "var(--grad)", WebkitBackgroundClip: "text", MozBackgroundClip: "text", WebkitTextFillColor: "transparent", backgroundClip: "text", color: "transparent" }}>Choose a plan that fits your search</h2>
          <p style={{ fontSize: 15, color: "var(--text-muted)", marginTop: 12 }}>Start free, then upgrade only when you want more depth and more outputs.</p>
        </div>
        <div className="pricing-grid" style={{ display: "grid", gridTemplateColumns: "repeat(4,minmax(0,1fr))", gap: 16, maxWidth: 1200, margin: "0 auto", alignItems: "stretch" }}>
          {[
            {
              name: "Free",
              price: "0",
              tag: "forever free",
              popular: false,
              note: "Try it out, no card needed",
              includes: [
                "1 free resume rewrite",
                { text: "Cover letter + Referral DM", cross: true },
                { text: "Interview Prep + Find & Reach", cross: true },
                { text: "Salary Negotiation tool", cross: true },
              ],
            },
            {
              name: "Starter",
              price: "99",
              tag: "1 Career kit",
              popular: false,
              note: "One strong application, fully built",
              includes: [
                "1 complete career kit",
                "Resume rewrite",
                "Cover letter + Referral DM",
                "Interview Prep + Find & Reach",
                "PDF & DOCX export",
                { text: "Salary Negotiation tool", cross: true },
              ],
            },
            {
              name: "Pro",
              price: "299",
              tag: "5 Career kits",
              popular: true,
              note: "Best for active job seekers",
              includes: [
                "5 complete career kits",
                "Resume, Cover letter, Referral DM",
                "Interview Prep + Find & Reach",
                "Priority generation",
                "Valid for 90 days",
                { text: "Salary Negotiation tool", cross: true },
              ],
            },
            {
              name: "Unlimited",
              price: "599",
              tag: "per month",
              popular: false,
              note: "Best for negotiation and ongoing search",
              includes: [
                "Unlimited career kits",
                "All 5 kit outputs",
                "Salary Negotiation tool",
                "Notice Period scripts",
                "First access to new features",
                "Cancel anytime",
              ],
            },
          ].map((p, i) => (
            <div key={i} style={{ background: p.popular ? "var(--grad)" : "var(--surface)", border: p.popular ? "none" : "1.5px solid var(--border)", borderRadius: 22, padding: "40px 24px 28px", position: "relative", boxShadow: p.popular ? "var(--shadow-lg)" : "var(--shadow-sm)", transition: "transform .25s ease, box-shadow .25s ease", display: "flex", flexDirection: "column", minHeight: "100%", height: "100%" }}
              onMouseEnter={e => { e.currentTarget.style.transform = "translateY(-6px)"; e.currentTarget.style.boxShadow = p.popular ? "0 20px 50px rgba(3,29,64,.32)" : "var(--shadow-lg)"; }}
              onMouseLeave={e => { e.currentTarget.style.transform = "translateY(0)"; e.currentTarget.style.boxShadow = p.popular ? "var(--shadow-lg)" : "var(--shadow-sm)"; }}
            >
              {p.popular && <div style={{ position: "absolute", top: -13, left: "50%", transform: "translateX(-50%)", background: PB, color: N1, fontSize: 11, fontWeight: 800, padding: "4px 16px", borderRadius: 20, letterSpacing: ".05em", whiteSpace: "nowrap" }}>BEST VALUE</div>}

              <div style={{ fontSize: 12, fontWeight: 700, color: p.popular ? "rgba(255,255,255,.65)" : "var(--text-muted)", textTransform: "uppercase", letterSpacing: ".1em", marginBottom: 10 }}>{p.name}</div>
              <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 44, fontWeight: 400, color: p.popular ? "white" : "var(--text)", lineHeight: 1, marginBottom: 4 }}>₹{p.price}</div>
              <div style={{ fontSize: 13, color: p.popular ? "rgba(255,255,255,.6)" : "var(--text-muted)", marginBottom: 10 }}>{p.tag}</div>
              <div style={{ fontSize: 12, color: p.popular ? "rgba(255,255,255,.78)" : AC, fontWeight: 700, marginBottom: 22 }}>{p.note}</div>
              <div style={{ flex: 1, marginBottom: 28, display: "grid", gap: 9 }}>
                {p.includes.map((f, j) => {
                  const isCross = typeof f === "object" && f.cross;
                  const text = typeof f === "object" ? f.text : f;
                  return (
                    <div key={j} style={{ display: "flex", gap: 9, alignItems: "flex-start", opacity: isCross ? 0.6 : 1 }}>
                      {isCross ? (
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={p.popular ? "rgba(255,255,255,.5)" : "var(--text-muted)"} strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0, marginTop: 2 }}><line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" /></svg>
                      ) : (
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={p.popular ? "#86efac" : G} strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0, marginTop: 2 }}><polyline points="20 6 9 17 4 12" /></svg>
                      )}
                      <span style={{ fontSize: 13, color: p.popular ? "rgba(255,255,255,.86)" : "var(--text-mid)", lineHeight: 1.6, textDecoration: isCross ? "line-through" : "none" }}>{text}</span>
                    </div>
                  )
                })}
              </div>
              {(() => {
                const currentPlan = (user?.plan || "Free").toLowerCase();
                const cardName = p.name.toLowerCase();
                const isCurrentPlan = user && (cardName === currentPlan || (cardName === "unlimited" && currentPlan === "unlimited_monthly"));
                return (
                  <button onClick={() => isCurrentPlan ? setPage("dashboard") : setPage(user ? "payment" : "auth")} className={isCurrentPlan ? "" : "btn-primary"} style={{ width: "100%", marginTop: "auto", background: isCurrentPlan ? (p.popular ? "rgba(255,255,255,.1)" : "var(--surface2)") : p.popular ? "rgba(255,255,255,.15)" : "var(--grad)", color: isCurrentPlan ? (p.popular ? "rgba(255,255,255,.6)" : "var(--text-muted)") : "white", border: isCurrentPlan ? (p.popular ? "1px solid rgba(255,255,255,.2)" : "1px solid var(--border)") : p.popular ? "2px solid rgba(255,255,255,.3)" : "none", borderRadius: 12, padding: "13px", fontSize: 14, fontWeight: 700, cursor: isCurrentPlan ? "default" : "pointer", fontFamily: "inherit", backdropFilter: p.popular ? "blur(8px)" : "none" }}>
                    {user ? (isCurrentPlan ? "Current Plan" : "Upgrade Now") : "Get Started"}
                  </button>
                );
              })()}
            </div>
          ))}
        </div>
      </div>

      {/* ── ARTICLES ── */}
      <div id="articles-section" style={{ background: "var(--bg)", borderTop: "1px solid var(--border)" }}>
        <div className="section-pad" style={{ padding: "80px clamp(20px,5vw,80px)" }}>
          <div style={{ textAlign: "center", marginBottom: 56 }}>
            <div style={{ fontSize: 12, fontWeight: 700, letterSpacing: ".16em", textTransform: "uppercase", color: AC, marginBottom: 14 }}>Career insights</div>
            <h2 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(28px,4vw,44px)", fontWeight: 400, letterSpacing: "-.02em", background: "var(--grad)", WebkitBackgroundClip: "text", MozBackgroundClip: "text", WebkitTextFillColor: "transparent", backgroundClip: "text", color: "transparent" }}>
              Practical guidance for a tougher job market
            </h2>
            <p style={{ fontSize: 15, color: "var(--text-muted)", marginTop: 14, maxWidth: 560, margin: "14px auto 0", lineHeight: 1.7 }}>
              You don’t always need more experience, sometimes you need clearer positioning, stronger proof, and sharper relevance.
            </p>
          </div>

          <div className="card-hover" style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 24, padding: "32px 32px", marginBottom: 22, position: "relative", overflow: "hidden", boxShadow: "var(--soft-shadow)" }}>

            <div className="two-col-md article-featured" style={{ display: "grid", gridTemplateColumns: "1.2fr .8fr", gap: 26, alignItems: "start" }}>
              <div>
                <div style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "var(--accent-soft)", border: "1px solid var(--border)", borderRadius: 20, padding: "4px 12px", marginBottom: 16 }}>
                  <div style={{ width: 6, height: 6, borderRadius: "50%", background: AC }} />
                  <span style={{ fontSize: 11, fontWeight: 700, color: AC }}>What matters most</span>
                </div>
                <h3 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(20px,3vw,30px)", fontWeight: 400, color: "var(--text)", letterSpacing: "-.02em", marginBottom: 14, lineHeight: 1.25 }}>
                  Why strong candidates still get filtered out
                </h3>
                <p style={{ fontSize: 14, color: "var(--text-muted)", lineHeight: 1.8, marginBottom: 16 }}>
                  In high-volume hiring, resumes are scanned before stories are heard. The first goal isn’t to explain everything, it’s to make relevance obvious in seconds.
                </p>
                <div style={{ display: "grid", gap: 10, marginBottom: 18 }}>
                  {[
                    "Match the language of the role where it's genuinely true.",
                    "Show outcomes, not only responsibilities.",
                    "Make the recruiter’s next decision feel easy.",
                  ].map(item => (
                    <div key={item} style={{ display: "flex", alignItems: "center", gap: 10, fontSize: 13, color: MID }}>
                      <span style={{ width: 20, height: 20, borderRadius: "50%", background: "var(--accent-soft)", color: AC, display: "inline-flex", alignItems: "center", justifyContent: "center", fontWeight: 700, flexShrink: 0 }}>✓</span>
                      <span>{item}</span>
                    </div>
                  ))}
                </div>
                <button onClick={() => { setSelectedArticle?.(1); setPage("articles"); }} style={{ background: "none", border: "none", padding: 0, color: AC, fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
                  Read the full article
                </button>
              </div>
              <div style={{ display: "grid", gap: 12 }}>
                <div style={{ background: "var(--grad)", borderRadius: 16, padding: "22px 18px", color: "white", textAlign: "center" }}>
                  <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 42, fontWeight: 400, lineHeight: 1, marginBottom: 6 }}>75%</div>
                  <div style={{ fontSize: 12, color: "rgba(255,255,255,.74)", lineHeight: 1.5 }}>Research suggests up to 75% of resumes are filtered before a recruiter reads them</div>
                </div>
                <div style={{ background: "var(--surface2)", borderRadius: 14, padding: "16px 18px", border: "1px solid var(--border)" }}>
                  <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 30, fontWeight: 400, color: AC, lineHeight: 1, marginBottom: 4 }}>6s</div>
                  <div style={{ fontSize: 12, color: "var(--text-muted)", lineHeight: 1.6 }}>Studies show recruiters spend an average of 6–7 seconds on a first scan</div>
                </div>
              </div>
            </div>
          </div>

          <div className="three-col" style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 18, marginBottom: 24 }}>
            {[
              {
                title: "Why ATS fit matters more than you think",
                summary: "Use the same language as the job description where it is true and proven.",
                stat: "Keywords",
                tag: "ATS",
                articleIndex: 0,
              },
              {
                title: "Proof beats promises",
                summary: "Numbers, outcomes, and measurable wins create faster trust than generic claims.",
                stat: "Impact",
                tag: "Positioning",
                articleIndex: 2,
              },
              {
                title: "Relevance creates momentum",
                summary: "A cleaner, more targeted resume improves the odds of better outreach and better interviews.",
                stat: "Clarity",
                tag: "Strategy",
                articleIndex: 5,
              },
            ].map((a, i) => (
              <button key={i} onClick={() => { setSelectedArticle?.(a.articleIndex); setPage("articles"); }} className="card-hover" style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 18, padding: "24px 22px", display: "flex", flexDirection: "column", textAlign: "left", cursor: "pointer", fontFamily: "inherit" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                  <span style={{ background: "var(--accent-soft)", color: AC, fontSize: 10, fontWeight: 700, padding: "3px 10px", borderRadius: 20, border: "1px solid var(--border)" }}>{a.tag}</span>
                  <span style={{ fontSize: 11, fontWeight: 700, color: FAINT }}>{a.stat}</span>
                </div>
                <h3 style={{ fontFamily: "'DM Serif Display', serif", fontSize: 18, fontWeight: 400, color: "var(--text)", marginBottom: 10, lineHeight: 1.35, letterSpacing: "-.01em" }}>{a.title}</h3>
                <p style={{ fontSize: 13, color: "var(--text-muted)", lineHeight: 1.75, marginBottom: 14 }}>{a.summary}</p>
                <span style={{ marginTop: "auto", fontSize: 12, fontWeight: 700, color: AC }}>Read more</span>
              </button>
            ))}
          </div>

          <div style={{ background: "var(--grad)", borderRadius: 20, padding: "34px 36px", textAlign: "center", position: "relative", overflow: "hidden" }}>
            <div className="blob" style={{ width: 200, height: 200, background: LB, top: -60, right: 40 }} />
            <div style={{ position: "relative", zIndex: 1 }}>
              <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: ".15em", textTransform: "uppercase", color: "rgba(255,255,255,.6)", marginBottom: 14 }}>The core formula</div>
              <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(14px,2.5vw,22px)", color: "white", fontWeight: 400, lineHeight: 1.6, marginBottom: 18 }}>
                Resume Effectiveness = <span style={{ color: PB }}>Keyword Match</span> × <span style={{ color: PB }}>Clarity</span> × <span style={{ color: PB }}>Measurable Impact</span> × <span style={{ color: PB }}>ATS Compatibility</span>
              </div>
              <p style={{ fontSize: 14, color: "rgba(255,255,255,.74)", marginBottom: 24 }}>If even one of these drops to zero, your chances drop with it. Rezolt is designed to strengthen all four together.</p>
              <button onClick={() => setPage(user ? "generate" : "auth")} className="hero-btn btn-primary" style={{ background: "linear-gradient(135deg,#E4BE47,#F2D46C,#E4BE47)", backgroundSize: "200% auto", color: N1, border: "none", borderRadius: 12, padding: "14px 32px", fontSize: 15, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" }}>
                {user ? "Build My Next Kit" : "Start Free"}
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* ── FUN CTA ── */}
      <div style={{ padding: "0 clamp(16px,5vw,80px) 80px", background: "var(--bg)" }} className="section-pad">
        <div className="hero-gradient" style={{ borderRadius: 28, padding: "72px 40px", textAlign: "center", position: "relative", overflow: "hidden" }}>
          <div className="blob" style={{ width: 300, height: 300, background: AC, top: -80, right: 40 }} />
          <div className="blob" style={{ width: 200, height: 200, background: LB, bottom: -60, left: 60 }} />
          <div style={{ position: "relative", zIndex: 1 }}>
            <div style={{ display: "inline-flex", alignItems: "center", gap: 8, background: "rgba(3,29,64,0.06)", border: "1px solid rgba(3,29,64,0.10)", borderRadius: 999, padding: "6px 14px", fontSize: 12, fontWeight: 700, color: AC, marginBottom: 16 }}>
              Ready when you are
            </div>
            <h2 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(30px,4vw,52px)", fontWeight: 400, letterSpacing: "-.02em", color: "var(--text)", marginBottom: 14 }}>
              Your next offer starts <em style={{ color: PB }}>here.</em>
            </h2>
            <p style={{ fontSize: 16, color: "var(--text-muted)", marginBottom: 36 }}>
              {user ? "You're already in. Generate your next Career kit in 60 seconds." : "Sign up now and get 3 free Career Kits instantly. No card required."}
            </p>
            <button onClick={() => setPage(user ? "generate" : "auth")} className="hero-btn btn-primary" style={{ background: "linear-gradient(135deg,#E4BE47,#F2D46C,#E4BE47)", backgroundSize: "200% auto", color: N1, border: "none", borderRadius: 12, padding: "16px 40px", fontSize: 16, fontWeight: 800, cursor: "pointer", fontFamily: "inherit", boxShadow: "0 6px 22px rgba(228,190,71,.32)" }}>
              {user ? "Generate a Kit" : "Get Started Free"}
            </button>
          </div>
        </div>
      </div>

      {/* ── FOOTER ── */}
      <div style={{ borderTop: "1px solid var(--border)", background: "var(--surface)", padding: "28px clamp(16px,5vw,80px)" }} className="section-pad">
        <div className="footer-grid" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 16 }}>
          <div style={{ display: "flex", alignItems: "center" }}>
            <BrandLogo height={60} />
          </div>
          <div style={{ fontSize: 13, color: "var(--text-muted)" }}>© 2026 Rezolt. Career Kit for Indian Job Seekers.</div>
          <div style={{ display: "flex", gap: 18, alignItems: "center", flexWrap: "wrap" }}>
            {[
              { label: "Privacy Policy", action: () => setPage("privacy") },
              { label: "Terms", action: () => setPage("terms") },
              { label: "FAQ", action: () => setPage("faq") },
              { label: "Contact Us", action: () => setPage("contact") },
            ].map(l => (
              <button key={l.label} onClick={l.action} style={{ background: "none", border: "none", padding: 0, fontSize: 13, color: "var(--text-muted)", cursor: "pointer", transition: "color .15s ease", fontFamily: "inherit" }}
                onMouseEnter={e => e.target.style.color = AC} onMouseLeave={e => e.target.style.color = "var(--text-muted)"}>{l.label}</button>
            ))}
            <a href="https://www.instagram.com/rezolt.in/" target="_blank" rel="noreferrer" style={{ fontSize: 13, color: "var(--text-muted)", textDecoration: "none", transition: "color .15s ease", display: "inline-flex", alignItems: "center", gap: 6 }}
              onMouseEnter={e => e.currentTarget.style.color = AC} onMouseLeave={e => e.currentTarget.style.color = "var(--text-muted)"}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><rect x="3" y="3" width="18" height="18" rx="5" ry="5" /><path d="M16.5 7.5h.01" /><circle cx="12" cy="12" r="4" /></svg>
              <span>Instagram</span>
            </a>
            <a href="https://www.linkedin.com/company/rezolt/" target="_blank" rel="noreferrer" style={{ fontSize: 13, color: "var(--text-muted)", textDecoration: "none", transition: "color .15s ease", display: "inline-flex", alignItems: "center", gap: 6 }}
              onMouseEnter={e => e.currentTarget.style.color = AC} onMouseLeave={e => e.currentTarget.style.color = "var(--text-muted)"}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M6.94 8.5a1.56 1.56 0 1 1 0-3.12 1.56 1.56 0 0 1 0 3.12ZM5.5 9.75h2.88V18H5.5V9.75Zm4.68 0h2.76v1.13h.04c.38-.73 1.32-1.5 2.72-1.5 2.92 0 3.46 1.92 3.46 4.42V18h-2.88v-3.72c0-.89-.02-2.03-1.24-2.03-1.24 0-1.43.97-1.43 1.97V18h-2.88V9.75Z" /></svg>
              <span>LinkedIn</span>
            </a>
            <a href="mailto:hello@rezolt.in" style={{ fontSize: 13, color: "var(--text-muted)", textDecoration: "none", transition: "color .15s ease" }}
              onMouseEnter={e => e.target.style.color = AC} onMouseLeave={e => e.target.style.color = "var(--text-muted)"}>
              hello@rezolt.in
            </a>
          </div>
        </div>
      </div>
    </div>
  );
}



function ArticlesPage({ setPage, user, selectedArticle = 0, setSelectedArticle = () => { } }) {
  const [activeArticle, setActiveArticle] = useState(selectedArticle || 0);

  useEffect(() => {
    setActiveArticle(selectedArticle || 0);
  }, [selectedArticle]);

  const featured = {
    title: "The Reality of the Indian Job Market — Resume as a Filtering Weapon",
    body: [
      "India's job market is not skill-first. It is filter-first. At scale, companies receive hundreds to thousands of applications per role, which forces evaluation to happen in layers rather than holistically.",
      "The resume becomes the first elimination gate, not just a representation tool. Employers rely heavily on ATS screening and quick recruiter scans simply because application volume is so high.",
      "A candidate can be highly capable and still lose out if their value is not easy to interpret within a few seconds. That is why clarity, relevance, and measurable outcomes matter so much.",
    ],
  };

  const articles = [
    {
      tag: "ATS & Algorithms",
      readTime: "4 min read",
      title: "How resumes actually work — ATS, keywords, and ranking logic",
      text: "Modern resumes operate inside an algorithm. ATS systems parse information, match keywords against job descriptions, and rank candidates based on relevance.",
      content: [
        "An ATS is not judging your potential. It is sorting for relevance. It scans job titles, skills, locations, dates, and repeated language from the job description so a recruiter can reduce a large pile of applicants into a smaller shortlist.",
        "That means strong candidates often lose before a human conversation begins. If your resume uses vague wording, hides important skills, or buries outcomes under generic responsibilities, the match can look weaker than it really is.",
        "The practical takeaway is simple: mirror the language of the role where it is genuinely true, keep formatting easy to parse, and make business impact visible in the first half of the page. Clarity is not cosmetic. It directly affects whether you survive the first filter.",
      ],
    },
    {
      tag: "Career Strategy",
      readTime: "3 min read",
      title: "Why strong candidates still do not get interview calls",
      text: "High skill does not automatically create visibility. Recruiters optimise for speed, and resumes that show immediate role-match move forward faster.",
      content: [
        "A lot of job seekers assume that good experience should speak for itself. In reality, hiring teams are reading under pressure. They are scanning for fit, not reading each application like a case study.",
        "This is why candidates with real capability can still feel invisible. Their work may be impressive, but the story is not arranged around the employer’s current need. A recruiter sees effort, but not relevance.",
        "The better strategy is to position your experience around the decision the recruiter is making today: Can this person solve the problems in this role quickly and credibly? When your resume answers that directly, interview rates usually improve.",
      ],
    },
    {
      tag: "Positioning",
      readTime: "3 min read",
      title: "Why measurable proof changes everything",
      text: "Recruiters trust outcomes more than effort. Numbers, improvements, reductions, scale, and business impact make your value easier to believe and easier to shortlist.",
      content: [
        "A line like ‘responsible for hiring’ sounds ordinary. A line like ‘reduced time-to-hire by 28% across 40+ roles’ instantly creates trust because it shows scope, ownership, and outcome in one sentence.",
        "Metrics do not have to be dramatic. They can be counts, turnaround time, revenue influenced, user base, team size, process improvement, customer satisfaction, or error reduction. The point is to make your contribution tangible.",
        "When a recruiter sees proof, they spend less energy guessing. That makes your profile easier to defend internally and easier to move to the next round.",
      ],
    },
    {
      tag: "Outcomes & ROI",
      readTime: "4 min read",
      title: "How resume quality affects the rest of your search",
      text: "Resume quality influences interview rate, salary positioning, career mobility, and the speed of your pipeline. Better framing creates better conversion.",
      content: [
        "A sharper resume does more than improve application quality. It affects the tone of recruiter conversations, the confidence of people referring you, and the way you introduce yourself in interviews.",
        "When your materials are well-positioned, you enter the process with stronger leverage. You are not spending the first ten minutes correcting confusion about your background. You are building momentum from the start.",
        "This is why resume work has real return on effort. Improving one document can make every application, referral message, and salary conversation more effective.",
      ],
    },
    {
      tag: "Future of Hiring",
      readTime: "4 min read",
      title: "Resume is no longer enough — but it is still the entry ticket",
      text: "LinkedIn, referrals, and portfolios matter more than before, but the resume still activates the rest of the process.",
      content: [
        "Hiring today is multi-signal. Recruiters may look at your LinkedIn, portfolio, GitHub, referrals, and interview presence. But the resume is still the document that travels most easily inside a company and anchors the first internal discussion.",
        "That is why candidates should think in systems. Your LinkedIn should reinforce your positioning, your outreach should support your story, and your resume should make your credibility obvious quickly.",
        "The resume is not the whole game, but it is still the ticket that gets you into the game. When it is weak, every other advantage has to work harder.",
      ],
    },
    {
      tag: "Practical takeaway",
      readTime: "3 min read",
      title: "What job seekers should optimise first",
      text: "The best first move is not sending more applications. It is improving the signal quality of the resume you are sending.",
      content: [
        "If your current application rate is high but response rate is low, more volume is rarely the answer. The first optimisation is usually sharper positioning: stronger headlines, clearer skill alignment, and better proof of impact.",
        "Once the resume improves, your cover letter, outreach message, and interview prep all become easier because the story is already more coherent. You stop starting from scratch each time.",
        "The smartest early move is simple: improve the asset you reuse most. One high-quality resume kit can influence dozens of future opportunities.",
      ],
    },
  ];

  return (
    <div className="fade-in">
      <div className="hero-gradient" style={{ padding: "56px clamp(16px,5vw,80px) 54px", position: "relative", overflow: "hidden" }}>
        <div style={{ position: "relative", zIndex: 1, maxWidth: 1040, margin: "0 auto", textAlign: "center" }}>
          <div style={{ display: "inline-flex", alignItems: "center", gap: 8, background: "var(--accent-soft)", border: "1px solid rgba(3,29,64,0.08)", borderRadius: 999, padding: "6px 14px", marginBottom: 18, color: AC, fontSize: 12, fontWeight: 700 }}>
            <i className="fa-solid fa-newspaper" />Articles & Insights
          </div>
          <h1 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(32px,4vw,50px)", fontWeight: 400, color: "var(--text)", marginBottom: 12 }}>
            Career guidance worth reading in full
          </h1>
          <p style={{ fontSize: 15, color: "var(--text-muted)", lineHeight: 1.8, maxWidth: 760, margin: "0 auto 22px" }}>
            Your earlier articles are kept here completely in a separate sheet-style page, so the homepage stays cleaner while the full guidance remains available.
          </p>
          <div style={{ display: "flex", justifyContent: "center", gap: 10, flexWrap: "wrap" }}>
            <button onClick={() => setPage("landing")} style={{ background: WHITE, color: AC, border: "1px solid var(--border)", borderRadius: 12, padding: "12px 24px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
              Back to home
            </button>
            <button onClick={() => setPage(user ? "generate" : "auth")} className="btn-primary" style={{ background: "var(--grad)", color: WHITE, border: "none", borderRadius: 12, padding: "12px 24px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
              {user ? "Build a kit" : "Get started"}
            </button>
          </div>
        </div>
      </div>

      <div className="section-pad" style={{ padding: "44px clamp(16px,5vw,80px) 80px", background: "var(--bg)" }}>
        <div style={{ maxWidth: 1180, margin: "0 auto" }}>
          <div className="card-hover" style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 24, padding: "30px 28px", marginBottom: 22, boxShadow: "var(--soft-shadow)", textAlign: "justify", textJustify: "inter-word" }}>
            <div style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "var(--accent-soft)", border: "1px solid var(--border)", borderRadius: 999, padding: "4px 12px", marginBottom: 16 }}>
              <div style={{ width: 6, height: 6, borderRadius: "50%", background: AC }} />
              <span style={{ fontSize: 11, fontWeight: 700, color: AC }}>Featured</span>
            </div>
            <h2 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(22px,3vw,34px)", fontWeight: 400, color: "var(--text)", lineHeight: 1.25, marginBottom: 14, textAlign: "center" }}>{featured.title}</h2>
            {featured.body.map(p => (
              <p key={p} style={{ fontSize: 14, color: "var(--text-muted)", lineHeight: 1.85, marginBottom: 14, textAlign: "justify", textJustify: "inter-word" }}>{p}</p>
            ))}
            <div style={{ background: "var(--surface2)", border: "1px solid var(--border)", borderRadius: 16, padding: "14px 16px", fontSize: 13, color: MID, lineHeight: 1.8, textAlign: "justify", textJustify: "inter-word" }}>
              A resume is not just a summary of your background. It is a document engineered to survive elimination layers and create a stronger first impression.
            </div>
          </div>

          <div className="three-col" style={{ display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 18, marginBottom: 24 }}>
            {articles.map((item, index) => (
              <button
                key={item.title}
                onClick={() => {
                  setActiveArticle(index);
                  setSelectedArticle(index);
                  setTimeout(() => document.getElementById("article-detail")?.scrollIntoView({ behavior: "smooth", block: "start" }), 80);
                }}
                className="card-hover"
                style={{ background: WHITE, border: `1px solid ${index === activeArticle ? "rgba(3,29,64,0.22)" : BORDER}`, borderRadius: 20, padding: "22px 20px", boxShadow: "var(--soft-shadow)", textAlign: "left", cursor: "pointer", fontFamily: "inherit" }}
              >
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
                  <div style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "var(--accent-soft)", border: "1px solid var(--border)", borderRadius: 999, padding: "3px 10px" }}>
                    <span style={{ fontSize: 10, fontWeight: 700, color: AC }}>{item.tag}</span>
                  </div>
                  <span style={{ fontSize: 11, color: MUTED, fontWeight: 600 }}>{item.readTime}</span>
                </div>
                <div style={{ fontSize: 17, fontWeight: 700, color: DARK, lineHeight: 1.45, marginBottom: 10 }}>{item.title}</div>
                <div style={{ fontSize: 13, color: MUTED, lineHeight: 1.75, marginBottom: 14 }}>{item.text}</div>
                <span style={{ fontSize: 12, fontWeight: 700, color: AC }}>Read full article</span>
              </button>
            ))}
          </div>

          <div id="article-detail" style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 24, padding: "30px 28px", boxShadow: "var(--soft-shadow)", textAlign: "justify", textJustify: "inter-word" }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 14 }}>
              <div style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "var(--accent-soft)", border: "1px solid var(--border)", borderRadius: 999, padding: "4px 12px" }}>
                <span style={{ fontSize: 11, fontWeight: 700, color: AC }}>{articles[activeArticle].tag}</span>
              </div>
              <span style={{ fontSize: 12, color: MUTED, fontWeight: 600 }}>{articles[activeArticle].readTime}</span>
            </div>
            <div style={{ maxWidth: 980, margin: "0 auto" }}>
              <h2 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(24px,3vw,36px)", fontWeight: 400, color: DARK, lineHeight: 1.3, marginBottom: 10, textAlign: "center" }}>{articles[activeArticle].title}</h2>
              <p style={{ fontSize: 14, color: MID, lineHeight: 1.8, marginBottom: 16, textAlign: "justify", textJustify: "inter-word" }}>{articles[activeArticle].text}</p>
              <div style={{ display: "grid", gap: 14 }}>
                {articles[activeArticle].content.map(paragraph => (
                  <p key={paragraph} style={{ fontSize: 14, color: MUTED, lineHeight: 1.9, textAlign: "justify", textJustify: "inter-word" }}>{paragraph}</p>
                ))}
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function AuthPage({ onAuth, setPage }) {
  const [mode, setMode] = useState("signup");
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [verifyMsg, setVerifyMsg] = useState("");
  const [resent, setResent] = useState(false);
  const [tcAccepted, setTcAccepted] = useState(false);

  const resendVerification = async () => {
    if (!email) return setError("Enter your email first so we can resend the confirmation link.");
    setLoading(true); setError(""); setResent(false);
    const { error: err } = await supabase.auth.resend({
      type: "signup",
      email,
      options: { emailRedirectTo: `${window.location.origin}` },
    });
    if (err) { setError(err.message); setLoading(false); return; }
    setVerifyMsg("Fresh confirmation link sent. If you do not see it, please check Spam or Promotions.");
    setResent(true);
    setLoading(false);
  };

  const handle = async () => {
    if (!email || !password) return setError("Please fill in all fields.");
    if (mode === "signup" && !name) return setError("Please enter your name.");
    if (mode === "signup" && !tcAccepted) return setError("Please accept the Terms of Service and Privacy Policy to continue.");
    setLoading(true); setError(""); setVerifyMsg(""); setResent(false);
    if (mode === "signup") {
      const { data, error: err } = await supabase.auth.signUp({
        email,
        password,
        options: {
          data: { full_name: name },
          emailRedirectTo: `${window.location.origin}`,
        },
      });
      if (err) { setError(err.message); setLoading(false); return; }
      if (data.user && !data.session) {
        setVerifyMsg("Check your inbox to confirm your email. If it is not visible, please look in Spam or Promotions.");
        setLoading(false);
        return;
      }
      if (data.user) { onAuth({ id: data.user.id, name: name, email, credits: 1, plan: "Free" }); setPage("dashboard"); }
    } else {
      const { data, error: err } = await supabase.auth.signInWithPassword({ email, password });
      if (err) {
        if (/email not confirmed/i.test(err.message || "")) {
          setVerifyMsg("Your account still needs email confirmation. Check Spam or Promotions, then come back here.");
          setLoading(false);
          return;
        }
        setError(err.message);
        setLoading(false);
        return;
      }
      if (data.user) { onAuth({ id: data.user.id, name: data.user.email.split("@")[0], email, credits: 1, plan: "Free" }, data.session?.access_token); setPage("dashboard"); }
    }
    setLoading(false);
  };

  return (
    <div className="fade-in" style={{ minHeight: "calc(100vh - 85px)", display: "flex", alignItems: "center", justifyContent: "center", padding: 24, background: "var(--bg)" }}>
      {/* Background decoration */}
      <div style={{ position: "fixed", top: 0, right: 0, width: 500, height: 500, background: "radial-gradient(circle, rgba(3,29,64,0.08) 0%, transparent 70%)", pointerEvents: "none", zIndex: 0 }} />

      <div style={{ width: "100%", maxWidth: 440, position: "relative", zIndex: 1 }}>
        {/* Header */}
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 16 }}>
            <BrandLogo height={90} />
          </div>
          <h1 style={{ fontFamily: "'DM Serif Display', serif", fontSize: 30, fontWeight: 400, color: "var(--text)", marginBottom: 6 }}>
            {mode === "signup" ? "Create your account" : "Welcome back"}
          </h1>
          <p style={{ fontSize: 14, color: "var(--text-muted)" }}>{mode === "signup" ? "Get 3 free Career Kits when you sign up" : "Sign in to your Rezolt account"}</p>
        </div>

        <div className="auth-card" style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 22, padding: "36px 32px", boxShadow: "var(--shadow-lg)" }}>
          {/* Mode switcher */}
          <div style={{ display: "flex", gap: 0, marginBottom: 28, background: "var(--surface2)", borderRadius: 12, padding: 4 }}>
            {["signup", "login"].map(m => (
              <button key={m} className="mode-toggle" onClick={() => { setMode(m); setError(""); setVerifyMsg(""); }} style={{
                flex: 1, padding: "10px", border: "none", borderRadius: 10, fontSize: 14,
                background: mode === m ? "var(--surface)" : "transparent",
                color: mode === m ? "var(--text)" : "var(--text-muted)",
                fontFamily: "inherit", fontWeight: mode === m ? 700 : 500, cursor: "pointer",
                boxShadow: mode === m ? "var(--shadow-sm)" : "none",
                transition: "all 0.2s ease",
              }}>{m === "signup" ? "Sign Up" : "Log In"}</button>
            ))}
          </div>

          {mode === "signup" && (
            <div style={{ marginBottom: 16 }}>
              <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", color: "var(--text-muted)", marginBottom: 7 }}>Full Name</div>
              <input style={inp} value={name} onChange={e => setName(e.target.value)} placeholder="Your name" />
            </div>
          )}
          <div style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", color: "var(--text-muted)", marginBottom: 7 }}>Email</div>
            <input style={inp} type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="you@email.com" />
          </div>
          <div style={{ marginBottom: 20 }}>
            <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", color: "var(--text-muted)", marginBottom: 7 }}>Password</div>
            <input style={inp} type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="••••••••" onKeyDown={e => e.key === "Enter" && handle()} />
          </div>

          {error && <div style={{ background: "rgba(239,68,68,0.06)", border: "1px solid rgba(239,68,68,0.2)", borderRadius: 10, padding: "11px 14px", fontSize: 13, color: ER, marginBottom: 16 }}>{error}</div>}
          {verifyMsg && (
            <div style={{ background: "var(--accent-soft)", border: "1px solid var(--border)", borderRadius: 10, padding: "11px 14px", fontSize: 13, color: AC, marginBottom: 16 }}>
              <div style={{ lineHeight: 1.7 }}>{verifyMsg}</div>
              <div style={{ fontSize: 12, color: MUTED, marginTop: 6 }}>Tip: search for <strong>Rezolt</strong> in Spam or Promotions if the email is delayed.</div>
              <button onClick={resendVerification} disabled={loading} style={{ marginTop: 10, background: WHITE, color: AC, border: `1px solid ${BORDER}`, borderRadius: 9, padding: "8px 12px", fontSize: 12, fontWeight: 700, cursor: loading ? "not-allowed" : "pointer", fontFamily: "inherit" }}>
                {loading ? "Sending..." : resent ? "Confirmation email resent ✓" : "Resend confirmation email"}
              </button>
            </div>
          )}

          {mode === "signup" && !verifyMsg && (
            <div style={{ background: "rgba(34,197,94,0.06)", border: "1px solid rgba(34,197,94,0.2)", borderRadius: 10, padding: "11px 14px", fontSize: 13, color: "#15803D", marginBottom: 18, display: "flex", gap: 8, alignItems: "center" }}>
              <span>🎁</span> Sign up free and get 1 resume kit instantly
            </div>
          )}

          {mode === "signup" && !verifyMsg && (
            <div style={{ marginBottom: 16, display: "flex", alignItems: "flex-start", gap: 9 }}>
              <input type="checkbox" id="tc-check" checked={tcAccepted} onChange={e => setTcAccepted(e.target.checked)} style={{ marginTop: 3, accentColor: AC, cursor: "pointer", flexShrink: 0, width: 15, height: 15 }} />
              <label htmlFor="tc-check" style={{ fontSize: 12, color: MUTED, lineHeight: 1.65, cursor: "pointer" }}>
                I agree to Rezolt&apos;s{" "}
                <button onClick={() => setPage("terms")} style={{ background: "none", border: "none", color: AC, cursor: "pointer", fontFamily: "inherit", fontSize: 12, padding: 0, textDecoration: "underline" }}>Terms of Service</button>
                {" "}and{" "}
                <button onClick={() => setPage("privacy")} style={{ background: "none", border: "none", color: AC, cursor: "pointer", fontFamily: "inherit", fontSize: 12, padding: 0, textDecoration: "underline" }}>Privacy Policy</button>.
                {" "}I confirm my resume reflects accurate information.
              </label>
            </div>
          )}

          {!verifyMsg && (
            <button onClick={handle} disabled={loading} className="btn-primary" style={{
              width: "100%", background: loading ? "var(--surface2)" : "var(--grad)",
              color: loading ? "var(--text-muted)" : "white",
              border: "none", borderRadius: 12, padding: "14px", fontSize: 15, fontWeight: 700,
              cursor: loading ? "not-allowed" : "pointer", fontFamily: "inherit",
              boxShadow: loading ? "none" : "var(--shadow-accent)", transition: "all 0.2s",
            }}>
              {loading ? "Please wait..." : mode === "signup" ? "Create Account" : "Sign In"}
            </button>
          )}

          {mode === "login" && !verifyMsg && (
            <button onClick={() => setPage("forgot")} style={{ width: "100%", background: "none", border: "none", color: "var(--text-muted)", fontSize: 13, cursor: "pointer", marginTop: 12, fontFamily: "inherit", transition: "color 0.15s ease" }}
              onMouseEnter={e => e.target.style.color = AC}
              onMouseLeave={e => e.target.style.color = "var(--text-muted)"}
            >Forgot your password?</button>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────

function Dashboard({ user, history, setPage, onBuyCredits, profileLoaded = true }) {
  const [viewingKit, setViewingKit] = useState(null);
  const [activeTab, setActiveTab] = useState("resume");
  const [copied, setCopied] = useState("");
  const showResumeBranding = (user?.plan ?? "starter") === "starter";

  const copy = (id) => {
    navigator.clipboard.writeText(viewingKit.outputs?.[id] || "").catch(() => { });
    setCopied(id); setTimeout(() => setCopied(""), 2000);
  };

  const downloadPdf = async (id) => {
    const tab = TABS.find(t => t.id === id);
    const label = tab ? tab.label : "Output";
    const baseName = label.replace(/\s+/g, "-").toLowerCase();
    const fileName = id === "resume" && !showResumeBranding ? `${baseName}.pdf` : `Rezolt-${baseName}.pdf`;
    const element = document.querySelector(".output-pad");
    if (element) {
      try {
        await exportElementToPdf(element, fileName);
        return;
      } catch (err) {
        console.warn("PDF export failed:", err);
      }
    }
    const doc = new jsPDF();
    const text = (viewingKit.outputs?.[id] || "").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    const lines = doc.splitTextToSize(text, 180);
    doc.text(lines, 14, 20);
    doc.save(fileName);
  };

  const downloadDocx = async (id) => {
    const tab = TABS.find(t => t.id === id);
    const label = tab ? tab.label : "Output";
    const baseName = label.replace(/\s+/g, "-").toLowerCase();
    const fileName = id === "resume" && !showResumeBranding ? `${baseName}.docx` : `Rezolt-${baseName}.docx`;
    await exportTextToDocx(viewingKit.outputs?.[id] || "", fileName, `Rezolt ${label}`);
  };

  if (viewingKit) return (
    <div className="fade-in page-pad" style={{ maxWidth: "100%", padding: "40px clamp(16px,5vw,80px) 80px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 24 }}>
        <div style={{ textAlign: "left", alignItems: "flex-start" }}>
          <button onClick={() => setViewingKit(null)} style={{ background: "none", border: "none", color: MUTED, fontSize: 13, cursor: "pointer", padding: 0, marginBottom: 6, fontFamily: "inherit", display: "block", textAlign: "left" }}>Back to Dashboard</button>
          <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 28, fontWeight: 500, color: DARK, textAlign: "left" }}>Your Career Kit</div>
          <div style={{ fontSize: 13, color: MUTED, marginTop: 3, textAlign: "left" }}><i className="fa-solid fa-bullseye" style={{ marginRight: 6 }} />{viewingKit.role} · {viewingKit.date}</div>
        </div>
      </div>

      {/* Tab pills */}
      <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
        {TABS.filter(t => viewingKit.outputs?.[t.id]).map(t => {
          const active = activeTab === t.id;
          return (
            <button key={t.id} className="tab-pill" onClick={() => setActiveTab(t.id)} style={{ display: "flex", alignItems: "center", gap: 7, padding: "9px 16px", borderRadius: 10, border: `1.5px solid ${active ? AC : BORDER}`, background: active ? "var(--accent-soft)" : WHITE, color: active ? AC : MUTED, fontSize: 13, fontWeight: active ? 700 : 500, cursor: "pointer", fontFamily: "inherit", transition: "all 0.15s" }}>
              <span style={{ fontSize: 14, lineHeight: 1, display: "inline-flex", alignItems: "center" }}><i className={t.iconClass} /></span>{t.label}
              <div style={{ width: 7, height: 7, borderRadius: "50%", background: G }} />
            </button>
          );
        })}
      </div>

      {/* Output panel */}
      <div style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 16, overflow: "hidden", boxShadow: "0 1px 4px rgba(0,0,0,0.06)" }}>
        <div className="output-header" style={{ padding: "16px 24px", borderBottom: `1px solid ${BORDER}`, display: "flex", justifyContent: "space-between", alignItems: "center", background: BG, flexWrap: "wrap", gap: 10 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 18, lineHeight: 1, display: "inline-flex", alignItems: "center" }}><i className={TABS.find(t => t.id === activeTab)?.iconClass} /></span>
            <span style={{ fontSize: 15, fontWeight: 700, color: DARK }}>{TABS.find(t => t.id === activeTab)?.label}</span>
            <span style={{ background: "var(--accent-soft)", color: O, border: `1px solid ${BORDER}`, borderRadius: 20, padding: "2px 10px", fontSize: 11, fontWeight: 700 }}>Saved</span>
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="copy-btn" onClick={() => copy(activeTab)} style={{ background: copied === activeTab ? "rgba(34,197,94,0.08)" : BG, border: `1px solid ${BORDER}`, borderRadius: 7, padding: "7px 16px", fontSize: 12, fontWeight: 600, color: copied === activeTab ? "#15803D" : MUTED, cursor: "pointer", fontFamily: "inherit", transition: "all 0.15s" }}>
              {copied === activeTab ? "Copied!" : "Copy text"}
            </button>
            <button onClick={() => downloadDocx(activeTab)} className="copy-btn" style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 7, padding: "7px 16px", fontSize: 12, fontWeight: 600, color: O, cursor: "pointer", fontFamily: "inherit" }}>
              Download DOCX
            </button>
            <button onClick={() => downloadPdf(activeTab)} className="copy-btn" style={{ background: "var(--grad)", border: "none", borderRadius: 7, padding: "7px 16px", fontSize: 12, fontWeight: 600, color: "white", cursor: "pointer", fontFamily: "inherit" }}>
              Download PDF
            </button>
          </div>
        </div>
        <div className="output-pad" style={{ padding: "28px 32px 40px" }}>
          {viewingKit.outputs?.[activeTab] ? (
            <>
              {(activeTab === "referral" || activeTab === "reach") && (
                <OutreachQuickActions type={activeTab} text={viewingKit.outputs[activeTab]} />
              )}
              <div className={(activeTab === "resume" || activeTab === "cover") ? "paper-preview" : ""} style={{ textAlign: "left" }}>
                {activeTab === "resume"
                  ? renderResumeWithTemplate(viewingKit.outputs[activeTab], viewingKit.template || "creative", null, showResumeBranding)
                  : renderOutput(viewingKit.outputs[activeTab])}
              </div>
            </>
          ) : (
            <div style={{ textAlign: "center", padding: "40px 0", color: FAINT, fontSize: 14 }}>Nothing is saved in this tab yet.</div>
          )}
        </div>
      </div>

      {/* Mobile sticky action bar */}
      {viewingKit.outputs?.[activeTab] && (
        <div className="mobile-only" style={{ position: "fixed", bottom: 65, left: 0, right: 0, background: "var(--surface)", borderTop: "1px solid var(--border)", padding: "10px 16px", display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8, zIndex: 199, boxShadow: "0 -2px 12px rgba(0,27,72,.08)" }}>
          <button onClick={() => copy(activeTab)} style={{ background: copied === activeTab ? "rgba(34,197,94,.1)" : "var(--surface2)", border: `1px solid ${copied === activeTab ? G : "var(--border)"}`, borderRadius: 10, padding: "11px 8px", fontSize: 12, fontWeight: 700, color: copied === activeTab ? G : "var(--text)", cursor: "pointer", fontFamily: "inherit" }}>
            {copied === activeTab ? "Copied" : "Copy"}
          </button>
          <button onClick={() => downloadDocx(activeTab)} style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 10, padding: "11px 8px", fontSize: 12, fontWeight: 700, color: O, cursor: "pointer", fontFamily: "inherit" }}>
            DOCX
          </button>
          <button onClick={() => downloadPdf(activeTab)} style={{ background: "var(--grad)", border: "none", borderRadius: 10, padding: "11px 8px", fontSize: 12, fontWeight: 700, color: "white", cursor: "pointer", fontFamily: "inherit" }}>
            PDF
          </button>
        </div>
      )}
    </div>
  );

  return (
    <div className="fade-in page-pad" style={{ maxWidth: "100%", padding: "44px clamp(16px,5vw,80px) 80px" }}>
      <div style={{ marginBottom: 36 }}>
        <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 34, fontWeight: 400, color: "var(--text)", marginBottom: 4 }}>
          Welcome back, {user?.name?.split(" ")[0] || "User"}<span style={{ color: AC }}>.</span>
        </div>
        <div style={{ fontSize: 14, color: "var(--text-muted)" }}>Your human-first career dashboard</div>
      </div>

      <div className="dash-stats" style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 18, marginBottom: 36 }}>
        {[
          { label: "Credits Left", value: !profileLoaded ? "—" : user?.plan === "unlimited" ? "∞" : user?.credits, color: O, sub: !profileLoaded ? "loading..." : user?.plan === "unlimited" ? "unlimited plan" : user?.plan === "Free" ? "free resume available" : "Career Kits available" },
          { label: "Kits Generated", value: !profileLoaded ? "—" : history.length, color: DARK, sub: "total" },
          { label: "Status", value: !profileLoaded ? "—" : (user?.credits ?? 0) > 0 || user?.plan === "unlimited" ? "Active" : "No Credits", color: !profileLoaded ? FAINT : (user?.credits ?? 0) > 0 || user?.plan === "unlimited" ? G : "#EF4444", sub: !profileLoaded ? "loading..." : (user?.credits ?? 0) > 0 || user?.plan === "unlimited" ? "ready to apply" : "buy credits to continue" },
        ].map(s => (
          <div key={s.label} style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 24, padding: "24px 26px", boxShadow: "var(--soft-shadow)" }}>
            <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", color: FAINT, marginBottom: 10 }}>{s.label}</div>
            <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 36, fontWeight: 500, color: s.color, marginBottom: 4, lineHeight: 1 }}>{s.value}</div>
            <div style={{ fontSize: 12, color: FAINT }}>{s.sub}</div>
          </div>
        ))}
      </div>

      <div className="dash-actions" style={{ display: "flex", gap: 14, marginBottom: 40 }}>
        <button onClick={() => setPage("generate")} style={{ background: O, color: WHITE, border: "none", borderRadius: 16, padding: "14px 30px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", boxShadow: "var(--soft-shadow)" }}>
          Start a New Kit
        </button>
        <button onClick={() => setPage("payment")} style={{ background: WHITE, color: O, border: `1px solid ${O}`, borderRadius: 16, padding: "13px 24px", fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: "inherit", boxShadow: "var(--soft-shadow)" }}>
          + Buy Credits
        </button>
      </div>

      <div>
        <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", color: FAINT, marginBottom: 14 }}>Recent Kits</div>
        {history.length === 0 ? (
          <div style={{ background: BG, border: `1px dashed ${BORDER}`, borderRadius: 24, padding: "56px 36px", textAlign: "center", boxShadow: "var(--soft-shadow)" }}>
            <div style={{ fontSize: 34, marginBottom: 14 }}><i className="fa-solid fa-wand-magic-sparkles" /></div>
            <div style={{ color: MUTED, fontSize: 15, marginBottom: 20 }}>Ready to land your dream role? Let’s build your first kit together.</div>
            <button onClick={() => setPage("generate")} style={{ background: O, color: WHITE, border: "none", borderRadius: 16, padding: "12px 26px", fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: "inherit", boxShadow: "var(--soft-shadow)" }}>Let’s build the first kit</button>
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {history.map((kit, i) => (
              <div key={i} className="card-hover" onClick={() => { setViewingKit(kit); setActiveTab(Object.keys(kit.outputs || {})[0] || "resume"); }} style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 22, padding: "18px 22px", display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer", transition: "all 0.15s", boxShadow: "var(--soft-shadow)" }}>
                <div>
                  <div style={{ fontSize: 14, fontWeight: 600, color: DARK, marginBottom: 3 }}><i className="fa-solid fa-bullseye" style={{ marginRight: 6 }} />{kit.role}</div>
                  <div style={{ fontSize: 12, color: FAINT }}>{kit.date} · {Object.keys(kit.outputs || {}).length} outputs saved</div>
                </div>
                <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
                  {TABS.filter(t => t.id !== "negotiate").map(t => <div key={t.id} style={{ width: 7, height: 7, borderRadius: "50%", background: kit.outputs?.[t.id] ? G : BORDER }} />)}
                  <span style={{ color: O, fontSize: 13, marginLeft: 8, fontWeight: 600 }}>View</span>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── NEGOTIATE TAB ────────────────────────────────────────────────────────────

function NegotiateTab() {
  const [tool, setTool] = useState("salary");
  const [sForm, setSForm] = useState({ offerCtc: "", expectedCtc: "", role: "", company: "", exp: "", skills: "" });
  const [nForm, setNForm] = useState({ notice: "", role: "", company: "", currentCompany: "", reason: "", buyout: "" });
  const [sOut, setSOut] = useState(""); const [nOut, setNOut] = useState("");
  const [sLoading, setSLoading] = useState(false); const [nLoading, setNLoading] = useState(false);
  const [sCopied, setSCopied] = useState(false); const [nCopied, setNCopied] = useState(false);
  const [sSection, setSSection] = useState("opener"); const [nSection, setNSection] = useState("verbal");

  const sF = (k, v) => setSForm(p => ({ ...p, [k]: v }));
  const nF = (k, v) => setNForm(p => ({ ...p, [k]: v }));
  const sSections = [{ id: "opener", label: "Opening" }, { id: "counter", label: "Counter offer" }, { id: "pushback", label: "If they push back" }, { id: "accept", label: "If you accept" }];
  const nSections = [{ id: "verbal", label: "Verbal" }, { id: "email", label: "Email" }, { id: "linkedin", label: "LinkedIn" }];

  const genSalary = async () => {
    if (!sForm.offerCtc || !sForm.expectedCtc) return;
    setSLoading(true); setSOut("");
    try {
      const res = await fetchWithAuth("/api/claude-generate", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ prompt: `You are a salary negotiation coach for the Indian job market. Plain text only, no markdown, no em dashes.\n\nOffered CTC: ${sForm.offerCtc} LPA\nExpected CTC: ${sForm.expectedCtc} LPA\nRole: ${sForm.role || "not specified"}\nCompany: ${sForm.company || "not specified"}\nExperience: ${sForm.exp || "not specified"}\nKey Skills: ${sForm.skills || "not specified"}\n\nGenerate 4 sections:\n\nSECTION 1: OPENING SCRIPT\n4-6 lines. Grateful, confident, sets up negotiation.\n\nSECTION 2: COUNTER OFFER\n4-6 lines. Specific ask with justification from skills and experience.\n\nSECTION 3: IF THEY PUSH BACK\n4-5 lines. Handle "budget is fixed." Ask for joining bonus, early appraisal, or extra leave.\n\nSECTION 4: IF YOU ACCEPT\n3-4 lines. Graceful close maintaining goodwill.` }) });
      const d = await res.json();
      if (!res.ok) throw new Error(d.error || "AI generation failed");
      setSOut(d.text || d.content?.[0]?.text || "Error. Try again.");
    } catch (err) { setSOut(err?.message || "Something went wrong."); }
    setSLoading(false);
  };

  const genNotice = async () => {
    if (!nForm.notice) return;
    setNLoading(true); setNOut("");
    try {
      const res = await fetchWithAuth("/api/claude-generate", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ prompt: `You are a career coach for Indian professionals. Plain text only, no markdown, no em dashes. Always frame the notice period positively.\n\nNotice Period: ${nForm.notice}\nTarget Role: ${nForm.role || "not specified"}\nTarget Company: ${nForm.company || "not specified"}\nCurrent Company: ${nForm.currentCompany || "not specified"}\nReason for leaving: ${nForm.reason || "not specified"}\nBuyout option: ${nForm.buyout || "unsure"}\n\nGenerate 3 sections:\n\nSECTION 1: VERBAL SCRIPT\n5-7 lines spoken answer for "What is your notice period?" Confident framing, mention buyout if applicable.\n\nSECTION 2: EMAIL TO HR\nSubject: [subject line]\n8-10 line professional email to new company HR communicating notice period and any flexibility.\n\nSECTION 3: LINKEDIN MESSAGE\n6-8 line message to recruiter after interview, communicating notice period warmly and positively.` }) });
      const d = await res.json();
      if (!res.ok) throw new Error(d.error || "AI generation failed");
      setNOut(d.text || d.content?.[0]?.text || "Error. Try again.");
    } catch (err) { setNOut(err?.message || "Something went wrong."); }
    setNLoading(false);
  };

  const extractSection = (text, id, type) => {
    if (!text) return "";
    const map = type === "salary" ? {
      opener: /SECTION 1[:\s]+OPENING SCRIPT\n([\s\S]*?)(?=SECTION 2|$)/i,
      counter: /SECTION 2[:\s]+COUNTER OFFER\n([\s\S]*?)(?=SECTION 3|$)/i,
      pushback: /SECTION 3[:\s]+IF THEY PUSH BACK\n([\s\S]*?)(?=SECTION 4|$)/i,
      accept: /SECTION 4[:\s]+IF YOU ACCEPT\n([\s\S]*?)$/i,
    } : {
      verbal: /SECTION 1[:\s]+VERBAL SCRIPT\n([\s\S]*?)(?=SECTION 2|$)/i,
      email: /SECTION 2[:\s]+EMAIL TO HR\n([\s\S]*?)(?=SECTION 3|$)/i,
      linkedin: /SECTION 3[:\s]+LINKEDIN MESSAGE\n([\s\S]*?)$/i,
    };
    const m = text.match(map[id]); return m ? m[1].trim() : text;
  };

  const gap = () => {
    const o = parseFloat(sForm.offerCtc), e = parseFloat(sForm.expectedCtc);
    if (!o || !e || e <= o) return null;
    return { lpa: (e - o).toFixed(1), pct: ((e - o) / o * 100).toFixed(1) };
  };
  const g = gap();

  const tinp = { fontFamily: "inherit", color: DARK, background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 16, padding: "14px 16px", fontSize: 13, width: "100%", outline: "none" };

  return (
    <div>
      {/* Tool switcher */}
      <div style={{ display: "flex", gap: 0, borderBottom: `1px solid ${BORDER}`, marginBottom: 24 }}>
        {[{ id: "salary", label: "Salary Negotiation" }, { id: "notice", label: "Notice Period" }].map(t => (
          <button key={t.id} onClick={() => setTool(t.id)} style={{ padding: "10px 20px", border: "none", background: "transparent", borderBottom: `2.5px solid ${tool === t.id ? O : "transparent"}`, marginBottom: -1, color: tool === t.id ? O : MUTED, fontSize: 13, fontWeight: tool === t.id ? 700 : 400, cursor: "pointer", fontFamily: "inherit", transition: "all .2s" }}>{t.label}</button>
        ))}
      </div>

      {tool === "salary" ? (
        <div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr auto 1fr", gap: 12, alignItems: "end", marginBottom: 14 }}>
            <div><label style={{ display: "block", fontSize: 11, fontWeight: 700, color: MUTED, marginBottom: 5 }}>What offer did they share?*</label><input type="number" value={sForm.offerCtc} onChange={e => sF("offerCtc", e.target.value)} placeholder="e.g. 10" style={tinp} /></div>
            <div style={{ paddingBottom: 10, color: MUTED, fontSize: 14, textAlign: "center", fontWeight: 600 }}>to</div>
            <div><label style={{ display: "block", fontSize: 11, fontWeight: 700, color: MUTED, marginBottom: 5 }}>What would feel right?*</label><input type="number" value={sForm.expectedCtc} onChange={e => sF("expectedCtc", e.target.value)} placeholder="e.g. 13" style={tinp} /></div>
          </div>
          {g && <div style={{ background: "#FFF7ED", border: "1px solid #FED7AA", borderRadius: 8, padding: "9px 14px", marginBottom: 14, display: "flex", justifyContent: "space-between", alignItems: "center" }}><span style={{ fontSize: 13, color: "#92400E" }}>Gap from offer</span><div style={{ display: "flex", gap: 10 }}><span style={{ fontWeight: 700, color: O }}>+{g.lpa} LPA</span><span style={{ background: "#FED7AA", color: "#92400E", fontSize: 11, fontWeight: 600, padding: "2px 8px", borderRadius: 20 }}>{g.pct}% above offer</span></div></div>}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
            <div><label style={{ display: "block", fontSize: 11, fontWeight: 700, color: MUTED, marginBottom: 5 }}>Which role is this for?</label><input value={sForm.role} onChange={e => sF("role", e.target.value)} placeholder="e.g. HRBP Manager" style={tinp} /></div>
            <div><label style={{ display: "block", fontSize: 11, fontWeight: 700, color: MUTED, marginBottom: 5 }}>Which company is it with?</label><input value={sForm.company} onChange={e => sF("company", e.target.value)} placeholder="e.g. Target Company" style={tinp} /></div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 18 }}>
            <div><label style={{ display: "block", fontSize: 11, fontWeight: 700, color: MUTED, marginBottom: 5 }}>How much experience are you bringing?</label>
              <select value={sForm.exp} onChange={e => sF("exp", e.target.value)} style={{ ...tinp, appearance: "none", cursor: "pointer" }}>
                <option value="">Select</option><option>1-3 years</option><option>3-5 years</option><option>5-8 years</option><option>8-12 years</option><option>12+ years</option>
              </select>
            </div>
            <div><label style={{ display: "block", fontSize: 11, fontWeight: 700, color: MUTED, marginBottom: 5 }}>What strengths should we highlight?</label><input value={sForm.skills} onChange={e => sF("skills", e.target.value)} placeholder="e.g. Power BI, SQL" style={tinp} /></div>
          </div>
          <button onClick={genSalary} disabled={!sForm.offerCtc || !sForm.expectedCtc || sLoading} style={{ width: "100%", padding: "14px", background: (!sForm.offerCtc || !sForm.expectedCtc || sLoading) ? "#E2E8F0" : O, color: "white", border: "none", borderRadius: 16, fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", transition: "all .2s", boxShadow: "var(--soft-shadow)" }}>
            {sLoading ? "Refining your tone..." : "Craft my negotiation plan"}
          </button>
          {sOut && (
            <div style={{ marginTop: 20 }}>
              <div style={{ display: "flex", borderBottom: `1px solid ${BORDER}`, marginBottom: 0 }}>
                {sSections.map(s => <button key={s.id} onClick={() => setSSection(s.id)} style={{ padding: "8px 14px", border: "none", background: "transparent", borderBottom: `2.5px solid ${sSection === s.id ? O : "transparent"}`, marginBottom: -1, color: sSection === s.id ? O : MUTED, fontSize: 12, fontWeight: sSection === s.id ? 700 : 400, cursor: "pointer", fontFamily: "inherit", whiteSpace: "nowrap" }}>{s.label}</button>)}
              </div>
              <div style={{ background: BG, border: `1px solid ${BORDER}`, borderTop: "none", borderRadius: "0 0 10px 10px", padding: "16px" }}>
                <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 8 }}>
                  <button onClick={() => { navigator.clipboard.writeText(extractSection(sOut, sSection, "salary")); setSCopied(true); setTimeout(() => setSCopied(false), 2000); }} style={{ padding: "5px 12px", border: `1px solid ${sCopied ? G : BORDER}`, borderRadius: 7, background: sCopied ? "#F0FDF4" : WHITE, color: sCopied ? G : MUTED, fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>{sCopied ? "Copied!" : "Copy"}</button>
                </div>
                <div style={{ fontSize: 13, lineHeight: 1.85, color: DARK, whiteSpace: "pre-wrap", fontFamily: "Georgia, serif" }}>{extractSection(sOut, sSection, "salary")}</div>
              </div>
            </div>
          )}
        </div>
      ) : (
        <div>
          <div style={{ marginBottom: 14 }}>
            <label style={{ display: "block", fontSize: 11, fontWeight: 700, color: MUTED, marginBottom: 5 }}>What notice period are you working with?*</label>
            <select value={nForm.notice} onChange={e => nF("notice", e.target.value)} style={{ ...tinp, appearance: "none", cursor: "pointer", fontSize: 14 }}>
              <option value="">Select notice period</option>
              <option>Immediate joiner</option>
              <option>15 days</option>
              <option>30 days</option>
              <option>60 days</option>
              <option>90 days</option>
            </select>
            <div style={{ fontSize: 11, color: FAINT, marginTop: 6 }}>Common in India: 30, 60, and 90 days — we’ll help you frame any option positively.</div>
          </div>
          {nForm.notice && <div style={{ background: nForm.notice === "Immediate joiner" ? "#F0FDF4" : nForm.notice === "90 days" ? "#FFFBEB" : "#EFF6FF", border: `1px solid ${nForm.notice === "Immediate joiner" ? "#BBF7D0" : nForm.notice === "90 days" ? "#FDE68A" : "#BFDBFE"}`, borderRadius: 8, padding: "9px 14px", marginBottom: 14, fontSize: 13, color: DARK }}>{nForm.notice === "Immediate joiner" ? "🟢 Immediate joiners are in high demand — we'll frame this as a major strength." : nForm.notice === "90 days" ? "🟡 90 days is long — we'll coach you on framing and explore buyout angles." : nForm.notice === "60 days" ? "🔵 60 days is common in India — we'll help you highlight flexibility and a smooth handover." : `🔵 ${nForm.notice} is standard in India. We'll help you present it with confidence.`}</div>}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
            <div><label style={{ display: "block", fontSize: 11, fontWeight: 700, color: MUTED, marginBottom: 5 }}>Which role are you targeting?</label><input value={nForm.role} onChange={e => nF("role", e.target.value)} placeholder="e.g. HRBP Manager" style={tinp} /></div>
            <div><label style={{ display: "block", fontSize: 11, fontWeight: 700, color: MUTED, marginBottom: 5 }}>Which company is on your mind?</label><input value={nForm.company} onChange={e => nF("company", e.target.value)} placeholder="e.g. Target Company" style={tinp} /></div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
            <div><label style={{ display: "block", fontSize: 11, fontWeight: 700, color: MUTED, marginBottom: 5 }}>Where are you working right now?</label><input value={nForm.currentCompany} onChange={e => nF("currentCompany", e.target.value)} placeholder="e.g. Current Company" style={tinp} /></div>
            <div><label style={{ display: "block", fontSize: 11, fontWeight: 700, color: MUTED, marginBottom: 5 }}>Is an early release possible?</label>
              <select value={nForm.buyout} onChange={e => nF("buyout", e.target.value)} style={{ ...tinp, appearance: "none", cursor: "pointer" }}>
                <option value="">Not sure</option><option>Yes, full buyout possible</option><option>Yes, partial buyout possible</option><option>No buyout option</option>
              </select>
            </div>
          </div>
          <div style={{ marginBottom: 18 }}>
            <label style={{ display: "block", fontSize: 11, fontWeight: 700, color: MUTED, marginBottom: 5 }}>What’s motivating the move?</label>
            <input value={nForm.reason} onChange={e => nF("reason", e.target.value)} placeholder="e.g. career growth, better domain fit" style={tinp} />
          </div>
          <button onClick={genNotice} disabled={!nForm.notice || nLoading} style={{ width: "100%", padding: "14px", background: (!nForm.notice || nLoading) ? "#E2E8F0" : O, color: "white", border: "none", borderRadius: 16, fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", transition: "all .2s", boxShadow: "var(--soft-shadow)" }}>
            {nLoading ? "Organizing your response..." : "Craft my notice period plan"}
          </button>
          {nOut && (
            <div style={{ marginTop: 20 }}>
              <div style={{ display: "flex", borderBottom: `1px solid ${BORDER}` }}>
                {nSections.map(s => <button key={s.id} onClick={() => setNSection(s.id)} style={{ padding: "8px 14px", border: "none", background: "transparent", borderBottom: `2.5px solid ${nSection === s.id ? O : "transparent"}`, marginBottom: -1, color: nSection === s.id ? O : MUTED, fontSize: 12, fontWeight: nSection === s.id ? 700 : 400, cursor: "pointer", fontFamily: "inherit", whiteSpace: "nowrap" }}>{s.label}</button>)}
              </div>
              <div style={{ background: BG, border: `1px solid ${BORDER}`, borderTop: "none", borderRadius: "0 0 10px 10px", padding: "16px" }}>
                <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 8 }}>
                  <button onClick={() => { navigator.clipboard.writeText(extractSection(nOut, nSection, "notice")); setNCopied(true); setTimeout(() => setNCopied(false), 2000); }} style={{ padding: "5px 12px", border: `1px solid ${nCopied ? G : BORDER}`, borderRadius: 7, background: nCopied ? "#F0FDF4" : WHITE, color: nCopied ? G : MUTED, fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>{nCopied ? "Copied!" : "Copy"}</button>
                </div>
                <div style={{ fontSize: 13, lineHeight: 1.85, color: DARK, whiteSpace: "pre-wrap", fontFamily: "Georgia, serif" }}>{extractSection(nOut, nSection, "notice")}</div>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ─── KIT GENERATOR ────────────────────────────────────────────────────────────

function KitGenerator({ user, setUser, onSaveKit, onUseCredit, setPage, selectedTemplate, setSelectedTemplate, sessionToken }) {
  const [resume, setResume] = useState("");
  const [jd, setJd] = useState("");
  const [outputs, setOutputs] = useState({});
  const [loading, setLoading] = useState({});
  const [activeTab, setActiveTab] = useState("resume");
  const [step, setStep] = useState("input");
  const [copied, setCopied] = useState("");
  const [uploading, setUploading] = useState(false);
  const [uploadedFileName, setUploadedFileName] = useState("");
  const [inputMode, setInputMode] = useState("upload");
  const [currentRole, setCurrentRole] = useState("");
  const [generating, setGenerating] = useState(false);
  const [loaderStage, setLoaderStage] = useState(0);
  const [pasteStatus, setPasteStatus] = useState({ resume: "", jd: "" });
  const [showSample, setShowSample] = useState(false);
  const [sampleTab, setSampleTab] = useState("resume");
  const [kitFeedback, setKitFeedback] = useState(null);
  const [subStep, setSubStep] = useState("");

  const sampleJdText = `Role: HRBP Manager
Company: Leading pharma analytics firm
Location: Bangalore
Responsibilities: Partner with business leaders, improve hiring velocity, build dashboards for workforce insights, manage campus engagement, and support stakeholder communication.
Must have: 4+ years in talent acquisition or HRBP, strong Excel/Power BI exposure, stakeholder management, and experience hiring for analytics or commercial roles.`;

  const loaderStages = [
    { title: "Parsing your resume", note: "Reviewing your experience, wins, and positioning." },
    { title: "Matching the job description", note: "Aligning keywords, tone, and relevance to the role." },
    { title: "Drafting your career kit", note: "Writing tailored resume, outreach, and interview assets." },
    { title: "Finalising the polish", note: "Cleaning clarity, structure, and recruiter-ready details." },
  ];

  const isReady = resume.trim().length > 50 && jd.trim().length > 50;
  const hasCredits = user?.plan === "unlimited" || user?.credits > 0;
  const anyLoading = Object.values(loading).some(Boolean);
  const doneCount = Object.keys(outputs).length;
  const rawProgress = Math.round((doneCount / Math.max(TABS.filter(t => canAccess(user?.plan, t.minPlan)).length, 1)) * 100);
  const progressPct = anyLoading ? Math.max(rawProgress, Math.min(92, 14 + loaderStage * 18)) : rawProgress;
  const activeLoader = loaderStages[Math.min(loaderStage, loaderStages.length - 1)];
  const showResumeBranding = (user?.plan ?? "starter") === "starter";

  const resumeTrimmed = resume.trim();
  const resumeChecklist = [
    { label: "Experience section", ok: /(experience|employment|work history)/i.test(resume) },
    { label: "Skills or tools", ok: /(skills|tools|competencies|expertise)/i.test(resume) },
    { label: "Education details", ok: /(education|degree|university|college|mba|b\.tech|btech|graduation)/i.test(resume) },
    { label: "Measurable wins", ok: /(\d+\+|\d+%|₹\s?\d+|hired\s+\d+|reduced|improved|increased|saved)/i.test(resume) },
  ];
  const resumeScore = resumeChecklist.filter(item => item.ok).length;
  const resumeQualityLabel = !resumeTrimmed
    ? "Add your resume to get an instant quality check"
    : resumeTrimmed.length < 120
      ? "Too short for strong tailoring"
      : resumeScore >= 4
        ? "Strong base for a tailored kit"
        : resumeScore >= 2
          ? "Good start — a few additions will help"
          : "Add more detail before generating";
  const resumeQualityColor = !resumeTrimmed ? FAINT : resumeScore >= 4 ? "#15803D" : resumeScore >= 2 ? O : ER;

  const accessibleTabs = TABS.filter(t => canAccess(user?.plan, t.minPlan));
  const generatableTabs = accessibleTabs.filter(t => PROMPTS[t.id]);
  const loadingTone = {
    resume: "Curating your experience...",
    cover: "Refining your tone...",
    referral: "Personalising your outreach...",
    interview: "Organising your stories...",
    reach: "Planning your next steps...",
    negotiate: "Preparing your negotiation notes...",
  };

  useEffect(() => {
    if (!generating) {
      setLoaderStage(0);
      return;
    }
    const timer = window.setInterval(() => {
      setLoaderStage(prev => Math.min(prev + 1, loaderStages.length - 1));
    }, 1800);
    return () => window.clearInterval(timer);
  }, [generating]);

  const generate = async () => {
    if (typeof window !== "undefined" && (window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1")) {
      alert("Note: AI generation requires the live server. Please test this on www.rezolt.in");
    }
    if (!isReady || !hasCredits || generating) return;
    setLoaderStage(0);
    setGenerating(true);
    setSubStep("1/3: Authenticating...");
    try {
      try { await Promise.race([supabase.auth.refreshSession(), new Promise((_, r) => setTimeout(r, 1500))]); } catch (e) { console.warn("Session refresh failed:", e); }

      let freshProfile = null;
      try {
        const profileRes = await Promise.race([
          fetchWithAuth("/api/get-profile", { method: "GET" }, sessionToken),
          new Promise((_, r) => setTimeout(() => r(new Error("timeout")), 4000)),
        ]);
        const profileData = await profileRes.json().catch(() => ({}));
        freshProfile = profileData?.profile ?? null;
      } catch (e) {
        console.warn("Profile fetch in generate() failed:", e);
      }
      const effectivePlan = normalizePlan(freshProfile?.plan) ?? user?.plan ?? "Free";
      if (freshProfile) setUser(prev => ({
        ...prev,
        credits: freshProfile.credits ?? prev.credits,
        plan: effectivePlan,
      }));
      const freshGeneratableTabs = TABS.filter(t => canAccess(effectivePlan, t.minPlan) && PROMPTS[t.id]);

      const safeResume = prepareInputForAi(resume, MAX_RESUME_CHARS, "Resume");
      const safeJd = prepareInputForAi(jd, MAX_JD_CHARS, "Job description");
      if (safeResume.trimmed || safeJd.trimmed) {
        setPasteStatus(prev => ({
          ...prev,
          resume: safeResume.trimmed ? "Large resume detected — Rezolt is using the most relevant text automatically." : prev.resume,
          jd: safeJd.trimmed ? "Large JD detected — Rezolt is using the key requirements automatically." : prev.jd,
        }));
      }

      const extractRoleLabel = (jdText) => {
        const lines = jdText.split("\n").map(l => l.trim()).filter(Boolean);
        const companyPatterns = [
          /(?:company|organisation|employer|hiring|at|with|for)[:\s]+([A-Z][A-Za-z\s&.,-]{2,40})/i,
          /([A-Z][A-Za-z\s&.]+(?:Pvt\.? Ltd\.?|Inc\.?|Corp\.?|Limited|India|Solutions|Technologies|Pharma|Analytics))/,
        ];
        const rolePatterns = [
          /(?:role|position|title|opening|vacancy)[:\s]+([A-Za-z\s&\/,-]{4,50})/i,
          /(?:we are hiring|looking for|seeking)[:\s]+(?:a\s+|an\s+)?([A-Za-z\s&\/,-]{4,50})/i,
        ];
        let role = "", company = "";
        for (const line of lines.slice(0, 15)) {
          if (!role) for (const p of rolePatterns) { const m = line.match(p); if (m) { role = m[1].trim().slice(0, 40); break; } }
          if (!company) for (const p of companyPatterns) { const m = line.match(p); if (m) { company = m[1].trim().slice(0, 30); break; } }
          if (role && company) break;
        }
        if (!role) role = lines[0]?.slice(0, 50) || "Target Role";
        return company ? `${role} · ${company}` : role;
      };

      const role = extractRoleLabel(safeJd.text || jd);
      setCurrentRole(role);
      setStep("output");
      setOutputs({});
      setActiveTab("resume");
      setLoading(Object.fromEntries(freshGeneratableTabs.map(t => [t.id, true])));

      setSubStep("2/3: Connecting to AI...");
      const results = {};
      await Promise.all(freshGeneratableTabs.map(async (tab) => {
        try {
          const res = await fetchWithAuth("/api/claude-generate", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ prompt: PROMPTS[tab.id](safeResume.text, safeJd.text, selectedTemplate) }),
          }, sessionToken);
          const data = await res.json().catch(() => ({}));
          if (!res.ok) {
            const e = new Error(data.error || "Generation failed");
            e.diagnostic = data.diagnostic;
            throw e;
          }
          results[tab.id] = data.text || "Something went wrong. Please try again.";
        } catch (err) {
          const msg = err?.message || "";
          const isOutage = /overloaded|capacity|service.*unavailable|too many requests/i.test(msg);
          results[tab.id] = isOutage
            ? "Claude AI is temporarily busy — please try again in a few minutes. Your credits have not been used."
            : /sign in|session/i.test(msg)
              ? "Your session expired. Please sign in again."
              : "Network error. Please try again.";
        }
        setOutputs(prev => ({ ...prev, [tab.id]: results[tab.id] }));
        setLoading(prev => ({ ...prev, [tab.id]: false }));
      }));

      setSubStep("3/3: Saving results...");
      const successfulCount = Object.values(results).filter(isUsableGeneration).length;
      if (successfulCount > 0) {
        if (user?.plan !== "unlimited") {
          try {
            await onUseCredit();
          } catch (err) {
            console.warn("Credit update failed:", err);
          }
        }
        onSaveKit({ role, outputs: results, date: new Date().toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" }) });
      }
    } finally {
      setGenerating(false);
    }
  };

  const copy = (id) => {
    navigator.clipboard.writeText(outputs[id] || "").catch(() => { });
    setCopied(id); setTimeout(() => setCopied(""), 2000);
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setUploading(true);
    setUploadedFileName(file.name);
    try {
      if (file.type === "application/pdf") {
        const reader = new FileReader();
        reader.onload = async (ev) => {
          const typedArray = new Uint8Array(ev.target.result);
          const pdfjsLib = window["pdfjs-dist/build/pdf"];
          if (pdfjsLib) {
            pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
            const pdf = await pdfjsLib.getDocument({ data: typedArray }).promise;
            let text = "";
            for (let i = 1; i <= pdf.numPages; i++) {
              const page = await pdf.getPage(i);
              const content = await page.getTextContent();
              text += content.items.map(item => item.str).join(" ") + "\n";
            }
            setResume(text.trim());
          }
          setUploading(false);
        };
        reader.readAsArrayBuffer(file);
      } else if (file.name.endsWith(".docx")) {
        const reader = new FileReader();
        reader.onload = async (ev) => {
          const mammoth = window.mammoth;
          if (mammoth) {
            const result = await mammoth.extractRawText({ arrayBuffer: ev.target.result });
            setResume(result.value.trim());
          }
          setUploading(false);
        };
        reader.readAsArrayBuffer(file);
      } else {
        const text = await file.text();
        setResume(text.trim());
        setUploading(false);
      }
    } catch (err) {
      console.error(err);
      setUploading(false);
    }
  };

  const downloadPdf = async (id) => {
    const tab = TABS.find(t => t.id === id);
    const label = tab ? tab.label : "Output";
    const baseName = label.replace(/\s+/g, "-").toLowerCase();
    const fileName = id === "resume" && !showResumeBranding ? `${baseName}.pdf` : `Rezolt-${baseName}.pdf`;
    const element = document.querySelector(".output-pad");
    if (element) {
      try {
        await exportElementToPdf(element, fileName);
        return;
      } catch (err) {
        console.warn("PDF export failed:", err);
      }
    }
    const doc = new jsPDF();
    const text = (outputs[id] || "").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    const lines = doc.splitTextToSize(text, 180);
    doc.text(lines, 14, 20);
    doc.save(fileName);
  };

  const downloadDocx = async (id) => {
    const tab = TABS.find(t => t.id === id);
    const label = tab ? tab.label : "Output";
    const baseName = label.replace(/\s+/g, "-").toLowerCase();
    const fileName = id === "resume" && !showResumeBranding ? `${baseName}.docx` : `Rezolt-${baseName}.docx`;
    await exportTextToDocx(outputs[id] || "", fileName, `Rezolt ${label}`);
  };

  const regenerateTab = async (tabId) => {
    if (loading[tabId] || generating) return;
    setLoading(prev => ({ ...prev, [tabId]: true }));
    try {
      try { await Promise.race([supabase.auth.refreshSession(), new Promise((_, r) => setTimeout(r, 1500))]); } catch (e) { console.warn(e); }
      const safeResume = prepareInputForAi(resume, MAX_RESUME_CHARS, "Resume");
      const safeJd = prepareInputForAi(jd, MAX_JD_CHARS, "Job description");
      const res = await fetchWithAuth("/api/claude-generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ prompt: PROMPTS[tabId](safeResume.text, safeJd.text, selectedTemplate) }),
      }, sessionToken);
      const data = await res.json().catch(() => ({}));
      if (!res.ok) {
        const e = new Error(data.error || "Generation failed");
        e.diagnostic = data.diagnostic;
        throw e;
      }
      setOutputs(prev => ({ ...prev, [tabId]: data.text || "Something went wrong. Please try again." }));
    } catch (err) {
      const msg = err?.message || "";
      const isOutage = /overloaded|capacity|service.*unavailable|too many requests/i.test(msg);
      setOutputs(prev => ({ ...prev, [tabId]: isOutage ? "Claude AI is temporarily busy — please try again in a few minutes." : "Couldn't regenerate. Please try again." }));
    }
    setLoading(prev => ({ ...prev, [tabId]: false }));
  };

  const handleClipboardPaste = async (target) => {
    try {
      const text = await navigator.clipboard.readText();
      if (!text.trim()) {
        setPasteStatus(prev => ({ ...prev, [target]: "Clipboard is empty — paste the full text or use the sample option." }));
        return;
      }
      if (target === "resume") {
        setInputMode("paste");
        setResume(text.trim());
      } else {
        setJd(text.trim());
      }
      setPasteStatus(prev => ({ ...prev, [target]: "Pasted from clipboard ✓" }));
      window.setTimeout(() => setPasteStatus(prev => ({ ...prev, [target]: "" })), 2200);
    } catch {
      setPasteStatus(prev => ({ ...prev, [target]: "Clipboard access was blocked — you can still paste manually." }));
    }
  };

  const applySampleJd = () => {
    setJd(sampleJdText);
    setPasteStatus(prev => ({ ...prev, jd: "Sample JD added — replace it anytime." }));
    window.setTimeout(() => setPasteStatus(prev => ({ ...prev, jd: "" })), 2200);
  };

  const taInp = { ...inp, background: "var(--surface3)", border: "1px solid var(--border)", resize: "vertical", lineHeight: 1.75, fontSize: 13, minHeight: 340, padding: "18px 20px", boxShadow: "inset 0 1px 2px rgba(0,27,72,0.04)" };

  return (
    <div className="fade-in page-pad" style={{ maxWidth: "100%", padding: "40px clamp(16px,5vw,80px) 80px" }}>
      {step === "input" && (
        <>
          <div className="marine-hero">
            <div style={{ position: "relative", zIndex: 1 }}>
              <div className="marine-kicker">01 / Build your career kit</div>
              <div style={{ fontFamily: "'Raleway', sans-serif", fontSize: "clamp(30px,3.4vw,42px)", fontWeight: 600, lineHeight: 1.08, color: "var(--text)", marginBottom: 10 }}>
                Build your next application with a calmer, more human touch.
              </div>
              <div style={{ fontSize: 14, color: "var(--text-muted)", maxWidth: 620, lineHeight: 1.8 }}>
                Share your resume and the role you're eyeing. Rezolt helps shape everything into a thoughtful, role-matched kit in under a minute.
              </div>
              <div className="marine-chip-row">
                {[[<i className="fa-solid fa-check" />, "ATS-aware"], [<i className="fa-solid fa-user-check" />, "Role-matched"], [<i className="fa-solid fa-clock" />, "~60 sec turnaround"]].map(([icon, label]) => (
                  <span key={label} className="marine-chip"><span>{icon}</span>{label}</span>
                ))}
              </div>
            </div>

            <div className="marine-hero-side">
              {[["01", "Resume"], ["02", "Job Description"], ["03", "Generate"]].map(([num, label]) => (
                <div key={num} className="marine-step">
                  <strong>{num}</strong>
                  <span>{label}</span>
                </div>
              ))}
            </div>
          </div>

          <div style={{ maxWidth: 1320, margin: "0 auto" }}>
            <div className="kit-inputs" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, alignItems: "stretch", marginBottom: 20 }}>
              <div className="panel-shell">
                <div className="panel-head" style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                  <div className="panel-label">Share your resume</div>
                  <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                    <button onClick={() => handleClipboardPaste("resume")} style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 999, padding: "7px 12px", fontSize: 11, fontWeight: 700, color: O, cursor: "pointer", fontFamily: "inherit" }}>
                      Paste from clipboard
                    </button>
                    <div style={{ display: "flex", gap: 4, background: "var(--surface2)", borderRadius: 10, padding: 4, border: "1px solid var(--border)" }}>
                      {["upload", "paste"].map(m => (
                        <button key={m} className="mode-toggle" onClick={() => setInputMode(m)} style={{ padding: "6px 12px", borderRadius: 8, border: "none", fontSize: 11, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", background: inputMode === m ? WHITE : "transparent", color: inputMode === m ? DARK : FAINT, boxShadow: inputMode === m ? "0 1px 2px rgba(0,0,0,0.08)" : "none" }}>
                          {m === "upload" ? "Upload" : "Paste"}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>

                <div className="panel-body">
                  {inputMode === "upload" ? (
                    <label style={{ display: "block", cursor: "pointer", flex: 1 }}>
                      <input type="file" accept=".pdf,.docx,.txt" onChange={handleFileUpload} style={{ display: "none" }} />
                      <div style={{ border: `2px dashed ${resume.length > 50 ? G : uploading ? O : BORDER}`, borderRadius: 14, padding: "36px 22px", textAlign: "center", background: resume.length > 50 ? "rgba(34,197,94,0.04)" : "var(--surface3)", transition: "all 0.2s", minHeight: 360, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 12, boxShadow: "inset 0 1px 2px rgba(0,27,72,0.03)" }}>
                        {uploading ? (
                          <>
                            <div style={{ width: 32, height: 32, borderRadius: "50%", border: `3px solid ${BORDER}`, borderTopColor: O, animation: "spin 0.8s linear infinite" }} />
                            <div style={{ fontSize: 13, color: MUTED }}>Organizing your achievements...</div>
                          </>
                        ) : resume.length > 50 ? (
                          <>
                            <div style={{ fontSize: 30 }}><i className="fa-solid fa-check" /></div>
                            <div style={{ fontSize: 14, fontWeight: 700, color: G }}>{uploadedFileName}</div>
                            <div style={{ fontSize: 12, color: MUTED }}>{resume.length} characters extracted</div>
                            <div style={{ fontSize: 12, color: O, textDecoration: "underline", cursor: "pointer" }}>Upload different file</div>
                          </>
                        ) : (
                          <>
                            <div style={{ fontSize: 34 }}><i className="fa-solid fa-wand-magic-sparkles" /></div>
                            <div style={{ fontSize: 14, fontWeight: 700, color: DARK }}>Bring in your current resume</div>
                            <div style={{ fontSize: 13, color: MUTED }}>PDF, DOCX or TXT — we’ll take it from here</div>
                            <div style={{ background: O, color: WHITE, borderRadius: 14, padding: "10px 22px", fontSize: 13, fontWeight: 600 }}>Select Resume</div>
                          </>
                        )}
                      </div>
                    </label>
                  ) : (
                    <textarea value={resume} onChange={e => setResume(e.target.value)}
                      placeholder="Paste the resume you already have — experience, wins, skills, education, and achievements..."
                      rows={18} style={taInp} />
                  )}

                  {pasteStatus.resume && <div style={{ marginTop: 10, fontSize: 12, color: O, fontWeight: 600 }}>{pasteStatus.resume}</div>}
                  {resume.length > MAX_RESUME_CHARS && <div style={{ marginTop: 8, fontSize: 11, color: MUTED, lineHeight: 1.7 }}>Very long resume detected. Rezolt will keep the most relevant sections so your credit is not wasted on overflow.</div>}

                  <div style={{ marginTop: 12, background: "var(--surface2)", border: `1px solid ${BORDER}`, borderRadius: 14, padding: "12px 14px" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center", marginBottom: 10, flexWrap: "wrap" }}>
                      <div style={{ fontSize: 12, fontWeight: 700, color: DARK }}>Resume quality check</div>
                      <div style={{ fontSize: 11, fontWeight: 700, color: resumeQualityColor }}>{resumeQualityLabel}</div>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 8 }}>
                      {resumeChecklist.map(item => (
                        <div key={item.label} style={{ fontSize: 12, color: item.ok ? "#166534" : MUTED, display: "flex", alignItems: "center", gap: 6 }}>
                          <span style={{ fontWeight: 800 }}>{item.ok ? "✓" : "•"}</span>
                          <span>{item.label}</span>
                        </div>
                      ))}
                    </div>
                    {resumeTrimmed && !resumeChecklist[3].ok && (
                      <div style={{ marginTop: 10, fontSize: 11, color: MUTED, lineHeight: 1.7 }}>
                        Tip: add outcomes like “reduced time-to-hire by 20%”, “closed 35 roles”, or “improved dashboard accuracy”.
                      </div>
                    )}
                  </div>
                </div>
              </div>

              <div className="panel-shell">
                <div className="panel-head" style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                  <div className="panel-label">Tell us about the role you’re eyeing</div>
                  <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                    <button onClick={() => handleClipboardPaste("jd")} style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 999, padding: "7px 12px", fontSize: 11, fontWeight: 700, color: O, cursor: "pointer", fontFamily: "inherit" }}>
                      Paste JD
                    </button>
                    <button onClick={applySampleJd} style={{ background: "var(--surface2)", border: `1px solid ${BORDER}`, borderRadius: 999, padding: "7px 12px", fontSize: 11, fontWeight: 700, color: DARK, cursor: "pointer", fontFamily: "inherit" }}>
                      Use sample JD
                    </button>
                    <div style={{ fontSize: 12, color: jd.length > 50 ? G : FAINT, fontWeight: 700 }}>{jd.length > 50 ? "✓ Ready" : `${jd.length}/50 min`}</div>
                  </div>
                </div>
                <div className="panel-body">
                  <textarea value={jd} onChange={e => setJd(e.target.value)}
                    placeholder="Share the role, responsibilities, expectations, and company context — the more detail, the more personal the output feels..."
                    rows={18} style={taInp} />
                  {pasteStatus.jd && <div style={{ marginTop: 10, fontSize: 12, color: O, fontWeight: 600 }}>{pasteStatus.jd}</div>}
                  {jd.length > MAX_JD_CHARS && <div style={{ marginTop: 8, fontSize: 11, color: MUTED, lineHeight: 1.7 }}>This JD is quite large. Rezolt will focus on the highest-signal requirements automatically.</div>}
                  <div style={{ marginTop: 10, fontSize: 11, color: MUTED, lineHeight: 1.7 }}>For better results, include the job title, must-have skills, responsibilities, and any company context you have.</div>
                </div>
              </div>
            </div>

            <div className="dock-bar">
              <div>
                <div className="panel-label" style={{ marginBottom: 8 }}>Included outputs</div>
                <div className="output-badges">
                  {TABS.map(t => (
                    <div key={t.id} style={{ fontSize: 12, color: MUTED, display: "flex", alignItems: "center", gap: 6, background: "var(--surface2)", border: "1px solid var(--border)", borderRadius: 999, padding: "6px 10px" }}>
                      <span style={{ color: G, fontSize: 11 }}>✓</span><i className={t.iconClass} style={{ fontSize: 12 }} /> {t.label}
                    </div>
                  ))}
                </div>
                <div style={{ marginTop: 12, background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 14, padding: "12px 14px" }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: DARK, marginBottom: 6 }}><i className="fa-solid fa-shield-halved" style={{ marginRight: 6 }} />Private by default</div>
                  <div style={{ fontSize: 12, color: MUTED, lineHeight: 1.7 }}>
                    Your resume and JD are only used to generate your kit. They are not posted publicly or shared with employers, and photos or styling from uploaded files are ignored.
                  </div>
                </div>
              </div>

              <div className="action-stack">
                <div style={{ fontSize: 12, color: MUTED, background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 10, padding: "10px 12px", lineHeight: 1.7, textAlign: "center" }}>
                  Text only, no photo storage.
                </div>
                <div style={{ fontSize: 13, color: MUTED }}>
                  <span style={{ color: hasCredits ? G : "#EF4444", fontWeight: 700 }}>
                    {user?.plan === "unlimited" ? "Unlimited kits" : `${user?.credits ?? 0} ${user?.plan === "starter" ? "free kit" : "credit"}${user?.credits !== 1 ? "s" : ""}`}
                  </span> left
                </div>
                <button onClick={() => setShowSample(true)} style={{ background: "none", border: "none", color: AC, fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", textDecoration: "underline", padding: 0 }}>See a sample output</button>
                <button className="gen-btn" onClick={generate} disabled={!isReady || !hasCredits || generating}
                  style={{ background: isReady && hasCredits && !generating ? O : "#E2E8F0", color: isReady && hasCredits && !generating ? WHITE : FAINT, border: "none", borderRadius: 12, padding: "13px 30px", fontSize: 15, fontWeight: 700, cursor: isReady && hasCredits && !generating ? "pointer" : "not-allowed", fontFamily: "inherit", boxShadow: isReady && hasCredits && !generating ? "0 8px 22px rgba(3,29,64,0.18)" : "none", transition: "all 0.2s", minWidth: 180 }}>
                  {generating ? (subStep || "Curating your experience...") : !hasCredits ? "No credits — upgrade to continue" : "Let's build your kit"}
                </button>
              </div>
            </div>
          </div>
          {!hasCredits && (
            <div style={{ background: "rgba(239,68,68,0.06)", border: "1px solid rgba(239,68,68,0.2)", borderRadius: 16, padding: "20px 24px", marginTop: 16, textAlign: "center" }}>
              <div style={{ fontSize: 16, fontWeight: 700, color: "#B91C1C", marginBottom: 6 }}>You've used your {user?.plan === "Free" ? "free resume rewrite" : "credits"}</div>
              <div style={{ fontSize: 13, color: MUTED, marginBottom: 14, lineHeight: 1.7 }}>
                {user?.plan === "Free" ? "Upgrade to Starter (₹99) for a full career kit — resume, cover letter, referral DM, interview prep, and outreach." : "Purchase more credits to continue generating kits."}
              </div>
              <button onClick={() => setPage("payment")} className="btn-primary" style={{ background: "var(--grad)", color: "white", border: "none", borderRadius: 12, padding: "12px 28px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
                View Plans & Upgrade
              </button>
            </div>
          )}
        </>
      )}

      {step === "output" && (
        <>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 24 }}>
            <div>
              <button onClick={() => { setStep("input"); setOutputs({}); setLoading({}); }} style={{ background: "none", border: "none", color: MUTED, fontSize: 13, cursor: "pointer", padding: 0, marginBottom: 6, fontFamily: "inherit" }}>Back</button>
              <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 28, fontWeight: 500, color: DARK }}>Your Career Kit</div>
              <div style={{ fontSize: 13, color: MUTED, marginTop: 3 }}><i className="fa-solid fa-bullseye" style={{ marginRight: 6 }} />{currentRole}</div>
            </div>
            <button onClick={() => { setStep("input"); setOutputs({}); setLoading({}); setResume(""); setJd(""); }}
              style={{ background: BG, border: `1px solid ${BORDER}`, borderRadius: 8, padding: "9px 18px", fontSize: 13, color: MUTED, cursor: "pointer", fontFamily: "inherit" }}>
              New Kit
            </button>
          </div>

          <div style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 24, padding: "18px 22px", marginBottom: 20, boxShadow: "var(--soft-shadow)" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <div style={{ fontSize: 13, fontWeight: 600, color: DARK }}>{anyLoading ? `${activeLoader.title}... ${doneCount} of ${accessibleTabs.length} ready` : "Your partner-crafted kit is ready"}</div>
              <div style={{ fontSize: 13, fontWeight: 700, color: anyLoading ? O : G }}>{progressPct}%</div>
            </div>
            <div style={{ fontSize: 12, color: MUTED, marginBottom: 12 }}>
              {anyLoading ? `${activeLoader.note} This usually takes under a minute — no need to refresh.` : "Everything is ready to review, copy, or download."}
            </div>
            <div style={{ height: 6, background: BORDER, borderRadius: 3, overflow: "hidden", marginBottom: anyLoading ? 14 : 12 }}>
              <div style={{ height: "100%", width: `${progressPct}%`, background: anyLoading ? O : G, borderRadius: 3, transition: "width 0.5s ease" }} />
            </div>
            {anyLoading && (
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(150px,1fr))", gap: 10, marginBottom: 14 }}>
                {loaderStages.map((stage, index) => {
                  const complete = index < loaderStage;
                  const current = index === loaderStage;
                  return (
                    <div key={stage.title} style={{ border: `1px solid ${current ? "rgba(3,29,64,0.18)" : BORDER}`, borderRadius: 14, padding: "10px 12px", background: complete || current ? "var(--surface2)" : WHITE }}>
                      <div style={{ fontSize: 11, fontWeight: 700, color: complete ? G : current ? O : FAINT, marginBottom: 4 }}>{complete ? "✓ Done" : current ? "In progress" : `Step ${index + 1}`}</div>
                      <div style={{ fontSize: 12, fontWeight: 700, color: DARK, lineHeight: 1.45 }}>{stage.title}</div>
                    </div>
                  );
                })}
              </div>
            )}
            <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
              {TABS.map(t => (
                <div key={t.id} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: outputs[t.id] ? MID : FAINT }}>
                  {loading[t.id] ? <div style={{ width: 8, height: 8, borderRadius: "50%", border: `2px solid ${O}`, borderTopColor: "transparent", animation: "spin 0.7s linear infinite" }} />
                    : outputs[t.id] ? <div style={{ width: 8, height: 8, borderRadius: "50%", background: G }} />
                      : <div style={{ width: 8, height: 8, borderRadius: "50%", background: BORDER }} />}
                  <i className={t.iconClass} style={{ fontSize: 12 }} /> {t.label}
                </div>
              ))}
            </div>
            {!anyLoading && doneCount > 0 && (
              <div style={{ marginTop: 14, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                <span style={{ fontSize: 13, color: MUTED }}>Was this kit helpful?</span>
                {kitFeedback ? (
                  <span style={{ fontSize: 13, color: G, fontWeight: 700 }}>Thanks for the feedback.</span>
                ) : (
                  <>
                    <button onClick={() => setKitFeedback("up")} style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 8, padding: "5px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }} title="Helpful">Helpful</button>
                    <button onClick={() => setKitFeedback("down")} style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 8, padding: "5px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }} title="Needs work">Needs work</button>
                  </>
                )}
              </div>
            )}
          </div>

          <div className="tab-pills" style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
            {TABS.map(t => {
              const active = activeTab === t.id;
              const locked = !canAccess(user?.plan, t.minPlan);
              const planNeeded = t.minPlan === "unlimited" ? "Unlimited" : "Starter";
              return (
                <button key={t.id} className="tab-pill" onClick={() => setActiveTab(t.id)}
                  style={{
                    display: "flex", alignItems: "center", gap: 7, padding: "9px 16px", borderRadius: 10,
                    border: `1.5px solid ${locked ? BORDER : active ? O : BORDER}`,
                    background: locked ? BG : active ? "var(--accent-soft)" : WHITE,
                    color: locked ? FAINT : active ? O : MUTED,
                    fontSize: 13, fontWeight: active ? 700 : 500, cursor: "pointer", fontFamily: "inherit", transition: "all 0.15s",
                    opacity: locked ? 0.7 : 1,
                  }}>
                  <span style={{ fontSize: 14, lineHeight: 1, display: "inline-flex", alignItems: "center" }}><i className={t.iconClass} /></span>
                  {t.label}
                  {locked && <span style={{ fontSize: 10, background: "#F1F5F9", color: FAINT, padding: "1px 6px", borderRadius: 10, fontWeight: 700 }}>{planNeeded}+</span>}
                  {!locked && loading[t.id] && <div style={{ width: 7, height: 7, borderRadius: "50%", border: `2px solid ${O}`, borderTopColor: "transparent", animation: "spin 0.7s linear infinite" }} />}
                  {!locked && outputs[t.id] && !loading[t.id] && <div style={{ width: 7, height: 7, borderRadius: "50%", background: G }} />}
                </button>
              );
            })}
          </div>

          <div style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 28, overflow: "hidden", boxShadow: "var(--soft-shadow)" }}>
            <div className="output-header" style={{ padding: "16px 24px", borderBottom: `1px solid ${BORDER}`, display: "flex", justifyContent: "space-between", alignItems: "center", background: BG, flexWrap: "wrap", gap: 10 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ fontSize: 20 }}>{TABS.find(t => t.id === activeTab)?.icon}</span>
                <span style={{ fontSize: 15, fontWeight: 700, color: DARK }}>{TABS.find(t => t.id === activeTab)?.label}</span>
                {outputs[activeTab] && <span style={{ background: "var(--accent-soft)", color: O, border: `1px solid ${BORDER}`, borderRadius: 20, padding: "2px 10px", fontSize: 11, fontWeight: 700 }}>Ready to review</span>}
                {outputs[activeTab] && !loading[activeTab] && canAccess(user?.plan, TABS.find(t => t.id === activeTab)?.minPlan ?? "starter") && (
                  <button onClick={() => regenerateTab(activeTab)} title="Regenerate this output" style={{ background: "none", border: `1px solid ${BORDER}`, borderRadius: 7, padding: "3px 9px", fontSize: 11, color: MUTED, cursor: "pointer", fontFamily: "inherit", transition: "all .15s" }}>↺ Retry</button>
                )}
                {loading[activeTab] && <span style={{ background: "var(--accent-soft)", color: O, border: `1px solid ${BORDER}`, borderRadius: 20, padding: "2px 10px", fontSize: 11, fontWeight: 700 }}>In progress</span>}
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                {/* Template switcher — only on resume tab */}
                {activeTab === "resume" && outputs[activeTab] && (
                  <div className="template-switcher" style={{ display: "flex", gap: 4, background: "var(--surface2)", borderRadius: 8, padding: 3, border: "1px solid var(--border)", flexWrap: "wrap" }}>
                    {["creative", "modern", "bold", "classic", "elegant", "compact", "minimal", "tech", "warm"].map(t => (
                      <button key={t} onClick={() => setSelectedTemplate(t)} style={{
                        padding: "3px 9px", borderRadius: 6, fontSize: 10, fontWeight: 600, cursor: "pointer", fontFamily: "inherit", border: "none", transition: "all .15s",
                        background: selectedTemplate === t ? N1 : "transparent",
                        color: selectedTemplate === t ? "white" : "var(--text-muted)",
                        textTransform: "capitalize",
                      }}>{t}</button>
                    ))}
                  </div>
                )}
                {outputs[activeTab] && (
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    <button className="copy-btn" onClick={() => copy(activeTab)}
                      style={{ background: copied === activeTab ? "rgba(34,197,94,0.08)" : BG, border: `1px solid ${BORDER}`, borderRadius: 7, padding: "9px 18px", fontSize: 13, fontWeight: 600, color: copied === activeTab ? "#15803D" : MUTED, cursor: "pointer", fontFamily: "inherit", transition: "all 0.15s" }}>
                      {copied === activeTab ? "✓ Copied!" : "Copy text"}
                    </button>
                    <button onClick={() => downloadDocx(activeTab)} className="copy-btn"
                      style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 7, padding: "9px 18px", fontSize: 13, fontWeight: 600, color: O, cursor: "pointer", fontFamily: "inherit" }}>
                      Download DOCX
                    </button>
                    <button onClick={() => downloadPdf(activeTab)} className="copy-btn"
                      style={{ background: "var(--grad)", border: "none", borderRadius: 7, padding: "9px 18px", fontSize: 13, fontWeight: 600, color: "white", cursor: "pointer", fontFamily: "inherit" }}>
                      Download PDF
                    </button>
                  </div>
                )}
              </div>
            </div>
            <div className="output-pad" style={{ padding: "28px 32px 40px" }}>
              {activeTab === "negotiate" && canAccess(user?.plan, "unlimited") ? (
                <NegotiateTab />
              ) : !canAccess(user?.plan, TABS.find(t => t.id === activeTab)?.minPlan ?? "starter") ? (
                <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "52px 0", gap: 14, textAlign: "center" }}>
                  <div style={{ width: 52, height: 52, borderRadius: "50%", background: "#FFF7ED", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22 }}>🔒</div>
                  <div style={{ fontSize: 16, fontWeight: 700, color: DARK }}>
                    {TABS.find(t => t.id === activeTab)?.label} is a {TABS.find(t => t.id === activeTab)?.minPlan === "unlimited" ? "Unlimited" : "Starter"} feature
                  </div>
                  <div style={{ fontSize: 13, color: MUTED, maxWidth: 320, lineHeight: 1.7 }}>
                    Upgrade your plan to unlock this output and generate a complete kit for every application.
                  </div>
                  <button onClick={() => setPage("payment")} style={{ background: O, color: WHITE, border: "none", borderRadius: 10, padding: "12px 28px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", boxShadow: "0 2px 10px rgba(249,115,22,0.25)", marginTop: 4 }}>
                    View Plans & Upgrade
                  </button>
                </div>
              ) : loading[activeTab] ? (
                <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "64px 0", gap: 12, textAlign: "center" }}>
                  <div style={{ width: 36, height: 36, borderRadius: "50%", border: `3px solid ${BORDER}`, borderTopColor: O, animation: "spin 0.8s linear infinite" }} />
                  <div style={{ color: DARK, fontSize: 15, fontWeight: 700 }}>{activeLoader.title}</div>
                  <div style={{ color: FAINT, fontSize: 14, maxWidth: 360, lineHeight: 1.7 }}>{loadingTone[activeTab] || "Shaping your next draft..."} {activeLoader.note}</div>
                  <div style={{ background: "var(--surface2)", border: `1px solid ${BORDER}`, borderRadius: 999, padding: "5px 12px", fontSize: 11, fontWeight: 700, color: O }}>{progressPct}% complete</div>
                </div>
              ) : outputs[activeTab] ? (
                <>
                  {(activeTab === "referral" || activeTab === "reach") && (
                    <OutreachQuickActions type={activeTab} text={outputs[activeTab]} />
                  )}
                  <div className={(activeTab === "resume" || activeTab === "cover") ? "paper-preview" : ""} style={{ textAlign: "left" }}>
                    {activeTab === "resume"
                      ? renderResumeWithTemplate(outputs[activeTab], selectedTemplate, null, showResumeBranding)
                      : renderOutput(outputs[activeTab])}
                  </div>
                </>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "64px 0", gap: 10, color: FAINT }}>
                  <span style={{ fontSize: 32 }}>{TABS.find(t => t.id === activeTab)?.icon}</span>
                  <span style={{ fontSize: 14 }}>Your next draft will appear here.</span>
                </div>
              )}
            </div>
          </div>

          {/* Mobile sticky action bar */}
          {outputs[activeTab] && (
            <div className="mobile-only" style={{ position: "fixed", bottom: 65, left: 0, right: 0, background: "var(--surface)", borderTop: "1px solid var(--border)", padding: "10px 16px", display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8, zIndex: 199, boxShadow: "0 -2px 12px rgba(0,27,72,.08)" }}>
              <button onClick={() => copy(activeTab)} style={{ background: copied === activeTab ? "rgba(34,197,94,.1)" : "var(--surface2)", border: `1px solid ${copied === activeTab ? G : "var(--border)"}`, borderRadius: 10, padding: "11px 8px", fontSize: 12, fontWeight: 700, color: copied === activeTab ? G : "var(--text)", cursor: "pointer", fontFamily: "inherit" }}>
                {copied === activeTab ? "✓ Copied" : "📋 Copy"}
              </button>
              <button onClick={() => downloadDocx(activeTab)} style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 10, padding: "11px 8px", fontSize: 12, fontWeight: 700, color: O, cursor: "pointer", fontFamily: "inherit" }}>
                DOCX
              </button>
              <button onClick={() => downloadPdf(activeTab)} style={{ background: "var(--grad)", border: "none", borderRadius: 10, padding: "11px 8px", fontSize: 12, fontWeight: 700, color: "white", cursor: "pointer", fontFamily: "inherit" }}>
                PDF
              </button>
            </div>
          )}
        </>
      )}

      {/* ── SAMPLE OUTPUT MODAL ── */}
      {showSample && (
        <>
          <div onClick={() => setShowSample(false)} style={{ position: "fixed", inset: 0, background: "rgba(3,29,64,0.5)", zIndex: 400 }} />
          <div className="modal-enter" style={{ position: "fixed", top: "50%", left: "50%", width: "min(720px,calc(100vw - 24px))", maxHeight: "88vh", overflowY: "auto", background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 24, padding: "22px 24px", boxShadow: "var(--shadow-lg)", zIndex: 401 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 18 }}>
              <div>
                <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: ".12em", textTransform: "uppercase", color: AC, marginBottom: 6 }}>Sample career kit</div>
                <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 24, color: DARK }}>See what Rezolt produces</div>
                <div style={{ fontSize: 13, color: MUTED, marginTop: 4 }}>HRBP Manager · pharma analytics hiring</div>
              </div>
              <button onClick={() => setShowSample(false)} style={{ background: "var(--surface2)", border: `1px solid ${BORDER}`, borderRadius: 10, width: 34, height: 34, cursor: "pointer", color: MUTED, fontSize: 16 }}><i className="fa-solid fa-xmark" /></button>
            </div>

            {/* Tab switcher */}
            <div style={{ display: "flex", gap: 8, marginBottom: 18, flexWrap: "wrap" }}>
              {[["resume", "Resume"], ["cover", "Cover Letter"], ["referral", "Referral DM"]].map(([id, label]) => (
                <button key={id} onClick={() => setSampleTab(id)} style={{ padding: "8px 18px", borderRadius: 30, fontSize: 13, fontWeight: sampleTab === id ? 700 : 600, cursor: "pointer", fontFamily: "inherit", border: "none", background: sampleTab === id ? N1 : "var(--surface2)", color: sampleTab === id ? "white" : MUTED }}>
                  {label}
                </button>
              ))}
            </div>

            {/* Resume sample */}
            {sampleTab === "resume" && (
              <div style={{ background: "var(--surface)", borderRadius: 16, overflow: "hidden", border: `1px solid ${BORDER}` }}>
                <div style={{ background: "linear-gradient(135deg,#001B48,#02457A)", padding: "12px 18px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <span style={{ color: "rgba(255,255,255,.65)", fontSize: 12 }}>resume_rewrite.pdf</span>
                  <div style={{ display: "flex", gap: 8 }}>
                    <span style={{ background: "rgba(34,197,94,.2)", color: "#86efac", fontSize: 10, fontWeight: 700, padding: "2px 10px", borderRadius: 10 }}>ATS 94/100</span>
                    <span style={{ background: "rgba(255,255,255,.1)", color: "#C9D6E4", fontSize: 10, fontWeight: 600, padding: "2px 10px", borderRadius: 10 }}>Tailored ✦</span>
                  </div>
                </div>
                <div style={{ padding: "20px 22px", display: "flex", gap: 18 }}>
                  <div style={{ width: 130, flexShrink: 0, background: "linear-gradient(180deg,#001B48,#02457A)", borderRadius: 8, padding: "14px 10px", color: "white" }}>
                    <div style={{ width: 36, height: 36, borderRadius: "50%", background: "#02457A", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, fontWeight: 700, marginBottom: 8 }}>P</div>
                    <div style={{ fontSize: 11, fontWeight: 700, marginBottom: 2 }}>Priya Sharma</div>
                    <div style={{ fontSize: 9, color: "#C9D6E4", marginBottom: 8 }}>Lead TA Specialist</div>
                    <div style={{ fontSize: 8, color: "rgba(255,255,255,.4)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 4 }}>Skills</div>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 3 }}>
                      {["Power BI", "Naukri", "HRBP", "SQL"].map(s => <span key={s} style={{ background: "rgba(3,29,64,.28)", color: "#C9D6E4", fontSize: 8, padding: "2px 5px", borderRadius: 3 }}>{s}</span>)}
                    </div>
                  </div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 8, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#02457A", marginBottom: 4 }}>Professional Summary</div>
                    <p style={{ fontSize: 10, color: "var(--text-mid)", lineHeight: 1.75, marginBottom: 12 }}>Lead Talent Acquisition Specialist with 4.5 years driving end-to-end recruitment for pharma analytics roles. Targeting the HRBP Manager position — bringing Power BI dashboards, stakeholder management, and MOU partnership execution.</p>
                    <div style={{ fontSize: 8, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: "#02457A", marginBottom: 6 }}>Experience</div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: N1, marginBottom: 2 }}>Lead TA Specialist · Current Company, Bangalore</div>
                    <div style={{ fontSize: 9, color: "var(--text-mid)", lineHeight: 1.65 }}>• Reduced time-to-hire by 28% across 40+ analytics and engineering roles<br />• Built Power BI WFH dashboard tracking 150+ headcount across 4 quarters<br />• Executed MOUs with Dayananda Sagar College and MIT Pune — 12 hires in FY2024</div>
                  </div>
                </div>
              </div>
            )}

            {/* Cover letter sample */}
            {sampleTab === "cover" && (
              <div style={{ background: "var(--surface)", borderRadius: 16, overflow: "hidden", border: `1px solid ${BORDER}` }}>
                <div style={{ background: "linear-gradient(135deg,#02457A,#02457A)", padding: "12px 18px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <span style={{ color: "rgba(255,255,255,.65)", fontSize: 12 }}>cover_letter.pdf</span>
                  <span style={{ background: "rgba(255,255,255,.1)", color: "#D6E8EE", fontSize: 10, fontWeight: 600, padding: "2px 10px", borderRadius: 10 }}>Tailored ✦</span>
                </div>
                <div style={{ padding: "28px 32px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 18 }}>
                    <div><div style={{ fontSize: 14, fontWeight: 700, color: "var(--text)" }}>Priya Sharma</div><div style={{ fontSize: 11, color: MUTED }}>priya@email.com · Bangalore, KA</div></div>
                    <div style={{ fontSize: 11, color: FAINT }}>5 April 2026</div>
                  </div>
                  <div style={{ width: 32, height: 3, background: "var(--grad)", borderRadius: 2, marginBottom: 16 }} />
                  <p style={{ fontSize: 13, color: "var(--text-mid)", lineHeight: 1.85, marginBottom: 12 }}>Dear Hiring Manager,</p>
                  <p style={{ fontSize: 13, color: "var(--text-mid)", lineHeight: 1.85, marginBottom: 12 }}>With 4.5 years building talent pipelines for pharma analytics functions, I bring exactly what the HRBP Manager role requires — domain fluency you cannot onboard, and execution speed your team can rely on from day one.</p>
                  <p style={{ fontSize: 13, color: "var(--text-mid)", lineHeight: 1.85, marginBottom: 12 }}>At my current organisation, I reduced time-to-hire by 28% across 40+ roles and built a Power BI dashboard tracking 150+ headcount decisions — giving me a dual lens on talent strategy and HR data storytelling.</p>
                  <p style={{ fontSize: 13, color: "var(--text-mid)", lineHeight: 1.85 }}>Warm regards,<br /><strong style={{ color: "var(--text)" }}>Priya Sharma</strong></p>
                </div>
              </div>
            )}

            {/* Referral DM sample */}
            {sampleTab === "referral" && (
              <div style={{ background: "var(--surface)", borderRadius: 16, overflow: "hidden", border: `1px solid ${BORDER}` }}>
                <div style={{ background: "linear-gradient(135deg,#031D40,#08284F)", padding: "12px 18px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <span style={{ color: "rgba(255,255,255,.55)", fontSize: 12 }}>referral_messages.txt</span>
                  <span style={{ background: "rgba(0,27,72,.12)", color: N1, fontSize: 10, fontWeight: 700, padding: "2px 10px", borderRadius: 10 }}>Outreach-ready</span>
                </div>
                <div className="output-pad" style={{ padding: "22px 26px" }}>
                  <div style={{ marginBottom: 18 }}>
                    <div style={{ display: "inline-block", background: "var(--accent-soft)", color: AC, fontSize: 9, fontWeight: 700, padding: "3px 10px", borderRadius: 4, marginBottom: 10, letterSpacing: ".06em" }}>VERSION 1 · CONNECTION REQUEST</div>
                    <div style={{ background: "var(--bg)", borderLeft: `3px solid ${AC}`, borderRadius: "0 10px 10px 0", padding: "12px 14px", fontSize: 13, color: "var(--text-mid)", lineHeight: 1.75 }}>
                      Hi Anjali, I noticed you&apos;re at [Company] — I&apos;m applying for the HRBP Manager role and have 4.5 years in pharma TA. Would love to connect!
                    </div>
                  </div>
                  <div style={{ height: 1, background: "var(--border)", marginBottom: 18 }} />
                  <div>
                    <div style={{ display: "inline-block", background: "var(--accent-soft)", color: AC, fontSize: 9, fontWeight: 700, padding: "3px 10px", borderRadius: 4, marginBottom: 10, letterSpacing: ".06em" }}>VERSION 2 · DIRECT MESSAGE</div>
                    <div style={{ background: "var(--bg)", borderLeft: `3px solid ${LB}`, borderRadius: "0 10px 10px 0", padding: "12px 14px", fontSize: 13, color: "var(--text-mid)", lineHeight: 1.85 }}>
                      Hi Anjali, hope you&apos;re doing well!<br /><br />
                      I came across the HRBP Manager opening at [Company] and I&apos;m genuinely excited — it aligns closely with what I&apos;ve been building over 4.5 years in pharma TA and HR analytics.<br /><br />
                      Would you be open to referring me or sharing tips on the process? Happy to send my resume directly!
                    </div>
                  </div>
                </div>
              </div>
            )}

            <div style={{ marginTop: 20, textAlign: "center" }}>
              <button onClick={() => setShowSample(false)} style={{ background: "var(--grad)", color: "white", border: "none", borderRadius: 12, padding: "12px 28px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
                Got it, build mine
              </button>
            </div>
          </div>
        </>
      )}
    </div>
  );
}

function ResetPasswordPage({ setPage }) {
  const [password, setPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [success, setSuccess] = useState(false);
  const [error, setError] = useState("");

  const handle = async () => {
    if (password.length < 8) return setError("Use at least 8 characters for your new password.");
    if (password !== confirmPassword) return setError("Passwords do not match.");
    setLoading(true); setError("");
    const { error: err } = await supabase.auth.updateUser({ password });
    if (err) { setError(err.message); setLoading(false); return; }
    setSuccess(true); setLoading(false);
  };

  return (
    <div className="fade-in" style={{ minHeight: "calc(100vh - 63px)", display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
      <div style={{ width: "100%", maxWidth: 420 }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 28, fontWeight: 500, color: DARK, marginBottom: 6 }}>Set a new password</div>
          <div style={{ fontSize: 14, color: MUTED }}>You are securely signed in for password recovery</div>
        </div>
        <div style={{ background: WHITE, border: `1.5px solid ${O}`, borderRadius: 14, padding: 36, boxShadow: "0 0 0 4px rgba(249,115,22,0.06)" }}>
          {success ? (
            <div style={{ textAlign: "center" }}>
              <div style={{ fontSize: 36, marginBottom: 16 }}><i className="fa-solid fa-check" /></div>
              <div style={{ fontSize: 15, fontWeight: 700, color: DARK, marginBottom: 8 }}>Password updated</div>
              <div style={{ fontSize: 13, color: MUTED, lineHeight: 1.7, marginBottom: 24 }}>Your password has been reset successfully. You can now sign in normally.</div>
              <button onClick={() => setPage("auth")} style={{ background: O, color: WHITE, border: "none", borderRadius: 10, padding: "12px 28px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>Back to Sign In</button>
            </div>
          ) : (
            <>
              <div style={{ marginBottom: 16 }}>
                <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", color: MUTED, marginBottom: 7 }}>New Password</div>
                <input style={inp} type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="At least 8 characters" />
              </div>
              <div style={{ marginBottom: 20 }}>
                <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", color: MUTED, marginBottom: 7 }}>Confirm Password</div>
                <input style={inp} type="password" value={confirmPassword} onChange={e => setConfirmPassword(e.target.value)} placeholder="Repeat your new password" onKeyDown={e => e.key === "Enter" && handle()} />
              </div>
              {error && <div style={{ background: "rgba(239,68,68,0.06)", border: "1px solid rgba(239,68,68,0.2)", borderRadius: 8, padding: "10px 14px", fontSize: 13, color: "#B91C1C", marginBottom: 16 }}>{error}</div>}
              <button onClick={handle} disabled={loading} style={{ width: "100%", background: loading ? "#E2E8F0" : O, color: loading ? FAINT : WHITE, border: "none", borderRadius: 10, padding: "14px", fontSize: 15, fontWeight: 700, cursor: loading ? "not-allowed" : "pointer", fontFamily: "inherit", transition: "all 0.2s" }}>
                {loading ? "Updating..." : "Save New Password"}
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

function ForgotPasswordPage({ setPage }) {
  const [email, setEmail] = useState("");
  const [loading, setLoading] = useState(false);
  const [sent, setSent] = useState(false);
  const [error, setError] = useState("");

  const handle = async () => {
    if (!email) return setError("Please enter your email.");
    setLoading(true); setError("");
    const { error: err } = await supabase.auth.resetPasswordForEmail(email, {
      redirectTo: `${window.location.origin}`,
    });
    if (err) { setError(err.message); setLoading(false); return; }
    setSent(true); setLoading(false);
  };

  return (
    <div className="fade-in" style={{ minHeight: "calc(100vh - 63px)", display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
      <div style={{ width: "100%", maxWidth: 420 }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 28, fontWeight: 500, color: DARK, marginBottom: 6 }}>Reset your password</div>
          <div style={{ fontSize: 14, color: MUTED }}>We'll send a reset link to your email</div>
        </div>
        <div style={{ background: WHITE, border: `1.5px solid ${O}`, borderRadius: 14, padding: 36, boxShadow: "0 0 0 4px rgba(249,115,22,0.06)" }}>
          {sent ? (
            <div style={{ textAlign: "center" }}>
              <div style={{ fontSize: 40, marginBottom: 16 }}>📬</div>
              <div style={{ fontSize: 15, fontWeight: 700, color: DARK, marginBottom: 8 }}>Check your inbox</div>
              <div style={{ fontSize: 13, color: MUTED, lineHeight: 1.7, marginBottom: 24 }}>We sent a password reset link to <strong>{email}</strong>. Open it on this device to set a new password. If you do not see it within a minute, please check Spam or Promotions.</div>
              <button onClick={() => setPage("auth")} style={{ background: O, color: WHITE, border: "none", borderRadius: 10, padding: "12px 28px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>Back to Sign In</button>
            </div>
          ) : (
            <>
              <div style={{ marginBottom: 20 }}>
                <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", color: MUTED, marginBottom: 7 }}>Email</div>
                <input style={inp} type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="you@email.com" onKeyDown={e => e.key === "Enter" && handle()} />
              </div>
              {error && <div style={{ background: "rgba(239,68,68,0.06)", border: "1px solid rgba(239,68,68,0.2)", borderRadius: 8, padding: "10px 14px", fontSize: 13, color: "#B91C1C", marginBottom: 16 }}>{error}</div>}
              <button onClick={handle} disabled={loading} style={{ width: "100%", background: loading ? "#E2E8F0" : O, color: loading ? FAINT : WHITE, border: "none", borderRadius: 10, padding: "14px", fontSize: 15, fontWeight: 700, cursor: loading ? "not-allowed" : "pointer", fontFamily: "inherit", transition: "all 0.2s" }}>
                {loading ? "Sending..." : "Send Reset Link"}
              </button>
              <button onClick={() => setPage("auth")} style={{ width: "100%", background: "none", border: "none", color: MUTED, fontSize: 13, cursor: "pointer", marginTop: 14, fontFamily: "inherit" }}>Back to Sign In</button>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

function FAQPage({ setPage, user }) {
  const [open, setOpen] = useState(0);
  const faqs = [
    {
      q: "How does Rezolt create my career kit?",
      a: "You share your current resume and the role you want. Rezolt then rewrites and tailors your resume, cover letter, referral message, interview prep, and outreach materials so they feel more aligned to the opportunity you are targeting.",
    },
    {
      q: "Will this keep my original resume style and fonts?",
      a: "Yes. Your resume output keeps the clean professional template direction already built into Rezolt, while the platform continues to use your chosen font system and brand palette.",
    },
    {
      q: "Is the resume ATS-friendly?",
      a: "Yes. Rezolt is designed to produce cleaner, role-matched content with stronger keyword alignment, measurable impact, and formatting that is easier for ATS systems to read.",
    },
    {
      q: "Which plan should I choose?",
      a: "Starter is best for a single application, Pro is better if you are actively applying across multiple roles, and Unlimited is ideal if you want access to negotiation and notice-period tools as well.",
    },
    {
      q: "Can I export my outputs?",
      a: "Yes. You can copy your drafts directly or download them as PDF from inside your saved career kit view.",
    },
    {
      q: "How do I reach Rezolt for help?",
      a: "Use the Contact page to send an enquiry. You can also write directly to hello@rezolt.in for support, feedback, or partnership discussions.",
    },
  ];

  const quickTopics = [
    { title: "Resume & tailoring", desc: "ATS, rewriting, and export help.", icon: <i className="fa-solid fa-file" /> },
    { title: "Plans & credits", desc: "Starter, Pro, Unlimited, and pricing.", icon: <i className="fa-solid fa-dollar-sign" /> },
    { title: "Support & enquiries", desc: "Billing, feedback, or partnership questions.", icon: <i className="fa-solid fa-handshake" /> },
  ];

  return (
    <div className="fade-in">
      <div className="hero-gradient" style={{ padding: "56px clamp(16px,5vw,80px) 52px", position: "relative", overflow: "hidden" }}>
        <div style={{ position: "relative", zIndex: 1, maxWidth: 1080, margin: "0 auto", textAlign: "center" }}>
          <div style={{ display: "inline-flex", alignItems: "center", gap: 8, background: "var(--accent-soft)", border: "1px solid rgba(3,29,64,0.08)", borderRadius: 999, padding: "6px 14px", marginBottom: 18, color: AC, fontSize: 12, fontWeight: 700 }}>
            <i className="fa-solid fa-circle-question" />Help Center
          </div>
          <h1 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(32px,4vw,50px)", fontWeight: 400, color: "var(--text)", marginBottom: 12 }}>
            Answers for every step of your application
          </h1>
          <p style={{ fontSize: 15, color: "var(--text-muted)", lineHeight: 1.8, maxWidth: 760, margin: "0 auto 24px" }}>
            A more polished help-center experience inspired by leading resume platforms, while still staying fully within Rezolt’s original template, fonts, and brand tone.
          </p>

          <div className="three-col" style={{ display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 12, maxWidth: 960, margin: "0 auto 24px" }}>
            {[
              { value: "6", label: "common questions" },
              { value: "ATS", label: "resume guidance" },
              { value: "24/7", label: "email enquiry option" },
            ].map(item => (
              <div key={item.label} style={{ background: "rgba(255,255,255,0.84)", border: "1px solid rgba(3,29,64,0.08)", borderRadius: 18, padding: "14px 16px", boxShadow: "var(--soft-shadow)" }}>
                <div style={{ fontSize: 22, fontWeight: 700, color: AC, marginBottom: 2 }}>{item.value}</div>
                <div style={{ fontSize: 12, color: MUTED }}>{item.label}</div>
              </div>
            ))}
          </div>

          <div style={{ display: "flex", justifyContent: "center", gap: 10, flexWrap: "wrap" }}>
            <button onClick={() => setPage(user ? "generate" : "auth")} className="btn-primary" style={{ background: "var(--grad)", color: WHITE, border: "none", borderRadius: 12, padding: "12px 24px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
              {user ? "Build a kit" : "Get started"}
            </button>
            <button onClick={() => setPage("contact")} style={{ background: WHITE, color: AC, border: "1px solid var(--border)", borderRadius: 12, padding: "12px 24px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
              Contact us
            </button>
          </div>
        </div>
      </div>

      <div className="section-pad" style={{ padding: "44px clamp(16px,5vw,80px) 80px", background: "var(--bg)" }}>
        <div style={{ maxWidth: 1180, margin: "0 auto" }}>
          <div className="three-col" style={{ display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 16, marginBottom: 26 }}>
            {quickTopics.map(card => (
              <div key={card.title} style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 22, padding: "20px 22px", boxShadow: "var(--soft-shadow)" }}>
                <div style={{ fontSize: 24, marginBottom: 10 }}>{card.icon}</div>
                <div style={{ fontSize: 14, fontWeight: 700, color: DARK, marginBottom: 6 }}>{card.title}</div>
                <div style={{ fontSize: 13, color: MUTED, lineHeight: 1.7 }}>{card.desc}</div>
              </div>
            ))}
          </div>

          <div className="two-col-md" style={{ display: "grid", gridTemplateColumns: "minmax(260px, 0.8fr) minmax(0, 1.2fr)", gap: 20, alignItems: "start" }}>
            <div style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 24, padding: "20px", boxShadow: "var(--soft-shadow)" }}>
              <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", color: AC, marginBottom: 10 }}>Popular topics</div>
              <div style={{ display: "grid", gap: 10, marginBottom: 18 }}>
                {[
                  "Resume improvements",
                  "Plan and pricing",
                  "Exporting PDFs",
                  "Negotiation tools",
                ].map(item => (
                  <div key={item} style={{ padding: "10px 12px", background: "var(--surface2)", borderRadius: 12, fontSize: 13, color: MID, border: "1px solid var(--border)" }}>
                    {item}
                  </div>
                ))}
              </div>
              <div style={{ fontSize: 13, color: MUTED, lineHeight: 1.8, marginBottom: 16 }}>
                Still not finding what you need? Reach out and we’ll point you in the right direction.
              </div>
              <button onClick={() => setPage("contact")} style={{ width: "100%", background: WHITE, color: AC, border: "1px solid var(--border)", borderRadius: 12, padding: "12px 16px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
                Ask a question
              </button>
            </div>

            <div style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 26, boxShadow: "var(--soft-shadow)", overflow: "hidden" }}>
              {faqs.map((item, i) => {
                const active = open === i;
                return (
                  <div key={item.q} style={{ borderBottom: i < faqs.length - 1 ? `1px solid ${BORDER}` : "none" }}>
                    <button onClick={() => setOpen(active ? -1 : i)} style={{ width: "100%", textAlign: "left", background: active ? "var(--accent-soft)" : WHITE, border: "none", padding: "18px 20px", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, fontFamily: "inherit" }}>
                      <span style={{ fontSize: 15, fontWeight: 700, color: DARK }}>{item.q}</span>
                      <span style={{ fontSize: 18, color: active ? AC : MUTED }}>{active ? "−" : "+"}</span>
                    </button>
                    {active && (
                      <div style={{ padding: "0 20px 18px", fontSize: 14, color: MUTED, lineHeight: 1.8, background: "var(--accent-soft)" }}>
                        {item.a}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function ContactPage({ setPage, user }) {
  const [form, setForm] = useState(() => ({ name: user?.name || "", email: user?.email || "", company: "", topic: "General enquiry", message: "" }));
  const [submitted, setSubmitted] = useState(false);
  const [error, setError] = useState("");

  const update = (key, value) => setForm(prev => ({ ...prev, [key]: value }));

  const submitEnquiry = () => {
    if (!form.name || !form.email || !form.message) {
      setError("Please fill in your name, email, and message.");
      return;
    }
    setError("");
    const subject = encodeURIComponent(`[Rezolt Enquiry] ${form.topic}`);
    const body = encodeURIComponent(
      `Name: ${form.name}\nEmail: ${form.email}\nCompany: ${form.company || "Not shared"}\nTopic: ${form.topic}\n\nMessage:\n${form.message}`
    );
    setSubmitted(true);
    window.location.href = `mailto:hello@rezolt.in?subject=${subject}&body=${body}`;
  };

  return (
    <div className="fade-in">
      <div className="hero-gradient" style={{ padding: "56px clamp(16px,5vw,80px) 52px", position: "relative", overflow: "hidden" }}>
        <div style={{ position: "relative", zIndex: 1, maxWidth: 1040, margin: "0 auto", textAlign: "center" }}>
          <div style={{ display: "inline-flex", alignItems: "center", gap: 8, background: "var(--accent-soft)", border: "1px solid rgba(3,29,64,0.08)", borderRadius: 999, padding: "6px 14px", marginBottom: 18, color: AC, fontSize: 12, fontWeight: 700 }}>
            <i className="fa-solid fa-envelope" />Contact Rezolt
          </div>
          <h1 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(32px,4vw,50px)", fontWeight: 400, color: "var(--text)", marginBottom: 12 }}>
            We’d love to hear from you
          </h1>
          <p style={{ fontSize: 15, color: "var(--text-muted)", lineHeight: 1.8, maxWidth: 760, margin: "0 auto 20px" }}>
            Need help with plans, billing, feedback, product questions, or a partnership enquiry? Share a few details and we’ll make it easy to reach the right team.
          </p>
          <div style={{ display: "flex", justifyContent: "center", gap: 10, flexWrap: "wrap" }}>
            {[
              "Billing support",
              "Plan questions",
              "Product feedback",
              "Partnerships",
            ].map(item => (
              <span key={item} style={{ padding: "7px 12px", borderRadius: 999, background: "rgba(255,255,255,0.84)", border: "1px solid rgba(3,29,64,0.08)", fontSize: 12, color: MID, fontWeight: 600 }}>
                {item}
              </span>
            ))}
          </div>
        </div>
      </div>

      <div className="section-pad" style={{ padding: "44px clamp(16px,5vw,80px) 80px", background: "var(--bg)" }}>
        <div className="two-col-md" style={{ maxWidth: 1180, margin: "0 auto", display: "grid", gridTemplateColumns: "minmax(0, 1.1fr) minmax(320px, 0.9fr)", gap: 22, alignItems: "start" }}>
          <div style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 26, padding: "24px", boxShadow: "var(--soft-shadow)" }}>
            <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", color: AC, marginBottom: 8 }}>Enquiry form</div>
            <div style={{ fontSize: 20, fontWeight: 700, color: DARK, marginBottom: 6 }}>Tell us what you need</div>
            <div style={{ fontSize: 13, color: MUTED, lineHeight: 1.7, marginBottom: 18 }}>Share a few details below and your email draft will open automatically for `hello@rezolt.in`.</div>

            {submitted && (
              <div style={{ background: "rgba(34,197,94,0.07)", border: "1px solid rgba(34,197,94,0.24)", borderRadius: 14, padding: "12px 14px", fontSize: 13, color: "#15803D", marginBottom: 16 }}>
                Your enquiry draft is ready. If your mail app didn’t open automatically, you can write directly to hello@rezolt.in.
              </div>
            )}
            {error && (
              <div style={{ background: "rgba(239,68,68,0.06)", border: "1px solid rgba(239,68,68,0.2)", borderRadius: 12, padding: "10px 12px", fontSize: 13, color: "#B91C1C", marginBottom: 16 }}>
                {error}
              </div>
            )}

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 12, marginBottom: 12, alignItems: "start" }}>
              <div>
                <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: MUTED, marginBottom: 7 }}>Name</div>
                <input style={inp} value={form.name} onChange={e => update("name", e.target.value)} placeholder="Your name" />
              </div>
              <div>
                <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: MUTED, marginBottom: 7 }}>Email</div>
                <input style={inp} type="email" value={form.email} onChange={e => update("email", e.target.value)} placeholder="you@email.com" />
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 12, marginBottom: 12, alignItems: "start" }}>
              <div>
                <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: MUTED, marginBottom: 7 }}>Company</div>
                <input style={inp} value={form.company} onChange={e => update("company", e.target.value)} placeholder="Optional" />
              </div>
              <div>
                <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: MUTED, marginBottom: 7 }}>Topic</div>
                <select style={{ ...inp, cursor: "pointer" }} value={form.topic} onChange={e => update("topic", e.target.value)}>
                  <option>General enquiry</option>
                  <option>Plan and pricing</option>
                  <option>Billing support</option>
                  <option>Partnerships</option>
                  <option>Product feedback</option>
                </select>
              </div>
            </div>

            <div style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: MUTED, marginBottom: 7 }}>Message</div>
              <textarea style={{ ...inp, minHeight: 170, resize: "vertical" }} value={form.message} onChange={e => update("message", e.target.value)} placeholder="Tell us how we can help, what you’re looking for, or what you’d like to improve..." />
            </div>

            <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
              <button onClick={submitEnquiry} className="btn-primary" style={{ background: "var(--grad)", color: WHITE, border: "none", borderRadius: 12, padding: "12px 24px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
                Send enquiry
              </button>
              <button onClick={() => setPage("faq")} style={{ background: WHITE, color: AC, border: "1px solid var(--border)", borderRadius: 12, padding: "12px 24px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
                View FAQ
              </button>
            </div>
          </div>

          <div style={{ display: "grid", gap: 16 }}>
            <div style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 22, padding: "18px 20px", boxShadow: "var(--soft-shadow)" }}>
              <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", color: AC, marginBottom: 8 }}>Why people write to us</div>
              <div style={{ display: "grid", gap: 10 }}>
                {[
                  "Questions about plans and credits",
                  "Billing or account support",
                  "Feature suggestions and feedback",
                  "Business and partnership enquiries",
                ].map(item => (
                  <div key={item} style={{ padding: "10px 12px", background: "var(--surface2)", borderRadius: 12, fontSize: 13, color: MID, border: "1px solid var(--border)" }}>
                    {item}
                  </div>
                ))}
              </div>
            </div>

            {[
              { title: "Response time", desc: "Usually within 1–2 business days.", icon: <i className="fa-solid fa-clock" /> },
              { title: "Email us", desc: "hello@rezolt.in", icon: <i className="fa-solid fa-envelope" /> },
              { title: "Best for", desc: "Plans, support, feedback, and partnership enquiries.", icon: <i className="fa-solid fa-handshake" /> },
            ].map(card => (
              <div key={card.title} style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 22, padding: "18px 20px", boxShadow: "var(--soft-shadow)" }}>
                <div style={{ fontSize: 22, marginBottom: 8 }}>{card.icon}</div>
                <div style={{ fontSize: 15, fontWeight: 700, color: DARK, marginBottom: 4 }}>{card.title}</div>
                <div style={{ fontSize: 13, color: MUTED, lineHeight: 1.7 }}>{card.desc}</div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── ADMIN PAGE ───────────────────────────────────────────────────────────────

function AdminPage({ user, setPage }) {
  const ADMIN_EMAIL = "hema.manoharan13@outlook.com";
  const [stats, setStats] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [lastStep, setLastStep] = useState("Initializing...");
  const [search, setSearch] = useState("");
  const [planFilter, setPlanFilter] = useState("all");

  useEffect(() => {
    const isAdmin = user?.email?.trim().toLowerCase() === ADMIN_EMAIL.trim().toLowerCase();
    console.log("🔐 Admin check:", { userEmail: user?.email, isAdmin, expectedEmail: ADMIN_EMAIL });
    
    if (!isAdmin) {
      console.warn("⚠ Unauthorized admin access attempt");
      setPage("dashboard");
      return;
    }

    let mounted = true;
    const timeout = setTimeout(() => {
      if (mounted && loading) {
        console.warn("AdminPage: Data load timed out after 30s.");
        setError(`Request timed out at step: "${lastStep}". This usually means the database or network is sluggish.`);
        setLoading(false);
      }
    }, 30000);

    const load = async () => {
      try {
        console.log("AdminPage: Checking auth state...");
        const { data: { session } } = await supabase.auth.getSession();
        console.log("AdminPage: Session found:", !!session, "Token length:", session?.access_token?.length || 0);

        const s = await fetchAdminStats((step) => {
          if (mounted) setLastStep(step);
        });
        if (!mounted) return;
        clearTimeout(timeout);
        console.log("Admin stats loaded successfully.");
        setStats(s);
        setError(null);
        setLoading(false);
      } catch (e) {
        if (!mounted) return;
        clearTimeout(timeout);
        console.error("Admin stats error:", e);
        setError(e?.message || "Failed to load admin data.");
        setLoading(false);
      }
    };

    load();
    return () => { mounted = false; clearTimeout(timeout); };
  }, [user]);

  const resetSession = async () => {
    console.log("AdminPage: Resetting session...");
    localStorage.clear();
    sessionStorage.clear();
    await supabase.auth.signOut();
    window.location.reload();
  };

  if (!user || user.email?.trim().toLowerCase() !== ADMIN_EMAIL.trim().toLowerCase()) return null;

  const PLAN_COLORS = { starter: FAINT, Pro: O, unlimited: "#7C3AED" };
  const PLAN_BG = { starter: "#F1F5F9", Pro: "#FFF7ED", unlimited: "#F5F3FF" };

  const filtered = (stats?.profiles || []).filter(p => {
    const matchSearch = !search || p.name?.toLowerCase().includes(search.toLowerCase()) || p.email?.toLowerCase().includes(search.toLowerCase());
    const matchPlan = planFilter === "all" || p.plan === planFilter;
    return matchSearch && matchPlan;
  });

  const revenue = (stats?.profiles || []).reduce((acc, p) => {
    const plan = p.plan?.toLowerCase();
    if (plan === "pro") return acc + 299;
    if (plan === "unlimited" || plan === "unlimited_monthly") return acc + 599;
    return acc;
  }, 0);

  return (
    <div className="fade-in page-pad" style={{ maxWidth: "100%", padding: "40px clamp(16px,5vw,80px) 80px" }}>
      {/* Header */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 32 }}>
        <div>
          <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", color: O, marginBottom: 6 }}>Admin</div>
          <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 30, fontWeight: 500, color: DARK }}>Rezolt Dashboard</div>
          <div style={{ fontSize: 13, color: MUTED, marginTop: 4 }}>Only visible to you</div>
        </div>
        <button onClick={() => setPage("dashboard")} style={{ background: BG, border: `1px solid ${BORDER}`, borderRadius: 8, padding: "9px 18px", fontSize: 13, color: MUTED, cursor: "pointer", fontFamily: "inherit" }}>Back</button>
      </div>

      {loading ? (
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "80px 0", gap: 16 }}>
          <div style={{ width: 32, height: 32, borderRadius: "50%", border: `3px solid ${BORDER}`, borderTopColor: O, animation: "spin 0.8s linear infinite" }} />
          <div style={{ fontSize: 13, color: MUTED, fontWeight: 500 }}>{lastStep}</div>
          <button onClick={resetSession} style={{ marginTop: 24, padding: "8px 16px", fontSize: 12, background: "none", border: `1px solid ${BORDER}`, borderRadius: 8, cursor: "pointer", color: MUTED }}>Reset Session & Log In Fresh</button>
        </div>
      ) : error ? (
        <div style={{ background: "rgba(239,68,68,0.06)", border: "1px solid rgba(239,68,68,0.2)", borderRadius: 12, padding: "32px 24px", color: "#B91C1C", textAlign: "center" }}>
          <div style={{ fontSize: 32, marginBottom: 16 }}>⚠</div>
          <div style={{ fontWeight: 700, fontSize: 18, marginBottom: 8 }}>Admin Loading Failed</div>
          <div style={{ fontSize: 14, opacity: 0.8, maxWidth: 400, margin: "0 auto 24px", lineHeight: 1.6 }}>{error}</div>
          <div style={{ display: "flex", gap: 12, justifyContent: "center" }}>
            <button onClick={() => window.location.reload()} style={{ background: "#B91C1C", color: "white", border: "none", borderRadius: 8, padding: "10px 24px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>Retry Load</button>
            <button onClick={resetSession} style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 8, padding: "10px 24px", fontSize: 14, color: MUTED, cursor: "pointer", fontFamily: "inherit" }}>Reset Session & Log In Fresh</button>
          </div>
        </div>
      ) : (
        <>
          {/* Stats cards */}
          <div className="dash-stats" style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 16, marginBottom: 36 }}>
            {[
              { label: "Total Users", value: stats?.totalUsers ?? 0, color: DARK, sub: "signed up" },
              { label: "Kits Generated", value: stats?.totalKits ?? 0, color: O, sub: "all time" },
              { label: "Paying Users", value: filtered.filter(p => p.plan?.toLowerCase() !== "starter" && p.plan?.toLowerCase() !== "free").length, color: G, sub: "Pro + unlimited" },
              { label: "Est. Revenue", value: `₹${revenue.toLocaleString("en-IN")}`, color: "#7C3AED", sub: "current month" },
            ].map(s => (
              <div key={s.label} style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 12, padding: "20px 22px", boxShadow: "0 1px 3px rgba(0,0,0,0.05)" }}>
                <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", color: FAINT, marginBottom: 10 }}>{s.label}</div>
                <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 32, fontWeight: 500, color: s.color, marginBottom: 4, lineHeight: 1 }}>{s.value}</div>
                <div style={{ fontSize: 12, color: FAINT }}>{s.sub}</div>
              </div>
            ))}
          </div>

          {/* Plan breakdown */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 32 }}>
            {["starter", "Pro", "unlimited"].map(planName => {
              const count = (stats?.profiles || []).filter(p => p.plan?.toLowerCase() === planName.toLowerCase()).length;
              const total = stats?.totalUsers || 1;
              const pct = Math.round(count / total * 100);
              return (
                <div key={planName} style={{ background: PLAN_BG[planName] || BG, border: `1px solid ${BORDER}`, borderRadius: 10, padding: "16px 20px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <div style={{ fontSize: 12, fontWeight: 700, color: PLAN_COLORS[planName] || AC, textTransform: "capitalize", marginBottom: 4 }}>{planName}</div>
                    <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 26, fontWeight: 500, color: DARK }}>{count}</div>
                  </div>
                  <div style={{ fontSize: 22, fontWeight: 700, color: PLAN_COLORS[planName] || AC, opacity: 0.4 }}>{pct}%</div>
                </div>
              );
            })}
          </div>

          {/* User table */}
          <div style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 14, overflow: "hidden", boxShadow: "0 1px 4px rgba(0,0,0,0.05)" }}>
            <div style={{ padding: "16px 20px", borderBottom: `1px solid ${BORDER}`, display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
              <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search by name or email..." style={{ ...inp, width: 260, padding: "9px 12px", fontSize: 13 }} />
              <div style={{ display: "flex", gap: 6 }}>
                {["all", "starter", "Pro", "unlimited"].map(p => (
                  <button key={p} onClick={() => setPlanFilter(p)} style={{ padding: "7px 14px", borderRadius: 8, border: `1px solid ${planFilter === p ? O : BORDER}`, background: planFilter === p ? "rgba(249,115,22,0.07)" : WHITE, color: planFilter === p ? O : MUTED, fontSize: 12, fontWeight: planFilter === p ? 700 : 400, cursor: "pointer", fontFamily: "inherit", textTransform: "capitalize" }}>
                    {p === "all" ? "All" : p}
                  </button>
                ))}
              </div>
              <div style={{ marginLeft: "auto", fontSize: 13, color: MUTED }}>{filtered.length} users</div>
            </div>

            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr style={{ background: BG }}>
                    {["Name", "Email", "Plan", "Credits", "Joined"].map(h => (
                      <th key={h} style={{ padding: "11px 20px", fontSize: 11, fontWeight: 700, color: FAINT, textAlign: "left", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${BORDER}` }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((p, i) => (
                    <tr key={p.id} style={{ borderBottom: i < filtered.length - 1 ? `1px solid ${BORDER}` : "none" }}>
                      <td style={{ padding: "13px 20px", fontSize: 14, fontWeight: 600, color: DARK }}>{p.name || "—"}</td>
                      <td style={{ padding: "13px 20px", fontSize: 13, color: MUTED }}>{p.email}</td>
                      <td style={{ padding: "13px 20px" }}>
                        <span style={{ fontSize: 11, fontWeight: 700, color: PLAN_COLORS[p.plan] || FAINT, background: PLAN_BG[p.plan] || "#F1F5F9", padding: "3px 10px", borderRadius: 20, textTransform: "capitalize" }}>{p.plan || "starter"}</span>
                      </td>
                      <td style={{ padding: "13px 20px", fontSize: 13, color: p.credits > 0 ? G : "#EF4444", fontWeight: 600 }}>{p.credits ?? 0}</td>
                      <td style={{ padding: "13px 20px", fontSize: 13, color: MUTED }}>{new Date(p.created_at).toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" })}</td>
                    </tr>
                  ))}
                  {filtered.length === 0 && (
                    <tr><td colSpan={5} style={{ padding: "40px", textAlign: "center", color: FAINT, fontSize: 14 }}>No users found</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
    </div>
  );
}

// ─── PAYMENT PAGE ─────────────────────────────────────────────────────────────

function PaymentPage({ user, setUser, setPage }) {
  const [loading, setLoading] = useState(null);
  const [success, setSuccess] = useState(null);
  const [showCancel, setShowCancel] = useState(false);
  const [cancelling, setCancelling] = useState(false);
  const [cancelDone, setCancelDone] = useState(false);

  const handleCancelSubscription = async () => {
    setCancelling(true);
    try {
      await fetch(`mailto:hello@rezolt.in`);
    } catch { }
    const subject = encodeURIComponent("Cancel Unlimited Subscription — " + (user?.email || "Account"));
    const body = encodeURIComponent(
      `Hi Rezolt team,\n\nPlease cancel my Unlimited subscription effective immediately.\n\nAccount email: ${user.email}\nUser ID: ${user.id}\n\nThank you.`
    );
    window.open(`mailto:hello@rezolt.in?subject=${subject}&body=${body}`, "_blank");
    setCancelDone(true);
    setCancelling(false);
  };

  const syncPaymentProfile = async (plan) => {
    for (let attempt = 0; attempt < 5; attempt++) {
      const profile = await fetchProfile(user.id).catch(() => null);
      const latestCredits = profile?.credits ?? user.credits ?? 0;
      const latestPlan = normalizePlan(profile?.plan) ?? user?.plan;
      const synced = plan.type === "subscription"
        ? latestPlan === "unlimited"
        : latestCredits > (user.credits ?? 0);

      if (synced) {
        setUser(prev => ({ ...prev, credits: latestCredits, plan: latestPlan, name: profile?.name || prev?.name }));
        setSuccess({ ...plan, synced: true });
        return true;
      }

      await new Promise(resolve => window.setTimeout(resolve, 1200));
    }

    setSuccess({ ...plan, synced: false });
    return false;
  };

  const PLANS = [
    {
      id: "starter_kit",
      type: "one-time",
      name: "Starter",
      price: 99,
      tag: "1 Career kit",
      credits: 1,
      plan: "starter",
      popular: false,
      includes: ["1 complete career kit", "Resume rewrite", "Cover letter + Referral DM", "Interview Prep + Find & Reach", "PDF & DOCX export"],
    },
    {
      id: "Pro_kit",
      type: "one-time",
      name: "Pro",
      price: 299,
      tag: "5 Career kits",
      credits: 5,
      plan: "Pro",
      popular: true,
      includes: ["5 complete career kits", "Resume rewrite", "Cover letter + Referral DM", "Interview Prep + Find & Reach", "PDF & DOCX export"],
    },
    {
      id: "unlimited_monthly",
      type: "subscription",
      name: "Unlimited",
      price: 599,
      tag: "per month",
      credits: 999,
      plan: "unlimited",
      popular: false,
      includes: ["Unlimited career kits", "All 5 kit outputs", "Salary Negotiation tool", "PDF & DOCX export", "Cancel anytime"],
    },
  ];

  const handlePay = async (plan) => {
    setLoading(plan.id);
    try {
      await openRazorpay({
        planId: plan.id,
        amount: plan.price,
        name: plan.id,
        description: `Rezolt ${plan.name} — ${plan.tag}`,
        prefill: { name: user.name, email: user.email, userId: user.id },
        onSuccess: async () => {
          const synced = await syncPaymentProfile(plan);
          if (!synced && plan.plan) {
            setUser(prev => ({
              ...prev,
              credits: (prev.credits ?? 0) + plan.credits,
              plan: plan.plan,
            }));
          }
        },
        onDismiss: async () => {
          const profile = await fetchProfile(user.id).catch(() => null);
          const maybeUpdated = plan.type === "subscription"
            ? (profile?.plan ?? user.plan) === "unlimited"
            : (profile?.credits ?? 0) > (user.credits ?? 0);
          if (maybeUpdated) {
            setUser(prev => ({ ...prev, credits: profile?.credits ?? prev?.credits ?? 0, plan: profile?.plan ?? prev?.plan ?? "starter", name: profile?.name || prev?.name }));
            setSuccess({ ...plan, synced: true });
          }
        },
      });
    } finally {
      setLoading(null);
    }
  };

  if (success) return (
    <div className="fade-in" style={{ minHeight: "calc(100vh - 63px)", display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
      <div style={{ textAlign: "center", maxWidth: 420 }}>
        <div style={{ fontSize: 52, marginBottom: 20 }}><i className="fa-solid fa-circle-check" /></div>
        <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 28, fontWeight: 500, color: DARK, marginBottom: 10 }}>Payment successful!</div>
        <div style={{ fontSize: 15, color: MUTED, marginBottom: 8 }}>{success.synced ? <>You're now on the <strong style={{ color: O }}>{success.name}</strong> plan.</> : <>We are verifying your <strong style={{ color: O }}>{success.name}</strong> payment.</>}</div>
        <div style={{ fontSize: 14, color: MUTED, marginBottom: 32 }}>{success.synced ? (success.type === "subscription" ? "Unlimited kits unlocked." : `${success.credits} credit${success.credits > 1 ? "s" : ""} added to your account.`) : "If the browser closed during checkout, your verified payment will still sync automatically when you reopen Rezolt."}</div>
        <button onClick={() => setPage("generate")} style={{ background: O, color: WHITE, border: "none", borderRadius: 10, padding: "14px 32px", fontSize: 15, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", boxShadow: "0 2px 12px rgba(249,115,22,0.3)" }}>
          Generate Your Kit
        </button>
      </div>
    </div>
  );

  return (
    <div className="fade-in page-pad" style={{ maxWidth: "100%", padding: "40px clamp(16px,5vw,80px) 80px" }}>
      {RZP_IS_TEST && (
        <div style={{ background: "rgba(239,68,68,0.06)", border: "1px solid rgba(239,68,68,0.25)", borderRadius: 12, padding: "12px 18px", marginBottom: 28, fontSize: 13, color: "#B91C1C", display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontWeight: 700 }}>⚠ Test mode active</span> — VITE_RAZORPAY_KEY_ID is a test key. Update it in Vercel and redeploy before accepting real payments.
        </div>
      )}
      <div style={{ marginBottom: 36, textAlign: "center" }}>
        <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.15em", textTransform: "uppercase", color: O, marginBottom: 12 }}>Upgrade</div>
        <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 32, fontWeight: 500, color: DARK, marginBottom: 8 }}>Choose your plan</div>
        <div style={{ fontSize: 14, color: MUTED }}>You currently have <strong style={{ color: (user?.credits ?? 0) > 0 ? G : "#EF4444" }}>{user?.credits ?? 0} credit{(user?.credits ?? 0) !== 1 ? "s" : ""}</strong> remaining</div>
      </div>

      <div className="pricing-grid" style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 20, maxWidth: 860, margin: "0 auto" }}>
        {PLANS.map((p) => (
          <div key={p.id} style={{ background: p.popular ? "rgba(249,115,22,0.04)" : WHITE, border: `1.5px solid ${p.popular ? O : BORDER}`, borderRadius: 16, padding: "32px 26px", position: "relative", display: "flex", flexDirection: "column" }}>
            {p.popular && <div style={{ position: "absolute", top: -13, left: "50%", transform: "translateX(-50%)", background: O, color: WHITE, fontSize: 11, fontWeight: 700, padding: "3px 14px", borderRadius: 20 }}>BEST VALUE</div>}
            <div style={{ fontSize: 12, fontWeight: 700, color: FAINT, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 10 }}>{p.name}</div>
            <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 38, fontWeight: 500, color: p.popular ? O : DARK, lineHeight: 1, marginBottom: 4 }}>₹{p.price}</div>
            <div style={{ fontSize: 13, color: FAINT, marginBottom: 20 }}>{p.tag}</div>
            <div style={{ flex: 1, marginBottom: 24 }}>
              {p.includes.map((f, j) => (
                <div key={j} style={{ display: "flex", gap: 8, marginBottom: 9 }}>
                  <span style={{ color: G, fontSize: 13, flexShrink: 0, marginTop: 1 }}>✓</span>
                  <span style={{ fontSize: 13, color: MID }}>{f}</span>
                </div>
              ))}
            </div>
            <button onClick={() => handlePay(p)} disabled={!!loading} style={{
              width: "100%", background: loading === p.id ? "#E2E8F0" : p.popular ? O : "transparent",
              color: loading === p.id ? FAINT : p.popular ? WHITE : O,
              border: `1.5px solid ${loading === p.id ? BORDER : O}`,
              borderRadius: 10, padding: "13px", fontSize: 14, fontWeight: 700,
              cursor: loading ? "not-allowed" : "pointer", fontFamily: "inherit",
              boxShadow: p.popular && !loading ? "0 2px 10px rgba(249,115,22,0.3)" : "none",
              transition: "all 0.2s",
            }}>
              {loading === p.id ? "Opening payment..." : `Pay ₹${p.price}`}
            </button>
            {p.type === "subscription" && <div style={{ textAlign: "center", fontSize: 11, color: FAINT, marginTop: 8 }}>Billed monthly · Cancel anytime</div>}
          </div>
        ))}
      </div>

      <div style={{ textAlign: "center", marginTop: 24, fontSize: 12, color: MUTED, lineHeight: 1.7 }}>
        If checkout closes midway, verified payments still sync automatically to your account.
      </div>

      {user?.plan === "unlimited" && (
        <div style={{ textAlign: "center", marginTop: 28 }}>
          {!showCancel ? (
            <button onClick={() => setShowCancel(true)} style={{ background: "none", border: "none", color: MUTED, fontSize: 12, cursor: "pointer", fontFamily: "inherit", textDecoration: "underline" }}>Cancel Unlimited subscription</button>
          ) : (
            <div style={{ background: WHITE, border: `1.5px solid ${BORDER}`, borderRadius: 14, padding: "20px 24px", maxWidth: 400, margin: "0 auto", boxShadow: "var(--soft-shadow)" }}>
              <div style={{ fontSize: 15, fontWeight: 700, color: DARK, marginBottom: 8 }}>Cancel your subscription?</div>
              <div style={{ fontSize: 13, color: MUTED, lineHeight: 1.7, marginBottom: 16 }}>Your plan stays active until the end of this billing cycle. Clicking below opens an email to our team — cancellations are processed within 1 business day.</div>
              {cancelDone ? (
                <div style={{ fontSize: 13, color: "#15803D", fontWeight: 700 }}>✓ Cancellation email opened. We will process it within 1 business day.</div>
              ) : (
                <div style={{ display: "flex", gap: 10, justifyContent: "center", flexWrap: "wrap" }}>
                  <button onClick={handleCancelSubscription} disabled={cancelling} style={{ background: "#EF4444", color: WHITE, border: "none", borderRadius: 9, padding: "10px 20px", fontSize: 13, fontWeight: 700, cursor: cancelling ? "not-allowed" : "pointer", fontFamily: "inherit" }}>
                    {cancelling ? "Opening email..." : "Send cancellation request"}
                  </button>
                  <button onClick={() => setShowCancel(false)} style={{ background: WHITE, color: MUTED, border: `1px solid ${BORDER}`, borderRadius: 9, padding: "10px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>Keep plan</button>
                </div>
              )}
            </div>
          )}
        </div>
      )}

      <div style={{ textAlign: "center", marginTop: 18 }}>
        <button onClick={() => setPage("dashboard")} style={{ background: "none", border: "none", color: MUTED, fontSize: 13, cursor: "pointer", fontFamily: "inherit" }}>Back to Dashboard</button>
      </div>
    </div>
  );
}



async function fetchAdminStats(onStep = () => { }) {
  const start = performance.now();
  try {
    const mark = (msg) => {
      const time = ((performance.now() - start) / 1000).toFixed(2);
      console.log(`[AdminFetch] ${time}s - ${msg}`);
      onStep(msg);
    };

    mark("Fetching user list...");
    const usersRes = await supabase.from("profiles").select("id");
    if (usersRes.error) console.error("Admin user list error:", usersRes.error);

    mark("Fetching kits list...");
    const kitsRes = await supabase.from("kits").select("id");
    if (kitsRes.error) console.error("Admin kits list error:", kitsRes.error);

    mark("Fetching profiles details...");
    const profilesRes = await supabase.from("profiles")
      .select("id, name, email, plan, credits, created_at")
      .order("created_at", { ascending: false })
      .limit(100);
    if (profilesRes.error) console.error("Admin profiles detail error:", profilesRes.error);

    mark("Load complete.");
    return {
      totalUsers: usersRes.data?.length || 0,
      totalKits: kitsRes.data?.length || 0,
      profiles: profilesRes.data || [],
    };
  } catch (e) {
    console.error("fetchAdminStats failed fatally:", e);
    onStep(`Error: ${e.message}`);
    return { totalUsers: 0, totalKits: 0, profiles: [] };
  }
}


async function fetchProfile(userId) {
  try {
    const { data, error } = await supabase.from("profiles").select("*").eq("id", userId).single();
    if (error) throw error;
    return data;
  } catch (e) {
    console.warn("Profile fetch failed:", e);
    return null;
  }
}

async function fetchKits(userId) {
  try {
    const { data, error } = await supabase.from("kits").select("*").eq("user_id", userId).order("created_at", { ascending: false });
    if (error) throw error;
    return data || [];
  } catch (e) {
    console.warn("Kits fetch failed:", e);
    return [];
  }
}

async function saveKit(userId, kit) {
  await supabase.from("kits").insert({ user_id: userId, role: kit.role, outputs: kit.outputs });
}

async function decrementCredits(userId, currentCredits) {
  const next = Math.max(0, currentCredits - 1);
  const { error } = await supabase.from("profiles").update({ credits: next }).eq("id", userId);
  if (error) console.error("Credit decrement failed:", error);
  return next;
}

// ─── ROOT ─────────────────────────────────────────────────────────────────────

// ─── PRIVACY POLICY PAGE ────────────────────────────────────────────────────

function PrivacyPage({ setPage }) {
  return (
    <div className="fade-in" style={{ background: "var(--bg)", minHeight: "calc(100vh - 85px)" }}>
      <div style={{ maxWidth: 760, margin: "0 auto", padding: "56px clamp(20px,5vw,40px) 80px" }}>
        <button onClick={() => setPage("landing")} style={{ background: "none", border: "none", color: MUTED, fontSize: 13, cursor: "pointer", padding: 0, marginBottom: 24, fontFamily: "inherit" }}>Back</button>
        <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: ".14em", textTransform: "uppercase", color: AC, marginBottom: 10 }}>Legal</div>
        <h1 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(28px,4vw,42px)", fontWeight: 400, color: "var(--text)", marginBottom: 8 }}>Privacy Policy</h1>
        <p style={{ fontSize: 13, color: MUTED, marginBottom: 36 }}>Last updated: April 2026 · Rezolt (operated by Hema Margam)</p>

        {[
          { heading: "What we collect", body: "When you create an account, we collect your name and email address. When you use the kit generator, we receive your resume text and job description text temporarily to generate your kit. We do not store your raw resume or JD text beyond the generation request." },
          { heading: "How we use your data", body: "Your email is used to authenticate your account and send transactional messages (email confirmation, password reset). Your resume and JD text is sent to Anthropic's API solely to produce your career kit output. Generated kit outputs are stored against your account so you can access them from your dashboard." },
          { heading: "Third-party processors", body: "We use Supabase for database and authentication, Anthropic for AI text generation, Razorpay for payment processing, and Vercel for hosting. Each processes only the data necessary for their function. Anthropic's own privacy policy applies to data submitted via their API." },
          { heading: "Data retention", body: "Your profile and kit outputs are retained while your account is active. If you request account deletion, we will remove your data within 30 days. Resume and JD text submitted for generation is not persisted beyond the API call." },
          { heading: "Your rights (DPDPA 2023)", body: "Under India's Digital Personal Data Protection Act 2023, you have the right to access the personal data we hold about you, correct inaccurate data, and request deletion of your data. To exercise any of these rights, write to us at hello@rezolt.in with the subject 'Data Request'." },
          { heading: "Cookies", body: "We use only essential session cookies required for authentication. We do not use advertising or analytics cookies." },
          { heading: "Contact", body: "For privacy queries or data principal requests, contact: hello@rezolt.in" },
        ].map(({ heading, body }) => (
          <div key={heading} style={{ marginBottom: 28, borderBottom: "1px solid var(--border)", paddingBottom: 28 }}>
            <h2 style={{ fontSize: 16, fontWeight: 700, color: "var(--text)", marginBottom: 10 }}>{heading}</h2>
            <p style={{ fontSize: 14, color: "var(--text-mid)", lineHeight: 1.8 }}>{body}</p>
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── TERMS OF SERVICE PAGE ───────────────────────────────────────────────────

function TermsPage({ setPage }) {
  return (
    <div className="fade-in" style={{ background: "var(--bg)", minHeight: "calc(100vh - 85px)" }}>
      <div style={{ maxWidth: 760, margin: "0 auto", padding: "56px clamp(20px,5vw,40px) 80px", textAlign: "left" }}>
        <button onClick={() => setPage("landing")} style={{ background: "none", border: "none", color: "var(--text-muted)", fontSize: 13, cursor: "pointer", padding: 0, marginBottom: 24, fontFamily: "inherit" }}>Back</button>
        <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: ".14em", textTransform: "uppercase", color: "var(--accent)", marginBottom: 10 }}>Legal</div>
        <h1 style={{ fontFamily: "'DM Serif Display', serif", fontSize: "clamp(28px,4vw,42px)", fontWeight: 400, color: "var(--text)", marginBottom: 8 }}>Terms of Service</h1>
        <p style={{ fontSize: 13, color: "var(--text-muted)", marginBottom: 36 }}>Last updated: April 2026 — By using Rezolt, you agree to these terms.</p>

        {[
          { heading: "1. Acceptance of terms", body: "By accessing or using Rezolt, you agree to be bound by these Terms of Service. If you do not agree to all terms and conditions, you may not access the platform or use our tools." },
          { heading: "2. The service", body: "Rezolt is an AI-assisted career kit generation platform. You provide your resume and a job description; our platform uses Anthropic's Claude API to produce tailored career documents including a resume rewrite, cover letter, referral messages, interview preparation, reach-out plans, and salary negotiation notes." },
          { heading: "3. Account registration", body: "You must provide accurate, current, and complete information during registration. You are responsible for all activities that occur under your account. Do not share your credentials." },
          { heading: "4. Accurate information only", body: "You agree that the resume and information you submit to Rezolt represents your genuine professional experience. You must not use Rezolt to fabricate, exaggerate, or misrepresent qualifications, roles, achievements, or credentials. Misuse for fraudulent applications is a violation of these terms." },
          { heading: "5. Credits and payments", body: "Credits are non-refundable once a kit generation has been attempted. The Starter plan includes 1 free kit. Paid credits and plans are non-transferable. Rezolt reserves the right to modify pricing with reasonable notice." },
          { heading: "6. No refund policy", body: "Due to the nature of AI-generated outputs, we do not offer refunds once generation has been initiated. If you experience a technical failure that genuinely prevents output delivery, contact hello@rezolt.in and we will review on a case-by-case basis." },
          { heading: "7. Service availability", body: "Rezolt relies on third-party APIs (Anthropic, Supabase, Razorpay) and hosting infrastructure (Vercel). We do not guarantee uninterrupted availability. Scheduled or unscheduled downtime does not entitle users to refunds or compensation." },
          { heading: "8. Intellectual property", body: "You retain ownership of your resume content and generated outputs. Rezolt retains ownership of the platform, prompts, templates, and interface. Generated content may not be used to train competing AI models." },
          { heading: "9. Acceptable use", body: "You may not use Rezolt to spam recruiters, generate bulk applications without genuine intent, or circumvent any employer's application systems. Accounts found abusing the platform will be suspended." },
          { heading: "10. Liability limitations", body: "In no event shall Rezolt or its operators be liable for indirect, incidental, special, or consequential damages (including loss of employment opportunities) arising out of the use or inability to use the platform." },
          { heading: "11. Governing law", body: "These terms are governed by the laws of India. Any disputes will be subject to the exclusive jurisdiction of courts in India." },
          { heading: "12. Contact", body: "For any queries related to these terms, write to us directly at hello@rezolt.in." },
        ].map(({ heading, body }) => (
          <div key={heading} style={{ marginBottom: 28, borderBottom: "1px solid var(--border)", paddingBottom: 28 }}>
            <h2 style={{ fontSize: 16, fontWeight: 700, color: "var(--text)", marginBottom: 10 }}>{heading}</h2>
            <p style={{ fontSize: 14, color: "var(--text-mid)", lineHeight: 1.8 }}>{body}</p>
          </div>
        ))}
      </div>
    </div>
  );
}

export default function App() {
  const [page, setPage] = useState("landing");
  const [user, setUser] = useState(null);
  const [history, setHistory] = useState([]);
  const [profileLoaded, setProfileLoaded] = useState(false);
  const [booting, setBooting] = useState(true);
  const [selectedTemplate, setSelectedTemplate] = useState("creative");
  const [selectedArticle, setSelectedArticle] = useState(0);
  const [sessionToken, setSessionToken] = useState(null);

  // Safety net: if user gets set while on the auth page, navigate to dashboard
  useEffect(() => {
    if (user && page === "auth") setPage("dashboard");
  }, [user]);

  // Refresh profile (credits + plan) whenever dashboard or generate page opens
  useEffect(() => {
    if (!["dashboard", "generate"].includes(page) || !user?.id || !sessionToken) return;
    const uid = user.id;
    fetchWithAuth("/api/get-profile", { method: "GET" }, sessionToken)
      .then(r => r.json())
      .then(data => {
        const profile = data?.profile;
        if (profile) setUser(prev => ({
          ...prev,
          credits: profile.credits ?? prev.credits,
          plan: normalizePlan(profile.plan) ?? prev.plan,
        }));
      })
      .catch(e => console.warn("Page-change profile refresh failed:", e));
    if (page === "dashboard") {
      fetchKits(uid).then(kits => {
        setHistory(kits.map(k => ({
          role: k.role,
          outputs: k.outputs,
          date: new Date(k.created_at).toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" }),
        })));
      });
    }
  }, [page, user?.id]);

  // Restore session on load
  useEffect(() => {
    let cancelled = false;
    // Never allow infinite loading if Supabase session lock stalls.
    const bootTimeoutId = setTimeout(() => {
      if (cancelled) return;
      setPage("landing");
      setBooting(false);
    }, 5000);

    const syncSessionUser = async (authUser, nextPage) => {
      try {
        // Yield execution thread by 150ms so Supabase can safely resolve and drop its native initialization locks before Database queries run
        await new Promise(r => setTimeout(r, 150));

        // Ensure login feels instant, then hydrate richer data in background.
        if (!cancelled) {
          setUser({ id: authUser?.id, name: authUser?.email?.split("@")[0], email: authUser?.email, credits: 1, plan: "Free" });
          if (nextPage) setPage(nextPage);
        }

        // Sequentialize initial fetches to prevent Supabase v2 client auto-refresh lock collision
        const profile = await fetchProfile(authUser?.id);
        const kits = await fetchKits(authUser?.id);
        if (cancelled) return;
        setUser(prev => ({
          ...(prev || {}),
          id: authUser.id,
          name: profile?.name || prev?.name || authUser.email.split("@")[0],
          email: authUser.email,
          credits: profile?.credits ?? prev?.credits ?? 1,
          plan: normalizePlan(profile?.plan) ?? prev?.plan ?? "Free",
        }));
        setHistory(kits.map(k => ({ role: k.role, outputs: k.outputs, date: new Date(k.created_at).toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" }) })));
        setProfileLoaded(true);
      } catch (e) {
        console.error("Session sync failed:", e);
        if (!cancelled) setPage(prev =>
          (prev === "dashboard" || prev === "generate" || prev === "admin") ? prev : "landing"
        );
      }
    };

    const recoveryFlow = typeof window !== "undefined" && /type=recovery/i.test(`${window.location.hash}${window.location.search}`);

    const { data: { subscription } } = supabase.auth.onAuthStateChange(async (event, session) => {
      if (session?.access_token) setSessionToken(session.access_token);
      // Clear boot lock the moment Supabase answers natively
      if (!cancelled) {
        clearTimeout(bootTimeoutId);
        setBooting(false);
      }

      if (event === "SIGNED_OUT") {
        setUser(null);
        setHistory([]);
        setPage("landing");
        return;
      }

      if (!session?.user) {
        if (event === "INITIAL_SESSION") setPage(prev => prev === "loading" ? "landing" : prev);
        return;
      }

      if (event === "PASSWORD_RECOVERY") {
        await syncSessionUser(session.user, "reset-password");
        return;
      }

      if (event === "INITIAL_SESSION" || event === "TOKEN_REFRESHED") {
        // SIGNED_IN is handled by handleAuth (called directly from AuthPage) which uses
        // the API route for profile fetching. Running syncSessionUser for SIGNED_IN too
        // causes a race where syncSessionUser overwrites the correctly-set plan with "Free".
        await syncSessionUser(session.user, event === "INITIAL_SESSION" ? (recoveryFlow ? "reset-password" : "dashboard") : undefined);
      }
    });

    // Fallback: If no cache exists, INITIAL_SESSION may silently not broadcast fast enough. 
    // This perfectly bypasses the old lock conflict because it fails silently backwards to 'landing'.
    supabase.auth.getSession().then(({ data: { session } }) => {
      if (session?.access_token) setSessionToken(session.access_token);
      if (session?.access_token) setSessionToken(session.access_token);
      if (!cancelled && !session?.user) setPage(p => p === "loading" ? "landing" : p);
    }).catch(() => {
      if (!cancelled) setPage(p => p === "loading" ? "landing" : p);
    });

    return () => {
      cancelled = true;
      clearTimeout(bootTimeoutId);
      subscription?.unsubscribe();
    };
  }, []);

  const handleAuth = async (u, token) => {
    // Render dashboard immediately; avoid blocking on history/profile reads.
    setUser(u);
    if (token) setSessionToken(token);
    setPage(prev => prev === "payment" ? "payment" : "dashboard");

    const authToken = token || sessionToken;
    let profile = null;
    try {
      if (authToken) {
        const res = await fetchWithAuth("/api/get-profile", { method: "GET" }, authToken);
        const data = await res.json().catch(() => ({}));
        profile = data?.profile ?? null;
      }
      if (!profile) profile = await fetchProfile(u.id);
    } catch (e) {
      console.warn("handleAuth profile fetch failed:", e);
    }
    const kits = await fetchKits(u.id);

    setHistory(kits.map(k => ({ role: k.role, outputs: k.outputs, date: new Date(k.created_at).toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" }) })));
    if (profile) {
      setUser(prev => ({
        ...prev,
        name: profile?.name || prev?.name,
        credits: profile?.credits ?? prev?.credits ?? 1,
        plan: normalizePlan(profile?.plan) ?? prev?.plan ?? "Free",
      }));
    }
    setProfileLoaded(true);
  };

  const handleSignOut = async () => {
    try {
      // Forcefully clear local state first so UI feels instant
      setUser(null);
      setHistory([]);
      setPage("landing");
      setSessionToken(null);
      setProfileLoaded(false);

      // Attempt to tell Supabase, but don't hang if the lock is stuck
      await Promise.race([
        supabase.auth.signOut(),
        new Promise(r => setTimeout(r, 1000))
      ]);

      // Manual cleanup of storage keys just in case
      Object.keys(localStorage).forEach(key => {
        if (key.includes("supabase") || key.includes("sb-")) localStorage.removeItem(key);
      });
    } catch (e) {
      console.warn("Sign out encountered an error, but local state was cleared.");
    }
  };

  const handleSaveKit = async (kit) => {
    setHistory(prev => [kit, ...prev]);
    if (!user?.id || !sessionToken) return;
    try {
      await fetchWithAuth("/api/save-kit", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ role: kit.role, outputs: kit.outputs }),
      }, sessionToken);
    } catch (err) {
      console.error("Kit save failed:", err.message);
    }
  };

  const handleUseCredit = async () => {
    if (!user?.id || user?.plan === "unlimited" || !sessionToken) return;
    try {
      const res = await fetchWithAuth("/api/use-credit", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({}),
      }, sessionToken);
      const data = await res.json();
      if (data.credits !== undefined && data.credits !== null) {
        setUser(prev => ({ ...prev, credits: data.credits }));
      }
    } catch (err) {
      console.error("Credit update failed:", err.message);
    }
  };

  const handleBuyCredits = async () => {
    setPage("payment");
  };

  if (booting) return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: N1 }}>
      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 20 }}>
        <div style={{ padding: "12px 18px", borderRadius: 18, background: "rgba(255,255,255,0.96)", boxShadow: "0 10px 30px rgba(0,0,0,0.12)" }}>
          <BrandLogo height={90} />
        </div>
        <div style={{ display: "flex", gap: 6 }}>
          {[0, 1, 2].map(i => <div key={i} style={{ width: 8, height: 8, borderRadius: "50%", background: LB, animation: `dotPulse 1.2s ease ${i * 0.18}s infinite` }} />)}
        </div>
      </div>
    </div>
  );

  return (
    <>
      <style>{GLOBAL_CSS}</style>
      <TopBar page={page} setPage={setPage} user={user} onSignOut={handleSignOut} />
      {page === "landing" && <LandingPage setPage={setPage} user={user} selectedTemplate={selectedTemplate} setSelectedTemplate={setSelectedTemplate} setSelectedArticle={setSelectedArticle} />}
      {page === "articles" && <ArticlesPage setPage={setPage} user={user} selectedArticle={selectedArticle} setSelectedArticle={setSelectedArticle} />}
      {page === "faq" && <FAQPage setPage={setPage} user={user} />}
      {page === "contact" && <ContactPage setPage={setPage} user={user} />}
      {page === "privacy" && <PrivacyPage setPage={setPage} />}
      {page === "terms" && <TermsPage setPage={setPage} />}
      {page === "auth" && <AuthPage onAuth={handleAuth} setPage={setPage} />}
      {page === "forgot" && <ForgotPasswordPage setPage={setPage} />}
      {page === "reset-password" && <ResetPasswordPage setPage={setPage} />}
      {page === "admin" && user && <AdminPage user={user} setPage={setPage} />}
      {page === "payment" && (user ? <PaymentPage user={user} setUser={setUser} setPage={setPage} /> : <AuthPage onAuth={handleAuth} setPage={setPage} />)}
      {page === "dashboard" && user && <Dashboard user={user} history={history} setPage={setPage} onBuyCredits={handleBuyCredits} profileLoaded={profileLoaded} />}
      {page === "generate" && user && <KitGenerator sessionToken={sessionToken} user={user} setUser={setUser} onSaveKit={handleSaveKit} onUseCredit={handleUseCredit} setPage={setPage} selectedTemplate={selectedTemplate} setSelectedTemplate={setSelectedTemplate} />}
      {(page === "dashboard" || page === "generate" || page === "admin") && !user && <AuthPage onAuth={handleAuth} setPage={setPage} />}

      {/* Mobile bottom nav — logged in users only */}
      {user && (
        <div className="mobile-only" style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: "var(--surface)", borderTop: "1px solid var(--border)", display: "flex", zIndex: 200, boxShadow: "0 -4px 20px rgba(0,27,72,.1)", paddingBottom: "env(safe-area-inset-bottom)" }}>
          {[
            { icon: <i className="fa-solid fa-house" />, label: "Home", pg: "landing" },
            { icon: <FiLayout />, label: "Dashboard", pg: "dashboard" },
            { icon: <FiEdit3 />, label: "Create", pg: "generate" },
            { icon: <i className="fa-solid fa-dollar-sign" />, label: "Plans", pg: "payment" },
          ].map(item => (
            <button key={item.pg} onClick={() => setPage(item.pg)} style={{
              flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
              padding: "10px 4px 8px", background: "none", border: "none", cursor: "pointer", fontFamily: "inherit",
              borderTop: page === item.pg ? `2px solid ${AC}` : "2px solid transparent",
              transition: "all .15s ease",
            }}>
              <span style={{ fontSize: 18, lineHeight: 1, marginBottom: 4, display: "inline-flex", alignItems: "center" }}>{item.icon}</span>
              <span style={{ fontSize: 10, fontWeight: 600, color: page === item.pg ? AC : "var(--text-muted)" }}>{item.label}</span>
            </button>
          ))}
        </div>
      )}
    </>
  );
}
