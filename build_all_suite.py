import os
import zipfile
import xml.etree.ElementTree as ET

DATE_PREFIX = "2026-11-09"
TEMPLATE_DOCX = "2026-09-22_Original_Resume_Santheepkumar_V.docx"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def qn(tag):
    return f"{{{W_NS}}}{tag}"

def qn_w14(tag):
    return f"{{{W14_NS}}}{tag}"

para_id_counter = 100

def next_para_id():
    global para_id_counter
    p_id = f"{para_id_counter:08X}"
    para_id_counter += 1
    return p_id

# ==============================================================================
# SHARED DATA & ROLE CONFIGURATIONS
# ==============================================================================

COMMON_SKILLS = {
    "Cloud, K8s & DevOps": "Kubernetes (K8s), Docker, AWS (ECS, Lambda, S3, RDS, CloudFront, IAM), GCP, GitHub Actions, Bitbucket Pipelines, Jenkins, Vercel, Terraform / IaC, Sentry, Axiom",
    "Architecture & Systems": "Microservices, Inter-Service Communication, Event-Driven Systems, Clean/Hexagonal Architecture, Monorepos (Turborepo), Auth/RBAC, Notification Engines, Redis, Kafka, RabbitMQ",
    "Frontend Engineering": "React, Next.js (App & Pages Router), TypeScript, JavaScript (ES6+), SvelteKit / Svelte 5, React Native, Expo, Tailwind CSS, State Management (Zustand, Redux)",
    "Backend & Languages": "Node.js, NestJS, Go (Golang), Bun, Express.js, RESTful APIs, GraphQL, WebCrypto / AES-256-GCM Encryption, IoT Hardware & POS Peripherals",
    "Databases & Storage": "PostgreSQL (Raw SQL, RLS, FTS/tsvector, SKIP LOCKED), MongoDB, Supabase, Drizzle ORM, TypeORM, Prisma, Sequelize, Mongoose, Query Tuning",
    "Testing, AI & Quality": "Playwright, Vitest, Jest, React Testing Library, LLM RAG & Bounded Agents, Automated Eval Harnesses, CI/CD Quality Gates"
}

STRAW_HAT_PLATFORM_BULLETS = [
    ("StrawHatDevs Learning Platform (strawhatdevs.com): ", "Conceived, architected, and built an interactive developer education platform from scratch (Next.js 15, React 19, TypeScript, Supabase, PostgreSQL, Sanity CMS, Axiom) teaching full-stack engineering, system design, and cloud/DevOps."),
    ("YouTube Educator (200+ Architecture & Coding Videos): ", "Produced and published 200+ deep-dive technical videos breaking down distributed systems, microservices, cloud deployments, and production engineering for thousands of developers.")
]

INDEPENDENT_PRODUCTS = [
    {
        "title": "ChimeFlow — Multi-Tenant Omnichannel Notification & Cadence Platform",
        "role_or_domain": "Creator & Architect",
        "date": "2025 – Present",
        "desc": "Engineered a multi-tenant notification & cadence orchestration platform supporting Email (SendGrid), SMS (Twilio), and Push (Firebase FCM) using Clean/Hexagonal Architecture. Built a write-ahead delivery ledger, fair-claim background worker queues, idempotency keys, per-tenant rate limiting, suppression/consent engines, encrypted credentials at rest, and Razorpay billing.",
        "stack": "Bun, TypeScript, SvelteKit / Svelte 5, PostgreSQL 17, Clerk Auth, SendGrid, Twilio, Firebase Admin, Docker, AWS"
    },
    {
        "title": "ENVault — Zero-Knowledge Encrypted .env Sync Platform & Native Go CLI",
        "role_or_domain": "Creator & Architect (envault.strawhatdevs.com)",
        "date": "2025 – Present",
        "desc": "Built an end-to-end client-encrypted team environment variable & secret synchronization platform with browser WebCrypto key derivation, atomic revision rollback, and a native cross-platform Go CLI (github.com/Santheepkumar/envault-cli) integrated with macOS Keychain, Windows Credential Manager, and Linux Secret Service.",
        "stack": "Go (Golang CLI), Next.js, TypeScript, PostgreSQL 17, Drizzle ORM, WebCrypto API, Better Auth, Docker"
    },
    {
        "title": "IncidentLens AI — Evidence-Grounded Incident Triage & Runbook RAG Engine",
        "role_or_domain": "Creator & Architect",
        "date": "2025 – Present",
        "desc": "Designed a multi-tenant AI production incident investigator connecting Sentry telemetry with Markdown runbooks (PostgreSQL tsvector + GIN Hybrid FTS RAG), AES-256-GCM secret vault, a bounded 6-step agent loop with deterministic [ev_xxx] citation verification, FOR UPDATE SKIP LOCKED job queues, and a 35-case AI evaluation harness (93.3% root-cause accuracy).",
        "stack": "TypeScript, Turborepo, Node.js, PostgreSQL (Raw SQL + RLS), Auth0 OAuth2 PKCE, LLM Tool-Calling & RAG, Vitest"
    },
    {
        "title": "StrawHatDevs Platform (strawhatdevs.com) & Route-Atlas Open-Source Tooling",
        "role_or_domain": "Founder & Creator",
        "date": "Mar 2025 – Present",
        "desc": "Architected strawhatdevs.com full-stack learning platform (200+ architecture/programming videos) alongside developer tools including Route-Atlas (published npm visual atlas for Next.js App Router) and SelectCheck (Tauri v2 + Cloudflare Workers AI assistant).",
        "stack": "Next.js 15, React 19, TypeScript, Supabase, PostgreSQL, Sanity CMS, Tauri v2, Cloudflare Workers, Vercel CI/CD"
    }
]

ENTERPRISE_PROJECTS = [
    {
        "title": "Science & Humans — Healthcare Platform & Microservices Ecosystem",
        "role_or_domain": "Lead Software Engineer (Contractor)",
        "date": "Jun 2025 – Present",
        "desc": "Built core platform architecture from the ground up with full end-to-end ownership: engineered the complete authentication & RBAC authorization system, event-driven notification engine, decoupled microservices & inter-service communication layer, and production Kubernetes (K8s) + CI/CD DevOps infrastructure while mentoring junior developers.",
        "stack": "TypeScript, Node.js, Microservices, Auth/RBAC, Notification System, Kubernetes (K8s), Docker, CI/CD"
    },
    {
        "title": "Pony Up — Android Self-Service IoT Retail Kiosk & Cloud Web Platform",
        "role_or_domain": "Staff Engineer | Tuna Software Solution",
        "date": "Jun 2025 – Present",
        "desc": "Engineered a React Native/Expo self-service kiosk deployed in shopping malls integrating physical IoT hardware (payment terminals, thermal receipt printers, cash drawers), and subsequently led its strategic rebuild into a cloud-hosted web platform.",
        "stack": "React Native, Expo, Next.js, IoT Peripherals, Android Kiosk Mode, REST APIs, Docker, GitHub Actions"
    },
    {
        "title": "Chuck E. Cheese — Customer Experience Web Platform (2-Month Solo MVP)",
        "role_or_domain": "Senior Engineer | Heady",
        "date": "2024 – 2025",
        "desc": "Rebuilt the Chuck E. Cheese kiosk into a mobile-responsive customer experience web platform enabling instant product interaction via QR code scanning—delivered the production MVP solo in 2 months vs. an 8-month original estimate (75% faster time-to-market).",
        "stack": "React, Next.js, TypeScript, Tailwind CSS, QR Integration, AWS CloudFront, Automated CI/CD"
    },
    {
        "title": "HeyDay Skin Care & The Lash Lounge — Enterprise Booking & Microservices",
        "role_or_domain": "Lead Frontend / Full-Stack Architect | Heady",
        "date": "May 2023 – Jun 2025",
        "desc": "Architected high-concurrency appointment booking widgets, e-commerce shop apps, and microservices backends integrated with Boulevard API, Redis caching, Sentry monitoring, and AWS infrastructure sustaining 99.9% uptime.",
        "stack": "React, Next.js, Node.js, NestJS, Redis, Microservices, AWS, Vercel, Bitbucket Pipelines, Sentry, Boulevard API"
    }
]

def get_experiences(role_variant):
    # Role-specific framing for Tuna, Science & Humans, Heady, HOD Digitals, Skematix
    return [
        {
            "role": "Staff Engineer" if role_variant != "fde" else "Staff Engineer / Forward Deployed Lead",
            "company": "Tuna Software Solution",
            "date": "Jun 2025 – Present",
            "bullets": [
                ("End-to-End Architecture & Delivery: ", "Owned full-stack architecture, cloud infrastructure, backend microservices, and frontend platforms across two concurrent enterprise initiatives, driving feature releases on schedule."),
                ("Pony Up IoT Kiosk & Cloud Web Migration: ", "Architected and delivered an Android self-service kiosk (React Native/Expo) integrating physical IoT peripherals (payment terminals, receipt printers, cash drawers), then led its rebuild as a cloud web platform to eliminate deployment friction and improve cross-device compatibility."),
                ("CI/CD Automation & Engineering Standards: ", "Established structured code review processes, automated CI/CD pipelines (GitHub Actions/Docker), and coding standards—boosting team delivery velocity by 35% and reducing production defect escapes."),
                ("Stakeholder & Product Alignment: ", "Collaborated directly with product stakeholders to translate business requirements into scalable technical solutions, acting as the primary technical point of contact.")
            ]
        },
        {
            "role": "Lead Software Engineer (Part-Time Contractor)",
            "company": "Science & Humans",
            "date": "Jun 2025 – Present",
            "bullets": [
                ("Ground-Up Ownership & Product Planning: ", "Contributed from the ground up with full end-to-end ownership—partnering directly with Product Managers to plan implementations, estimate modules, architect solutions, and ship production releases."),
                ("Authentication, Authorization & Notification Engine: ", "Architected and built the entire authentication and role-based authorization (RBAC) system alongside a centralized multi-channel notification system for the platform."),
                ("Microservices & Communication Layer: ", "Set up decoupled backend microservices and engineered the inter-service communication layer from scratch while resolving critical system bottlenecks and bugs to harden platform stability."),
                ("Kubernetes (K8s) & DevOps Configuration: ", "Configured and managed containerized application deployments on Kubernetes (K8s) and built automated CI/CD DevOps workflows for reliable, zero-downtime releases."),
                ("Mentoring & Team Ramp-Up: ", "Guided, onboarded, and mentored junior developers joining the project, conducting code reviews and unblocking engineering execution across modules.")
            ]
        },
        {
            "role": "Senior Software Engineer" if role_variant != "fde" else "Senior Forward Deployed Engineer & Client Lead",
            "company": "Heady",
            "date": "May 2023 – Jun 2025",
            "bullets": [
                ("Enterprise Client Leadership: ", "Served as primary technical lead interfacing with tier-1 enterprise clients (HeyDay Skin Care, Chuck E. Cheese, The Lash Lounge), coordinating with PMs and stakeholders to define scope, set milestones, and ship production systems."),
                ("Chuck E. Cheese Web MVP (75% Faster Delivery): ", "Delivered the Chuck E. Cheese web kiosk MVP in 2 months as a solo contributor, compressing an 8-month original project timeline by 75% through focused full-stack engineering and rapid iteration."),
                ("HeyDay Skin Care Microservices Ecosystem: ", "Architected a microservices-based backend integration with Boulevard (booking platform), Redis caching, Sentry error monitoring, and AWS infrastructure sustaining 99.9% reliability."),
                ("AI-Assisted Workflows & Mentorship: ", "Mentored 6+ junior and mid-level developers on modern development practices, AI-assisted coding workflows, and architectural patterns that improved code quality and reduced regression rates.")
            ]
        },
        {
            "role": "Principal Software Engineer (Promoted from Senior Frontend Engineer)",
            "company": "HOD Digitals",
            "date": "Sep 2021 – May 2023",
            "bullets": [
                ("Fast-Track Promotion (Senior Engineer → Principal Engineer): ", "Joined as Senior Frontend Engineer contributing to Project Gen and concurrent client systems; promoted to Principal Software Engineer (Nov 2022) to lead technical direction and architecture across multiple product lines."),
                ("E-Commerce & Scalable System Architecture: ", "Architected and shipped full-stack e-commerce platforms (Tamil Mart web & iOS/Android mobile apps with Next.js, Magento GraphQL, React Native), CRM subscription systems (Vmeals on AWS), and Playwright CI/CD test automation (AMS)."),
                ("Cross-Functional Leadership & Mentorship: ", "Oversaw project planning, system design, and cloud infrastructure decisions while acting as primary technical liaison between engineering teams and business stakeholders.")
            ]
        },
        {
            "role": "Full Stack Developer (Converted from Intern in <3 Months)",
            "company": "Skematix Technologies",
            "date": "Dec 2019 – Sep 2021",
            "bullets": [
                ("Rapid Intern-to-Full-Time Conversion: ", "Joined during college in 2019 as an Intern Software Engineer and earned conversion to full-time Full Stack Developer within 2–3 months due to rapid execution and deep ownership."),
                ("End-to-End Web Application Delivery: ", "Developed, optimized, and deployed multiple client-facing web applications across React, Node.js, Express, and relational/NoSQL databases from initial schema design to production hosting.")
            ]
        }
    ]

RESUMES = [
    {
        "slug": "Principal_Software_Engineer",
        "headline": "Principal Software Engineer | Enterprise Systems Architect & Technical Leader",
        "accent_color": "#4338ca",
        "accent_dark": "#3730a3",
        "highlight_bg": "#eef2ff",
        "overview_p1": (
            "I’m a self-taught, obsessive software builder with 7+ years of experience taking complex systems from a blank repository to production cloud infrastructure. "
            "My journey started in a government school in Tamil Nadu on a 2GB RAM state-issued laptop, teaching myself programming through documentation and YouTube—a drive that took me from converting a college internship into a full-time role in under 3 months (Skematix), to fast-tracking from Senior Engineer to Principal Software Engineer (HOD Digitals), leading tier-1 enterprise deliveries at Heady (shipping a Chuck E. Cheese web MVP solo in 2 months vs. an 8-month estimate), and now owning end-to-end architecture as a Staff Engineer at Tuna Software Solution and Lead Software Engineer (Part-Time Contractor) at Science & Humans."
        ),
        "overview_p2": (
            "As a Principal-level architect and hands-on builder, I own the entire technical lifecycle: designing distributed microservices & inter-service communication layers, engineering authentication/RBAC and multi-channel notification engines from the ground up, orchestrating Kubernetes (K8s) and AWS infrastructure with automated CI/CD, and mentoring engineering teams to ship with high velocity (+35%). "
            "Outside of client work, I constantly build and ship my own production platforms—including StrawHatDevs (strawhatdevs.com + 200+ architecture and programming videos on YouTube), ChimeFlow (a multi-tenant omnichannel notification & cadence platform), ENVault (a zero-knowledge encrypted .env sync platform with a native Go CLI), and IncidentLens AI (an evidence-grounded incident triage & Runbook RAG engine)."
        ),
        "competencies": [
            "End-to-End System Architecture", "Kubernetes (K8s) & Cloud Infra", "Microservices & Comm Layers",
            "Auth, RBAC & Notification Engines", "Automated CI/CD & DevOps", "Multi-Pod Technical Leadership",
            "Rapid 0-to-1 MVP Execution", "IoT & Hardware Integrations", "Product Planning & Estimation",
            "Engineering Mentorship & Reviews", "Clean / Hexagonal Architecture", "AI Agents, RAG & Eval Harnesses"
        ],
        "experiences": get_experiences("principal")
    },
    {
        "slug": "Technical_Lead",
        "headline": "Technical Lead | Staff Full-Stack, Microservices & Cloud Architect",
        "accent_color": "#0f766e",
        "accent_dark": "#115e59",
        "highlight_bg": "#f0fdfa",
        "overview_p1": (
            "I’m a self-taught, hands-on Technical Lead and Staff Software Engineer with 7+ years of experience building software end-to-end and leading engineering teams from concept to production deployment. "
            "My love for programming started in a Tamil Nadu government school on a 2GB RAM state-scheme laptop, teaching myself code via Google and YouTube—a foundation that propelled me from converting a college internship into a full-time role in under 3 months (Skematix), to earning promotion from Senior Engineer to Principal Engineer (HOD Digitals), leading enterprise accounts at Heady (delivering Chuck E. Cheese’s 8-month web MVP solo in 2 months), and now serving as Staff Engineer at Tuna Software Solution and Lead Software Engineer (Part-Time Contractor) at Science & Humans."
        ),
        "overview_p2": (
            "I lead by example across the full stack: partnering with Product Managers to plan and estimate modules, architecting reactive frontends (React, Next.js) and decoupled microservices (Node.js, NestJS, Go, Bun), building authentication/RBAC & notification systems from the ground up, configuring Kubernetes (K8s) and CI/CD pipelines, and mentoring junior developers to become autonomous contributors. "
            "Driven by a fanatic love for building products, I also architected StrawHatDevs (strawhatdevs.com, with 200+ programming & architecture videos on YouTube), ChimeFlow (multi-tenant notification & cadence platform), ENVault (encrypted .env sync + native Go CLI), and IncidentLens AI."
        ),
        "competencies": [
            "Technical Leadership & Mentorship", "End-to-End System Architecture", "Kubernetes (K8s) & Cloud DevOps",
            "Microservices & Comm Layers", "Auth, RBAC & Notification Systems", "Product Planning & Estimation",
            "Automated CI/CD & Quality Gates", "Rapid MVP & Iterative Delivery", "IoT & Hardware Integrations",
            "Clean Code & Review Governance", "Cross-Functional PM Alignment", "AI-Assisted Engineering Workflows"
        ],
        "experiences": get_experiences("tech_lead")
    },
    {
        "slug": "Forward_Deployed_Engineer",
        "headline": "Forward Deployed Software Engineer | Enterprise Solutions & Full-Stack Architect",
        "accent_color": "#2563eb",
        "accent_dark": "#1d4ed8",
        "highlight_bg": "#eff6ff",
        "overview_p1": (
            "I’m a high-agency Forward Deployed Software Engineer and Full-Stack Architect with 7+ years of experience embedding directly with product teams and enterprise stakeholders to turn ambiguous problems into live production systems at 4x speed. "
            "My journey began on a 2GB RAM government-scheme laptop in Tamil Nadu, where I taught myself programming from scratch—a relentless resourcefulness that took me from converting a college internship into a full-time engineering role in under 3 months (Skematix), to fast-tracking to Principal Engineer (HOD Digitals), leading tier-1 enterprise deliveries at Heady (compressing Chuck E. Cheese’s 8-month web MVP into 2 months solo), and now owning end-to-end delivery as a Staff Engineer at Tuna Software Solution and Lead Software Engineer (Part-Time Contractor) at Science & Humans."
        ),
        "overview_p2": (
            "In customer-facing and greenfield environments, I own whatever it takes to ship: scoping and estimating modules with PMs, building frontends (React, Next.js, React Native) and microservices communication layers (Node.js, NestJS, Go, Bun), engineering ground-up authentication/RBAC and notification systems, integrating physical IoT hardware (payment terminals, printers, cash drawers), and deploying on Kubernetes (K8s) and AWS. "
            "Outside client work, I’m a prolific builder—shipping StrawHatDevs (strawhatdevs.com + 200+ architecture and coding videos on YouTube), ChimeFlow (multi-tenant notification platform), ENVault (zero-knowledge secret sync + native Go CLI), and IncidentLens AI."
        ),
        "competencies": [
            "Customer-Facing Technical Delivery", "Rapid 0-to-1 MVP Execution (4x Speed)", "End-to-End Full-Stack Ownership",
            "Microservices & Auth/RBAC Systems", "Kubernetes (K8s) & Cloud Infra", "IoT Hardware & POS Peripherals",
            "Third-Party & Enterprise APIs", "Automated CI/CD & Zero-Downtime", "PM Collaboration & Scoping",
            "Production Triage & Bug Resolution", "Developer Education (200+ Videos)", "AI Agents & RAG Workflows"
        ],
        "experiences": get_experiences("fde")
    },
    {
        "slug": "Staff_Software_Engineer",
        "headline": "Staff Software Engineer | Full-Stack, Microservices & Cloud/Kubernetes Architect",
        "accent_color": "#0f766e",
        "accent_dark": "#115e59",
        "highlight_bg": "#f0fdfa",
        "overview_p1": (
            "I’m a self-taught, fanatic software engineer with 7+ years of experience owning applications end-to-end—from initial whiteboard architecture and frontend/backend code to Kubernetes infrastructure and automated CI/CD pipelines. "
            "My passion for programming started in 11th/12th standard at a government school in Tamil Nadu on a 2GB RAM state-issued laptop. That obsession carried me from converting a college internship into a full-time role in under 3 months (Skematix), to rising from Senior Engineer to Principal Software Engineer (HOD Digitals), delivering enterprise systems at Heady (shipping a Chuck E. Cheese web MVP solo in 2 months vs. an 8-month estimate), and now serving as Staff Engineer at Tuna Software Solution and Lead Software Engineer (Part-Time Contractor) at Science & Humans."
        ),
        "overview_p2": (
            "I specialize in building complete, production-hardened systems: architecting decoupled microservices and communication layers, engineering authentication/authorization (RBAC) and notification engines from the ground up, deploying containerized workloads on Kubernetes (K8s) and AWS, and mentoring junior engineers in collaboration with Product Managers. "
            "Because I love building products, I also created StrawHatDevs (strawhatdevs.com + 200+ programming and architecture videos on YouTube), ChimeFlow (a multi-tenant omnichannel notification & cadence platform), ENVault (an encrypted .env sync platform with a native Go CLI), and IncidentLens AI."
        ),
        "competencies": [
            "End-to-End System Architecture", "Kubernetes (K8s) & Cloud DevOps", "Microservices & Comm Layers",
            "Auth, RBAC & Notification Engines", "Full-Stack TypeScript / Go / Bun", "Automated CI/CD Pipelines",
            "Rapid MVP & Production Delivery", "IoT & Hardware Integrations", "Technical Mentorship & Reviews",
            "Product Planning & Estimation", "Clean / Hexagonal Architecture", "AI-Assisted Engineering & RAG"
        ],
        "experiences": get_experiences("staff")
    },
    {
        "slug": "Senior_Software_Engineer",
        "headline": "Senior Software Engineer | Full-Stack, Microservices, Cloud & Kubernetes",
        "accent_color": "#1d4ed8",
        "accent_dark": "#1e40af",
        "highlight_bg": "#eff6ff",
        "overview_p1": (
            "I’m a hands-on Senior Full-Stack Software Engineer with 7+ years of experience building, scaling, and deploying production web, mobile, and cloud applications end-to-end. "
            "My programming journey started in a government school in Tamil Nadu on a 2GB RAM state-issued laptop, teaching myself to code through documentation and YouTube. Since converting my college internship into a full-time role in under 3 months in 2019 (Skematix), I’ve progressed through Senior and Principal engineering roles (HOD Digitals), led enterprise product deliveries at Heady (shipping a Chuck E. Cheese web MVP solo in 2 months vs. an 8-month estimate), and currently build mission-critical systems at Tuna Software Solution and as a Lead Software Engineer (Part-Time Contractor) at Science & Humans."
        ),
        "overview_p2": (
            "I bring true full-lifecycle ownership to every team I join: building performant frontends (React, Next.js, TypeScript, React Native), architecting backend microservices and communication layers (Node.js, NestJS, Go, Bun, PostgreSQL), implementing complete authentication/RBAC and notification systems from scratch, configuring Kubernetes (K8s) and CI/CD pipelines, and mentoring junior developers alongside Product Managers. "
            "Outside of work, I’m an active builder and educator—creator of StrawHatDevs (strawhatdevs.com + 200+ programming and architecture videos on YouTube), ChimeFlow (multi-tenant notification platform), ENVault (encrypted .env sync + Go CLI), and IncidentLens AI."
        ),
        "competencies": [
            "Full-Stack Web & Mobile (React/Next)", "Backend Microservices (Node/Nest/Go)", "Authentication, RBAC & Notifications",
            "Kubernetes (K8s), Docker & AWS", "Automated CI/CD Pipelines", "PostgreSQL, Redis & Data Modeling",
            "Rapid MVP & Feature Execution", "IoT & Third-Party API Integrations", "Module Estimation & PM Alignment",
            "Mentoring Junior Engineers", "End-to-End Testing (Playwright/Vitest)", "AI-Assisted Development & RAG"
        ],
        "experiences": get_experiences("senior")
    }
]

# ==============================================================================
# MARKDOWN GENERATOR
# ==============================================================================

def generate_resume_md(cfg):
    lines = [
        "# SANTHEEPKUMAR VEERASAMY",
        f"**{cfg['headline']}**  ",
        "+91 86676 62978 • [devsantheep@gmail.com](mailto:devsantheep@gmail.com) • [LinkedIn](https://linkedin.com/in/santheepkumar) • [GitHub](https://github.com/Santheepkumar) • [strawhatdevs.com](https://strawhatdevs.com) • [Portfolio](https://sk.strawhatdevs.com) • [YouTube: Straw Hat Devs](https://youtube.com/@StrawHatDevs)",
        "",
        "---",
        "",
        "## OVERVIEW",
        cfg["overview_p1"],
        "",
        cfg["overview_p2"],
        "",
        "---",
        "",
        "## CORE TECHNICAL COMPETENCIES",
        ""
    ]
    for cat, val in COMMON_SKILLS.items():
        lines.append(f"- **{cat}**: {val}")

    lines.extend([
        "",
        "---",
        "",
        "## PROFESSIONAL EXPERIENCE",
        ""
    ])

    for exp in cfg["experiences"]:
        lines.append(f"### **{exp['role']}** | {exp['company']}")
        lines.append(f"*{exp['date']}*")
        for lead, text in exp["bullets"]:
            lines.append(f"- **{lead.strip()}** {text}")
        lines.append("")

    lines.extend([
        "---",
        "",
        "## SELF-BUILT PRODUCTS, PLATFORMS & COMMUNITY LEADERSHIP",
        ""
    ])

    for prod in INDEPENDENT_PRODUCTS:
        lines.append(f"### **{prod['title']}** | *{prod['role_or_domain']} ({prod['date']})*")
        lines.append(f"- {prod['desc']}")
        lines.append(f"- **Tech Stack**: {prod['stack']}")
        lines.append("")

    lines.extend([
        "---",
        "",
        "## KEY ENTERPRISE PROJECTS & PRODUCTION DEPLOYMENTS",
        ""
    ])

    for proj in ENTERPRISE_PROJECTS:
        lines.append(f"### **{proj['title']}** | *{proj['role_or_domain']} ({proj['date']})*")
        lines.append(f"- {proj['desc']}")
        lines.append(f"- **Tech Stack**: {proj['stack']}")
        lines.append("")

    lines.extend([
        "---",
        "",
        "## EDUCATION",
        "",
        "**Bachelor of Engineering (B.E.) — Computer Science & Engineering**  ",
        "Anna University, Tamil Nadu, India • *2015 – 2019*",
        ""
    ])
    return "\n".join(lines)

# ==============================================================================
# HTML GENERATOR
# ==============================================================================

def generate_resume_html(cfg):
    skills_html = "\n".join([
        f'        <div class="skill-item"><span class="skill-title">{cat}:</span><span class="skill-desc">{val}</span></div>'
        for cat, val in COMMON_SKILLS.items()
    ])

    exp_html_parts = []
    for exp in cfg["experiences"]:
        bullets_html = "\n".join([
            f'          <li><strong>{lead}</strong>{text}</li>'
            for lead, text in exp["bullets"]
        ])
        exp_html_parts.append(f"""      <div class="exp-item">
        <div class="exp-header">
          <div class="role-company">{exp['role']} <span>| {exp['company']}</span></div>
          <div class="exp-date">{exp['date']}</div>
        </div>
        <ul>
{bullets_html}
        </ul>
      </div>""")
    exp_html = "\n".join(exp_html_parts)

    ind_cards = "\n".join([
        f"""        <div class="project-card">
          <div class="project-title">{p['title']}</div>
          <div class="project-desc">{p['desc']}</div>
          <div class="project-tech">{p['stack']}</div>
        </div>"""
        for p in INDEPENDENT_PRODUCTS
    ])

    ent_cards = "\n".join([
        f"""        <div class="project-card">
          <div class="project-title">{p['title']}</div>
          <div class="project-desc">{p['desc']}</div>
          <div class="project-tech">{p['stack']}</div>
        </div>"""
        for p in ENTERPRISE_PROJECTS
    ])

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Santheepkumar Veerasamy - {cfg['headline']}</title>
  <style>
    :root {{
      --primary: #0f172a;
      --accent: {cfg['accent_color']};
      --accent-dark: {cfg['accent_dark']};
      --text: #334155;
      --text-dark: #0f172a;
      --text-light: #64748b;
      --border: #e2e8f0;
      --bg: #f8fafc;
      --card-bg: #ffffff;
    }}
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
      line-height: 1.48;
      color: var(--text);
      background-color: var(--bg);
      padding: 28px 16px;
    }}
    .container {{
      max-width: 900px;
      margin: 0 auto;
      background: var(--card-bg);
      padding: 40px 48px;
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.06);
      border-radius: 8px;
    }}
    .action-bar {{
      max-width: 900px;
      margin: 0 auto 14px auto;
      display: flex;
      justify-content: flex-end;
      gap: 12px;
    }}
    .btn {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      background: var(--accent);
      color: white;
      padding: 8px 16px;
      font-size: 14px;
      font-weight: 600;
      border-radius: 6px;
      text-decoration: none;
      border: none;
      cursor: pointer;
    }}
    .btn:hover {{ background: var(--accent-dark); }}
    header {{
      border-bottom: 2px solid var(--primary);
      padding-bottom: 14px;
      margin-bottom: 18px;
    }}
    h1 {{
      font-size: 27px;
      font-weight: 800;
      letter-spacing: -0.5px;
      color: var(--primary);
      text-transform: uppercase;
      margin-bottom: 3px;
    }}
    .subtitle {{
      font-size: 14.5px;
      font-weight: 700;
      color: var(--accent);
      text-transform: uppercase;
      letter-spacing: 0.4px;
      margin-bottom: 9px;
    }}
    .contact-bar {{
      display: flex;
      flex-wrap: wrap;
      gap: 6px 12px;
      font-size: 12.5px;
      color: var(--text-light);
    }}
    .contact-bar a {{ color: var(--text); text-decoration: none; font-weight: 500; }}
    .contact-bar a:hover {{ color: var(--accent); text-decoration: underline; }}
    .separator {{ color: #cbd5e1; }}
    section {{ margin-bottom: 18px; }}
    h2 {{
      font-size: 13.5px;
      font-weight: 800;
      color: var(--primary);
      text-transform: uppercase;
      letter-spacing: 0.9px;
      border-bottom: 1.5px solid var(--border);
      padding-bottom: 3px;
      margin-bottom: 10px;
    }}
    p {{
      font-size: 13px;
      color: var(--text);
      margin-bottom: 8px;
      text-align: justify;
    }}
    .skills-grid {{
      display: flex;
      flex-direction: column;
      gap: 6px;
      font-size: 12.5px;
    }}
    .skill-item {{
      display: grid;
      grid-template-columns: 205px 1fr;
      line-height: 1.42;
    }}
    .skill-title {{ font-weight: 700; color: var(--text-dark); }}
    .skill-desc {{ color: var(--text); }}
    .exp-item {{ margin-bottom: 14px; }}
    .exp-header {{
      display: flex;
      justify-content: space-between;
      align-items: baseline;
      margin-bottom: 3px;
    }}
    .role-company {{ font-size: 14px; font-weight: 700; color: var(--text-dark); }}
    .role-company span {{ font-weight: 600; color: var(--accent); }}
    .exp-date {{ font-size: 12px; font-weight: 600; color: var(--text-light); }}
    ul {{ list-style-type: none; padding-left: 0; }}
    li {{
      position: relative;
      padding-left: 16px;
      font-size: 12.5px;
      line-height: 1.45;
      margin-bottom: 4px;
      color: var(--text);
    }}
    li::before {{
      content: "▸";
      position: absolute;
      left: 0;
      color: var(--accent);
      font-weight: bold;
    }}
    li strong {{ color: var(--text-dark); }}
    .projects-grid {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
    }}
    .project-card {{
      background: #f8fafc;
      border: 1px solid var(--border);
      border-radius: 6px;
      padding: 9px 11px;
    }}
    .project-title {{
      font-size: 13px;
      font-weight: 700;
      color: var(--text-dark);
      margin-bottom: 3px;
    }}
    .project-desc {{
      font-size: 11.8px;
      color: var(--text);
      line-height: 1.38;
      margin-bottom: 5px;
      text-align: left;
    }}
    .project-tech {{
      font-size: 11px;
      color: var(--accent-dark);
      font-weight: 600;
    }}
    .edu-item {{
      display: flex;
      justify-content: space-between;
      align-items: baseline;
      font-size: 12.8px;
    }}
    .edu-degree {{ font-weight: 700; color: var(--text-dark); }}
    .edu-year {{ color: var(--text-light); font-weight: 600; }}
    @media print {{
      body {{ background: #fff; padding: 0; }}
      .action-bar {{ display: none !important; }}
      .container {{ padding: 0; box-shadow: none; max-width: 100%; }}
      a {{ text-decoration: none; color: inherit !important; }}
      .exp-item, .project-card {{ break-inside: avoid; }}
      section {{ margin-bottom: 14px; }}
    }}
  </style>
</head>
<body>
  <div class="action-bar">
    <button class="btn" onclick="window.print()">🖨️ Print / Save as PDF</button>
  </div>
  <div class="container">
    <header>
      <h1>Santheepkumar Veerasamy</h1>
      <div class="subtitle">{cfg['headline']}</div>
      <div class="contact-bar">
        <span>+91 86676 62978</span>
        <span class="separator">•</span>
        <span><a href="mailto:devsantheep@gmail.com">devsantheep@gmail.com</a></span>
        <span class="separator">•</span>
        <span><a href="https://linkedin.com/in/santheepkumar" target="_blank">linkedin.com/in/santheepkumar</a></span>
        <span class="separator">•</span>
        <span><a href="https://github.com/Santheepkumar" target="_blank">github.com/Santheepkumar</a></span>
        <span class="separator">•</span>
        <span><a href="https://strawhatdevs.com" target="_blank"><strong>strawhatdevs.com</strong></a></span>
        <span class="separator">•</span>
        <span><a href="https://youtube.com/@StrawHatDevs" target="_blank"><strong>YouTube (@StrawHatDevs - 200+ Videos)</strong></a></span>
        <span class="separator">•</span>
        <span><a href="https://sk.strawhatdevs.com" target="_blank">sk.strawhatdevs.com</a></span>
      </div>
    </header>

    <section>
      <h2>Overview</h2>
      <p>{cfg['overview_p1']}</p>
      <p style="margin-bottom: 0;">{cfg['overview_p2']}</p>
    </section>

    <section>
      <h2>Core Technical Competencies</h2>
      <div class="skills-grid">
{skills_html}
      </div>
    </section>

    <section>
      <h2>Professional Experience</h2>
{exp_html}
    </section>

    <section>
      <h2>Self-Built Products, Platforms & Developer Community</h2>
      <div class="projects-grid">
{ind_cards}
      </div>
    </section>

    <section>
      <h2>Key Enterprise Projects & Production Deployments</h2>
      <div class="projects-grid">
{ent_cards}
      </div>
    </section>

    <section>
      <h2>Education</h2>
      <div class="edu-item">
        <div>
          <span class="edu-degree">Bachelor of Engineering (B.E.) — Computer Science & Engineering</span> • 
          <span>Anna University, Tamil Nadu, India</span>
        </div>
        <div class="edu-year">2015 – 2019</div>
      </div>
    </section>
  </div>
</body>
</html>
"""

# ==============================================================================
# DOCX OPENXML BUILDER
# ==============================================================================

def create_p_bdr():
    pBdr = ET.Element(qn("pBdr"))
    for side in ["top", "left", "bottom", "right", "between"]:
        bdr = ET.SubElement(pBdr, qn(side))
        bdr.attrib[qn("color")] = "e3e3e3"
        bdr.attrib[qn("space")] = "0"
        bdr.attrib[qn("sz")] = "0"
        bdr.attrib[qn("val")] = "none"
    return pBdr

def make_run(text, bold=False, italic=False, color="141414", size="15", underline=False):
    r = ET.Element(qn("r"))
    r.attrib[qn("rsidDel")] = "00000000"
    r.attrib[qn("rsidR")] = "00000000"
    r.attrib[qn("rsidRPr")] = "00000000"

    rPr = ET.SubElement(r, qn("rPr"))
    if bold:
        b = ET.SubElement(rPr, qn("b"))
        b.attrib[qn("val")] = "1"
        bCs = ET.SubElement(rPr, qn("bCs"))
        bCs.attrib[qn("val")] = "1"
    if italic:
        i = ET.SubElement(rPr, qn("i"))
        i.attrib[qn("val")] = "1"
        iCs = ET.SubElement(rPr, qn("iCs"))
        iCs.attrib[qn("val")] = "1"
    if underline:
        u = ET.SubElement(rPr, qn("u"))
        u.attrib[qn("val")] = "single"
    
    col = ET.SubElement(rPr, qn("color"))
    col.attrib[qn("val")] = color

    sz = ET.SubElement(rPr, qn("sz"))
    sz.attrib[qn("val")] = str(size)
    szCs = ET.SubElement(rPr, qn("szCs"))
    szCs.attrib[qn("val")] = str(size)

    rtl = ET.SubElement(rPr, qn("rtl"))
    rtl.attrib[qn("val")] = "0"

    if text:
        t = ET.SubElement(r, qn("t"))
        t.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"
        t.text = text

    return r

def make_section_header(title, is_first=False):
    p = ET.Element(qn("p"))
    p.attrib[qn_w14("paraId")] = next_para_id()
    pPr = ET.SubElement(p, qn("pPr"))
    pPr.append(create_p_bdr())
    shd = ET.SubElement(pPr, qn("shd"))
    shd.attrib[qn("fill")] = "ffffff"
    shd.attrib[qn("val")] = "clear"

    spc = ET.SubElement(pPr, qn("spacing"))
    spc.attrib[qn("before")] = "0" if is_first else "220"
    spc.attrib[qn("after")] = "70"
    spc.attrib[qn("lineRule")] = "auto"

    p.append(make_run(title, bold=True, color="193c66", size="17"))
    return p

def make_job_header(role, company, date):
    p = ET.Element(qn("p"))
    p.attrib[qn_w14("paraId")] = next_para_id()
    pPr = ET.SubElement(p, qn("pPr"))
    pPr.append(create_p_bdr())
    shd = ET.SubElement(pPr, qn("shd"))
    shd.attrib[qn("fill")] = "ffffff"
    shd.attrib[qn("val")] = "clear"

    spc = ET.SubElement(pPr, qn("spacing"))
    spc.attrib[qn("before")] = "140"
    spc.attrib[qn("after")] = "35"
    spc.attrib[qn("lineRule")] = "auto"

    p.append(make_run(role, bold=True, color="141414", size="17"))
    p.append(make_run("  |  ", bold=True, color="434343", size="15"))
    p.append(make_run(company, bold=True, color="2560a7", size="16"))
    p.append(make_run(f"   {date}", bold=True, italic=True, color="434343", size="14"))
    return p

def make_bullet(text_or_runs, before=50, after=0):
    p = ET.Element(qn("p"))
    p.attrib[qn_w14("paraId")] = next_para_id()
    pPr = ET.SubElement(p, qn("pPr"))
    numPr = ET.SubElement(pPr, qn("numPr"))
    ilvl = ET.SubElement(numPr, qn("ilvl"))
    ilvl.attrib[qn("val")] = "0"
    numId = ET.SubElement(numPr, qn("numId"))
    numId.attrib[qn("val")] = "2"

    pPr.append(create_p_bdr())
    shd = ET.SubElement(pPr, qn("shd"))
    shd.attrib[qn("fill")] = "ffffff"
    shd.attrib[qn("val")] = "clear"

    spc = ET.SubElement(pPr, qn("spacing"))
    spc.attrib[qn("before")] = str(before)
    spc.attrib[qn("after")] = str(after)
    spc.attrib[qn("lineRule")] = "auto"

    ind = ET.SubElement(pPr, qn("ind"))
    ind.attrib[qn("left")] = "720"
    ind.attrib[qn("hanging")] = "360"

    if isinstance(text_or_runs, str):
        p.append(make_run(text_or_runs, bold=False, color="141414", size="15"))
    else:
        for r in text_or_runs:
            p.append(r)
    return p

def make_para(text_or_runs, bold=False, italic=False, color="141414", size="15", before=0, after=60):
    p = ET.Element(qn("p"))
    p.attrib[qn_w14("paraId")] = next_para_id()
    pPr = ET.SubElement(p, qn("pPr"))
    pPr.append(create_p_bdr())
    shd = ET.SubElement(pPr, qn("shd"))
    shd.attrib[qn("fill")] = "ffffff"
    shd.attrib[qn("val")] = "clear"

    spc = ET.SubElement(pPr, qn("spacing"))
    spc.attrib[qn("before")] = str(before)
    spc.attrib[qn("after")] = str(after)
    spc.attrib[qn("lineRule")] = "auto"

    if isinstance(text_or_runs, str):
        p.append(make_run(text_or_runs, bold=bold, italic=italic, color=color, size=size))
    else:
        for r in text_or_runs:
            p.append(r)
    return p

def make_competencies_table(items):
    tbl = ET.Element(qn("tbl"))
    tblPr = ET.SubElement(tbl, qn("tblPr"))
    tblStyle = ET.SubElement(tblPr, qn("tblStyle"))
    tblStyle.attrib[qn("val")] = "Table1"
    tblW = ET.SubElement(tblPr, qn("tblW"))
    tblW.attrib[qn("w")] = "7020.0"
    tblW.attrib[qn("type")] = "dxa"
    jc = ET.SubElement(tblPr, qn("jc"))
    jc.attrib[qn("val")] = "left"

    tblBorders = ET.SubElement(tblPr, qn("tblBorders"))
    for b in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        el = ET.SubElement(tblBorders, qn(b))
        el.attrib[qn("color")] = "000000"
        el.attrib[qn("space")] = "0"
        el.attrib[qn("sz")] = "0"
        el.attrib[qn("val")] = "nil"

    tblLayout = ET.SubElement(tblPr, qn("tblLayout"))
    tblLayout.attrib[qn("type")] = "fixed"
    tblLook = ET.SubElement(tblPr, qn("tblLook"))
    tblLook.attrib[qn("val")] = "0600"

    tblGrid = ET.SubElement(tbl, qn("tblGrid"))
    for _ in range(3):
        gc = ET.SubElement(tblGrid, qn("gridCol"))
        gc.attrib[qn("w")] = "2340"

    for i in range(0, len(items), 3):
        tr = ET.SubElement(tbl, qn("tr"))
        trPr = ET.SubElement(tr, qn("trPr"))
        cs = ET.SubElement(trPr, qn("cantSplit"))
        cs.attrib[qn("val")] = "0"
        trH = ET.SubElement(trPr, qn("trHeight"))
        trH.attrib[qn("val")] = "320"
        trH.attrib[qn("hRule")] = "atLeast"

        row_items = items[i:i+3]
        for item in row_items:
            tc = ET.SubElement(tr, qn("tc"))
            tcPr = ET.SubElement(tc, qn("tcPr"))
            tcBorders = ET.SubElement(tcPr, qn("tcBorders"))
            for b in ["top", "left", "bottom", "right"]:
                el = ET.SubElement(tcBorders, qn(b))
                el.attrib[qn("color")] = "000000"
                el.attrib[qn("space")] = "0"
                el.attrib[qn("sz")] = "6"
                el.attrib[qn("val")] = "single"
            tcMar = ET.SubElement(tcPr, qn("tcMar"))
            for m in ["top", "left", "bottom", "right"]:
                el = ET.SubElement(tcMar, qn(m))
                el.attrib[qn("w")] = "60.0"
                el.attrib[qn("type")] = "dxa"
            va = ET.SubElement(tcPr, qn("vAlign"))
            va.attrib[qn("val")] = "top"

            p = ET.SubElement(tc, qn("p"))
            p.attrib[qn_w14("paraId")] = next_para_id()
            pPr = ET.SubElement(p, qn("pPr"))
            pPr.append(create_p_bdr())
            shd = ET.SubElement(pPr, qn("shd"))
            shd.attrib[qn("fill")] = "ffffff"
            shd.attrib[qn("val")] = "clear"
            p.append(make_run(f"▸ {item}", bold=True, color="141414", size="14"))
    return tbl

def make_skills_table(skills_dict):
    tbl = ET.Element(qn("tbl"))
    tblPr = ET.SubElement(tbl, qn("tblPr"))
    tblStyle = ET.SubElement(tblPr, qn("tblStyle"))
    tblStyle.attrib[qn("val")] = "Table2"
    tblW = ET.SubElement(tblPr, qn("tblW"))
    tblW.attrib[qn("w")] = "7020.0"
    tblW.attrib[qn("type")] = "dxa"
    jc = ET.SubElement(tblPr, qn("jc"))
    jc.attrib[qn("val")] = "left"

    tblBorders = ET.SubElement(tblPr, qn("tblBorders"))
    for b in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        el = ET.SubElement(tblBorders, qn(b))
        el.attrib[qn("color")] = "000000"
        el.attrib[qn("space")] = "0"
        el.attrib[qn("sz")] = "0"
        el.attrib[qn("val")] = "nil"

    tblLayout = ET.SubElement(tblPr, qn("tblLayout"))
    tblLayout.attrib[qn("type")] = "fixed"
    tblLook = ET.SubElement(tblPr, qn("tblLook"))
    tblLook.attrib[qn("val")] = "0600"

    tblGrid = ET.SubElement(tbl, qn("tblGrid"))
    gc1 = ET.SubElement(tblGrid, qn("gridCol"))
    gc1.attrib[qn("w")] = "1650"
    gc2 = ET.SubElement(tblGrid, qn("gridCol"))
    gc2.attrib[qn("w")] = "5370"

    for cat, val in skills_dict.items():
        tr = ET.SubElement(tbl, qn("tr"))
        trPr = ET.SubElement(tr, qn("trPr"))
        cs = ET.SubElement(trPr, qn("cantSplit"))
        cs.attrib[qn("val")] = "0"
        trH = ET.SubElement(trPr, qn("trHeight"))
        trH.attrib[qn("val")] = "280"
        trH.attrib[qn("hRule")] = "atLeast"

        tc1 = ET.SubElement(tr, qn("tc"))
        tcPr1 = ET.SubElement(tc1, qn("tcPr"))
        tcBorders1 = ET.SubElement(tcPr1, qn("tcBorders"))
        for b in ["top", "left", "bottom", "right"]:
            el = ET.SubElement(tcBorders1, qn(b))
            el.attrib[qn("color")] = "000000"
            el.attrib[qn("space")] = "0"
            el.attrib[qn("sz")] = "6"
            el.attrib[qn("val")] = "single"
        tcMar1 = ET.SubElement(tcPr1, qn("tcMar"))
        for m in ["top", "left", "bottom", "right"]:
            el = ET.SubElement(tcMar1, qn(m))
            el.attrib[qn("w")] = "60.0"
            el.attrib[qn("type")] = "dxa"
        va1 = ET.SubElement(tcPr1, qn("vAlign"))
        va1.attrib[qn("val")] = "top"
        p1 = ET.SubElement(tc1, qn("p"))
        p1.attrib[qn_w14("paraId")] = next_para_id()
        p1.append(make_run(cat, bold=True, color="193c66", size="14"))

        tc2 = ET.SubElement(tr, qn("tc"))
        tcPr2 = ET.SubElement(tc2, qn("tcPr"))
        tcBorders2 = ET.SubElement(tcPr2, qn("tcBorders"))
        for b in ["top", "left", "bottom", "right"]:
            el = ET.SubElement(tcBorders2, qn(b))
            el.attrib[qn("color")] = "000000"
            el.attrib[qn("space")] = "0"
            el.attrib[qn("sz")] = "6"
            el.attrib[qn("val")] = "single"
        tcMar2 = ET.SubElement(tcPr2, qn("tcMar"))
        for m in ["top", "left", "bottom", "right"]:
            el = ET.SubElement(tcMar2, qn(m))
            el.attrib[qn("w")] = "60.0"
            el.attrib[qn("type")] = "dxa"
        va2 = ET.SubElement(tcPr2, qn("vAlign"))
        va2.attrib[qn("val")] = "top"
        p2 = ET.SubElement(tc2, qn("p"))
        p2.attrib[qn_w14("paraId")] = next_para_id()
        p2.append(make_run(val, bold=False, color="141414", size="14"))

    return tbl

def build_resume_docx_body(cfg):
    body = ET.Element(qn("body"))

    # Name
    body.append(make_para([make_run("SANTHEEPKUMAR VEERASAMY", bold=True, color="193c66", size="33")], before=0, after=20))
    # Headline
    body.append(make_para([make_run(cfg["headline"], bold=True, color="2560a7", size="17")], before=0, after=30))
    # Contacts
    body.append(make_para([
        make_run("+91 86676 62978  |  devsantheep@gmail.com  |  ", bold=True, color="434343", size="14"),
        make_run("linkedin.com/in/santheepkumar", bold=True, color="2560a7", size="14", underline=True),
        make_run("  |  ", bold=True, color="434343", size="14"),
        make_run("github.com/Santheepkumar", bold=True, color="1155cc", size="14", underline=True),
        make_run("  |  ", bold=True, color="434343", size="14"),
        make_run("strawhatdevs.com", bold=True, color="2560a7", size="14", underline=True),
        make_run("  |  ", bold=True, color="434343", size="14"),
        make_run("youtube.com/@StrawHatDevs", bold=True, color="1155cc", size="14", underline=True),
        make_run("  |  ", bold=True, color="434343", size="14"),
        make_run("sk.strawhatdevs.com", bold=True, color="2560a7", size="14", underline=True)
    ], before=0, after=100))

    # OVERVIEW
    body.append(make_section_header("OVERVIEW", is_first=True))
    body.append(make_para(cfg["overview_p1"], color="141414", size="15", before=40, after=60))
    body.append(make_para(cfg["overview_p2"], color="141414", size="15", before=0, after=100))

    # CORE COMPETENCIES
    body.append(make_section_header("CORE COMPETENCIES"))
    body.append(make_competencies_table(cfg["competencies"]))

    # TECHNICAL SKILLS
    body.append(make_section_header("TECHNICAL SKILLS"))
    body.append(make_skills_table(COMMON_SKILLS))

    # PROFESSIONAL EXPERIENCE
    body.append(make_section_header("PROFESSIONAL EXPERIENCE"))
    for exp in cfg["experiences"]:
        body.append(make_job_header(exp["role"], exp["company"], exp["date"]))
        for lead, text in exp["bullets"]:
            body.append(make_bullet([
                make_run(lead, bold=True, color="141414", size="15"),
                make_run(text, color="141414", size="15")
            ]))

    # SELF-BUILT PRODUCTS & PLATFORMS
    body.append(make_section_header("SELF-BUILT PRODUCTS, PLATFORMS & COMMUNITY LEADERSHIP"))
    for prod in INDEPENDENT_PRODUCTS:
        body.append(make_job_header(prod["title"], prod["role_or_domain"], prod["date"]))
        body.append(make_bullet(prod["desc"]))
        body.append(make_bullet([
            make_run("Tech Stack: ", bold=True, color="141414", size="14"),
            make_run(prod["stack"], color="434343", size="14")
        ]))

    # KEY ENTERPRISE PROJECTS
    body.append(make_section_header("KEY ENTERPRISE PROJECTS & PRODUCTION DEPLOYMENTS"))
    for proj in ENTERPRISE_PROJECTS:
        body.append(make_job_header(proj["title"], proj["role_or_domain"], proj["date"]))
        body.append(make_bullet(proj["desc"]))
        body.append(make_bullet([
            make_run("Tech Stack: ", bold=True, color="141414", size="14"),
            make_run(proj["stack"], color="434343", size="14")
        ]))

    # EDUCATION
    body.append(make_section_header("EDUCATION"))
    body.append(make_job_header("Bachelor of Engineering (B.E.) — Computer Science & Engineering", "Anna University", "2015 – 2019"))
    body.append(make_para("Tamil Nadu, India", color="434343", size="14", before=20, after=60))

    sectPr = ET.SubElement(body, qn("sectPr"))
    pgSz = ET.SubElement(sectPr, qn("pgSz"))
    pgSz.attrib[qn("h")] = "15840"
    pgSz.attrib[qn("w")] = "12240"
    pgSz.attrib[qn("orient")] = "portrait"

    pgMar = ET.SubElement(sectPr, qn("pgMar"))
    pgMar.attrib[qn("bottom")] = "1080"
    pgMar.attrib[qn("top")] = "1080"
    pgMar.attrib[qn("left")] = "1080"
    pgMar.attrib[qn("right")] = "1080"
    pgMar.attrib[qn("header")] = "720"
    pgMar.attrib[qn("footer")] = "720"

    pgNumType = ET.SubElement(sectPr, qn("pgNumType"))
    pgNumType.attrib[qn("start")] = "1"
    return body

def write_docx_from_body(dest_filename, new_body):
    with zipfile.ZipFile(TEMPLATE_DOCX, "r") as z_in:
        orig_doc_xml = z_in.read("word/document.xml")
        orig_root = ET.fromstring(orig_doc_xml)
        orig_body = orig_root.find(qn("body"))
        if orig_body is not None:
            orig_root.remove(orig_body)
        orig_root.append(new_body)
        new_doc_xml = ET.tostring(orig_root, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(dest_filename, "w", compression=zipfile.ZIP_DEFLATED) as z_out:
            for item in z_in.infolist():
                if item.filename == "word/document.xml":
                    z_out.writestr(item.filename, new_doc_xml)
                else:
                    z_out.writestr(item.filename, z_in.read(item.filename))
    print(f"Generated DOCX: {dest_filename}")

# ==============================================================================
# COVER LETTER GENERATOR (MD, HTML, DOCX)
# ==============================================================================

COVER_LETTER_PARAS = [
    "I’m writing to express my interest in joining [Company Name] as a [Role Title]. My love for programming started in 11th standard at a government school in Tamil Nadu, where a teacher first showed me what code could do, and I spent the next years teaching myself software engineering on a 2GB RAM state-issued laptop. That obsession with building real software carried me from converting a college internship into a full-time engineering role in under 3 months (Skematix) to rising from Senior Engineer to Principal Software Engineer (HOD Digitals), leading enterprise deliveries at Heady, and now serving as a Staff Engineer at Tuna Software Solution and Lead Software Engineer (Part-Time Contractor) at Science & Humans.",
    "Across 7+ years of engineering, I’ve gravitated toward owning hard problems end-to-end—from product discovery and module estimation with Product Managers to frontend UI, backend microservices, and production DevOps. At Science & Humans, I contributed from the ground up: architecting the entire authentication, RBAC authorization, and notification systems, setting up the microservices communication layer, mentoring junior engineers, and configuring deployments on Kubernetes (K8s). At Heady, I delivered the Chuck E. Cheese customer experience web MVP solo in 2 months (compressing an 8-month estimate by 75%), and at Tuna Software Solution, I engineered an Android IoT retail kiosk (Pony Up) with payment terminals and receipt printers before leading its cloud web migration.",
    "Beyond my day-to-day roles, I’m a relentless product builder and educator. I founded StrawHatDevs (strawhatdevs.com), where I’ve published 200+ programming and architecture videos on YouTube, and independently architected production platforms including ChimeFlow (a multi-tenant omnichannel notification & cadence engine), ENVault (a zero-knowledge encrypted .env sync platform with a native Go CLI), and IncidentLens AI (an evidence-grounded incident triage & Runbook RAG investigator). Building and teaching at this depth has shaped how I write clean code, reason about trade-offs, and help engineering teams ship faster without fragile complexity.",
    "I admire the engineering culture and product vision at [Company Name], and I would love the opportunity to discuss how my end-to-end full-stack, microservices, and Kubernetes/cloud experience can help your team build and ship resilient software."
]

def generate_cover_letter_all():
    base = f"{DATE_PREFIX}_Common_Cover_Letter_Santheepkumar_V"
    # MD
    md_content = f"""# Santheepkumar Veerasamy
**Staff Software Engineer | Principal / Technical Lead | Full-Stack & Cloud Architect**  
+91 86676 62978 • [devsantheep@gmail.com](mailto:devsantheep@gmail.com) • [LinkedIn](https://linkedin.com/in/santheepkumar) • [GitHub](https://github.com/Santheepkumar) • [strawhatdevs.com](https://strawhatdevs.com) • [sk.strawhatdevs.com](https://sk.strawhatdevs.com)

---

**Date:** November 9, 2026  
**To:** Hiring Manager / Engineering Leadership at [Company Name]  
**Re:** Application for [Role Title]  

---

Hi [Hiring Manager / Team Name],

{COVER_LETTER_PARAS[0]}

{COVER_LETTER_PARAS[1]}

{COVER_LETTER_PARAS[2]}

{COVER_LETTER_PARAS[3]}

Thank you for your time and consideration.

Warm regards,

**Santheepkumar Veerasamy**  
Staff / Lead Software Engineer & Full-Stack Cloud Architect  
+91 86676 62978 • devsantheep@gmail.com  
[strawhatdevs.com](https://strawhatdevs.com) • [linkedin.com/in/santheepkumar](https://linkedin.com/in/santheepkumar) • [github.com/Santheepkumar](https://github.com/Santheepkumar)
"""
    with open(f"{base}.md", "w") as f:
        f.write(md_content)

    # HTML
    paras_html = "\n".join([f"    <p>{p}</p>" for p in COVER_LETTER_PARAS])
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Cover Letter - Santheepkumar Veerasamy</title>
  <style>
    :root {{
      --primary: #0f172a;
      --accent: #0f766e;
      --accent-dark: #115e59;
      --text: #334155;
      --text-dark: #0f172a;
      --text-light: #64748b;
      --border: #e2e8f0;
      --bg: #f8fafc;
    }}
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
      line-height: 1.65;
      color: var(--text);
      background-color: var(--bg);
      padding: 36px 16px;
    }}
    .container {{
      max-width: 820px;
      margin: 0 auto;
      background: #ffffff;
      padding: 48px 56px;
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.06);
      border-radius: 8px;
    }}
    .action-bar {{
      max-width: 820px;
      margin: 0 auto 16px auto;
      display: flex;
      justify-content: flex-end;
    }}
    .btn {{
      background: var(--accent);
      color: white;
      padding: 8px 16px;
      font-size: 14px;
      font-weight: 600;
      border-radius: 6px;
      border: none;
      cursor: pointer;
    }}
    header {{
      border-bottom: 2px solid var(--primary);
      padding-bottom: 16px;
      margin-bottom: 24px;
    }}
    h1 {{ font-size: 26px; font-weight: 800; color: var(--primary); text-transform: uppercase; margin-bottom: 4px; }}
    .subtitle {{ font-size: 14px; font-weight: 700; color: var(--accent); text-transform: uppercase; margin-bottom: 10px; }}
    .contact-bar {{ display: flex; flex-wrap: wrap; gap: 8px 14px; font-size: 13px; color: var(--text-light); }}
    .contact-bar a {{ color: var(--text); text-decoration: none; font-weight: 500; }}
    .meta-box {{ margin-bottom: 22px; font-size: 14px; color: var(--text-light); line-height: 1.6; }}
    .meta-box strong {{ color: var(--text-dark); }}
    .salutation {{ font-size: 15px; font-weight: 700; color: var(--text-dark); margin-bottom: 16px; }}
    p {{ font-size: 14px; color: var(--text); margin-bottom: 15px; text-align: justify; }}
    .closing {{ margin-top: 24px; font-size: 14px; }}
    .signature-name {{ font-size: 16px; font-weight: 800; color: var(--primary); margin-top: 10px; }}
    .signature-title {{ font-size: 13px; color: var(--accent); font-weight: 600; margin-bottom: 4px; }}
    .signature-contacts {{ font-size: 12.5px; color: var(--text-light); }}
    @media print {{
      body {{ background: #fff; padding: 0; }}
      .action-bar {{ display: none !important; }}
      .container {{ padding: 0; box-shadow: none; max-width: 100%; }}
    }}
  </style>
</head>
<body>
  <div class="action-bar">
    <button class="btn" onclick="window.print()">🖨️ Print / Save as PDF</button>
  </div>
  <div class="container">
    <header>
      <h1>Santheepkumar Veerasamy</h1>
      <div class="subtitle">Staff Software Engineer | Principal / Technical Lead | Full-Stack & Cloud Architect</div>
      <div class="contact-bar">
        <span>+91 86676 62978</span> •
        <span><a href="mailto:devsantheep@gmail.com">devsantheep@gmail.com</a></span> •
        <span><a href="https://linkedin.com/in/santheepkumar">linkedin.com/in/santheepkumar</a></span> •
        <span><a href="https://github.com/Santheepkumar">github.com/Santheepkumar</a></span> •
        <span><a href="https://strawhatdevs.com"><strong>strawhatdevs.com</strong></a></span> •
        <span><a href="https://sk.strawhatdevs.com">sk.strawhatdevs.com</a></span>
      </div>
    </header>
    <div class="meta-box">
      <div><strong>Date:</strong> November 9, 2026</div>
      <div><strong>To:</strong> Hiring Manager / Engineering Leadership at <strong>[Company Name]</strong></div>
      <div><strong>Re:</strong> Application for <strong>[Role Title]</strong></div>
    </div>
    <div class="salutation">Hi [Hiring Manager / Team Name],</div>
{paras_html}
    <div class="closing">
      <div>Thank you for your time and consideration.</div>
      <div style="margin-top: 8px;">Warm regards,</div>
      <div class="signature-name">Santheepkumar Veerasamy</div>
      <div class="signature-title">Staff / Lead Software Engineer & Full-Stack Cloud Architect</div>
      <div class="signature-contacts">+91 86676 62978 • devsantheep@gmail.com • strawhatdevs.com</div>
    </div>
  </div>
</body>
</html>
"""
    with open(f"{base}.html", "w") as f:
        f.write(html_content)

    # DOCX
    body = ET.Element(qn("body"))
    body.append(make_para([make_run("SANTHEEPKUMAR VEERASAMY", bold=True, color="193c66", size="33")], before=0, after=30))
    body.append(make_para([make_run("Staff Software Engineer | Principal / Technical Lead | Full-Stack & Cloud Architect", bold=True, color="2560a7", size="17")], before=0, after=50))
    body.append(make_para([
        make_run("+91 86676 62978  |  devsantheep@gmail.com  |  linkedin.com/in/santheepkumar  |  github.com/Santheepkumar  |  strawhatdevs.com", bold=True, color="434343", size="14")
    ], before=0, after=200))
    body.append(make_para([make_run("Date: ", bold=True, size="16"), make_run("November 9, 2026", color="434343", size="16")], before=0, after=30))
    body.append(make_para([make_run("To: ", bold=True, size="16"), make_run("Hiring Manager / Engineering Leadership at [Company Name]", color="434343", size="16")], before=0, after=30))
    body.append(make_para([make_run("Re: ", bold=True, size="16"), make_run("Application for [Role Title]", bold=True, size="16")], before=0, after=160))
    body.append(make_para([make_run("Hi [Hiring Manager / Team Name],", bold=True, size="16")], before=0, after=100))
    for p in COVER_LETTER_PARAS:
        body.append(make_para([make_run(p, size="16")], before=0, after=120))
    body.append(make_para([make_run("Thank you for your time and consideration.", size="16")], before=0, after=60))
    body.append(make_para([make_run("Warm regards,", size="16")], before=0, after=60))
    body.append(make_para([make_run("Santheepkumar Veerasamy", bold=True, color="193c66", size="18")], before=0, after=20))
    body.append(make_para([make_run("Staff / Lead Software Engineer & Full-Stack Cloud Architect", bold=True, color="2560a7", size="15")], before=0, after=20))
    body.append(make_para([make_run("+91 86676 62978  •  devsantheep@gmail.com  •  strawhatdevs.com", color="434343", size="14")], before=0, after=0))

    sectPr = ET.SubElement(body, qn("sectPr"))
    pgSz = ET.SubElement(sectPr, qn("pgSz"))
    pgSz.attrib[qn("h")] = "15840"
    pgSz.attrib[qn("w")] = "12240"
    pgSz.attrib[qn("orient")] = "portrait"
    pgMar = ET.SubElement(sectPr, qn("pgMar"))
    for k in ["bottom", "top", "left", "right"]:
        pgMar.attrib[qn(k)] = "1440"
    pgMar.attrib[qn("header")] = "720"
    pgMar.attrib[qn("footer")] = "720"
    write_docx_from_body(f"{base}.docx", body)

# ==============================================================================
# LINKEDIN PROFILE OPTIMIZATION GUIDE GENERATOR (MD & HTML)
# ==============================================================================

def generate_linkedin_profile():
    base = f"{DATE_PREFIX}_LinkedIn_Profile_Santheepkumar_V"
    md = """# LinkedIn Profile Optimization — Santheepkumar Veerasamy
*Designed to immediately convert Technical Recruiters, Engineering Managers, VPs of Engineering, and Technical Leads across Staff, Principal, Tech Lead, Forward Deployed, and Senior Software Engineer roles.*

---

## 1. LinkedIn Headline Options (Pick 1 — 220 Character Limit)

- **Option A (Recommended All-Rounder — Covers Staff, Principal, Tech Lead & Senior Search Queries):**
  `Staff / Lead Software Engineer | Full-Stack, Microservices & Kubernetes (React, Next.js, Node.js, NestJS, Go, AWS) | Creator of StrawHatDevs (200+ Tech Videos), ChimeFlow & ENVault`
- **Option B (Builder & Architect Focus):**
  `Staff Software Engineer & Systems Architect | End-to-End Full-Stack, Microservices & Cloud/K8s | Building StrawHatDevs (200+ Videos), ChimeFlow & ENVault | Ex-Heady, Principal Eng`
- **Option C (High-Velocity Product & Forward Deployed Focus):**
  `Lead Full-Stack & Cloud Architect | 0-to-1 Systems, Microservices, Auth/RBAC & Kubernetes | Shipped 8-Month Enterprise MVP in 2 Months | Founder @ StrawHatDevs (200+ Videos)`

---

## 2. LinkedIn "About" / Summary Section (Your Authentic Story — Non-Generic)

```text
My programming journey didn’t start in a Silicon Valley lab—it started in an 11th-standard classroom at a government school in Tamil Nadu, where a computer science teacher showed us a few lines of code and what they could actually do. That single class planted a seed. In 12th standard, I received a state-scheme laptop with 2GB of RAM and no SSD. Slow as it was, that machine became my window into software engineering: I spent hours on Google and YouTube teaching myself how to build things from scratch.

In college (Computer Science & Engineering), I quickly realized that memorizing theory for exams wasn't real engineering. Inspired by seniors who took me to developer conferences in my second year, I went all-in on building real software. In 2019—while still in college—I joined Skematix Technologies as a Software Engineer Intern. Within 2–3 months, I converted to a full-time engineer because programming wasn't just a job to me; I was obsessed with how fast I could learn and ship.

Over the past 7+ years, that same obsession has driven every step of my career:
• Fast-Tracked to Principal Engineer (HOD Digitals): Joined as a Senior Frontend Engineer contributing to Project Gen and quickly earned promotion to Principal Software Engineer, architecting full-stack e-commerce platforms, mobile apps, and cloud systems.
• High-Velocity Enterprise Delivery (Heady): Led engineering for tier-1 brands including HeyDay Skin Care, The Lash Lounge, and Chuck E. Cheese—where I solo-architected and shipped a customer-facing web platform MVP in just 2 months (compressing an 8-month estimate by 75%).
• End-to-End Staff & Lead Ownership (Tuna Software Solution & Science and Humans): Currently, as a Staff Engineer at Tuna, I architected the Pony Up Android IoT retail kiosk (integrating payment terminals, receipt printers, and cash drawers) and led its cloud web migration. Concurrently, as a part-time Lead Software Engineer contractor at Science & Humans, I built core systems from the ground up—owning authentication & RBAC authorization, the notification engine, microservices communication layers, and Kubernetes (K8s) DevOps deployments while mentoring junior engineers and partnering with PMs on module estimation.

Because I’m a builder at heart, I don’t stop coding after work. I build and ship my own production products:
▸ StrawHatDevs (strawhatdevs.com) — An interactive developer learning platform I built from scratch, paired with 200+ programming and system architecture videos on YouTube.
▸ ChimeFlow — A multi-tenant omnichannel notification & cadence orchestration platform (Bun, SvelteKit, PostgreSQL 17, SendGrid, Twilio, FCM) with a write-ahead ledger and fair-claim worker queues.
▸ ENVault (envault.strawhatdevs.com) — A zero-knowledge client-encrypted .env & secret sync platform with a native Go CLI integrated with macOS Keychain, Windows Credential Manager, and Linux Secret Service.
▸ IncidentLens AI — An evidence-grounded multi-tenant incident investigator with PostgreSQL Hybrid FTS Runbook RAG, AES-256-GCM redaction, and a 35-case AI evaluation harness.

Core Stack:
• Frontend: React, Next.js, TypeScript, SvelteKit, React Native, Expo, Tailwind CSS
• Backend & Systems: Node.js, NestJS, Go (Golang), Bun, Express, Microservices, Auth/RBAC, Notification Engines, REST, GraphQL
• Data & Queues: PostgreSQL, MongoDB, Redis, Kafka, RabbitMQ, Supabase, Drizzle, TypeORM
• Cloud & DevOps: Kubernetes (K8s), Docker, AWS, GCP, GitHub Actions, Bitbucket Pipelines, Vercel, Playwright

If you’re a hiring manager, tech lead, or recruiter looking for an engineer who can take a complex product from a blank repo all the way to Kubernetes production infrastructure—let’s connect: devsantheep@gmail.com
```

---

## 3. LinkedIn Experience Entries (Copy-Paste Ready)

### **Staff Engineer** — Tuna Software Solution
*Jun 2025 – Present • Full-time*
- Owned end-to-end architecture, cloud infrastructure, backend microservices, and frontend platforms across two simultaneous enterprise projects.
- Architected and delivered the **Pony Up** Android self-service kiosk (React Native/Expo) integrating physical IoT hardware (payment terminals, thermal receipt printers, cash drawers), and subsequently led its rebuild as a cloud-hosted web platform.
- Established structured code review gates, automated CI/CD pipelines (GitHub Actions/Docker), and engineering standards, improving team delivery velocity by 35%.
- Collaborated directly with product stakeholders to translate business requirements into scalable technical specifications.
- **Skills**: React Native, Next.js, TypeScript, Node.js, IoT Integration, Docker, CI/CD, System Architecture

### **Lead Software Engineer (Part-Time Contractor)** — Science & Humans
*Jun 2025 – Present • Contract*
- Contributed from the ground up with full end-to-end ownership—collaborating directly with Product Managers on roadmap planning, module estimation, architecture, and production shipping.
- Architected and built the complete authentication & role-based authorization (RBAC) system and an event-driven multi-channel notification system from scratch.
- Stood up decoupled backend microservices and engineered the inter-service communication layer, resolving critical bottlenecks and improving system stability.
- Configured and managed containerized application deployments on **Kubernetes (K8s)** and established automated CI/CD DevOps pipelines.
- Guided, onboarded, and mentored junior developers ramping up on the platform.
- **Skills**: Kubernetes (K8s), Microservices, Authentication/RBAC, Notification Systems, TypeScript, Node.js, DevOps, Technical Leadership

### **Founder, Architect & Educator** — Straw Hat Devs (`strawhatdevs.com`)
*Mar 2025 – Present • Self-employed*
- Conceived, architected, and built **strawhatdevs.com**, a full-stack learning platform where developers master software engineering, system design, and cloud architecture.
- Published **200+ programming and system architecture videos** on YouTube, breaking down distributed systems, microservices, and full-stack development.
- Built and launched independent developer products under the Straw Hat Devs ecosystem, including **ChimeFlow** (multi-tenant notification platform), **ENVault** (`envault.strawhatdevs.com`, zero-knowledge `.env` sync + native Go CLI), **IncidentLens AI**, and **Route-Atlas**.
- **Skills**: Next.js, Go, Bun, PostgreSQL, Supabase, Sanity, System Design, Technical Education

### **Senior Software Engineer** — Heady
*May 2023 – Jun 2025 • Full-time*
- Served as primary technical lead interfacing with tier-1 enterprise clients (**HeyDay Skin Care**, **Chuck E. Cheese**, **The Lash Lounge**), translating business goals into resilient production systems.
- Delivered the **Chuck E. Cheese** customer experience web kiosk MVP in **2 months as a solo contributor**, compressing an 8-month original project estimate by 75%.
- Architected a microservices-based backend integration for **HeyDay Skin Care** with Boulevard API, Redis caching, Sentry monitoring, and AWS infrastructure (99.9% uptime).
- Mentored 6+ junior and mid-level developers on modern development practices, AI-assisted coding workflows, and clean architecture.
- **Skills**: React, Next.js, Node.js, NestJS, Redis, Microservices, AWS, Sentry

### **Principal Software Engineer** — HOD Digitals
*Nov 2022 – May 2023 • Full-time (Promoted from Senior Frontend Engineer, Sep 2021 – Nov 2022)*
- Joined as Senior Frontend Engineer contributing to **Project Gen** and client systems; earned fast-track promotion to **Principal Software Engineer** leading architecture across multiple product lines.
- Architected full-stack e-commerce platforms (**Tamil Mart** web & mobile apps with Next.js, Magento GraphQL, Expo), CRM subscription systems (**Vmeals** on AWS), and Playwright CI/CD test automation (**AMS**).
- Led engineering pods, oversaw system design and cloud cost optimization, and served as primary technical liaison to business stakeholders.
- **Skills**: Next.js, React Native, Node.js, Express.js, GraphQL, AWS, Playwright, Team Leadership

### **Full Stack Developer** — Skematix Technologies
*Dec 2019 – Sep 2021 • Full-time (Converted from Intern in <3 Months)*
- Joined during college in 2019 as an Intern Software Engineer and converted to full-time within 2–3 months due to rapid learning velocity and production output.
- Built, optimized, and deployed multiple client-facing web applications across React, Node.js, Express, and relational/NoSQL databases from database design to server provisioning.

---

## 4. LinkedIn Featured & Projects Section Setup
Pin these 4 items in your LinkedIn **Featured** section:
1. **StrawHatDevs Learning Platform & 200+ Architecture Videos** (`https://strawhatdevs.com` & `https://youtube.com/@StrawHatDevs`)
2. **ENVault — Zero-Knowledge Encrypted `.env` Sync & Go CLI** (`https://envault.strawhatdevs.com` / `https://github.com/Santheepkumar/envault-cli`)
3. **ChimeFlow — Multi-Tenant Omnichannel Notification & Cadence Engine**
4. **Portfolio Website** (`https://sk.strawhatdevs.com`)
"""
    with open(f"{base}.md", "w") as f:
        f.write(md)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>LinkedIn Profile Optimization - Santheepkumar Veerasamy</title>
  <style>
    body {{
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      line-height: 1.55;
      color: #1e293b;
      background: #f8fafc;
      padding: 32px 16px;
    }}
    .container {{
      max-width: 880px;
      margin: 0 auto;
      background: #fff;
      padding: 44px 52px;
      border-radius: 8px;
      box-shadow: 0 4px 20px rgba(0,0,0,0.06);
    }}
    h1 {{ font-size: 24px; color: #0f172a; margin-bottom: 6px; }}
    h2 {{ font-size: 16px; color: #0f766e; text-transform: uppercase; border-bottom: 1.5px solid #e2e8f0; padding-bottom: 4px; margin-top: 24px; margin-bottom: 12px; }}
    h3 {{ font-size: 14.5px; color: #0f172a; margin-top: 14px; margin-bottom: 4px; }}
    p, li {{ font-size: 13.5px; color: #334155; margin-bottom: 6px; }}
    pre {{
      background: #f1f5f9;
      padding: 16px;
      border-radius: 6px;
      white-space: pre-wrap;
      font-family: inherit;
      font-size: 13px;
      line-height: 1.55;
      border-left: 4px solid #0f766e;
    }}
    code {{ background: #f1f5f9; padding: 2px 5px; border-radius: 4px; font-size: 12.5px; }}
  </style>
</head>
<body>
  <div class="container">
    <h1>LinkedIn Profile Optimization — Santheepkumar Veerasamy</h1>
    <p><em>Tailored for Technical Recruiters, Hiring Managers, and Engineering Leaders across Principal, Staff, Tech Lead, Forward Deployed, and Senior Software Engineer roles.</em></p>
    <h2>1. Headline Options</h2>
    <ul>
      <li><strong>Option A (Recommended):</strong> <code>Staff / Lead Software Engineer | Full-Stack, Microservices &amp; Kubernetes (React, Next.js, Node.js, NestJS, Go, AWS) | Creator of StrawHatDevs (200+ Tech Videos), ChimeFlow &amp; ENVault</code></li>
      <li><strong>Option B (Architect Focus):</strong> <code>Staff Software Engineer &amp; Systems Architect | End-to-End Full-Stack, Microservices &amp; Cloud/K8s | Building StrawHatDevs (200+ Videos), ChimeFlow &amp; ENVault | Ex-Heady, Principal Eng</code></li>
      <li><strong>Option C (Velocity Focus):</strong> <code>Lead Full-Stack &amp; Cloud Architect | 0-to-1 Systems, Microservices, Auth/RBAC &amp; Kubernetes | Shipped 8-Month Enterprise MVP in 2 Months | Founder @ StrawHatDevs (200+ Videos)</code></li>
    </ul>
    <h2>2. Authentic "About" Section (Your Real Story)</h2>
    <pre>My programming journey didn’t start in a Silicon Valley lab—it started in an 11th-standard classroom at a government school in Tamil Nadu, where a computer science teacher showed us a few lines of code and what they could actually do. That single class planted a seed. In 12th standard, I received a state-scheme laptop with 2GB of RAM and no SSD. Slow as it was, that machine became my window into software engineering: I spent hours on Google and YouTube teaching myself how to build things from scratch.

In college (Computer Science &amp; Engineering), I quickly realized that memorizing theory for exams wasn't real engineering. Inspired by seniors who took me to developer conferences in my second year, I went all-in on building real software. In 2019—while still in college—I joined Skematix Technologies as a Software Engineer Intern. Within 2–3 months, I converted to a full-time engineer because programming wasn't just a job to me; I was obsessed with how fast I could learn and ship.

Over the past 7+ years, that same obsession has driven every step of my career:
• Fast-Tracked to Principal Engineer (HOD Digitals): Joined as a Senior Frontend Engineer contributing to Project Gen and quickly earned promotion to Principal Software Engineer, architecting full-stack e-commerce platforms, mobile apps, and cloud systems.
• High-Velocity Enterprise Delivery (Heady): Led engineering for tier-1 brands including HeyDay Skin Care, The Lash Lounge, and Chuck E. Cheese—where I solo-architected and shipped a customer-facing web platform MVP in just 2 months (compressing an 8-month estimate by 75%).
• End-to-End Staff &amp; Lead Ownership (Tuna Software Solution &amp; Science and Humans): Currently, as a Staff Engineer at Tuna, I architected the Pony Up Android IoT retail kiosk (integrating payment terminals, receipt printers, and cash drawers) and led its cloud web migration. Concurrently, as a part-time Lead Software Engineer contractor at Science &amp; Humans, I built core systems from the ground up—owning authentication &amp; RBAC authorization, the notification engine, microservices communication layers, and Kubernetes (K8s) DevOps deployments while mentoring junior engineers and partnering with PMs on module estimation.

Because I’m a builder at heart, I don’t stop coding after work. I build and ship my own production products:
▸ StrawHatDevs (strawhatdevs.com) — An interactive developer learning platform I built from scratch, paired with 200+ programming and system architecture videos on YouTube.
▸ ChimeFlow — A multi-tenant omnichannel notification &amp; cadence orchestration platform (Bun, SvelteKit, PostgreSQL 17, SendGrid, Twilio, FCM) with a write-ahead ledger and fair-claim worker queues.
▸ ENVault (envault.strawhatdevs.com) — A zero-knowledge client-encrypted .env &amp; secret sync platform with a native Go CLI integrated with macOS Keychain, Windows Credential Manager, and Linux Secret Service.
▸ IncidentLens AI — An evidence-grounded multi-tenant incident investigator with PostgreSQL Hybrid FTS Runbook RAG, AES-256-GCM redaction, and a 35-case AI evaluation harness.

Core Stack:
• Frontend: React, Next.js, TypeScript, SvelteKit, React Native, Expo, Tailwind CSS
• Backend &amp; Systems: Node.js, NestJS, Go (Golang), Bun, Express, Microservices, Auth/RBAC, Notification Engines, REST, GraphQL
• Data &amp; Queues: PostgreSQL, MongoDB, Redis, Kafka, RabbitMQ, Supabase, Drizzle, TypeORM
• Cloud &amp; DevOps: Kubernetes (K8s), Docker, AWS, GCP, GitHub Actions, Bitbucket Pipelines, Vercel, Playwright

If you’re a hiring manager, tech lead, or recruiter looking for an engineer who can take a complex product from a blank repo all the way to Kubernetes production infrastructure—let’s connect: devsantheep@gmail.com</pre>
  </div>
</body>
</html>"""
    with open(f"{base}.html", "w") as f:
        f.write(html)
    print(f"Generated LinkedIn Profile Guide: {base}.md and {base}.html")

def main():
    # Clean old 2026-09-22 generated files (except Original template)
    for fname in os.listdir("."):
        if fname.startswith("2026-09-22_") and fname != TEMPLATE_DOCX:
            os.remove(fname)
            print(f"Removed old version: {fname}")

    # Generate all 5 resumes
    for cfg in RESUMES:
        base = f"{DATE_PREFIX}_{cfg['slug']}_Resume_Santheepkumar_V"
        with open(f"{base}.md", "w") as f:
            f.write(generate_resume_md(cfg))
        with open(f"{base}.html", "w") as f:
            f.write(generate_resume_html(cfg))
        write_docx_from_body(f"{base}.docx", build_resume_docx_body(cfg))
        print(f"Completed MD/HTML/DOCX for: {base}")

    # Generate Cover Letter
    generate_cover_letter_all()

    # Generate LinkedIn Profile
    generate_linkedin_profile()

if __name__ == "__main__":
    main()
