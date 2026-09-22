import zipfile
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

namespaces = {
    "w": W_NS,
    "w14": W14_NS,
    "r": R_NS
}

def qn(tag):
    return f"{{{W_NS}}}{tag}"

def qn_w14(tag):
    return f"{{{W14_NS}}}{tag}"

def qn_r(tag):
    return f"{{{R_NS}}}{tag}"

para_id_counter = 1

def next_para_id():
    global para_id_counter
    p_id = f"{para_id_counter:08X}"
    para_id_counter += 1
    return p_id

def create_p_bdr():
    pBdr = ET.Element(qn("pBdr"))
    for side in ["top", "left", "bottom", "right", "between"]:
        bdr = ET.SubElement(pBdr, qn(side))
        bdr.attrib[qn("color")] = "e3e3e3"
        bdr.attrib[qn("space")] = "0"
        bdr.attrib[qn("sz")] = "0"
        bdr.attrib[qn("val")] = "none"
    return pBdr

def make_run(text, bold=False, italic=False, color="141414", size="15", underline=False, br=False):
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

    if br:
        ET.SubElement(r, qn("br"))

    if text:
        t = ET.SubElement(r, qn("t"))
        t.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"
        t.text = text

    return r

def make_section_header(title, is_first=False):
    p = ET.Element(qn("p"))
    p.attrib[qn("rsidR")] = "00000000"
    p.attrib[qn("rsidDel")] = "00000000"
    p.attrib[qn("rsidP")] = "00000000"
    p.attrib[qn("rsidRDefault")] = "00000000"
    p.attrib[qn("rsidRPr")] = "00000000"
    p.attrib[qn_w14("paraId")] = next_para_id()

    pPr = ET.SubElement(p, qn("pPr"))
    pPr.append(create_p_bdr())
    shd = ET.SubElement(pPr, qn("shd"))
    shd.attrib[qn("fill")] = "ffffff"
    shd.attrib[qn("val")] = "clear"

    spc = ET.SubElement(pPr, qn("spacing"))
    spc.attrib[qn("before")] = "0" if is_first else "240"
    spc.attrib[qn("after")] = "80"
    spc.attrib[qn("lineRule")] = "auto"

    rPr = ET.SubElement(pPr, qn("rPr"))
    b = ET.SubElement(rPr, qn("b"))
    b.attrib[qn("val")] = "1"
    bCs = ET.SubElement(rPr, qn("bCs"))
    bCs.attrib[qn("val")] = "1"
    col = ET.SubElement(rPr, qn("color"))
    col.attrib[qn("val")] = "193c66"
    sz = ET.SubElement(rPr, qn("sz"))
    sz.attrib[qn("val")] = "17"
    szCs = ET.SubElement(rPr, qn("szCs"))
    szCs.attrib[qn("val")] = "17"

    run = make_run(title, bold=True, color="193c66", size="17")
    p.append(run)
    return p

def make_job_header(role, company, date):
    p = ET.Element(qn("p"))
    p.attrib[qn("rsidR")] = "00000000"
    p.attrib[qn("rsidDel")] = "00000000"
    p.attrib[qn("rsidP")] = "00000000"
    p.attrib[qn("rsidRDefault")] = "00000000"
    p.attrib[qn("rsidRPr")] = "00000000"
    p.attrib[qn_w14("paraId")] = next_para_id()

    pPr = ET.SubElement(p, qn("pPr"))
    pPr.append(create_p_bdr())
    shd = ET.SubElement(pPr, qn("shd"))
    shd.attrib[qn("fill")] = "ffffff"
    shd.attrib[qn("val")] = "clear"

    spc = ET.SubElement(pPr, qn("spacing"))
    spc.attrib[qn("before")] = "160"
    spc.attrib[qn("after")] = "40"
    spc.attrib[qn("lineRule")] = "auto"

    # Role
    p.append(make_run(role, bold=True, color="141414", size="17"))
    # Separator
    p.append(make_run("  |  ", bold=True, color="434343", size="15"))
    # Company
    p.append(make_run(company, bold=True, color="2560a7", size="16"))
    # Date
    p.append(make_run(f"   {date}", bold=True, italic=True, color="434343", size="14"))
    return p

def make_bullet(text_or_runs, before=60, after=0):
    p = ET.Element(qn("p"))
    p.attrib[qn("rsidR")] = "00000000"
    p.attrib[qn("rsidDel")] = "00000000"
    p.attrib[qn("rsidP")] = "00000000"
    p.attrib[qn("rsidRDefault")] = "00000000"
    p.attrib[qn("rsidRPr")] = "00000000"
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

    rPr = ET.SubElement(pPr, qn("rPr"))
    col = ET.SubElement(rPr, qn("color"))
    col.attrib[qn("val")] = "141414"
    sz = ET.SubElement(rPr, qn("sz"))
    sz.attrib[qn("val")] = "15"
    szCs = ET.SubElement(rPr, qn("szCs"))
    szCs.attrib[qn("val")] = "15"

    if isinstance(text_or_runs, str):
        p.append(make_run(text_or_runs, bold=False, color="141414", size="15"))
    else:
        for r in text_or_runs:
            p.append(r)
    return p

def make_para(text, bold=False, italic=False, color="141414", size="15", before=0, after=60):
    p = ET.Element(qn("p"))
    p.attrib[qn("rsidR")] = "00000000"
    p.attrib[qn("rsidDel")] = "00000000"
    p.attrib[qn("rsidP")] = "00000000"
    p.attrib[qn("rsidRDefault")] = "00000000"
    p.attrib[qn("rsidRPr")] = "00000000"
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

    p.append(make_run(text, bold=bold, italic=italic, color=color, size=size))
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

    # Fill rows 3 per row
    for i in range(0, len(items), 3):
        tr = ET.SubElement(tbl, qn("tr"))
        trPr = ET.SubElement(tr, qn("trPr"))
        cs = ET.SubElement(trPr, qn("cantSplit"))
        cs.attrib[qn("val")] = "0"
        trH = ET.SubElement(trPr, qn("trHeight"))
        trH.attrib[qn("val")] = "340"
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
            p.attrib[qn("rsidR")] = "00000000"
            p.attrib[qn("rsidDel")] = "00000000"
            p.attrib[qn("rsidP")] = "00000000"
            p.attrib[qn("rsidRDefault")] = "00000000"
            p.attrib[qn("rsidRPr")] = "00000000"
            p.attrib[qn_w14("paraId")] = next_para_id()

            pPr = ET.SubElement(p, qn("pPr"))
            pPr.append(create_p_bdr())
            shd = ET.SubElement(pPr, qn("shd"))
            shd.attrib[qn("fill")] = "ffffff"
            shd.attrib[qn("val")] = "clear"

            p.append(make_run(f"▸ {item}", bold=True, color="141414", size="14"))

        # Pad row if less than 3
        while len(row_items) < 3:
            row_items.append("")
            tc = ET.SubElement(tr, qn("tc"))
            tcPr = ET.SubElement(tc, qn("tcPr"))
            tcBorders = ET.SubElement(tcPr, qn("tcBorders"))
            for b in ["top", "left", "bottom", "right"]:
                el = ET.SubElement(tcBorders, qn(b))
                el.attrib[qn("color")] = "000000"
                el.attrib[qn("space")] = "0"
                el.attrib[qn("sz")] = "6"
                el.attrib[qn("val")] = "single"
            p = ET.SubElement(tc, qn("p"))
            p.attrib[qn_w14("paraId")] = next_para_id()
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
    gc1.attrib[qn("w")] = "1550"
    gc2 = ET.SubElement(tblGrid, qn("gridCol"))
    gc2.attrib[qn("w")] = "5470"

    for cat, val in skills_dict.items():
        tr = ET.SubElement(tbl, qn("tr"))
        trPr = ET.SubElement(tr, qn("trPr"))
        cs = ET.SubElement(trPr, qn("cantSplit"))
        cs.attrib[qn("val")] = "0"
        trH = ET.SubElement(trPr, qn("trHeight"))
        trH.attrib[qn("val")] = "300"
        trH.attrib[qn("hRule")] = "atLeast"

        # Col 1
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
        pPr1 = ET.SubElement(p1, qn("pPr"))
        pPr1.append(create_p_bdr())
        shd1 = ET.SubElement(pPr1, qn("shd"))
        shd1.attrib[qn("fill")] = "ffffff"
        shd1.attrib[qn("val")] = "clear"
        p1.append(make_run(cat, bold=True, color="193c66", size="14"))

        # Col 2
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
        pPr2 = ET.SubElement(p2, qn("pPr"))
        pPr2.append(create_p_bdr())
        shd2 = ET.SubElement(pPr2, qn("shd"))
        shd2.attrib[qn("fill")] = "ffffff"
        shd2.attrib[qn("val")] = "clear"
        p2.append(make_run(val, bold=False, color="141414", size="14"))

    return tbl

def build_body():
    body = ET.Element(qn("body"))

    # Name
    p_name = ET.Element(qn("p"))
    p_name.attrib[qn_w14("paraId")] = next_para_id()
    pPr = ET.SubElement(p_name, qn("pPr"))
    pPr.append(create_p_bdr())
    shd = ET.SubElement(pPr, qn("shd"))
    shd.attrib[qn("fill")] = "ffffff"
    shd.attrib[qn("val")] = "clear"
    p_name.append(make_run("SANTHEEPKUMAR VEERASAMY", bold=True, color="193c66", size="33"))
    body.append(p_name)

    # Subtitle
    p_sub = ET.Element(qn("p"))
    p_sub.attrib[qn_w14("paraId")] = next_para_id()
    pPr = ET.SubElement(p_sub, qn("pPr"))
    pPr.append(create_p_bdr())
    shd = ET.SubElement(pPr, qn("shd"))
    shd.attrib[qn("fill")] = "ffffff"
    shd.attrib[qn("val")] = "clear"
    p_sub.append(make_run("Staff Software Engineer | Full-Stack & Cloud Solutions Architect", bold=True, color="2560a7", size="17"))
    body.append(p_sub)

    # Contacts
    p_contact = ET.Element(qn("p"))
    p_contact.attrib[qn_w14("paraId")] = next_para_id()
    pPr = ET.SubElement(p_contact, qn("pPr"))
    pPr.append(create_p_bdr())
    shd = ET.SubElement(pPr, qn("shd"))
    shd.attrib[qn("fill")] = "ffffff"
    shd.attrib[qn("val")] = "clear"
    
    p_contact.append(make_run("+91 86676 62978  |  devsantheep@gmail.com  |  ", bold=True, color="434343", size="14"))
    p_contact.append(make_run("linkedin.com/in/santheepkumar", bold=True, color="2560a7", size="14", underline=True))
    p_contact.append(make_run("  |  ", bold=True, color="434343", size="14"))
    p_contact.append(make_run("github.com/Santheepkumar", bold=True, color="1155cc", size="14", underline=True))
    p_contact.append(make_run("  |  ", bold=True, color="434343", size="14"))
    p_contact.append(make_run("strawhatdevs.com", bold=True, color="2560a7", size="14", underline=True))
    p_contact.append(make_run("  |  ", bold=True, color="434343", size="14"))
    p_contact.append(make_run("youtube.com/@StrawHatDevs", bold=True, color="1155cc", size="14", underline=True))
    p_contact.append(make_run("  |  ", bold=True, color="434343", size="14"))
    p_contact.append(make_run("sk.strawhatdevs.com", bold=True, color="2560a7", size="14", underline=True))
    body.append(p_contact)

    # PROFESSIONAL SUMMARY
    body.append(make_section_header("PROFESSIONAL SUMMARY"))
    summary_text = (
        "Results-driven Staff Software Engineer and End-to-End Solutions Architect with 7+ years of experience "
        "designing, building, and deploying mission-critical systems across enterprise and high-growth environments. "
        "Comprehensive hands-on mastery spanning frontend platforms, distributed microservices backends, cloud infrastructure "
        "provisioning, and automated CI/CD pipelines. Proven track record of taking complex products from initial whiteboard "
        "architecture through production launch at 4x speed—including compressing an 8-month enterprise MVP into 2 months as a sole "
        "contributor. Founder and Chief Architect of Straw Hat Devs (strawhatdevs.com) and technical educator with 200+ architecture "
        "and programming videos, demonstrating deep domain authority in system design, distributed systems, and modern cloud ecosystems. "
        "Experienced technical leader who establishes engineering rigor, mentors cross-functional teams, and leverages AI-assisted "
        "development workflows to maximize velocity and code quality."
    )
    body.append(make_para(summary_text, bold=False, color="141414", size="15", before=60, after=120))

    # CORE COMPETENCIES
    body.append(make_section_header("CORE COMPETENCIES"))
    comps = [
        "End-to-End System Architecture", "Cloud Infra & DevOps (AWS/GCP)", "Automated CI/CD Pipelines",
        "Microservices & Distributed Systems", "Full-Stack Product Delivery", "Rapid MVP & Iterative Prototyping",
        "IoT & Hardware Peripherals", "Cross-Functional Leadership", "Technical Mentorship & Code Reviews",
        "AI-Assisted Engineering Workflows", "Developer Education & Content", "High Availability & Scalability"
    ]
    body.append(make_competencies_table(comps))

    # TECHNICAL SKILLS
    body.append(make_section_header("TECHNICAL SKILLS"))
    skills = {
        "Cloud & DevOps": "Docker, Kubernetes, AWS (ECS, Lambda, S3, RDS, CloudFront), GCP, GitHub Actions, Bitbucket Pipelines, Jenkins, Vercel, Terraform, Sentry, Axiom",
        "Architecture": "Microservices, Monorepo (Turborepo), Event-Driven, RESTful APIs, GraphQL, Redis, Kafka, RabbitMQ, Zero-Downtime Releases",
        "Frontend": "React, Next.js (App & Pages Router), TypeScript, JavaScript, React Native, Expo, Tailwind CSS, SCSS, State Management (Zustand, Redux)",
        "Backend": "Node.js, Express.js, NestJS, GoLang, REST APIs, GraphQL, Microservices Architecture",
        "Databases": "PostgreSQL, MongoDB, Supabase, Mongoose, Sequelize, TypeORM, Prisma, Query Tuning",
        "Testing & QA": "Playwright, Jest, React Testing Library, End-to-End Automated Testing, Static Analysis",
        "CMS & Tools": "Sanity, WordPress, Supabase, Axiom, Sentry, Boulevard, Webflow"
    }
    body.append(make_skills_table(skills))

    # FOUNDER, CREATOR & COMMUNITY LEADERSHIP
    body.append(make_section_header("FOUNDER, CREATOR & COMMUNITY LEADERSHIP"))
    body.append(make_job_header("Founder & Chief Architect", "Straw Hat Devs (strawhatdevs.com)", "Mar 2025 – Present"))
    body.append(make_bullet([
        make_run("Platform Architecture & Delivery: ", bold=True, color="141414", size="15"),
        make_run("Architected and built strawhatdevs.com from scratch as a modern developer learning platform featuring structured courses in full-stack engineering, system design, and cloud infrastructure.", color="141414", size="15")
    ]))
    body.append(make_bullet([
        make_run("Full-Stack & Cloud Infrastructure: ", bold=True, color="141414", size="15"),
        make_run("Engineered the system using Next.js (App Router), Supabase, PostgreSQL, Sanity CMS, and Axiom with fully automated CI/CD pipelines deploying to serverless cloud infrastructure.", color="141414", size="15")
    ]))
    body.append(make_bullet([
        make_run("Curriculum & Mentorship: ", bold=True, color="141414", size="15"),
        make_run("Designed comprehensive, project-based engineering curricula empowering hundreds of developers to build and deploy production-grade software.", color="141414", size="15")
    ]))

    body.append(make_job_header("Technical Content Creator & Educator", "Straw Hat Devs (YouTube)", "200+ Videos Published"))
    body.append(make_bullet([
        make_run("Technical Thought Leadership: ", bold=True, color="141414", size="15"),
        make_run("Produced and published 200+ deep-dive videos on software architecture, microservices, cloud deployments, full-stack frameworks, and distributed system patterns.", color="141414", size="15")
    ]))
    body.append(make_bullet([
        make_run("Developer Community Impact: ", bold=True, color="141414", size="15"),
        make_run("Built an active global audience of software developers through pragmatic, code-first breakdowns of complex real-world enterprise architectures.", color="141414", size="15")
    ]))

    # PROFESSIONAL EXPERIENCE
    body.append(make_section_header("PROFESSIONAL EXPERIENCE"))

    # Tuna Software Solution
    body.append(make_job_header("Staff Engineer", "Tuna Software Solution", "Jun 2025 – Present"))
    body.append(make_bullet([
        make_run("End-to-End Ownership: ", bold=True, color="141414", size="15"),
        make_run("Owned end-to-end architecture, cloud infrastructure, backend microservices, and frontend platforms across two concurrent enterprise projects, driving key feature releases ahead of schedule.", color="141414", size="15")
    ]))
    body.append(make_bullet([
        make_run("Pony Up Retail Kiosk & Web Migration: ", bold=True, color="141414", size="15"),
        make_run("Architected and delivered an Android Kiosk application (React Native/Expo) with IoT peripherals (payment gateways, receipt printers, cash drawers), then spearheaded its rebuild as a cloud web platform, reducing deployment complexity.", color="141414", size="15")
    ]))
    body.append(make_bullet([
        make_run("DevOps & CI/CD Pipelines: ", bold=True, color="141414", size="15"),
        make_run("Established engineering best practices including structured code review processes, automated CI/CD pipelines (GitHub Actions/Docker), and coding standards—improving delivery velocity by 35% and code quality.", color="141414", size="15")
    ]))
    body.append(make_bullet([
        make_run("Stakeholder Alignment: ", bold=True, color="141414", size="15"),
        make_run("Collaborated directly with product stakeholders to translate customer requirements into technical solutions, acting as the primary technical point of contact for project alignment.", color="141414", size="15")
    ]))

    # Heady
    body.append(make_job_header("Senior Software Engineer", "Heady", "May 2023 – Jun 2025"))
    body.append(make_bullet([
        make_run("Enterprise Client Leadership: ", bold=True, color="141414", size="15"),
        make_run("Served as primary technical lead interfacing with enterprise clients (HeyDay Skin Care, The Lash Lounge, Chuck E. Cheese), translating business requirements into scalable, production-ready systems.", color="141414", size="15")
    ]))
    body.append(make_bullet([
        make_run("Chuck E. Cheese Web MVP (75% Faster Delivery): ", bold=True, color="141414", size="15"),
        make_run("Delivered Chuck E. Cheese web kiosk MVP in 2 months as a solo contributor, compressing an 8-month original project timeline by 75% through focused engineering and rapid iteration.", color="141414", size="15")
    ]))
    body.append(make_bullet([
        make_run("HeyDay Skin Care Microservices Ecosystem: ", bold=True, color="141414", size="15"),
        make_run("Architected a microservices-based backend integration with Boulevard (booking platform), Redis caching, Sentry error monitoring, and AWS infrastructure for enterprise-scale 99.9% reliability.", color="141414", size="15")
    ]))
    body.append(make_bullet([
        make_run("CI/CD Automation & Booking Platforms: ", bold=True, color="141414", size="15"),
        make_run("Engineered Next.js platforms with automated Bitbucket Pipelines and Vercel deployments for The Lash Lounge and Vmeals subscription management.", color="141414", size="15")
    ]))
    body.append(make_bullet([
        make_run("AI-Assisted Workflows & Mentorship: ", bold=True, color="141414", size="15"),
        make_run("Mentored junior and mid-level developers on modern development practices including AI-assisted coding, conducting code reviews, and establishing architectural patterns that reduced regression rates.", color="141414", size="15")
    ]))

    # HOD Digitals - Principal
    body.append(make_job_header("Principal Software Engineer", "HOD Digitals", "Nov 2022 – May 2023"))
    body.append(make_bullet([
        make_run("Technical Direction & Architecture: ", bold=True, color="141414", size="15"),
        make_run("Led a team of engineers across multiple product lines, providing technical direction, architectural guidance, and hands-on mentorship across full-stack systems.", color="141414", size="15")
    ]))
    body.append(make_bullet([
        make_run("Scalability & System Design: ", bold=True, color="141414", size="15"),
        make_run("Oversaw project planning, system design, and implementation; made key architectural decisions that improved system scalability and performance while cutting cloud infrastructure costs.", color="141414", size="15")
    ]))
    body.append(make_bullet([
        make_run("Business Stakeholder Liaison: ", bold=True, color="141414", size="15"),
        make_run("Acted as the primary technical liaison between engineering and business stakeholders, ensuring delivery alignment and proactive risk communication.", color="141414", size="15")
    ]))

    # HOD Digitals - Software Engineer
    body.append(make_job_header("Software Engineer", "HOD Digitals", "Sep 2021 – Nov 2022"))
    body.append(make_bullet([
        make_run("Full-Lifecycle Delivery: ", bold=True, color="141414", size="15"),
        make_run("Contributed across the full software development lifecycle—from requirements analysis through cloud deployment—on multiple concurrent client projects.", color="141414", size="15")
    ]))
    body.append(make_bullet([
        make_run("Feature Engineering & Optimization: ", bold=True, color="141414", size="15"),
        make_run("Developed new features and enhanced existing applications based on direct client feedback, improving product adoption, performance, and user satisfaction.", color="141414", size="15")
    ]))

    # Skematix Technologies
    body.append(make_job_header("Full Stack Developer", "Skematix Technologies", "Dec 2019 – Sep 2021"))
    body.append(make_bullet([
        make_run("Full-Stack Web Engineering: ", bold=True, color="141414", size="15"),
        make_run("Developed and maintained multiple client-facing web applications with React, Node.js, Express, and databases, with a focus on user experience and performance optimization.", color="141414", size="15")
    ]))
    body.append(make_bullet([
        make_run("End-to-End Execution: ", bold=True, color="141414", size="15"),
        make_run("Handled both frontend and backend responsibilities across the stack, gaining deep exposure to end-to-end product ownership and server provisioning early in career.", color="141414", size="15")
    ]))

    # KEY PROJECTS & DEPLOYMENTS
    body.append(make_section_header("KEY PROJECTS & DEPLOYMENTS"))

    # Project 1: Straw Hat Devs
    body.append(make_job_header("Straw Hat Devs — Learning Platform", "strawhatdevs.com", "Mar 2025 – Present"))
    body.append(make_bullet("Architected and launched full-stack edtech platform featuring interactive developer courses, role-based auth, and video delivery."))
    body.append(make_bullet("Tech Stack: Next.js (App Router), TypeScript, PostgreSQL, Supabase, Sanity, Axiom, Vercel CI/CD"))

    # Project 2: Pony Up
    body.append(make_job_header("Pony Up — Android Self-Service Kiosk", "Staff Engineer", "Jun 2025 – Present"))
    body.append(make_bullet("Engineered a React Native/Expo self-service kiosk deployed in shopping mall environments, integrating IoT peripherals including payment terminals, receipt printers, and cash drawers."))
    body.append(make_bullet("Implemented secure end-to-end payment and order processing workflows, connecting kiosk hardware with backend services and external payment providers."))
    body.append(make_bullet("Tech Stack: React Native, Expo, IoT Integrations, Android Kiosk Mode, REST APIs"))

    # Project 3: Chuck E. Cheese
    body.append(make_job_header("Chuck E. Cheese — Web Experience Platform", "Senior Engineer", "Jun 2025 – Present"))
    body.append(make_bullet("Rebuilt the Chuck E. Cheese kiosk as a web-based customer experience platform enabling product interaction via QR code scanning — delivered MVP in 2 months vs. an 8-month original estimate."))
    body.append(make_bullet("Functioned as a sole contributor handling full frontend development and system implementation, coordinating directly with design and product teams for real-time requirement validation."))
    body.append(make_bullet("Tech Stack: React, Next.js, QR Integration, Responsive Web, AWS CloudFront"))

    # Project 4: HeyDay Skin Care
    body.append(make_job_header("HeyDay Skin Care — Shop App & Booking Widget", "Lead Frontend Developer", "May 2023 – Present"))
    body.append(make_bullet("Built and maintained a customer-facing skin health shop app and booking widget enabling facial appointment scheduling, treatment tracking, and client profile management."))
    body.append(make_bullet("Architected a microservices-based backend integration with Boulevard (booking platform), Sentry (error monitoring), and AWS infrastructure for enterprise-scale reliability."))
    body.append(make_bullet("Tech Stack: ReactJS, Node.js, NestJS, Redis, Microservices, AWS, Sentry, Boulevard"))

    # Project 5: The Lash Lounge
    body.append(make_job_header("The Lash Lounge — Appointment Booking Platform", "Lead Frontend Developer", "Mar 2024 – Present"))
    body.append(make_bullet("Delivered a seamless booking experience for eyelash services with appointment tracking and treatment record management for a national beauty services chain."))
    body.append(make_bullet("Tech Stack: Next.js, Vercel, Bitbucket Pipelines, Sentry"))

    # Project 6: Vmeals
    body.append(make_job_header("Vmeals — CRM & Subscription Platform", "Team Lead", "Jan 2023 – Feb 2023"))
    body.append(make_bullet("Led development of an enhanced CRM system automating customer interactions and subscription management, plus a customer-facing app for subscription plan customization."))
    body.append(make_bullet("Tech Stack: Next.js, Node.js, Express.js, AWS, GitHub Actions"))

    # Project 7: AMS Testing Automation
    body.append(make_job_header("AMS Testing Automation Platform", "Lead Developer", "Jun 2022 – May 2023"))
    body.append(make_bullet("Implemented end-to-end test automation for a CI/CD-integrated platform managing app release schedules, beta builds, and developer tooling account management."))
    body.append(make_bullet("Tech Stack: Playwright, Bitbucket Pipelines, Docker"))

    # EDUCATION
    body.append(make_section_header("EDUCATION"))
    body.append(make_job_header("Bachelor of Engineering (B.E.) — Computer Science & Engineering", "Anna University", "2015 – 2019"))
    body.append(make_para("Tamil Nadu, India", color="434343", size="14", before=20, after=60))

    # Add sectPr at end
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

def main():
    src_docx = "2026-09-22_Original_Resume_Santheepkumar_V.docx"
    dest_docx = "2026-09-22_Staff_Software_Engineer_Resume_Santheepkumar_V.docx"

    with zipfile.ZipFile(src_docx, "r") as z_in:
        # Read the existing document.xml to obtain exact namespaces and root wrapper
        orig_doc_xml = z_in.read("word/document.xml")
        orig_root = ET.fromstring(orig_doc_xml)

        # Clear existing body and replace with our new built body
        orig_body = orig_root.find(qn("body"))
        if orig_body is not None:
            orig_root.remove(orig_body)
        
        new_body = build_body()
        orig_root.append(new_body)

        new_doc_xml = ET.tostring(orig_root, encoding="utf-8", xml_declaration=True)

        # Write to destination docx
        with zipfile.ZipFile(dest_docx, "w", compression=zipfile.ZIP_DEFLATED) as z_out:
            for item in z_in.infolist():
                if item.filename == "word/document.xml":
                    z_out.writestr(item.filename, new_doc_xml)
                else:
                    z_out.writestr(item.filename, z_in.read(item.filename))

    print(f"Successfully generated {dest_docx}!")

if __name__ == "__main__":
    main()
