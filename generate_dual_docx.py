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

para_id_counter = 100

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

    p.append(make_run(role, bold=True, color="141414", size="17"))
    p.append(make_run("  |  ", bold=True, color="434343", size="15"))
    p.append(make_run(company, bold=True, color="2560a7", size="16"))
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
            p.attrib[qn_w14("paraId")] = next_para_id()
            pPr = ET.SubElement(p, qn("pPr"))
            pPr.append(create_p_bdr())
            shd = ET.SubElement(pPr, qn("shd"))
            shd.attrib[qn("fill")] = "ffffff"
            shd.attrib[qn("val")] = "clear"

            p.append(make_run(f"▸ {item}", bold=True, color="141414", size="14"))

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
    gc1.attrib[qn("w")] = "1650"
    gc2 = ET.SubElement(tblGrid, qn("gridCol"))
    gc2.attrib[qn("w")] = "5370"

    for cat, val in skills_dict.items():
        tr = ET.SubElement(tbl, qn("tr"))
        trPr = ET.SubElement(tr, qn("trPr"))
        cs = ET.SubElement(trPr, qn("cantSplit"))
        cs.attrib[qn("val")] = "0"
        trH = ET.SubElement(trPr, qn("trHeight"))
        trH.attrib[qn("val")] = "300"
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
        pPr1 = ET.SubElement(p1, qn("pPr"))
        pPr1.append(create_p_bdr())
        shd1 = ET.SubElement(pPr1, qn("shd"))
        shd1.attrib[qn("fill")] = "ffffff"
        shd1.attrib[qn("val")] = "clear"
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
        pPr2 = ET.SubElement(p2, qn("pPr"))
        pPr2.append(create_p_bdr())
        shd2 = ET.SubElement(pPr2, qn("shd"))
        shd2.attrib[qn("fill")] = "ffffff"
        shd2.attrib[qn("val")] = "clear"
        p2.append(make_run(val, bold=False, color="141414", size="14"))

    return tbl

def build_resume_body(config):
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

    # Subtitle / Role Headline
    p_sub = ET.Element(qn("p"))
    p_sub.attrib[qn_w14("paraId")] = next_para_id()
    pPr = ET.SubElement(p_sub, qn("pPr"))
    pPr.append(create_p_bdr())
    shd = ET.SubElement(pPr, qn("shd"))
    shd.attrib[qn("fill")] = "ffffff"
    shd.attrib[qn("val")] = "clear"
    p_sub.append(make_run(config["headline"], bold=True, color="2560a7", size="17"))
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
    body.append(make_para(config["summary"], bold=False, color="141414", size="15", before=60, after=120))

    # CORE COMPETENCIES
    body.append(make_section_header("CORE COMPETENCIES"))
    body.append(make_competencies_table(config["competencies"]))

    # TECHNICAL SKILLS
    body.append(make_section_header("TECHNICAL SKILLS"))
    body.append(make_skills_table(config["skills"]))

    # FOUNDER, CREATOR & COMMUNITY LEADERSHIP
    body.append(make_section_header("FOUNDER, CREATOR & COMMUNITY LEADERSHIP"))
    body.append(make_job_header("Founder & Chief Architect", "Straw Hat Devs (strawhatdevs.com)", "Mar 2025 – Present"))
    for lead, text in config["straw_hat_platform_bullets"]:
        body.append(make_bullet([
            make_run(lead, bold=True, color="141414", size="15"),
            make_run(text, color="141414", size="15")
        ]))

    body.append(make_job_header("Technical Content Creator & Educator", "Straw Hat Devs (YouTube)", "200+ Videos Published"))
    for lead, text in config["straw_hat_youtube_bullets"]:
        body.append(make_bullet([
            make_run(lead, bold=True, color="141414", size="15"),
            make_run(text, color="141414", size="15")
        ]))

    # PROFESSIONAL EXPERIENCE
    body.append(make_section_header("PROFESSIONAL EXPERIENCE"))
    for exp in config["experiences"]:
        body.append(make_job_header(exp["role"], exp["company"], exp["date"]))
        for lead, text in exp["bullets"]:
            body.append(make_bullet([
                make_run(lead, bold=True, color="141414", size="15"),
                make_run(text, color="141414", size="15")
            ]))

    # KEY PROJECTS & DEPLOYMENTS
    body.append(make_section_header(config["projects_header"]))
    for proj in config["projects"]:
        body.append(make_job_header(proj["title"], proj["role_or_domain"], proj["date"]))
        for b in proj["bullets"]:
            body.append(make_bullet(b))

    # EDUCATION
    body.append(make_section_header("EDUCATION"))
    body.append(make_job_header("Bachelor of Engineering (B.E.) — Computer Science & Engineering", "Anna University", "2015 – 2019"))
    body.append(make_para("Tamil Nadu, India", color="434343", size="14", before=20, after=60))

    # sectPr
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

def generate_docx_file(src_template, dest_filename, config):
    with zipfile.ZipFile(src_template, "r") as z_in:
        orig_doc_xml = z_in.read("word/document.xml")
        orig_root = ET.fromstring(orig_doc_xml)

        orig_body = orig_root.find(qn("body"))
        if orig_body is not None:
            orig_root.remove(orig_body)
        
        new_body = build_resume_body(config)
        orig_root.append(new_body)

        new_doc_xml = ET.tostring(orig_root, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(dest_filename, "w", compression=zipfile.ZIP_DEFLATED) as z_out:
            for item in z_in.infolist():
                if item.filename == "word/document.xml":
                    z_out.writestr(item.filename, new_doc_xml)
                else:
                    z_out.writestr(item.filename, z_in.read(item.filename))
    print(f"Generated {dest_filename} successfully.")

# Configuration 1: Technical Lead
tech_lead_config = {
    "headline": "Technical Lead | Staff Full-Stack & Cloud Solutions Architect",
    "summary": (
        "Impact-driven Technical Lead and Staff Software Architect with 7+ years of experience leading the technical "
        "vision, architectural design, and end-to-end delivery of complex enterprise software systems. Comprehensive "
        "hands-on expertise spanning reactive frontend architectures, distributed microservices backends, cloud infrastructure "
        "orchestration (AWS/GCP/Docker/K8s), and automated CI/CD release engineering. Proven engineering leader and multiplier "
        "who establishes engineering best practices, implements automated quality gates, and mentors cross-functional engineering "
        "pods to boost team delivery velocity by 35% while reducing defect escapes. Known for solving high-complexity architectural "
        "challenges and accelerating time-to-market—including delivering an enterprise web platform MVP in 2 months (compressing "
        "an 8-month baseline by 75%). Founder and Chief Architect of Straw Hat Devs (strawhatdevs.com) and technical educator with "
        "200+ architecture and programming videos, recognized for thought leadership in system design, distributed systems, and "
        "modern AI-assisted software engineering."
    ),
    "competencies": [
        "Architectural Governance & System Design", "Cloud Infrastructure (AWS/GCP/K8s)", "Automated CI/CD & Release Engineering",
        "Microservices & Event-Driven Systems", "Technical Mentorship & Code Standards", "Engineering Velocity (+35% Boost)",
        "Zero-Downtime Releases & Reliability", "Cross-Functional Stakeholder Alignment", "IoT & Hardware System Integration",
        "AI-Assisted Development Workflows", "Developer Education & Architecture Content", "High Availability & Scalability"
    ],
    "skills": {
        "Architectural Leadership": "Distributed Systems, Microservices, Event-Driven Architecture, Monorepos (Turborepo), RESTful APIs, GraphQL, Redis, Kafka, RabbitMQ, Zero-Downtime Releases",
        "Cloud & CI/CD": "Docker, Kubernetes, AWS (ECS, Lambda, S3, RDS, CloudFront), GCP, GitHub Actions, Bitbucket Pipelines, Jenkins, Vercel, Infrastructure as Code (IaC), Sentry, Axiom",
        "Engineering Governance": "Architectural Roadmapping, Technical Mentorship, Code Review Standards, Automated Quality Gates, Cross-Functional Alignment, AI-Assisted Development Standards",
        "Frontend": "React, Next.js (App & Pages Router), TypeScript, JavaScript, React Native, Expo, Tailwind CSS, State Management (Zustand, Redux)",
        "Backend & Data": "Node.js, Express.js, NestJS, GoLang, PostgreSQL, MongoDB, Supabase, TypeORM, Prisma, Mongoose, Data Modeling, Query Tuning",
        "Testing & QA": "Playwright, Jest, React Testing Library, End-to-End Automated Testing, Static Analysis, CI Regression Suites",
        "CMS & Tools": "Sanity, WordPress, Supabase, Axiom, Sentry, Boulevard, Webflow"
    },
    "straw_hat_platform_bullets": [
        ("Platform Architecture & Delivery: ", "Conceived, architected, and built strawhatdevs.com from scratch as a comprehensive developer learning platform offering structured courses in full-stack engineering, system design, and cloud infrastructure."),
        ("Cloud Architecture & DevOps: ", "Engineered the platform with Next.js (App Router), Supabase, PostgreSQL, Sanity CMS, and Axiom, managed with fully automated CI/CD pipelines deploying to serverless cloud infrastructure."),
        ("Curriculum & Mentorship at Scale: ", "Authored comprehensive project-based system design and cloud architecture curricula, mentoring hundreds of software developers in building enterprise-grade applications.")
    ],
    "straw_hat_youtube_bullets": [
        ("Architectural Thought Leadership: ", "Produced and published 200+ deep-dive videos on software architecture, microservices, cloud deployments, full-stack frameworks, and distributed system design patterns."),
        ("Global Community Impact: ", "Built an engaged global audience of developers through code-first architectural breakdowns of real-world enterprise systems.")
    ],
    "experiences": [
        {
            "role": "Staff Engineer / Technical Lead",
            "company": "Tuna Software Solution",
            "date": "Jun 2025 – Present",
            "bullets": [
                ("Architectural Ownership: ", "Owned the technical roadmap, architectural blueprints, and full delivery lifecycle across frontend, backend microservices, and DevOps for two concurrent enterprise software platforms."),
                ("Engineering Velocity & CI/CD Governance: ", "Instituted modern engineering standards including automated CI/CD deployment pipelines (GitHub Actions/Docker), branch protection rules, automated test runners, and structured code reviews—boosting team velocity by 35% and reducing defects."),
                ("Platform Modernization & Web Migration: ", "Led the architectural transformation of an Android kiosk platform into a cloud-native web application, streamlining multi-device deployments and eliminating hardware operational friction."),
                ("Cross-Functional Leadership: ", "Served as principal technical liaison between engineering teams and executive product stakeholders, translating business goals into robust technical specifications and de-risking dependencies.")
            ]
        },
        {
            "role": "Senior Software Engineer & Lead Architect",
            "company": "Heady",
            "date": "May 2023 – Jun 2025",
            "bullets": [
                ("Enterprise Client Technical Leadership: ", "Served as lead architect and client-facing technical lead for tier-1 enterprise partnerships (HeyDay Skin Care, Chuck E. Cheese, The Lash Lounge), translating high-level business goals into scalable, production-grade applications."),
                ("Rapid MVP Delivery (75% Timeline Compression): ", "Architected and delivered the Chuck E. Cheese customer experience web platform MVP in 2 months as a sole contributor, compressing an estimated 8-month delivery timeline by 75% through modular system design and rapid iteration."),
                ("High-Availability Microservices Ecosystem: ", "Architected a microservices-based booking and eCommerce infrastructure for HeyDay Skin Care integrated with Boulevard API, Redis caching, and resilient AWS cloud services, sustaining 99.9% uptime during peak holiday booking surges."),
                ("DevOps & Pipeline Automation: ", "Engineered high-conversion booking engines and custom CRM subscription systems (The Lash Lounge & Vmeals) backed by automated Bitbucket/GitHub Actions deployment pipelines to AWS and Vercel."),
                ("Technical Mentorship & AI-Assisted Workflows: ", "Mentored 6+ junior and mid-level engineers in distributed system design, clean architecture, and modern AI-assisted engineering workflows, significantly improving code quality and team autonomy.")
            ]
        },
        {
            "role": "Principal Software Engineer",
            "company": "HOD Digitals",
            "date": "Nov 2022 – May 2023",
            "bullets": [
                ("Team Leadership & System Scalability: ", "Provided technical direction and architectural governance to a distributed team of engineers across multiple concurrent client product lines."),
                ("Decoupled Architecture: ", "Architected decoupled microservices, relational database schemas (PostgreSQL), and cloud deployment blueprints that lowered cloud infrastructure expenditures while improving horizontal scalability under traffic spikes."),
                ("Stakeholder Governance: ", "Acted as primary technical bridge between engineering pods and executive business stakeholders, managing sprint commitments, mitigating delivery bottlenecks, and driving engineering predictability.")
            ]
        },
        {
            "role": "Software Engineer",
            "company": "HOD Digitals",
            "date": "Sep 2021 – Nov 2022",
            "bullets": [
                ("Full-Lifecycle Engineering: ", "Delivered scalable web applications from requirements analysis through deployment across multiple client projects."),
                ("API & UI Development: ", "Engineered high-throughput RESTful APIs and modern reactive user interfaces based on direct client feedback, accelerating feature velocity and user adoption."),
                ("CI/CD Integration: ", "Integrated automated testing suites into deployment pipelines, ensuring zero-downtime releases and high code coverage.")
            ]
        },
        {
            "role": "Full Stack Developer",
            "company": "Skematix Technologies",
            "date": "Dec 2019 – Sep 2021",
            "bullets": [
                ("Full-Stack Application Delivery: ", "Developed and maintained multiple responsive client web applications utilizing React, Node.js, Express, and databases, with an emphasis on seamless frontend-backend integration."),
                ("End-to-End Ownership: ", "Managed complete software development lifecycle from initial data schema design to UI execution, performance tuning, and server provisioning.")
            ]
        }
    ],
    "projects_header": "FEATURED ARCHITECTURAL HIGHLIGHTS & SYSTEMS",
    "projects": [
        {
            "title": "Straw Hat Devs Developer Platform",
            "role_or_domain": "strawhatdevs.com",
            "date": "Mar 2025 – Present",
            "bullets": [
                "Complete end-to-end edtech platform featuring interactive courses, user authentication, streaming integrations, and observability.",
                "Tech Stack: Next.js (App Router), TypeScript, PostgreSQL, Supabase, Sanity, Axiom, Vercel CI/CD"
            ]
        },
        {
            "title": "Pony Up — Self-Service IoT Retail Kiosk",
            "role_or_domain": "Staff Engineer",
            "date": "Jun 2025 – Present",
            "bullets": [
                "Hardware-integrated Android kiosk system handling point-of-sale payments, barcode/QR scanning, cash drawers, and thermal receipt printing.",
                "Tech Stack: React Native, Expo, Android Kiosk Mode, IoT Peripherals, REST APIs"
            ]
        },
        {
            "title": "Chuck E. Cheese — Web Experience Platform",
            "role_or_domain": "Senior Engineer",
            "date": "Jun 2025 – Present",
            "bullets": [
                "Rapid-turnaround customer interaction web app featuring frictionless QR code scanning, dynamic menus, and responsiveness. Built in 2 months.",
                "Tech Stack: Next.js, React, Tailwind CSS, CloudFront/AWS, Automated CI/CD"
            ]
        },
        {
            "title": "HeyDay Skin Care — Booking & Shop Engine",
            "role_or_domain": "Lead Frontend Developer",
            "date": "May 2023 – Present",
            "bullets": [
                "High-concurrency booking widget and eCommerce app integrating third-party APIs (Boulevard), real-time inventory, and Redis caching.",
                "Tech Stack: React, Node.js, NestJS, Redis, AWS ECS/Lambda, Sentry, Boulevard API"
            ]
        },
        {
            "title": "Vmeals — CRM & Subscription Management Platform",
            "role_or_domain": "Team Lead",
            "date": "Jan 2023 – Feb 2023",
            "bullets": [
                "Automated meal subscription workflow system with custom plan builders, billing cycles, and automated customer notifications.",
                "Tech Stack: Next.js, Node.js, Express, PostgreSQL, AWS, GitHub Actions"
            ]
        },
        {
            "title": "AMS Test Automation & Release Platform",
            "role_or_domain": "Lead Developer",
            "date": "Jun 2022 – May 2023",
            "bullets": [
                "End-to-end regression automation suite integrating with Bitbucket Pipelines to validate pre-release mobile and web builds.",
                "Tech Stack: Playwright, TypeScript, Bitbucket Pipelines, Docker"
            ]
        }
    ]
}

# Configuration 2: Forward Deployed Engineer
fde_config = {
    "headline": "Forward Deployed Software Engineer | Enterprise Solutions Architect",
    "summary": (
        "High-agency Forward Deployed Software Engineer and Enterprise Solutions Architect with 7+ years of experience "
        "embedding with enterprise clients to scope, architect, and deploy production software systems at 4x industry speed. "
        "Hands-on polymath with end-to-end mastery across the entire technical stack—spanning client-facing applications, "
        "backend microservices, cloud infrastructure provisioning (AWS/GCP/Docker/K8s), complex hardware/IoT integrations, "
        "and automated CI/CD pipelines. Proven track record of turning highly ambiguous business requirements into mission-critical, "
        "production-ready systems under aggressive deadlines—including delivering an enterprise customer-experience platform MVP "
        "in 2 months as a sole contributor (compressing an 8-month baseline by 75%). Battle-tested in customer-facing technical delivery "
        "for enterprise clients (HeyDay Skin Care, Chuck E. Cheese, The Lash Lounge), solving edge cases in the field, integrating legacy "
        "and third-party APIs, and maintaining 99.9% uptime. Founder of Straw Hat Devs (strawhatdevs.com) and technical educator with "
        "200+ architecture and programming videos, exemplifying autonomous 0-to-1 execution and clear technical communication across "
        "technical and executive audiences."
    ),
    "competencies": [
        "Customer-Facing Technical Delivery", "Rapid MVP & 0-to-1 Execution", "End-to-End Enterprise Full-Stack",
        "IoT & Hardware Peripherals (POS/Printers)", "Third-Party & Legacy API Integration", "Cloud Infra Provisioning (AWS/GCP)",
        "Automated CI/CD & Field Deployments", "Field Troubleshooting & On-Call Triage", "Stakeholder Discovery & Scoping",
        "Microservices & Distributed Systems", "High Availability & 99.9% Uptime", "AI-Assisted Development Velocity"
    ],
    "skills": {
        "Customer Delivery": "Enterprise Client Embedding, Requirements Discovery, Rapid MVP & 0-to-1 Execution, Technical Stakeholder Alignment, Production Troubleshooting & On-Call Triage",
        "Full-Stack Architecture": "React, Next.js (App & Pages Router), TypeScript, JavaScript, Node.js, NestJS, Go (Golang), Express.js, React Native, Expo, Tailwind CSS, Responsive Web",
        "Cloud & CI/CD": "AWS (ECS, Lambda, S3, RDS, CloudFront), GCP, Docker, Kubernetes, GitHub Actions, Bitbucket Pipelines, Jenkins, Vercel, Infrastructure as Code, Sentry, Axiom",
        "Integrations & Hardware": "IoT Peripheral Integration (Payment Terminals, Thermal Printers, Cash Drawers), Android Kiosk Mode, Third-Party APIs (Boulevard, Stripe), QR/Barcode Flows",
        "Distributed Systems": "Microservices, Event-Driven Architecture, PostgreSQL, MongoDB, Supabase, Redis (Caching/Rate Limiting), Kafka, RabbitMQ, Zero-Downtime Releases",
        "Testing & Automation": "Playwright, Jest, React Testing Library, End-to-End Regression Automation, CI/CD Deployment Gates",
        "CMS & Tools": "Sanity, WordPress, Supabase, Axiom, Sentry, Boulevard, Webflow"
    },
    "straw_hat_platform_bullets": [
        ("Autonomous 0-to-1 Execution: ", "Independently conceived, architected, and deployed strawhatdevs.com, an interactive learning platform offering full-stack engineering, system design, and cloud infrastructure courses."),
        ("End-to-End Stack & Cloud Infra: ", "Engineered the entire platform using Next.js (App Router), Supabase, PostgreSQL, Sanity CMS, and Axiom with fully automated CI/CD pipelines deploying to serverless cloud infrastructure."),
        ("Technical Communication & Enablement: ", "Designed real-world project curricula and interactive workflows, empowering developers worldwide to build production-grade, enterprise-scale software.")
    ],
    "straw_hat_youtube_bullets": [
        ("Clear Technical Communication: ", "Authored and published 200+ in-depth videos breaking down complex distributed systems, backend architectures, cloud configurations, and modern full-stack frameworks."),
        ("Community Authority: ", "Built an active international developer audience by demonstrating how to translate complex system architectures into clean, pragmatic, production-ready code.")
    ],
    "experiences": [
        {
            "role": "Staff Engineer / Forward Deployed Solutions Lead",
            "company": "Tuna Software Solution",
            "date": "Jun 2025 – Present",
            "bullets": [
                ("Enterprise Client Delivery: ", "Owned end-to-end technical delivery across frontend, backend, cloud infrastructure, and DevOps for two concurrent high-priority enterprise client engagements."),
                ("Pony Up IoT Retail Kiosk (Field Deployment): ", "Architected, built, and deployed an Android self-service kiosk using React Native/Expo integrated with physical IoT hardware peripherals (credit card payment terminals, thermal receipt printers, cash drawers) operating in high-traffic shopping malls."),
                ("Zero-Downtime Hardware & Web Migration: ", "Spearheaded the strategic rebuild and migration of the retail kiosk architecture into a cross-platform, cloud-hosted web system, eliminating physical on-site deployment bottlenecks and ensuring multi-device operational stability."),
                ("Continuous Deployment & Infra Setup: ", "Built automated CI/CD pipelines (GitHub Actions/Docker) and observability dashboards, accelerating feature rollout cycles and enabling instant, zero-downtime remote updates to kiosks in the field."),
                ("Client Discovery & Alignment: ", "Collaborated directly with client-side business and operational stakeholders to translate on-ground operational hurdles into robust, automated software workflows.")
            ]
        },
        {
            "role": "Senior Forward Deployed Engineer & Client Lead",
            "company": "Heady",
            "date": "May 2023 – Jun 2025",
            "bullets": [
                ("Enterprise Account Ownership: ", "Served as the lead customer-facing software architect for major enterprise clients (HeyDay Skin Care, Chuck E. Cheese, The Lash Lounge), translating ambiguous client requirements into high-impact, production-ready systems."),
                ("Chuck E. Cheese Web Platform (75% Faster Time-to-Market): ", "Sole engineer responsible for architecting and deploying an interactive customer experience web platform featuring QR code scanning. Delivered the production MVP in 2 months versus an initial 8-month client estimate, reducing time-to-market by 75%."),
                ("HeyDay Skin Care Microservices & Enterprise Integration: ", "Architected a robust microservices layer integrating the client's eCommerce and booking workflows with the Boulevard API, Redis caching, and resilient AWS infrastructure, sustaining 99.9% uptime during nationwide holiday booking spikes."),
                ("Rapid Client Rollouts: ", "Engineered high-converting booking workflows and automated subscription CRM engines for The Lash Lounge and Vmeals, deploying via automated Bitbucket/GitHub Actions pipelines to AWS and Vercel."),
                ("Accelerating Delivery via AI-Assisted Workflows: ", "Leveraged and coached team members on AI-assisted coding and testing workflows to accelerate client delivery cycles and unblock complex edge cases in record time.")
            ]
        },
        {
            "role": "Principal Software Engineer / Solutions Architect",
            "company": "HOD Digitals",
            "date": "Nov 2022 – May 2023",
            "bullets": [
                ("Client Solution Architecture: ", "Led solution design and execution across multiple concurrent client product engagements, defining system architecture and technology stacks tailored to each client's business goals."),
                ("Infrastructure Cost Optimization & Scalability: ", "Redesigned client backend services and database architectures (PostgreSQL/MongoDB), lowering cloud hosting costs while hardening systems against sudden traffic surges."),
                ("Technical Liaison: ", "Acted as the primary technical point of contact between enterprise client stakeholders and engineering pods, resolving technical ambiguities, mitigating implementation risks, and ensuring on-time delivery.")
            ]
        },
        {
            "role": "Software Engineer",
            "company": "HOD Digitals",
            "date": "Sep 2021 – Nov 2022",
            "bullets": [
                ("Rapid Client Feature Engineering: ", "Developed full-stack web applications across multiple concurrent client accounts from initial scoping to cloud deployment."),
                ("Direct Feedback Iteration: ", "Partnered with product managers and clients to rapidly iterate on feedback, releasing performant RESTful APIs, modern reactive user interfaces, and automated data processing pipelines."),
                ("CI/CD Quality Control: ", "Implemented automated testing routines within deployment pipelines, ensuring zero-regression releases to client production environments.")
            ]
        },
        {
            "role": "Full Stack Developer",
            "company": "Skematix Technologies",
            "date": "Dec 2019 – Sep 2021",
            "bullets": [
                ("Client Web Application Delivery: ", "Developed, optimized, and deployed multiple client-facing responsive web applications with React, Node.js, Express, and databases."),
                ("Full Lifecycle Ownership: ", "Managed full application lifecycle from initial data modeling and frontend implementation to server provisioning and deployment.")
            ]
        }
    ],
    "projects_header": "FEATURED CLIENT DEPLOYMENTS & SYSTEMS",
    "projects": [
        {
            "title": "Chuck E. Cheese — Customer Experience Platform",
            "role_or_domain": "Solo Lead Engineer",
            "date": "Jun 2025 – Present",
            "bullets": [
                "Rapid-turnaround web platform featuring seamless QR-code interaction and mobile menus. Architected, coded, and deployed end-to-end in 2 months.",
                "Tech Stack: Next.js, React, Tailwind CSS, CloudFront/AWS, Automated CI/CD"
            ]
        },
        {
            "title": "Pony Up — IoT Retail Kiosk System",
            "role_or_domain": "Staff Solutions Engineer",
            "date": "Jun 2025 – Present",
            "bullets": [
                "End-to-end hardware-integrated Android kiosk system handling point-of-sale payments, barcode/QR scanning, cash drawers, and thermal receipt printing.",
                "Tech Stack: React Native, Expo, Android Kiosk Mode, IoT Peripherals, REST APIs"
            ]
        },
        {
            "title": "HeyDay Skin Care — Booking & Shop Engine",
            "role_or_domain": "Lead Enterprise Architect",
            "date": "May 2023 – Present",
            "bullets": [
                "High-concurrency booking widget and eCommerce app integrating third-party booking APIs (Boulevard), real-time inventory, and Redis caching on AWS.",
                "Tech Stack: React, Node.js, NestJS, Redis, AWS ECS/Lambda, Sentry, Boulevard API"
            ]
        },
        {
            "title": "Straw Hat Devs Developer Platform",
            "role_or_domain": "strawhatdevs.com",
            "date": "Mar 2025 – Present",
            "bullets": [
                "Complete full-stack edtech platform with real-time auth, payments/courses, video delivery, and automated CI/CD.",
                "Tech Stack: Next.js, TypeScript, PostgreSQL, Supabase, Sanity, Axiom, Vercel CI/CD"
            ]
        },
        {
            "title": "Vmeals — CRM & Subscription Management Platform",
            "role_or_domain": "Client Delivery Lead",
            "date": "Jan 2023 – Feb 2023",
            "bullets": [
                "Automated meal subscription workflow system with custom plan builders, billing cycles, and automated customer notifications.",
                "Tech Stack: Next.js, Node.js, Express, PostgreSQL, AWS, GitHub Actions"
            ]
        },
        {
            "title": "AMS Test Automation & Release Platform",
            "role_or_domain": "Lead Developer",
            "date": "Jun 2022 – May 2023",
            "bullets": [
                "End-to-end regression automation suite integrating with Bitbucket Pipelines to validate pre-release mobile and web builds prior to app store distribution.",
                "Tech Stack: Playwright, TypeScript, Bitbucket Pipelines, Docker"
            ]
        }
    ]
}

def main():
    src_template = "2026-09-22_Original_Resume_Santheepkumar_V.docx"
    generate_docx_file(src_template, "2026-09-22_Technical_Lead_Resume_Santheepkumar_V.docx", tech_lead_config)
    generate_docx_file(src_template, "2026-09-22_Forward_Deployed_Engineer_Resume_Santheepkumar_V.docx", fde_config)

if __name__ == "__main__":
    main()
