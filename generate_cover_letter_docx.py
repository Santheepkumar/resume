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

para_id_counter = 500

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

def make_run(text, bold=False, italic=False, color="141414", size="21", underline=False):
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

def make_para(runs, before=0, after=120, line_spacing=276):
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
    spc.attrib[qn("line")] = str(line_spacing)
    spc.attrib[qn("lineRule")] = "auto"

    for r in runs:
        p.append(r)
    return p

def build_cover_letter_body():
    body = ET.Element(qn("body"))

    # Name Header
    body.append(make_para([
        make_run("SANTHEEPKUMAR VEERASAMY", bold=True, color="193c66", size="33")
    ], before=0, after=40))

    # Subtitle
    body.append(make_para([
        make_run("Staff Software Engineer | Technical Lead | Solutions Architect", bold=True, color="2560a7", size="18")
    ], before=0, after=60))

    # Contacts
    body.append(make_para([
        make_run("+91 86676 62978  |  devsantheep@gmail.com  |  ", bold=True, color="434343", size="15"),
        make_run("linkedin.com/in/santheepkumar", bold=True, color="2560a7", size="15", underline=True),
        make_run("  |  ", bold=True, color="434343", size="15"),
        make_run("github.com/Santheepkumar", bold=True, color="1155cc", size="15", underline=True),
        make_run("  |  ", bold=True, color="434343", size="15"),
        make_run("strawhatdevs.com", bold=True, color="2560a7", size="15", underline=True),
        make_run("  |  ", bold=True, color="434343", size="15"),
        make_run("sk.strawhatdevs.com", bold=True, color="2560a7", size="15", underline=True)
    ], before=0, after=240))

    # Metadata
    body.append(make_para([
        make_run("Date: ", bold=True, color="141414", size="16"),
        make_run("September 22, 2026", color="434343", size="16")
    ], before=0, after=40))

    body.append(make_para([
        make_run("To: ", bold=True, color="141414", size="16"),
        make_run("Hiring Team / Engineering Leadership at ", color="434343", size="16"),
        make_run("[Company Name]", bold=True, color="141414", size="16")
    ], before=0, after=40))

    body.append(make_para([
        make_run("Re: ", bold=True, color="141414", size="16"),
        make_run("Application for ", color="434343", size="16"),
        make_run("[Role Title]", bold=True, color="141414", size="16")
    ], before=0, after=180))

    # Salutation
    body.append(make_para([
        make_run("Hi [Hiring Team / Name],", bold=True, color="141414", size="17")
    ], before=0, after=120))

    # Paragraph 1
    p1 = (
        "I’m writing to express my strong interest in joining [Company Name] as a [Role Title]. Over the past 7+ years, "
        "I’ve built software across the complete engineering lifecycle—from whiteboard system architecture and responsive frontends "
        "to distributed microservices backends, cloud infrastructure provisioning (AWS/GCP), and automated CI/CD pipelines. What excites "
        "me about this role is the opportunity to bring that end-to-end ownership to your team and help turn ambitious product goals "
        "into dependable, production-ready systems."
    )
    body.append(make_para([make_run(p1, color="141414", size="16")], before=0, after=140))

    # Paragraph 2
    p2 = (
        "Whether embedding directly with enterprise stakeholders or leading engineering initiatives, I gravitate toward solving "
        "complex technical problems with velocity and pragmatic engineering. At Heady, I delivered a customer experience web platform "
        "for Chuck E. Cheese in just 2 months as a sole contributor, compressing an estimated 8-month timeline by 75%. At Tuna Software Solution, "
        "I architected an Android retail kiosk integrating physical IoT peripherals (payment gateways, receipt printers, cash drawers), "
        "and then led its migration to a cloud-native web system with automated CI/CD pipelines to streamline deployment across devices. "
        "I don’t just write code; I care deeply about how systems are provisioned, monitored, and scaled in production."
    )
    body.append(make_para([make_run(p2, color="141414", size="16")], before=0, after=140))

    # Paragraph 3
    p3 = (
        "Beyond day-to-day engineering, I’m deeply passionate about developer enablement and system design. I founded and built "
        "Straw Hat Devs (strawhatdevs.com) from scratch as an interactive learning platform for software developers, and I have published "
        "over 200 videos on YouTube breaking down full-stack architecture, microservices, and modern cloud deployment patterns. "
        "Teaching complex systems to thousands of developers has honed my ability to communicate technical trade-offs clearly, mentor teammates, "
        "and ensure engineering teams build with clarity rather than unnecessary complexity."
    )
    body.append(make_para([make_run(p3, color="141414", size="16")], before=0, after=140))

    # Paragraph 4
    p4 = (
        "I admire the work [Company Name] is doing, and I’d love the chance to discuss how my hands-on full-stack skills, infrastructure "
        "mindset, and bias for action can help your team ship faster and build systems that scale."
    )
    body.append(make_para([make_run(p4, color="141414", size="16")], before=0, after=160))

    # Closing
    body.append(make_para([
        make_run("Thank you for your time and consideration.", color="141414", size="16")
    ], before=0, after=80))

    body.append(make_para([
        make_run("Warm regards,", color="141414", size="16")
    ], before=0, after=80))

    body.append(make_para([
        make_run("Santheepkumar Veerasamy", bold=True, color="193c66", size="18")
    ], before=0, after=30))

    body.append(make_para([
        make_run("Staff Software Engineer | Full-Stack & Cloud Solutions Architect", bold=True, color="2560a7", size="15")
    ], before=0, after=30))

    body.append(make_para([
        make_run("+91 86676 62978  •  devsantheep@gmail.com  •  strawhatdevs.com", color="434343", size="14")
    ], before=0, after=0))

    # Page layout
    sectPr = ET.SubElement(body, qn("sectPr"))
    pgSz = ET.SubElement(sectPr, qn("pgSz"))
    pgSz.attrib[qn("h")] = "15840"
    pgSz.attrib[qn("w")] = "12240"
    pgSz.attrib[qn("orient")] = "portrait"

    pgMar = ET.SubElement(sectPr, qn("pgMar"))
    pgMar.attrib[qn("bottom")] = "1440"
    pgMar.attrib[qn("top")] = "1440"
    pgMar.attrib[qn("left")] = "1440"
    pgMar.attrib[qn("right")] = "1440"
    pgMar.attrib[qn("header")] = "720"
    pgMar.attrib[qn("footer")] = "720"

    pgNumType = ET.SubElement(sectPr, qn("pgNumType"))
    pgNumType.attrib[qn("start")] = "1"

    return body

def main():
    src_template = "2026-09-22_Original_Resume_Santheepkumar_V.docx"
    dest_filename = "2026-09-22_Common_Cover_Letter_Santheepkumar_V.docx"

    with zipfile.ZipFile(src_template, "r") as z_in:
        orig_doc_xml = z_in.read("word/document.xml")
        orig_root = ET.fromstring(orig_doc_xml)

        orig_body = orig_root.find(qn("body"))
        if orig_body is not None:
            orig_root.remove(orig_body)
        
        new_body = build_cover_letter_body()
        orig_root.append(new_body)

        new_doc_xml = ET.tostring(orig_root, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(dest_filename, "w", compression=zipfile.ZIP_DEFLATED) as z_out:
            for item in z_in.infolist():
                if item.filename == "word/document.xml":
                    z_out.writestr(item.filename, new_doc_xml)
                else:
                    z_out.writestr(item.filename, z_in.read(item.filename))
    print(f"Generated {dest_filename} successfully.")

if __name__ == "__main__":
    main()
