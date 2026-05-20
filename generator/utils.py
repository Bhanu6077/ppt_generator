import os
import copy
import tempfile
from pptx import Presentation
from pptx.text.text import _Paragraph
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Pt

# --- HELPER 1: Get All Text (Deep Search) ---
def get_all_text(shape):
    text = ""
    if shape.has_text_frame:
        text += "".join([p.text for p in shape.text_frame.paragraphs])
    if hasattr(shape, "shapes"): 
        for s in shape.shapes:
            text += get_all_text(s)
    return text

# --- HELPER 2: ULTRA-AGGRESSIVE TEXT REPLACE ---
def replace_text(shape, tag, replacement):
    if hasattr(shape, "shapes"):
        for s in shape.shapes: replace_text(s, tag, replacement)
    if not shape.has_text_frame: return
    for p in shape.text_frame.paragraphs:
        if tag in p.text:
            replaced_in_run = False
            for run in p.runs:
                if tag in run.text:
                    run.text = run.text.replace(tag, replacement)
                    replaced_in_run = True
            if not replaced_in_run:
                p.text = p.text.replace(tag, replacement)

# --- HELPER 3: PERFECT BULLETS ---
def replace_with_bullets(shape, tag, bullet_list):
    if hasattr(shape, "shapes"):
        for s in shape.shapes: replace_with_bullets(s, tag, bullet_list)
    if not shape.has_text_frame: return
    has_tag = False
    for p in shape.text_frame.paragraphs:
        if tag in p.text:
            has_tag = True
            break
    if not has_tag: return
    tf = shape.text_frame
    tf.vertical_anchor = MSO_ANCHOR.TOP 
    if not bullet_list:
        tf.clear()
        return
    template_p_xml = copy.deepcopy(tf.paragraphs[0]._p)
    tf.clear()
    p0 = tf.paragraphs[0]
    p0._p.getparent().replace(p0._p, template_p_xml)
    p0 = _Paragraph(template_p_xml, tf)
    if len(p0.runs) > 0:
        p0.runs[0].text = bullet_list[0]
        for i in range(len(p0.runs)-1, 0, -1): p0._p.remove(p0.runs[i]._r)
    else:
        p0.text = bullet_list[0]
    p0.space_before, p0.space_after = Pt(0), Pt(0)
    parent = p0._p.getparent()
    for point in bullet_list[1:]:
        new_p_xml = copy.deepcopy(template_p_xml)
        parent.append(new_p_xml)
        new_p = _Paragraph(new_p_xml, tf)
        if len(new_p.runs) > 0:
            new_p.runs[0].text = point
            for i in range(len(new_p.runs)-1, 0, -1): new_p._p.remove(new_p.runs[i]._r)
        else:
            new_p.text = point
        new_p.space_before, new_p.space_after = Pt(0), Pt(0)

# --- HELPER 4: Image Replacement (Clean & Simple) ---
def replace_image(slide, shape, img_file):
    x, y, cx, cy = shape.left, shape.top, shape.width, shape.height
    with tempfile.NamedTemporaryFile(delete=False) as t:
        for chunk in img_file.chunks(): t.write(chunk)
        p = t.name
    slide.shapes.add_picture(p, x, y, cx, cy)
    shape._element.getparent().remove(shape._element)
    os.unlink(p)

# ==========================================
# MAIN GENERATOR FUNCTION (Slides 1 - 17)
# ==========================================
def generate_pptx(form_data, uploaded_images, uploaded_videos, template_path, output_path):
    prs = Presentation(template_path)

    def gl(prefix): return [form_data.get(k, "").strip() for k in sorted([k for k in form_data.keys() if k.startswith(prefix)], key=lambda x: int(x.split("_")[1])) if form_data.get(k, "").strip()]
    
    ps_data = {f"[psh{i}]": form_data.get(f"psh_{i}", "").strip() for i in range(1, 5)}
    ps_data.update({f"[ds{i}]": form_data.get(f"ds_{i}", "").strip() for i in range(1, 5)})
    
    # 1. Master Dictionary for ALL Simple Text Mappings (Including Slide 17)
    st = {
        "[erp_main]": form_data.get("erp_main", "").strip(),
        "[sol_main]": form_data.get("sol_main", "").strip(),
        "[sol_used_main]": form_data.get("sol_used_main", "").strip(),
        "[pre_post_main]": form_data.get("pre_post_main", "").strip(),
        "[ai_cap_main]": form_data.get("ai_cap_main", "").strip(),
        "[ai_cap_h1]": form_data.get("ai_cap_h1", "").strip(),
        "[ai_cap_h2]": form_data.get("ai_cap_h2", "").strip(),
        "[dash_main]": form_data.get("dash_main", "").strip(),
        "[demo_main]": form_data.get("demo_main", "").strip(),
        "[heat_map_main]": form_data.get("heat_map_main", "").strip(),
        "[cloud_main]": form_data.get("cloud_main", "").strip(),
        "[cloud_sub]": form_data.get("cloud_sub", "").strip(),
        "[cloud_feat_h]": form_data.get("cloud_feat_h", "").strip(),
        "[mobile_main]": form_data.get("mobile_main", "").strip(),
        "[mobile_sub]": form_data.get("mobile_sub", "").strip(),
        "[work_main]": form_data.get("work_main", "").strip(),
        "[work_sub]": form_data.get("work_sub", "").strip(),
        "[partner_main]": form_data.get("partner_main", "").strip(),
        
        # SLIDE 17 ADDITIONS (Ye aapke pichle code mein missing the!)
        "[proj_main]": form_data.get("proj_main", "").strip(),
        "[proj_banner]": form_data.get("proj_banner", "").strip(),

        # Slide 18 Mappings
        "[us_main]": form_data.get("us_main", "").strip(),
        "[recog_main]": form_data.get("recog_main", "").strip(),
        "[recog_cap]": form_data.get("recog_cap", "").strip(),
    }

    # 2. LOOPS FOR MULTI-BLOCK SECTIONS
    for i in range(1, 5): st[f"[solh{i}]"] = form_data.get(f"solh_{i}", "").strip()
    for i in range(1, 5): st[f"[sold{i}]"] = form_data.get(f"sold_{i}", "").strip()
    for i in range(1, 9): st[f"[suh{i}]"] = form_data.get(f"suh_{i}", "").strip()
    for i in range(1, 9): st[f"[sud{i}]"] = form_data.get(f"sud_{i}", "").strip()
    for i in range(1, 8): st[f"[cloud_f{i}]"] = form_data.get(f"cloud_f{i}", "").strip()
    for i in range(1, 6): st[f"[cf{i}]"] = form_data.get(f"cf{i}", "").strip()
    for i in range(1, 8): st[f"[step{i}]"] = form_data.get(f"step_{i}", "").strip()
    for i in range(1, 5): st[f"[wh{i}]"] = form_data.get(f"work_h{i}", "").strip()
    for i in range(1, 5): st[f"[wd{i}]"] = form_data.get(f"work_d{i}", "").strip()
    
    
    # SLIDE 17 LOOPS
    for i in range(1, 5):
        st[f"[pf{i}]"] = form_data.get(f"proj_f{i}", "").strip()
        st[f"[ph{i}]"] = form_data.get(f"proj_h{i}", "").strip()
        st[f"[pd{i}]"] = form_data.get(f"proj_d{i}", "").strip()

    # Slide 18 Strategy Boxes
    for i in range(1, 7):
        st[f"[usb{i}]"] = form_data.get(f"us_b{i}", "").strip()


    # Slide 19 Awards (8 Boxes)
    for i in range(1, 9):
        st[f"[ri{i}]"] = form_data.get(f"recog_i{i}", "").strip()
        st[f"[rh{i}]"] = form_data.get(f"recog_h{i}", "").strip()
        st[f"[rd{i}]"] = form_data.get(f"recog_d{i}", "").strip()


    # 3. Bullet Mappings
    b_pts = {
        "[ai_cap_b1]": [p.strip() for p in form_data.get("ai_cap_b1", "").split('\n') if p.strip()],
        "[ai_cap_b2]": [p.strip() for p in form_data.get("ai_cap_b2", "").split('\n') if p.strip()],
        "[dash_stat_b]": [p.strip() for p in form_data.get("dash_stat_b", "").split('\n') if p.strip()],
        "[mobile_b]": [p.strip() for p in form_data.get("mobile_b", "").split('\n') if p.strip()]
    }

    # VIP Slide logic
    vip_indices = set(range(21))

    # --- EXECUTION ENGINE ---
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            txt = get_all_text(shape)
            
            # Replace Simple Text Tags
            for tag, val in st.items(): replace_text(shape, tag, val)
            for tag, val in ps_data.items(): replace_text(shape, tag, val)
            
            # Replace Bullets
            if "[who_we_are_1]" in txt: replace_with_bullets(shape, "[who_we_are_1]", gl("who_"))
            elif "[our_goals_1]" in txt: replace_with_bullets(shape, "[our_goals_1]", gl("goal_"))
            for tag, pts in b_pts.items():
                if tag in txt: replace_with_bullets(shape, tag, pts)

            # Replace Images
            if hasattr(shape, "shape_type") and shape.shape_type == 13:
                alt = shape._element._nvXxPr.cNvPr.attrib.get('descr', shape.name)
                tags = [
                    "erp_img_1", "erp_img_2", "ai_cap_img_1", "ai_cap_img_2", 
                    "dash_img_main", "demo_img_1", "heat_map_img_1", "cloud_img", 
                    "mobile_img", "work_img_1", "work_img_2", "work_img_3", "work_img_4",
                    "partner_globe_img", "proj_map_img" # <-- Map image added here!
                ]
                for t in range(1, 5): tags.append(f"pre_post_img_{t}")
                for t in tags:
                    if t in alt and uploaded_images.get(t): replace_image(slide, shape, uploaded_images[t])

    # Delete unused slides
    for j in range(len(prs.slides._sldIdLst)-1, -1, -1):
        if j not in vip_indices: prs.slides._sldIdLst.remove(prs.slides._sldIdLst[j])
        
    prs.save(output_path)