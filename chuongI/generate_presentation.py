r"""
Script: generate_presentation.py
Description: Minimalist, Flat, Card-Free Academic Presentation Generator for Chapter I.
Faithfully mirrors the exact design style of Chuong-II-Cortex-M85:
  - Uses figs/bg.jpg (Cyan banner + PTIT logo at top-right + pure white body)
  - No card layouts, no rectangles/boxes around content, no dashboard styling
  - Text and bullets directly on white background
  - Two columns separated solely by position
  - Overleaf LaTeX font (Latin Modern / \usepackage{lmodern})
Generates:
  1. Chuong_I_Overview_Embedded_Systems.pptx
  2. Chuong_I_Beamer.zip (with main.tex, figs/bg.jpg, and all images)
"""

import os
import sys
import shutil
import zipfile
import pypdf
from PIL import Image

import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Set stdout encoding for Windows
sys.stdout.reconfigure(encoding='utf-8')

# ==============================================================================
# 1. PREPARE ASSETS (IMAGES & BACKGROUND)
# ==============================================================================
def prepare_assets():
    print("[1/4] Preparing image assets and background...")
    os.makedirs("images", exist_ok=True)
    os.makedirs("figs", exist_ok=True)

    # 1. Extract bg.jpg from Chuong_II_Beamer.zip if needed
    if not os.path.exists("figs/bg.jpg") and os.path.exists("Chuong_II_Beamer.zip"):
        with zipfile.ZipFile("Chuong_II_Beamer.zip", 'r') as z:
            if "figs/bg.jpg" in z.namelist():
                with open("figs/bg.jpg", "wb") as f:
                    f.write(z.read("figs/bg.jpg"))
                print("  Extracted figs/bg.jpg from Chuong_II_Beamer.zip")

    # 2. Extract from user image zip
    user_zips = [f for f in os.listdir(".") if f.endswith(".zip") and "beamer" not in f.lower() and "chuong_ii" not in f.lower()]
    if user_zips:
        u_zip = user_zips[0]
        print(f"  Extracting images from user zip: {u_zip}")
        with zipfile.ZipFile(u_zip, 'r') as z:
            for n in z.namelist():
                if not n.endswith('/'):
                    base = os.path.basename(n)
                    with open(os.path.join("images", base), "wb") as f:
                        f.write(z.read(n))

    # Standard clean filename mapping
    mapping = {
        'chương 1.jpg': 'fig_cover.jpg',
        'hình 1.1.1 a typical embedded system.jpg': 'fig_1_1_typical_system.jpg',
        'hình 1.1.2 hardware of embedded system.jpg': 'fig_1_1_2_hardware.jpg',
        'hình 1.1.3 edge computing.png': 'fig_1_1_3_edge.png',
        'hình 1.2 charactareistics of embedded system.jpg': 'fig_1_2_characteristics.jpg',
        'hình 1.3 general computing systems vs embedded systems.jpg': 'fig_1_3_gpc_vs_embedded.jpg',
        'hình 1.4 Embedded Software Layered Architecture.jpg': 'fig_1_4_layered_arch.jpg',
    }

    # Verify or extract from PDF if missing
    pdf_source = "Chuong_I_overview_2.pdf" if os.path.exists("Chuong_I_overview_2.pdf") else "Chuong_I_overview.pdf"
    if os.path.exists(pdf_source):
        reader = pypdf.PdfReader(pdf_source)
        pdf_img_map = {
            1: 'fig_cover.jpg',
            2: 'fig_1_1_typical_system.jpg',
            3: 'fig_1_1_2_hardware.jpg',
            5: 'fig_1_1_3_edge.png',
            6: 'fig_1_2_characteristics.jpg',
            9: 'fig_1_3_gpc_vs_embedded.jpg',
            11: 'fig_1_4_layered_arch.jpg'
        }
        for page_idx, target_name in pdf_img_map.items():
            target_path = os.path.join("images", target_name)
            if not os.path.exists(target_path) and page_idx - 1 < len(reader.pages):
                page = reader.pages[page_idx - 1]
                for img in page.images:
                    if img.name not in ['X7.png', 'X8.png']:
                        with open(target_path, "wb") as f:
                            f.write(img.data)
                        break

    # Sync to both images/ and figs/
    for src_name, dst_name in mapping.items():
        src_path = os.path.join("images", src_name)
        dst_path = os.path.join("images", dst_name)
        if os.path.exists(src_path) and not os.path.exists(dst_path):
            shutil.copy(src_path, dst_path)
        if os.path.exists(dst_path):
            shutil.copy(dst_path, os.path.join("figs", dst_name))

    print("  Verified assets ready in images/ and figs/.")


# ==============================================================================
# 2. BUILD MINIMALIST POWERPOINT PRESENTATION (.PPTX)
# ==============================================================================
PTIT_BLUE = RGBColor(0, 169, 206)      # Exact PTIT Cyan #00A9CE
TEXT_BLACK = RGBColor(20, 20, 20)      # High contrast clean text
WHITE = RGBColor(255, 255, 255)
GRAY_TEXT = RGBColor(120, 120, 120)

FONT_LATEX = "Latin Modern Sans"       # Overleaf font

def create_base_slide(prs):
    """Creates a blank slide with figs/bg.jpg as background (Cyan bar + PTIT logo)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    if os.path.exists("figs/bg.jpg"):
        bg = slide.shapes.add_picture("figs/bg.jpg", Inches(0), Inches(0), prs.slide_width, prs.slide_height)
        # Send to back
        slide.shapes._spTree.remove(bg._element)
        slide.shapes._spTree.insert(2, bg._element)
    return slide

def add_pptx_title(slide, title_text):
    """Adds title text directly onto the cyan banner."""
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.12), Inches(10.2), Inches(0.75))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = title_text.upper()
    p.font.name = FONT_LATEX
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.LEFT

def add_pptx_footer(slide, slide_num):
    """Adds ONLY the page number in the bottom right corner."""
    tb = slide.shapes.add_textbox(Inches(12.0), Inches(7.05), Inches(0.9), Inches(0.35))
    tf = tb.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = str(slide_num)
    p.font.name = FONT_LATEX
    p.font.size = Pt(11)
    p.font.color.rgb = GRAY_TEXT
    p.alignment = PP_ALIGN.RIGHT

def add_pptx_bullet(tf, keyword, description, is_first=False, pt_size=13, space_after=12):
    """Adds a clean bullet point directly on white background."""
    p = tf.paragraphs[0] if is_first else tf.add_paragraph()
    p.font.name = FONT_LATEX
    p.font.size = Pt(pt_size)
    p.space_after = Pt(space_after)
    p.alignment = PP_ALIGN.LEFT

    # Bullet symbol
    r_b = p.add_run()
    r_b.text = "▶  "
    r_b.font.name = FONT_LATEX
    r_b.font.bold = True
    r_b.font.color.rgb = PTIT_BLUE
    r_b.font.size = Pt(pt_size - 1.5)

    # Keyword
    if keyword:
        r_k = p.add_run()
        r_k.text = keyword + ": "
        r_k.font.name = FONT_LATEX
        r_k.font.bold = True
        r_k.font.color.rgb = TEXT_BLACK

    # Description
    r_d = p.add_run()
    r_d.text = description
    r_d.font.name = FONT_LATEX
    r_d.font.bold = False
    r_d.font.color.rgb = TEXT_BLACK

def build_pptx():
    print("[2/4] Building Minimalist PowerPoint Presentation (.pptx)...")
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # --------------------------------------------------------------------------
    # SLIDE 1: TITLE SLIDE (MINIMALIST, NO BOXES)
    # --------------------------------------------------------------------------
    s1 = create_base_slide(prs)
    add_pptx_title(s1, "CHAPTER I. OVERVIEW OF EMBEDDED SYSTEMS")
    add_pptx_footer(s1, 1)

    # Left text directly on white background
    tb1 = s1.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(7.5), Inches(5.2))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = tf1.margin_top = tf1.margin_right = tf1.margin_bottom = 0

    p_sub = tf1.paragraphs[0]
    p_sub.text = "Foundations, Core Characteristics & Layered Architecture"
    p_sub.font.name = FONT_LATEX
    p_sub.font.size = Pt(17)
    p_sub.font.bold = True
    p_sub.font.color.rgb = TEXT_BLACK
    p_sub.space_after = Pt(20)

    s1_items = [
        ("Curriculum Scope", "Core Embedded Systems & Edge Computing Architecture"),
        ("Engineering Paradigms", "Deep hardware-software co-design, deterministic timing & tight resource budgets"),
        ("Target Platform", "32-bit Microcontroller Architecture (ARM Cortex-M85 / Renesas RA8M1)"),
        ("Implementation Standards", "C99/POSIX firmware patterns, Hard RTOS, standardized MCAL/HAL stacks")
    ]
    for kw, desc in s1_items:
        add_pptx_bullet(tf1, kw, desc, is_first=False, pt_size=13, space_after=14)

    # Right image directly on slide
    if os.path.exists("figs/fig_cover.jpg"):
        s1.shapes.add_picture("figs/fig_cover.jpg", Inches(8.8), Inches(1.6), Inches(3.6), Inches(4.8))

    # --------------------------------------------------------------------------
    # SLIDE 2: AGENDA / TABLE OF CONTENTS (MATCHING CHAPTER II SLIDE 1)
    # --------------------------------------------------------------------------
    s2 = create_base_slide(prs)
    add_pptx_title(s2, "TABLE OF CONTENTS")
    add_pptx_footer(s2, 2)

    tb2 = s2.shapes.add_textbox(Inches(0.8), Inches(1.35), Inches(11.7), Inches(5.5))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    tf2.margin_left = tf2.margin_top = tf2.margin_right = tf2.margin_bottom = 0

    agenda_sections = [
        ("1. Concept & Nature of Embedded Systems", [
            "1.1. Standard Definition, 3-Element Triad, and Physical Embedded Nature",
            "1.2. Role of Embedded Systems in IoT & Edge Computing"
        ]),
        ("2. Core Technical Characteristics", [
            "2.1. Dedicated Functionality & Severe Resource Constraints",
            "2.2. Deterministic Real-Time Execution & High Reliability (WDT, BOD, CRC)",
            "2.3. Ultra-Low Power, Embedded Security & Firmware Maintenance (OTA)"
        ]),
        ("3. Embedded Systems vs. General-Purpose Computers", [
            "3.1. Architectural Paradigms: Specialization vs. Versatility",
            "3.2. Detailed Multi-Criteria Technical Comparison Matrix"
        ]),
        ("4. Layered Embedded Software Architecture", [
            "4.1. The 5-Layer Stack Model: Hardware -> MCAL -> HAL -> RTOS -> Middleware -> Application",
            "4.2. RTOS Layer: Multitasking vs. Super-Loop & Inter-Process Communication",
            "4.3. Middleware & Application Layer: Decoupling and Host Testability"
        ])
    ]

    first_sec = True
    for sec_title, sub_items in agenda_sections:
        p_sec = tf2.paragraphs[0] if first_sec else tf2.add_paragraph()
        first_sec = False
        p_sec.text = sec_title
        p_sec.font.name = FONT_LATEX
        p_sec.font.size = Pt(14.5)
        p_sec.font.bold = True
        p_sec.font.color.rgb = TEXT_BLACK
        p_sec.space_after = Pt(2)

        for sub in sub_items:
            p_sub = tf2.add_paragraph()
            p_sub.text = "        " + sub
            p_sub.font.name = FONT_LATEX
            p_sub.font.size = Pt(11.5)
            p_sub.font.color.rgb = GRAY_TEXT
            p_sub.space_after = Pt(2)

        p_sec.space_before = Pt(8)

    # Guiding Question at bottom
    p_q = tf2.add_paragraph()
    p_q.space_before = Pt(14)
    r_arr = p_q.add_run()
    r_arr.text = "⇒  "
    r_arr.font.bold = True
    r_arr.font.color.rgb = PTIT_BLUE
    r_qt = p_q.add_run()
    r_qt.text = "Guiding Question: What fundamentally distinguishes embedded software from desktop/PC software?"
    r_qt.font.name = FONT_LATEX
    r_qt.font.bold = True
    r_qt.font.color.rgb = TEXT_BLACK
    r_qt.font.size = Pt(12.5)

    # Helper for standard 2-column slide with image on left, bullets on right
    def add_flat_two_column_slide(slide_num, title, img_path, caption, bullets, takeaway=None):
        slide = create_base_slide(prs)
        add_pptx_title(slide, title)
        add_pptx_footer(slide, slide_num)

        # Left Column: Image placed directly on slide
        img_w = Inches(5.6)
        if os.path.exists(img_path):
            with Image.open(img_path) as im:
                aspect = im.width / im.height
            max_h = 4.8
            calc_h = 5.6 / aspect
            if calc_h > max_h:
                calc_h = max_h
                calc_w = calc_h * aspect
                left_x = Inches(0.7) + (img_w - Inches(calc_w)) / 2
            else:
                calc_w = 5.6
                left_x = Inches(0.7)

            slide.shapes.add_picture(img_path, left_x, Inches(1.4) + (Inches(max_h) - Inches(calc_h)) / 2, Inches(calc_w), Inches(calc_h))

            # Caption below image
            tb_cap = slide.shapes.add_textbox(Inches(0.7), Inches(6.35), img_w, Inches(0.45))
            tf_cap = tb_cap.text_frame
            tf_cap.word_wrap = True
            p_c = tf_cap.paragraphs[0]
            p_c.text = caption
            p_c.font.name = FONT_LATEX
            p_c.font.size = Pt(10)
            p_c.font.italic = True
            p_c.font.color.rgb = GRAY_TEXT
            p_c.alignment = PP_ALIGN.CENTER

        # Right Column: Bullets placed directly on white background
        tb_r = slide.shapes.add_textbox(Inches(6.6), Inches(1.35), Inches(6.0), Inches(5.5))
        tf_r = tb_r.text_frame
        tf_r.word_wrap = True
        tf_r.margin_left = tf_r.margin_top = tf_r.margin_right = tf_r.margin_bottom = 0

        first = True
        for kw, desc in bullets:
            add_pptx_bullet(tf_r, kw, desc, is_first=first, pt_size=12.5, space_after=11)
            first = False

        if takeaway:
            p_t = tf_r.add_paragraph()
            p_t.space_before = Pt(8)
            r_tarr = p_t.add_run()
            r_tarr.text = "⇒  "
            r_tarr.font.bold = True
            r_tarr.font.color.rgb = PTIT_BLUE
            r_txt = p_t.add_run()
            r_txt.text = takeaway
            r_txt.font.name = FONT_LATEX
            r_txt.font.bold = True
            r_txt.font.color.rgb = TEXT_BLACK
            r_txt.font.size = Pt(11.5)

    # --------------------------------------------------------------------------
    # SLIDE 3: 1.1 DEFINITION & TRIAD NATURE
    # --------------------------------------------------------------------------
    add_flat_two_column_slide(
        3, "1.1. EMBEDDED SYSTEM DEFINITION & NATURE",
        "figs/fig_1_1_typical_system.jpg",
        "Figure 1.1: Architecture of a Typical Embedded System.",
        [
            ("Definition", "Specialized computer system integrated within a host device to perform dedicated, pre-defined functions."),
            ("Triad Nature", "Inseparable co-design between Computation (Hardware) — Processing Logic (Firmware) — Real-world Interaction (Sensors/Actuators)."),
            ("Design Goal", "Fulfill dedicated tasks under strict constraints: deterministic latency, ultra-low power, compact size, high MTBF, and low unit cost."),
            ("Distinction", "Not a general-purpose PC — strictly eliminates software bloat in favor of deep optimization for a single mission."),
            ("Real Example", "Smart Environmental Edge Node: Continuous multi-sensor sampling, local wildfire/anomaly detection, and 24/7 telemetry.")
        ]
    )

    # --------------------------------------------------------------------------
    # SLIDE 4: 1.1 HARDWARE ARCHITECTURE & SUBSYSTEMS
    # --------------------------------------------------------------------------
    add_flat_two_column_slide(
        4, "1.1. HARDWARE ARCHITECTURE & SUBSYSTEMS",
        "figs/fig_1_1_2_hardware.jpg",
        "Figure 1.2: Hardware Subsystems of an Embedded Device.",
        [
            ("Core Processing (MCU)", "Silicon microcontroller integrating CPU core, hardware timers, DMA controllers, and low-latency NVIC."),
            ("Memory Hierarchy", "Embedded Flash Memory for persistent firmware; tight on-chip SRAM for execution stack, heap, and buffers."),
            ("Physical Transducers", "High-precision analog/digital Sensors paired with robust Actuators (relays, buzzers, solenoid valves)."),
            ("Communication Buses", "Local serial interfaces (UART, SPI, I2C, CAN) alongside networked transceivers (Ethernet, Wi-Fi, Cellular)."),
            ("Power & Protection", "Integrated LDO regulators, Brown-Out Detector (BOD), and hardware Watchdog Timers.")
        ]
    )

    # --------------------------------------------------------------------------
    # SLIDE 5: 1.1 EMBEDDED NATURE & EDGE COMPUTING IN IOT
    # --------------------------------------------------------------------------
    add_flat_two_column_slide(
        5, "1.1. EMBEDDED NATURE & EDGE COMPUTING IN IOT",
        "figs/fig_1_1_3_edge.png",
        "Figure 1.3: Embedded Edge Node in Cloud-IoT Topology.",
        [
            ("The 'Embedded' Nature", "Deeply encapsulated inside host machinery; user interacts solely via HMI or cloud; requires bare-metal register manipulation."),
            ("Frontline IoT Edge Node", "Serves as the vital physical bridge collecting field telemetry and driving deterministic real-time actuation."),
            ("On-Device Edge Computing", "Performs local filtering, anomaly classification, and inference at the edge, drastically reducing cloud bandwidth and latency."),
            ("Disconnection Resilience", "Operates 100% autonomously during WAN network blackouts; maintains local safety cutoffs independently."),
            ("Data Fallback", "Retains critical telemetry on local SD/Flash memory during signal loss and automatically synchronizes upon reconnection.")
        ]
    )

    # --------------------------------------------------------------------------
    # SLIDE 6: 1.2 DEDICATED FUNCTIONALITY & RESOURCE CONSTRAINTS
    # --------------------------------------------------------------------------
    add_flat_two_column_slide(
        6, "1.2. DEDICATED FUNCTIONALITY & RESOURCE CONSTRAINTS",
        "figs/fig_1_2_characteristics.jpg",
        "Figure 1.4: Multi-dimensional Constraints in Embedded Design.",
        [
            ("Dedicated Functionality", "Hardware and firmware are strictly trimmed to the target workload -> minimizes BOM unit cost and eliminates software bloat."),
            ("Engineering Trade-off", "High specialization reduces generic software reusability; major specification shifts necessitate PCB redesign."),
            ("Scarce Flash Memory", "Constrained to tens of KB to few MB (vs. hundreds of GB SSD on PCs) -> demands aggressive code footprint minimization."),
            ("Extremely Tight SRAM", "Limited to 16 KB – 512 KB -> dynamic heap allocation (malloc) is strictly prohibited; static allocation is mandatory."),
            ("Clock Speed & Pinout", "Conservative MHz clock rates to conserve battery; tightly budgeted GPIO pinout requires pin multiplexing.")
        ]
    )

    # Helper for flat 2-column text slide (NO boxes, NO cards)
    def add_flat_dual_text_slide(slide_num, title, col1_header, col1_bullets, col2_header, col2_bullets):
        slide = create_base_slide(prs)
        add_pptx_title(slide, title)
        add_pptx_footer(slide, slide_num)

        # Left Column text box
        tb_l = slide.shapes.add_textbox(Inches(0.7), Inches(1.35), Inches(5.7), Inches(5.5))
        tf_l = tb_l.text_frame
        tf_l.word_wrap = True
        tf_l.margin_left = tf_l.margin_top = tf_l.margin_right = tf_l.margin_bottom = 0

        p_h1 = tf_l.paragraphs[0]
        p_h1.text = col1_header
        p_h1.font.name = FONT_LATEX
        p_h1.font.size = Pt(15)
        p_h1.font.bold = True
        p_h1.font.color.rgb = TEXT_BLACK
        p_h1.space_after = Pt(10)

        for kw, desc in col1_bullets:
            add_pptx_bullet(tf_l, kw, desc, is_first=False, pt_size=12, space_after=10)

        # Right Column text box
        tb_r = slide.shapes.add_textbox(Inches(6.8), Inches(1.35), Inches(5.7), Inches(5.5))
        tf_r = tb_r.text_frame
        tf_r.word_wrap = True
        tf_r.margin_left = tf_r.margin_top = tf_r.margin_right = tf_r.margin_bottom = 0

        p_h2 = tf_r.paragraphs[0]
        p_h2.text = col2_header
        p_h2.font.name = FONT_LATEX
        p_h2.font.size = Pt(15)
        p_h2.font.bold = True
        p_h2.font.color.rgb = TEXT_BLACK
        p_h2.space_after = Pt(10)

        for kw, desc in col2_bullets:
            add_pptx_bullet(tf_r, kw, desc, is_first=False, pt_size=12, space_after=10)

    # --------------------------------------------------------------------------
    # SLIDE 7: 1.2 REAL-TIME DETERMINISM & SYSTEM RELIABILITY
    # --------------------------------------------------------------------------
    add_flat_dual_text_slide(
        7, "1.2. REAL-TIME DETERMINISM & SYSTEM RELIABILITY",
        "Real-Time Determinism & Deadlines",
        [
            ("Correctness Criteria", "System correctness requires BOTH logical accuracy AND completion strictly before defined Deadlines."),
            ("Hard Real-Time", "Missing a deadline causes catastrophic system destruction or loss of human life (automotive airbag, ABS braking, medical ventilators)."),
            ("Soft Real-Time", "Deadline misses degrade Quality of Service without catastrophic harm (UI display refresh, periodic logging)."),
            ("Deterministic Timing", "Interrupt latency and execution times must have mathematical upper bounds; zero tolerance for unbounded jitter."),
            ("Architectural Support", "Enforced via low-latency Nested Vectored Interrupt Controllers (NVIC) and preemptive priority RTOS scheduling.")
        ],
        "Autonomous Reliability & Self-Healing",
        [
            ("Continuous 24/7 Uptime", "Must operate unattended for decades without human manual intervention or reboot."),
            ("Watchdog Timer (WDT)", "Independent free-running counter resets the MCU automatically if firmware deadlocks or enters an infinite loop."),
            ("Brown-Out Detector (BOD)", "Continually monitors supply voltage droop, forcing safe reset before Flash memory is corrupted."),
            ("Cyclic Redundancy Checks", "Hardware CRC engine verifies Flash code integrity and communication packet validity."),
            ("Fail-Safe Defaults", "Automatic fallback into safe benign states (opening relays, sounding alarms) upon critical faults.")
        ]
    )

    # --------------------------------------------------------------------------
    # SLIDE 8: 1.2 POWER EFFICIENCY, SECURITY & MAINTENANCE
    # --------------------------------------------------------------------------
    add_flat_dual_text_slide(
        8, "1.2. POWER EFFICIENCY, SECURITY & MAINTENANCE",
        "Ultra-Low Power Optimization",
        [
            ("Deep Sleep States", "Advanced low-power modes (Sleep, Standby, Deep Software Standby) dropping quiescent current to microamperes (µA)."),
            ("Interrupt-Driven Wakeup", "CPU core remains powered down, waking up solely upon external sensor edge interrupts or periodic RTC timers."),
            ("Autonomous DMA Transfers", "Direct Memory Access transfers peripheral data to RAM buffers without CPU intervention, maximizing sleep duty cycles."),
            ("Energy Harvesting", "Supports ultra-low-power duty cycles compatible with solar, thermal, and RF energy-harvesting nodes.")
        ],
        "Embedded Security & Remote Maintenance",
        [
            ("Hardware Secure Boot", "Silicon cryptographic engine validates digital signatures (RSA/ECC) of firmware before allowing boot execution."),
            ("Arm TrustZone Isolation", "Partitions memory and peripherals into Secure and Non-Secure domains, isolating encryption keys from user code."),
            ("Line-Rate Encryption", "Hardware-accelerated AES, SHA-256, and ECC engines enable secure TLS/DTLS encryption for cloud telemetry."),
            ("Dual-Bank Flash OTA", "Downloads and stages new firmware images into secondary Flash bank without downtime; automated rollback on failure.")
        ]
    )

    # --------------------------------------------------------------------------
    # SLIDE 9: 1.3 EMBEDDED SYSTEMS VS. GENERAL-PURPOSE PCS
    # --------------------------------------------------------------------------
    add_flat_two_column_slide(
        9, "1.3. EMBEDDED SYSTEMS VS. GENERAL-PURPOSE PCS",
        "figs/fig_1_3_gpc_vs_embedded.jpg",
        "Figure 1.5: Paradigm Comparison: General-Purpose PC vs. Embedded System.",
        [
            ("Specialization vs. Versatility", "General-Purpose Computers (GPC) run arbitrary multi-user software, office suites, and gaming. Embedded systems are deeply pruned for one dedicated workload."),
            ("Monolithic SoC Integration", "Embedded MCUs integrate CPU, SRAM, Flash ROM, ADC/DAC, timers, and serial controllers on a single die. PCs rely on discrete motherboards with socketed CPUs and RAM sticks."),
            ("Environmental Robustness", "Embedded systems survive extreme industrial temperatures (-40°C to +85°C/125°C), high vibration, dust, and electrical noise. PCs require climate-controlled rooms and active fans."),
            ("Cost & Unit Optimization", "MCU Bill-of-Materials cost is measured in single-digit dollars for mass production, whereas PC platforms cost hundreds to thousands of dollars.")
        ]
    )

    # --------------------------------------------------------------------------
    # SLIDE 10: 1.3 DETAILED TECHNICAL COMPARISON MATRIX (FLAT TABLE)
    # --------------------------------------------------------------------------
    s10 = create_base_slide(prs)
    add_pptx_title(s10, "1.3. DETAILED TECHNICAL COMPARISON MATRIX")
    add_pptx_footer(s10, 10)

    rows, cols = 9, 3
    tbl_shape = s10.shapes.add_table(rows, cols, Inches(0.7), Inches(1.3), Inches(11.933), Inches(4.8))
    table = tbl_shape.table
    table.columns[0].width = Inches(2.6)
    table.columns[1].width = Inches(4.65)
    table.columns[2].width = Inches(4.68)

    headers = ["Evaluation Metric", "Embedded Systems (MCU-based)", "General-Purpose Computers (PC-based)"]
    for c_idx, h_text in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = PTIT_BLUE
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.name = FONT_LATEX
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

    data = [
        ("Design Objective", "Application-specific; deeply optimized for a single mission", "General-purpose; supports arbitrary concurrent user applications"),
        ("Processor Architecture", "Low-power MCU (e.g., ARM Cortex-M), SoC, DSP (MHz)", "High-performance multi-core CPU (x86/x64, GHz), discrete GPU"),
        ("Memory Budget", "Flash: 32 KB – 4 MB; SRAM: 8 KB – 1 MB (On-chip)", "RAM: 8 GB – 64+ GB; NVMe SSD Storage: 512 GB – Multiple TBs"),
        ("Operating System", "Bare-metal firmware, Hard RTOS (FreeRTOS, Zephyr)", "General-Purpose OS (Windows, Linux, macOS)"),
        ("Timing Constraints", "Deterministic hard/soft real-time; bounded µs latency", "Best-effort execution; non-deterministic task scheduling"),
        ("Power Consumption", "Ultra-low power (mW to µA); multi-year battery operation", "High power draw (50W – 500W+); requires continuous AC mains"),
        ("Operating Environment", "Harsh industrial (-40°C to 85°C), vibration, moisture, dust", "Benign indoor climate; office/datacenter controlled environments"),
        ("Maintenance & Lifecycle", "Secure remote OTA firmware; decades-long operational life", "Frequent OS patches, dynamic package installations, short lifespan")
    ]

    for r_idx, (m_col, e_col, p_col) in enumerate(data, start=1):
        bg_c = WHITE if r_idx % 2 == 1 else RGBColor(245, 249, 252)
        for c_idx, text_val in enumerate([m_col, e_col, p_col]):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_c
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = text_val
            p.font.name = FONT_LATEX
            p.font.size = Pt(10.5)
            p.font.bold = (c_idx == 0)
            p.font.color.rgb = TEXT_BLACK
            p.alignment = PP_ALIGN.LEFT

    # Flat takeaway text at bottom (NO card)
    tb_tk = s10.shapes.add_textbox(Inches(0.7), Inches(6.3), Inches(11.933), Inches(0.5))
    tf_tk = tb_tk.text_frame
    p_tk = tf_tk.paragraphs[0]
    r_tk_arr = p_tk.add_run()
    r_tk_arr.text = "⇒  Conclusion: "
    r_tk_arr.font.bold = True
    r_tk_arr.font.color.rgb = PTIT_BLUE
    r_tk_txt = p_tk.add_run()
    r_tk_txt.text = "For dedicated IoT edge monitoring, dedicated microcontrollers (e.g. ARM Cortex-M85 / RA8M1) vastly outperform PCs in determinism, reliability, power efficiency, and unit cost."
    r_tk_txt.font.name = FONT_LATEX
    r_tk_txt.font.bold = True
    r_tk_txt.font.color.rgb = TEXT_BLACK
    r_tk_txt.font.size = Pt(11)

    # --------------------------------------------------------------------------
    # SLIDE 11: 1.4 THE 5-LAYER SOFTWARE ARCHITECTURE
    # --------------------------------------------------------------------------
    add_flat_two_column_slide(
        11, "1.4. THE 5-LAYER EMBEDDED SOFTWARE ARCHITECTURE",
        "figs/fig_1_4_layered_arch.jpg",
        "Figure 1.6: Standard 5-Layer Embedded Software Stack.",
        [
            ("Separation of Concerns", "Each architectural tier encapsulates specific functional responsibilities, hiding lower-level hardware quirks from upper layers."),
            ("Hardware Decoupling", "Application logic never manipulates raw chip registers; all hardware interactions pass through standardized APIs."),
            ("Concurrent Engineering", "Hardware schematics and software develop in parallel; allows extensive Unit Testing natively on host PCs."),
            ("Silicon Portability", "When migrating to an alternative MCU vendor, only the MCAL/HAL drivers are rewritten—preserving 100% of Middleware and Application code."),
            ("Extensibility", "Modifying hardware components (e.g., swapping an analog sensor for an I2C sensor) requires zero changes to core business rules.")
        ]
    )

    # --------------------------------------------------------------------------
    # SLIDE 12: 1.4.1 – 1.4.2 HARDWARE LAYER & MCAL / HAL
    # --------------------------------------------------------------------------
    add_flat_dual_text_slide(
        12, "1.4.1 – 1.4.2. HARDWARE LAYER & MCAL / HAL",
        "Physical Hardware Layer",
        [
            ("Target Silicon", "MCU silicon (e.g., Renesas RA8M1 / Cortex-M85), analog front-ends, relay drivers, communication transceivers."),
            ("PCB Layout Integrity", "Physical trace routing directly dictates Signal-to-Noise Ratio (SNR), ADC sampling precision, and parasitic leakage currents."),
            ("Hardware Robustness", "In-line PTC resettable fuses, TVS suppression diodes, and galvanic isolation prevent field electrical destruction."),
            ("Engineering Reality", "Silicon layout bugs and hardware flaws cannot be solved by software patches alone—demands rigorous design validation.")
        ],
        "MCAL vs. HAL Driver Layers",
        [
            ("MCAL (Microcontroller Abstraction)", "Hardware-dependent register manipulation; configures internal silicon clocks, pin multiplexing, interrupt routing, and DMA channels."),
            ("HAL (Hardware Abstraction Layer)", "Vendor-neutral API abstraction wrapping register complexity into uniform C calls (HAL_GPIO_WritePin(), HAL_ADC_Read())."),
            ("Device Driver Interface", "Provides standard bus interfaces allowing external sensor/actuator drivers to operate seamlessly across different chips."),
            ("Portability Boundary", "Decouples silicon-specific CMSIS headers from higher-level middleware and application stacks.")
        ]
    )

    # --------------------------------------------------------------------------
    # SLIDE 13: 1.4.3 REAL-TIME OPERATING SYSTEM (RTOS LAYER)
    # --------------------------------------------------------------------------
    add_flat_dual_text_slide(
        13, "1.4.3. REAL-TIME OPERATING SYSTEM (RTOS LAYER)",
        "Preemptive Multitasking vs. Super-Loop",
        [
            ("Super-Loop Pitfalls", "Sequential while(1) polling suffers catastrophic latency blockage when long I/O tasks (e.g., TCP network handshakes) stall the CPU."),
            ("Preemptive Priority Scheduling", "Critical high-priority tasks (e.g., fire detection, over-temperature alarm) preempt the CPU within microseconds upon interrupt."),
            ("Deterministic Software Timers", "Non-blocking software timers drive periodic sensor acquisitions with millisecond determinism."),
            ("Tickless Low-Power Idle", "Automatically transitions the MCU core into deep sleep whenever all tasks are blocked in wait states.")
        ],
        "IPC & Resource Synchronization",
        [
            ("Message Queues", "Safe FIFO message queues pass sensor telemetry to processing tasks without memory corruption or locks."),
            ("Mutexes & Semaphores", "Guard shared communication buses (I2C, SPI, UART) against race conditions and priority inversions."),
            ("Task Priority Hierarchy", "Emergency Safety Alarms (High) -> Sensor Filtering (Medium) -> Cloud MQTT Telemetry (Low)."),
            ("Leading Industrial RTOS", "FreeRTOS (ubiquitous, minimal footprint), Zephyr RTOS (modular, modern IoT stack), Azure RTOS / ThreadX.")
        ]
    )

    # --------------------------------------------------------------------------
    # SLIDE 14: 1.4.4 MIDDLEWARE & SYSTEM SERVICES LAYER
    # --------------------------------------------------------------------------
    add_flat_dual_text_slide(
        14, "1.4.4. MIDDLEWARE & SYSTEM SERVICES LAYER",
        "IoT Networking & Protocol Stacks",
        [
            ("Embedded TCP/IP Stack", "Lightweight network stacks (LwIP, FreeRTOS-Plus-TCP) powering low-overhead Ethernet and Wi-Fi connectivity."),
            ("MQTT Protocol", "Lightweight Publish/Subscribe messaging protocol minimizing transmission packet overhead and radio active time."),
            ("CoAP & HTTP Micro-Servers", "Embedded micro-web servers for on-site device commissioning, status diagnostics, and RESTful APIs."),
            ("DHCP & DNS Clients", "Automated IP dynamic configuration and central cloud hostname resolution over dynamic local networks.")
        ],
        "Security, Storage & Maintenance",
        [
            ("Cryptographic Library (MbedTLS)", "Implements industry-grade TLS 1.3 / DTLS encryption, safeguarding telemetry against eavesdropping and MITM attacks."),
            ("Embedded File System (FatFS)", "Hardware-independent FAT12/16/32 and exFAT implementation for local microSD and external SPI Flash storage."),
            ("Robust OTA Engine", "Handles remote firmware download, cryptographic hash validation (SHA-256), and dual-bank Flash staging."),
            ("Diagnostics & Logging", "Tracks runtime CPU load, stack watermarks, memory leaks, and records multi-level error logs for post-mortem analysis.")
        ]
    )

    # --------------------------------------------------------------------------
    # SLIDE 15: 1.4.5 APPLICATION LAYER & DOMAIN LOGIC
    # --------------------------------------------------------------------------
    add_flat_dual_text_slide(
        15, "1.4.5. APPLICATION LAYER & DOMAIN LOGIC",
        "Core Domain Tasks & Signal Processing",
        [
            ("Periodic Sampling", "Orchestrates multi-sensor data acquisition (ambient temperature, relative humidity, VOC/toxic gas, structural vibration)."),
            ("Digital Signal Processing", "Implements Moving Average or Kalman filters to eliminate electrical noise spikes and sensor drift."),
            ("Threshold Evaluation", "Classifies environmental health across multi-tiered levels: Normal -> Early Warning -> Critical Alarm."),
            ("Autonomous Emergency Actuation", "Drives safety actuators immediately upon anomaly detection (opening relays, running fans, sirens)."),
            ("Telemetry Serialization", "Encodes validated telemetry into compact standard formats (JSON, CBOR, Protocol Buffers).")
        ],
        "Design Decoupling & Host Testability",
        [
            ("Total Hardware Independence", "Pure business logic interacts strictly through generic domain APIs (get_sensor_data(), trigger_alarm())."),
            ("Transducer Modularity", "Replacing an analog sensor with an I2C sensor leaves 100% of threshold comparison logic completely untouched."),
            ("Host-Based Simulation", "Entire application state machine compiles and runs on a host PC using simulated sensor streams for rapid verification."),
            ("Finite State Machine (FSM)", "Structured as deterministic states: Boot -> Normal Operation -> Warning -> Emergency Fault.")
        ]
    )

    # --------------------------------------------------------------------------
    # SLIDE 16: CHAPTER SUMMARY & TRANSITION TO CHAPTER II (FLAT LIST)
    # --------------------------------------------------------------------------
    s16 = create_base_slide(prs)
    add_pptx_title(s16, "CHAPTER I SUMMARY & TRANSITION TO CHAPTER II")
    add_pptx_footer(s16, 16)

    tb16 = s16.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(11.7), Inches(5.5))
    tf16 = tb16.text_frame
    tf16.word_wrap = True
    tf16.margin_left = tf16.margin_top = tf16.margin_right = tf16.margin_bottom = 0

    s16_bullets = [
        ("Embedded Systems Defined", "Tight co-design synergy between specialized MCU silicon, deterministic firmware, and physical world transducers."),
        ("The Core Triad", "Stringent resource boundaries, deterministic real-time deadlines, and 24/7 unassisted autonomous reliability."),
        ("The 5-Layer Software Stack", "Golden architectural standard separating silicon registers from business logic, ensuring portability and maintainability."),
        ("Escalating Edge Demands", "Modern IoT edge stations demand on-device ML/DSP, sub-microsecond interrupt latency, and silicon-enforced domain security."),
        ("Bridge to Chapter II", "Transitioning from legacy 8/16-bit or basic 32-bit cores to advanced vector-accelerated architectures.")
    ]

    first = True
    for kw, desc in s16_bullets:
        add_pptx_bullet(tf16, kw, desc, is_first=first, pt_size=13, space_after=14)
        first = False

    p_nxt = tf16.add_paragraph()
    p_nxt.space_before = Pt(14)
    r_arr = p_nxt.add_run()
    r_arr.text = "⇒  "
    r_arr.font.bold = True
    r_arr.font.color.rgb = PTIT_BLUE
    r_txt = p_nxt.add_run()
    r_txt.text = "Chapter II: Embedded Processor Architectures & ARM Cortex-M85 on Renesas RA8M1"
    r_txt.font.name = FONT_LATEX
    r_txt.font.bold = True
    r_txt.font.color.rgb = TEXT_BLACK
    r_txt.font.size = Pt(14)

    # Save
    pptx_path = "Chuong_I_Overview_Embedded_Systems.pptx"
    prs.save(pptx_path)
    print(f"  Successfully generated: {pptx_path}")


# ==============================================================================
# 3. BUILD MINIMALIST LATEX BEAMER PROJECT (main.tex)
# ==============================================================================
def build_beamer():
    print("[3/4] Building Minimalist LaTeX Beamer Project (main.tex)...")

    tex_content = r"""% ============================================================================
% CHAPTER I: OVERVIEW OF EMBEDDED SYSTEMS
% Academic & Technical Presentation (LaTeX Beamer)
% Exact style match with Chapter II:
% - Template background: figs/bg.jpg (Cyan top banner + PTIT logo + 16:9 white body)
% - Minimal flat layout: NO cards, NO bounding boxes, NO shadows
% - Latin Modern font (\usepackage{lmodern})
% - Text & bullets directly on white background
% ============================================================================
\documentclass[aspectratio=169]{beamer}

\usepackage[T1]{fontenc}
\usepackage[utf8]{inputenc}
\usepackage{lmodern}
\usepackage{graphicx}
\usepackage{xcolor}
\usepackage{tikz}
\usepackage{booktabs}
\usepackage{array}

\definecolor{ptitblue}{HTML}{00A9CE}
\setbeamertemplate{navigation symbols}{}
\setbeamertemplate{footline}{%
  \hfill{\color{gray}\footnotesize\insertframenumber}\hspace{0.4cm}\vspace{0.25cm}}
\setbeamertemplate{background}{%
  \includegraphics[width=\paperwidth,height=\paperheight]{figs/bg.jpg}}
\setbeamertemplate{itemize item}{\color{ptitblue}$\blacktriangleright$}
\setbeamertemplate{itemize subitem}{\color{ptitblue}--}
\setlength{\leftmargini}{0.55cm}
\setlength{\itemsep}{0.35cm}

% Slide title inside cyan banner (matches Chapter II)
\newcommand{\slidetitle}[1]{%
  \begin{tikzpicture}[remember picture,overlay]
    \node[anchor=west, text=white, font=\bfseries\fontsize{17}{21}\selectfont,
      text width=0.82\paperwidth, inner sep=0pt]
      at ([xshift=0.5cm,yshift=-0.58cm]current page.north west) {\uppercase{#1}};
  \end{tikzpicture}%
  \vspace*{1.55cm}}

% Left column figure with caption
\newcommand{\leftfig}[2]{%
  \includegraphics[width=\linewidth]{#1}\\[0.15cm]
  {\footnotesize\color{black!70}\textit{#2}}}

\begin{document}

% ===================== 1. TITLE SLIDE =====================
\begin{frame}[plain,t]
  \slidetitle{CHAPTER I. OVERVIEW OF EMBEDDED SYSTEMS}
  \par\vspace{0.1cm}
  \begin{columns}[T]
    \column{0.60\textwidth}
      {\normalsize\textbf{Foundations, Core Characteristics \& Software Architecture}}\\[0.25cm]
      \begin{itemize}
        \item \textbf{Curriculum Scope:} Core Embedded Systems \& Edge Computing Architecture
        \item \textbf{Engineering Paradigms:} Deep hardware-software co-design, deterministic timing, and tight resource budgets
        \item \textbf{Target Platform:} 32-bit Microcontroller Core (ARM Cortex-M85 / Renesas RA8M1)
        \item \textbf{Implementation Standards:} C99/POSIX firmware patterns, Hard RTOS, standardized MCAL/HAL stacks
      \end{itemize}
    \column{0.38\textwidth}
      \begin{center}
      \includegraphics[height=5.2cm]{figs/fig_cover.jpg}
      \end{center}
  \end{columns}
\end{frame}

% ===================== 2. TABLE OF CONTENTS =====================
\begin{frame}[plain,t]
  \slidetitle{TABLE OF CONTENTS}
  \par\vspace{0.05cm}
  {\normalsize\textbf{1. Concept \& Nature of Embedded Systems}}\\[0.05cm]
  {\footnotesize\hspace{0.8cm}1.1. Standard Definition, 3-Element Triad, and Physical Embedded Nature\\
  \hspace{0.8cm}1.2. Role of Embedded Systems in IoT \& Edge Computing}\\[0.15cm]
  {\normalsize\textbf{2. Core Technical Characteristics}}\\[0.05cm]
  {\footnotesize\hspace{0.8cm}2.1. Dedicated Functionality \& Severe Resource Constraints\\
  \hspace{0.8cm}2.2. Deterministic Real-Time Execution \& High Reliability (WDT, BOD, CRC)\\
  \hspace{0.8cm}2.3. Ultra-Low Power, Embedded Security \& Firmware Maintenance (OTA)}\\[0.15cm]
  {\normalsize\textbf{3. Embedded Systems vs. General-Purpose Computers}}\\[0.05cm]
  {\footnotesize\hspace{0.8cm}3.1. Architectural Paradigms: Specialization vs. Versatility\\
  \hspace{0.8cm}3.2. Detailed Multi-Criteria Technical Comparison Matrix}\\[0.15cm]
  {\normalsize\textbf{4. Layered Embedded Software Architecture}}\\[0.05cm]
  {\footnotesize\hspace{0.8cm}4.1. The 5-Layer Stack Model: Hardware $\rightarrow$ MCAL $\rightarrow$ HAL $\rightarrow$ RTOS $\rightarrow$ Middleware $\rightarrow$ Application\\
  \hspace{0.8cm}4.2. RTOS Layer: Multitasking vs. Super-Loop \& Inter-Process Communication\\
  \hspace{0.8cm}4.3. Middleware \& Application Layer: Decoupling and Host Testability}\\[0.22cm]
  \textbf{$\Rightarrow$ Guiding Question: What fundamentally distinguishes embedded software from desktop/PC software?}
\end{frame}

% ===================== 3. DEFINITION & NATURE =====================
\begin{frame}
  \slidetitle{1.1. Embedded System Definition \& Nature}
  \begin{columns}[T]
    \column{0.49\textwidth}
      \leftfig{figs/fig_1_1_typical_system.jpg}{Figure 1.1: Architecture of a Typical Embedded System}
    \column{0.49\textwidth}
      \begin{itemize}
        \item \textbf{Definition:} Specialized computer system integrated within a host device to perform dedicated, pre-defined functions.
        \item \textbf{Triad Nature:} Inseparable co-design between \textbf{Computation} (Hardware) --- \textbf{Processing Logic} (Firmware) --- \textbf{Physical World Interaction} (Sensors/Actuators).
        \item \textbf{Design Goal:} Fulfill dedicated tasks under strict constraints: deterministic latency, ultra-low power, compact size, high MTBF, and low unit cost.
        \item \textbf{Distinction:} Not a general-purpose PC --- strictly eliminates software bloat in favor of deep optimization for a single mission.
        \item \textbf{Real Example:} Environmental monitoring edge node: continuous sensing, local anomaly detection, and 24/7 telemetry.
      \end{itemize}
  \end{columns}
\end{frame}

% ===================== 4. HARDWARE SUBSYSTEMS =====================
\begin{frame}
  \slidetitle{1.1. Hardware Architecture \& Subsystems}
  \begin{columns}[T]
    \column{0.49\textwidth}
      \leftfig{figs/fig_1_1_2_hardware.jpg}{Figure 1.2: Hardware Subsystems of an Embedded Device}
    \column{0.49\textwidth}
      \begin{itemize}
        \item \textbf{Core Processing (MCU):} Silicon microcontroller integrating CPU core, hardware timers, DMA controllers, and low-latency NVIC.
        \item \textbf{Memory Hierarchy:} Embedded \textbf{Flash Memory} for persistent firmware; tight on-chip \textbf{SRAM} for execution stack, heap, and buffers.
        \item \textbf{Physical Transducers:} High-precision \textbf{Sensors} paired with robust \textbf{Actuators} (relays, buzzers, solenoid valves).
        \item \textbf{Communication Buses:} Local serial interfaces (\textbf{UART, SPI, I2C, CAN}) alongside networked transceivers (\textbf{Ethernet, Wi-Fi, Cellular}).
        \item \textbf{Power \& Protection:} Integrated \textbf{LDO} regulators, Brown-Out Detector (\textbf{BOD}), and hardware \textbf{Watchdog Timers}.
      \end{itemize}
  \end{columns}
\end{frame}

% ===================== 5. EDGE COMPUTING IN IOT =====================
\begin{frame}
  \slidetitle{1.1. Embedded Nature \& Edge Computing in IoT}
  \begin{columns}[T]
    \column{0.49\textwidth}
      \leftfig{figs/fig_1_1_3_edge.png}{Figure 1.3: Embedded Edge Node in Cloud-IoT Topology}
    \column{0.49\textwidth}
      \begin{itemize}
        \item \textbf{The `Embedded' Nature:} Deeply encapsulated inside host machinery; user interacts solely via HMI or cloud; requires bare-metal register manipulation.
        \item \textbf{Frontline IoT Edge Node:} Serves as the vital physical bridge collecting field telemetry and driving deterministic real-time actuation.
        \item \textbf{On-Device Edge Computing:} Performs local filtering, anomaly classification, and inference at the edge, drastically reducing cloud bandwidth and latency.
        \item \textbf{Disconnection Resilience:} Operates 100\% autonomously during WAN network blackouts; maintains local safety cutoffs independently.
        \item \textbf{Data Fallback:} Retains critical telemetry on local SD/Flash memory during signal loss and automatically synchronizes upon reconnection.
      \end{itemize}
  \end{columns}
\end{frame}

% ===================== 6. RESOURCE CONSTRAINTS =====================
\begin{frame}
  \slidetitle{1.2. Dedicated Functionality \& Resource Constraints}
  \begin{columns}[T]
    \column{0.49\textwidth}
      \leftfig{figs/fig_1_2_characteristics.jpg}{Figure 1.4: Multi-dimensional Constraints in Embedded Design}
    \column{0.49\textwidth}
      \begin{itemize}
        \item \textbf{Dedicated Functionality:} Hardware and firmware are strictly trimmed to the target workload $\rightarrow$ minimizes BOM unit cost and eliminates software bloat.
        \item \textbf{Engineering Trade-off:} High specialization reduces generic software reusability; major specification shifts necessitate PCB redesign.
        \item \textbf{Scarce Flash Memory:} Constrained to tens of KB to few MB (vs. hundreds of GB SSD on PCs) $\rightarrow$ demands aggressive code footprint minimization.
        \item \textbf{Extremely Tight SRAM:} Limited to 16 KB -- 512 KB $\rightarrow$ \textbf{dynamic heap allocation (\texttt{malloc}) is strictly prohibited}; static allocation is mandatory.
        \item \textbf{Clock Speed \& Pinout:} Conservative MHz clock rates to conserve battery; tightly budgeted GPIO pinout requires pin multiplexing.
      \end{itemize}
  \end{columns}
\end{frame}

% ===================== 7. REAL-TIME DETERMINISM & RELIABILITY =====================
\begin{frame}
  \slidetitle{1.2. Real-Time Determinism \& System Reliability}
  \begin{columns}[T]
    \column{0.49\textwidth}
      {\normalsize\textbf{Real-Time Determinism \& Deadlines}}\\[0.15cm]
      \begin{itemize}
        \item \textbf{Correctness Criteria:} System correctness requires \textbf{both} logical accuracy AND completion strictly before defined \textbf{Deadlines}.
        \item \textbf{Hard Real-Time:} Missing a deadline causes catastrophic destruction or loss of human life (\textit{airbags, ABS braking, medical ventilators}).
        \item \textbf{Soft Real-Time:} Deadline misses degrade Quality of Service without catastrophic harm (\textit{UI display refresh, periodic logging}).
        \item \textbf{Deterministic Timing:} Interrupt latency and execution times must have mathematical upper bounds; zero tolerance for unbounded jitter.
        \item \textbf{Architectural Support:} Enforced via low-latency NVIC and preemptive priority RTOS scheduling.
      \end{itemize}
    \column{0.49\textwidth}
      {\normalsize\textbf{Autonomous Reliability \& Self-Healing}}\\[0.15cm]
      \begin{itemize}
        \item \textbf{Continuous 24/7 Uptime:} Must operate unattended for decades without human manual intervention or reboot.
        \item \textbf{Watchdog Timer (WDT):} Independent free-running counter resets the MCU automatically if firmware deadlocks or enters an infinite loop.
        \item \textbf{Brown-Out Detector (BOD):} Continually monitors supply voltage droop, forcing safe reset before Flash memory is corrupted.
        \item \textbf{CRC Integrity:} Hardware CRC engine verifies Flash code integrity and communication packet validity.
        \item \textbf{Fail-Safe Defaults:} Automatic fallback into safe benign states (\textit{opening relays, sounding alarms}) upon critical faults.
      \end{itemize}
  \end{columns}
\end{frame}

% ===================== 8. POWER & SECURITY =====================
\begin{frame}
  \slidetitle{1.2. Power Efficiency, Security \& Maintenance}
  \begin{columns}[T]
    \column{0.49\textwidth}
      {\normalsize\textbf{Ultra-Low Power Optimization}}\\[0.15cm]
      \begin{itemize}
        \item \textbf{Deep Sleep States:} Advanced low-power modes (Sleep, Standby, Deep Software Standby) dropping quiescent current to microamperes ($\mu\text{A}$).
        \item \textbf{Interrupt-Driven Wakeup:} CPU core remains powered down, waking up solely upon external sensor edge interrupts or periodic RTC timers.
        \item \textbf{Autonomous DMA Transfers:} Direct Memory Access transfers peripheral data to RAM buffers without CPU intervention, maximizing sleep duty cycles.
        \item \textbf{Energy Harvesting:} Supports ultra-low-power duty cycles compatible with solar, thermal, and RF energy-harvesting nodes.
      \end{itemize}
    \column{0.49\textwidth}
      {\normalsize\textbf{Embedded Security \& Maintenance}}\\[0.15cm]
      \begin{itemize}
        \item \textbf{Hardware Secure Boot:} Silicon cryptographic engine validates digital signatures (RSA/ECC) of firmware before allowing boot execution.
        \item \textbf{Arm TrustZone Isolation:} Partitions memory and peripherals into Secure and Non-Secure domains, isolating encryption keys from user code.
        \item \textbf{Line-Rate Encryption:} Hardware-accelerated AES, SHA-256, and ECC engines enable secure TLS/DTLS encryption for cloud telemetry.
        \item \textbf{Dual-Bank Flash OTA:} Downloads and stages new firmware images into secondary Flash bank without downtime; automated rollback on failure.
      \end{itemize}
  \end{columns}
\end{frame}

% ===================== 9. GPC VS EMBEDDED =====================
\begin{frame}
  \slidetitle{1.3. Embedded Systems vs. General-Purpose PCs}
  \begin{columns}[T]
    \column{0.49\textwidth}
      \leftfig{figs/fig_1_3_gpc_vs_embedded.jpg}{Figure 1.5: Paradigm Comparison: General-Purpose PC vs. Embedded System}
    \column{0.49\textwidth}
      \begin{itemize}
        \item \textbf{Specialization vs. Versatility:} General-Purpose Computers (GPC) run arbitrary multi-user software, office suites, and gaming. Embedded systems are deeply pruned for one dedicated workload.
        \item \textbf{Monolithic SoC Integration:} Embedded MCUs integrate CPU, SRAM, Flash ROM, ADC/DAC, timers, and serial controllers on a single die. PCs rely on discrete motherboards with socketed CPUs and RAM sticks.
        \item \textbf{Environmental Robustness:} Embedded systems survive extreme industrial temperatures ($-40^\circ\text{C}$ to $+85^\circ\text{C}$ / $125^\circ\text{C}$), high vibration, dust, and electrical noise. PCs require climate-controlled rooms and active fans.
        \item \textbf{Cost \& Unit Optimization:} MCU Bill-of-Materials cost is measured in single-digit dollars for mass production, whereas PC platforms cost hundreds to thousands of dollars.
      \end{itemize}
  \end{columns}
\end{frame}

% ===================== 10. COMPARISON MATRIX =====================
\begin{frame}
  \slidetitle{1.3. Detailed Technical Comparison Matrix}
  \vspace{0.15cm}
  \begin{center}
  {\footnotesize
  \setlength{\tabcolsep}{5pt}
  \renewcommand{\arraystretch}{1.15}
  \begin{tabular}{p{0.22\textwidth}p{0.36\textwidth}p{0.36\textwidth}}
    \toprule
    \textbf{Evaluation Metric} & \textbf{Embedded Systems (MCU-based)} & \textbf{General-Purpose Computers (PC-based)} \\
    \midrule
    Design Objective & Dedicated, deeply optimized for a single mission & Versatile, executes arbitrary concurrent user applications \\
    Processor Core & Low-power MCU (e.g., ARM Cortex-M), SoC, DSP (MHz) & High-speed multi-core CPU (x86/x64, GHz), discrete GPU \\
    Memory Hierarchy & Flash: 32 KB -- 4 MB; SRAM: 8 KB -- 1 MB (On-chip) & RAM: 8 GB -- 64+ GB; Storage: 512 GB -- Multiple TB SSD \\
    Operating System & Bare-metal firmware, Hard RTOS (\textit{FreeRTOS, Zephyr}) & General-Purpose OS (\textit{Windows, Linux, macOS}) \\
    Timing Constraints & Deterministic hard/soft real-time; bounded $\mu$s latency & Best-effort execution; non-deterministic task scheduling \\
    Power Consumption & Ultra-low power (mW to $\mu$A); multi-year battery operation & High power draw (50W -- 500W+); continuous AC mains \\
    Operating Environment & Harsh industrial ($-40^\circ\text{C}$ to $85^\circ\text{C}$), vibration, dust & Benign indoor climate; office/datacenter environments \\
    Maintenance & Secure remote OTA firmware; decades-long operational life & Frequent OS patches, dynamic package installations \\
    \bottomrule
  \end{tabular}}
  \end{center}
  \vspace{0.1cm}
  \textbf{$\Rightarrow$ Conclusion:} For dedicated IoT edge monitoring, specialized MCUs (Cortex-M85 / RA8M1) vastly outperform PCs in determinism, reliability, power efficiency, and unit cost.
\end{frame}

% ===================== 11. 5-LAYER ARCHITECTURE =====================
\begin{frame}
  \slidetitle{1.4. The 5-Layer Embedded Software Architecture}
  \begin{columns}[T]
    \column{0.49\textwidth}
      \leftfig{figs/fig_1_4_layered_arch.jpg}{Figure 1.6: Standard 5-Layer Embedded Software Stack}
    \column{0.49\textwidth}
      \begin{itemize}
        \item \textbf{Separation of Concerns:} Each architectural tier encapsulates specific functional responsibilities, hiding lower-level hardware quirks from upper layers.
        \item \textbf{Hardware Decoupling:} Application logic never manipulates raw chip registers; all hardware interactions pass through standardized APIs.
        \item \textbf{Concurrent Engineering:} Hardware schematics and software develop in parallel; allows extensive Unit Testing natively on host PCs.
        \item \textbf{Silicon Portability:} When migrating to an alternative MCU vendor, only the MCAL/HAL drivers are rewritten---preserving \textbf{100\% of Middleware and Application code}.
        \item \textbf{Extensibility:} Modifying hardware components (\textit{e.g., swapping an analog sensor for an I2C sensor}) requires zero changes to core business rules.
      \end{itemize}
  \end{columns}
\end{frame}

% ===================== 12. HARDWARE & MCAL / HAL =====================
\begin{frame}
  \slidetitle{1.4.1 -- 1.4.2. Hardware Layer \& MCAL / HAL}
  \begin{columns}[T]
    \column{0.49\textwidth}
      {\normalsize\textbf{Physical Hardware Layer}}\\[0.15cm]
      \begin{itemize}
        \item \textbf{Target Silicon:} MCU silicon (e.g., Renesas RA8M1 / Cortex-M85), analog front-ends, relay drivers, communication transceivers.
        \item \textbf{PCB Layout Integrity:} Physical trace routing directly dictates Signal-to-Noise Ratio (SNR), ADC sampling precision, and parasitic leakage currents.
        \item \textbf{Hardware Robustness:} In-line PTC resettable fuses, TVS suppression diodes, and galvanic isolation prevent field electrical destruction.
        \item \textbf{Engineering Reality:} Silicon layout bugs and hardware flaws cannot be solved by software patches alone---demands rigorous design validation.
      \end{itemize}
    \column{0.49\textwidth}
      {\normalsize\textbf{MCAL vs. HAL Driver Layers}}\\[0.15cm]
      \begin{itemize}
        \item \textbf{MCAL (Microcontroller Abstraction):} Hardware-dependent register manipulation; configures internal silicon clocks, pin multiplexing, interrupt routing, and DMA channels.
        \item \textbf{HAL (Hardware Abstraction Layer):} Vendor-neutral API abstraction wrapping register complexity into uniform C calls (\texttt{HAL\_GPIO\_WritePin()}, \texttt{HAL\_ADC\_Read()}).
        \item \textbf{Device Driver Interface:} Provides standard bus interfaces allowing external sensor/actuator drivers to operate seamlessly across different chips.
        \item \textbf{Portability Boundary:} Decouples silicon-specific CMSIS headers from higher-level middleware and application stacks.
      \end{itemize}
  \end{columns}
\end{frame}

% ===================== 13. RTOS LAYER =====================
\begin{frame}
  \slidetitle{1.4.3. Real-Time Operating System (RTOS Layer)}
  \begin{columns}[T]
    \column{0.49\textwidth}
      {\normalsize\textbf{Preemptive Multitasking vs. Super-Loop}}\\[0.15cm]
      \begin{itemize}
        \item \textbf{Super-Loop Pitfalls:} Sequential \texttt{while(1)} polling suffers catastrophic latency blockage when long I/O tasks (\textit{e.g., TCP network handshakes}) stall the CPU.
        \item \textbf{Preemptive Priority Scheduling:} Critical high-priority tasks (\textit{e.g., fire detection, over-temperature alarm}) preempt the CPU within microseconds upon interrupt.
        \item \textbf{Deterministic Software Timers:} Non-blocking software timers drive periodic sensor acquisitions with millisecond determinism.
        \item \textbf{Tickless Low-Power Idle:} Automatically transitions the MCU core into deep sleep whenever all tasks are blocked in wait states.
      \end{itemize}
    \column{0.49\textwidth}
      {\normalsize\textbf{IPC \& Resource Synchronization}}\\[0.15cm]
      \begin{itemize}
        \item \textbf{Message Queues:} Safe FIFO message queues pass sensor telemetry to processing tasks without memory corruption or locks.
        \item \textbf{Mutexes \& Semaphores:} Guard shared communication buses (\textit{I2C, SPI, UART}) against race conditions and priority inversions.
        \item \textbf{Task Priority Hierarchy:} Emergency Safety Alarms (High) $\rightarrow$ Sensor Filtering (Medium) $\rightarrow$ Cloud MQTT Telemetry (Low).
        \item \textbf{Leading Industrial RTOS:} \textbf{FreeRTOS} (ubiquitous, minimal footprint), \textbf{Zephyr RTOS} (modular, modern IoT stack), \textbf{Azure RTOS / ThreadX}.
      \end{itemize}
  \end{columns}
\end{frame}

% ===================== 14. MIDDLEWARE LAYER =====================
\begin{frame}
  \slidetitle{1.4.4. Middleware \& System Services Layer}
  \begin{columns}[T]
    \column{0.49\textwidth}
      {\normalsize\textbf{IoT Networking \& Protocol Stacks}}\\[0.15cm]
      \begin{itemize}
        \item \textbf{Embedded TCP/IP Stack:} Lightweight network stacks (\textbf{LwIP}, FreeRTOS-Plus-TCP) powering low-overhead Ethernet and Wi-Fi connectivity.
        \item \textbf{MQTT Protocol:} Lightweight Publish/Subscribe messaging protocol minimizing transmission packet overhead and radio active time.
        \item \textbf{CoAP \& HTTP Micro-Servers:} Embedded micro-web servers for on-site device commissioning, status diagnostics, and RESTful APIs.
        \item \textbf{DHCP \& DNS Clients:} Automated IP dynamic configuration and central cloud hostname resolution over dynamic local networks.
      \end{itemize}
    \column{0.49\textwidth}
      {\normalsize\textbf{Security, Storage \& Maintenance}}\\[0.15cm]
      \begin{itemize}
        \item \textbf{Cryptographic Library (MbedTLS):} Implements industry-grade TLS 1.3 / DTLS encryption, safeguarding telemetry against eavesdropping and MITM attacks.
        \item \textbf{Embedded File System (FatFS):} Hardware-independent FAT12/16/32 and exFAT implementation for local microSD and external SPI Flash storage.
        \item \textbf{Robust OTA Engine:} Handles remote firmware download, cryptographic hash validation (SHA-256), and dual-bank Flash staging.
        \item \textbf{Diagnostics \& Logging:} Tracks runtime CPU load, stack watermarks, memory leaks, and records multi-level error logs for post-mortem analysis.
      \end{itemize}
  \end{columns}
\end{frame}

% ===================== 15. APPLICATION LAYER =====================
\begin{frame}
  \slidetitle{1.4.5. Application Layer \& Domain Logic}
  \begin{columns}[T]
    \column{0.49\textwidth}
      {\normalsize\textbf{Core Domain Tasks \& Signal Processing}}\\[0.15cm]
      \begin{itemize}
        \item \textbf{Periodic Sampling:} Orchestrates multi-sensor data acquisition (ambient temperature, relative humidity, VOC/toxic gas, structural vibration).
        \item \textbf{Digital Signal Processing:} Implements Moving Average or Kalman filters to eliminate electrical noise spikes and sensor drift.
        \item \textbf{Threshold Evaluation:} Classifies environmental health across multi-tiered levels: Normal $\rightarrow$ Warning $\rightarrow$ Critical Alarm.
        \item \textbf{Autonomous Emergency Actuation:} Drives safety actuators immediately upon anomaly detection (\textit{opening relays, running fans, sirens}).
        \item \textbf{Telemetry Serialization:} Encodes validated telemetry into compact standard formats (\textit{JSON, CBOR, Protocol Buffers}).
      \end{itemize}
    \column{0.49\textwidth}
      {\normalsize\textbf{Design Decoupling \& Host Testability}}\\[0.15cm]
      \begin{itemize}
        \item \textbf{Total Hardware Independence:} Pure business logic interacts strictly through generic domain APIs (\texttt{get\_sensor\_data()}, \texttt{trigger\_alarm()}).
        \item \textbf{Transducer Modularity:} Replacing an analog sensor with an I2C sensor leaves 100% of threshold comparison logic completely untouched.
        \item \textbf{Host-Based Simulation:} Entire application state machine compiles and runs on a host PC using simulated sensor streams for rapid verification.
        \item \textbf{Finite State Machine (FSM):} Structured as deterministic states: Boot $\rightarrow$ Normal Operation $\rightarrow$ Warning $\rightarrow$ Emergency Fault.
      \end{itemize}
  \end{columns}
\end{frame}

% ===================== 16. CHAPTER SUMMARY =====================
\begin{frame}
  \slidetitle{Chapter I Summary \& Transition to Chapter II}
  \begin{itemize}
    \item \textbf{Embedded Co-Design:} Tight synergy between specialized MCU silicon, deterministic firmware, and physical world transducers.
    \item \textbf{The Core Triad:} Stringent resource boundaries, deterministic real-time deadlines, and 24/7 unassisted autonomous reliability.
    \item \textbf{The 5-Layer Stack:} Golden architectural standard separating silicon registers from business logic, ensuring portability and maintainability.
    \item \textbf{Escalating Edge Demands:} Modern IoT edge stations demand on-device ML/DSP, sub-microsecond interrupt latency, and silicon-enforced domain security.
    \item \textbf{Bridge to Chapter II:} Transitioning from legacy 8/16-bit or basic 32-bit cores to advanced vector-accelerated architectures.
  \end{itemize}
  \vspace{0.35cm}
  \textbf{$\Rightarrow$ Chapter II: Embedded Processor Architectures \& ARM Cortex-M85 on Renesas RA8M1.}
\end{frame}

\end{document}
"""

    tex_path = "main.tex"
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex_content)
    print(f"  Successfully generated: {tex_path}")

    # Package into Chuong_I_Beamer.zip
    zip_path = "Chuong_I_Beamer.zip"
    print(f"[4/4] Packaging LaTeX Beamer project into {zip_path}...")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(tex_path, arcname="main.tex")
        # Include all files in figs/
        for root, dirs, files in os.walk("figs"):
            for file in files:
                full_p = os.path.join(root, file)
                arc_p = os.path.join("figs", file)
                z.write(full_p, arcname=arc_p)
        # Also include images/
        for root, dirs, files in os.walk("images"):
            for file in files:
                if file.startswith("fig_"):
                    full_p = os.path.join(root, file)
                    arc_p = os.path.join("images", file)
                    z.write(full_p, arcname=arc_p)
    print(f"  Successfully created: {zip_path}")


# ==============================================================================
# MAIN EXECUTION
# ==============================================================================
if __name__ == "__main__":
    print("=" * 70)
    print("MINIMALIST ACADEMIC PRESENTATION GENERATOR (CHAPTER II STYLE)")
    print("=" * 70)
    prepare_assets()
    build_pptx()
    build_beamer()
    print("=" * 70)
    print("ALL OUTPUTS GENERATED SUCCESSFULLY!")
    print("  1. PowerPoint Presentation : Chuong_I_Overview_Embedded_Systems.pptx")
    print("  2. LaTeX Beamer Project    : Chuong_I_Beamer.zip")
    print("=" * 70)
