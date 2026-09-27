from __future__ import annotations

from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    BaseDocTemplate,
    Flowable,
    Frame,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "reports"
PDF_PATH = OUT / "Low_Light_Image_Enhancement_Project_Report.pdf"
DOCX_PATH = OUT / "Low_Light_Image_Enhancement_Project_Report.docx"
TODAY = date(2026, 9, 27).strftime("%d %B %Y")
GROUP_MEMBERS = (
    "Nasim Ahmed Kazi (20244111)",
    "Mujahid Kalanthar (20244510)",
    "Narendra Prajapat (20244110)",
    "Muhammed Sabab Rawoot (20244509)",
)
NAVY = colors.HexColor("#0d1b2a")
TEAL = colors.HexColor("#168a82")
PALE = colors.HexColor("#edf5f5")
INK = colors.HexColor("#202b35")
MUTED = colors.HexColor("#64717e")
LINE = colors.HexColor("#d7e0e5")


class ArchitectureDiagram(Flowable):
    def __init__(self, width=172 * mm, height=56 * mm):
        super().__init__()
        self.width = width
        self.height = height

    def draw(self):
        c = self.canv
        w, h = self.width, self.height
        c.setStrokeColor(LINE)
        c.setLineWidth(1)
        c.roundRect(1, 1, w - 2, h - 2, 5, fill=0, stroke=1)

        boxes = [
            (8, 23, 34, 24, "Browser UI", "Upload • preview"),
            (51, 23, 35, 24, "Flask API", "Validate • route"),
            (95, 23, 34, 24, "LiteRT", "Zero-DCE TFLite"),
            (138, 23, 26, 24, "Result", "PNG • compare"),
        ]
        scale = mm
        for x, y, bw, bh, title, detail in boxes:
            x *= scale
            y *= scale
            bw *= scale
            bh *= scale
            c.setFillColor(PALE)
            c.setStrokeColor(TEAL)
            c.roundRect(x, y, bw, bh, 4, fill=1, stroke=1)
            c.setFillColor(NAVY)
            c.setFont("Helvetica-Bold", 8)
            c.drawCentredString(x + bw / 2, y + bh * 0.62, title)
            c.setFillColor(MUTED)
            c.setFont("Helvetica", 6.4)
            c.drawCentredString(x + bw / 2, y + bh * 0.33, detail)
        c.setStrokeColor(TEAL)
        c.setLineWidth(1.5)
        for x1, x2 in ((42, 51), (86, 95), (129, 138)):
            y = 35 * scale
            c.line(x1 * scale, y, x2 * scale, y)
            c.line(x2 * scale - 2, y + 2, x2 * scale, y)
            c.line(x2 * scale - 2, y - 2, x2 * scale, y)
        c.setFillColor(MUTED)
        c.setFont("Helvetica-Oblique", 7)
        c.drawCentredString(w / 2, 9, "POST multipart image + strength                 PNG response")


def make_styles():
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "ReportTitle", parent=base["Title"], fontName="Helvetica-Bold",
            fontSize=25, leading=31, textColor=NAVY, alignment=TA_CENTER,
            spaceAfter=10,
        ),
        "subtitle": ParagraphStyle(
            "ReportSubtitle", parent=base["Normal"], fontName="Helvetica",
            fontSize=14, leading=20, textColor=TEAL, alignment=TA_CENTER,
        ),
        "h1": ParagraphStyle(
            "SectionHeading", parent=base["Heading1"], fontName="Helvetica-Bold",
            fontSize=18, leading=23, textColor=NAVY, spaceAfter=12,
        ),
        "h2": ParagraphStyle(
            "SubHeading", parent=base["Heading2"], fontName="Helvetica-Bold",
            fontSize=11, leading=14, textColor=TEAL, spaceBefore=7, spaceAfter=5,
        ),
        "body": ParagraphStyle(
            "ReportBody", parent=base["BodyText"], fontName="Helvetica",
            fontSize=9.4, leading=13.4, textColor=INK, spaceAfter=7,
        ),
        "small": ParagraphStyle(
            "ReportSmall", parent=base["BodyText"], fontName="Helvetica",
            fontSize=8, leading=10.5, textColor=MUTED, spaceAfter=4,
        ),
        "bullet": ParagraphStyle(
            "ReportBullet", parent=base["BodyText"], fontName="Helvetica",
            fontSize=9.2, leading=12.5, textColor=INK, leftIndent=13,
            firstLineIndent=-8, spaceAfter=4,
        ),
        "equation": ParagraphStyle(
            "Equation", parent=base["BodyText"], fontName="Helvetica-Bold",
            fontSize=12, leading=18, textColor=NAVY, alignment=TA_CENTER,
            backColor=PALE, borderColor=LINE, borderWidth=0.5,
            borderPadding=9, spaceBefore=5, spaceAfter=10,
        ),
        "center": ParagraphStyle(
            "Center", parent=base["BodyText"], fontName="Helvetica",
            fontSize=10, leading=15, textColor=INK, alignment=TA_CENTER,
        ),
    }


def P(text, style):
    return Paragraph(text, style)


def bullet(text, styles):
    return P(f"&bull; {text}", styles["bullet"])


def report_pages(styles):
    s = styles

    page1 = [
        Spacer(1, 25 * mm),
        P("PROJECT REPORT", s["subtitle"]),
        Spacer(1, 11 * mm),
        P("Low-Light Image<br/>Enhancement with Zero-DCE", s["title"]),
        Spacer(1, 4 * mm),
        P("LumaLift: an interactive image-enhancement application", s["subtitle"]),
        Spacer(1, 18 * mm),
        Table(
            [
                ["Submitted by", Paragraph("<br/>".join(GROUP_MEMBERS), s["small"])],
                ["Course / laboratory", "Digital Signal Processing Lab — Group 5"],
                ["Submitted to", "________________________________"],
                ["Department / institution", "________________________________"],
                ["Submission date", TODAY],
            ],
            colWidths=[48 * mm, 92 * mm],
            style=TableStyle([
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("TEXTCOLOR", (0, 0), (0, -1), TEAL),
                ("TEXTCOLOR", (1, 0), (1, -1), INK),
                ("BACKGROUND", (0, 0), (-1, -1), PALE),
                ("GRID", (0, 0), (-1, -1), 0.5, LINE),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ]),
        ),
        Spacer(1, 20 * mm),
        P("A practical implementation report on a browser-based interface and a lightweight inference service for enhancing low-illumination images.", s["center"]),
        Spacer(1, 18 * mm),
        P("Project repository: github.com/NasimAhmedKazi/Low-Light-Image-and-Video-Enhancement-", s["small"]),
        P("Deployed application: https://nasimahmedkazi-solid-disco.vercel.app", s["small"]),
        PageBreak(),
    ]

    page2 = [
        P("1. Abstract and Project Overview", s["h1"]),
        P("<b>Abstract.</b> Low-light photographs often contain insufficient visible detail, poor local contrast, and color shifts. This project presents a working image-enhancement application built around Zero-Reference Deep Curve Estimation (Zero-DCE), a learning-based approach designed to enhance low-light images without requiring paired low-light/normal-light training examples. The repository combines a browser interface, a Flask API, and a pre-trained TensorFlow Lite model executed through LiteRT. A user selects an image, previews the original, adjusts enhancement strength, submits the image for inference, compares the result with a draggable divider, and downloads the enhanced PNG.", s["body"]),
        P("The model is loaded lazily by the backend and reused for later requests. The API validates image data and settings, resizes input to the model's 512 × 512 inference size, invokes the TFLite model, resizes the enhanced image back to its original dimensions, and blends it with the input according to the selected strength. The same Flask application serves both the frontend and API routes, which enables a same-origin deployment on Vercel.", s["body"]),
        P("1.1 Problem Statement", s["h2"]),
        P("Images captured in dim environments can be difficult to interpret or present. Conventional global brightness adjustments may lift shadows while also washing out highlights, reducing contrast, or exaggerating sensor noise. The project explores a curve-estimation model that adapts image intensity through repeated pixel-wise transformations, then makes this process accessible through a simple upload-and-preview workflow.", s["body"]),
        P("1.2 Objectives", s["h2"]),
        bullet("Apply a pre-trained Zero-DCE model to user-provided images.", s),
        bullet("Provide a responsive interface for image selection, original preview, comparison, and download.", s),
        bullet("Expose inference through a small HTTP API with validation and explicit error responses.", s),
        bullet("Package a practical local run path and a hosted Vercel deployment.", s),
        P("1.3 Project Scope", s["h2"]),
        P("The demonstrated web inference path supports still images. The interface can preview video files, and the repository also contains a separate C++ video-processing program using OpenCV, TensorFlow C++, and FFmpeg; however, the browser's Enhance action does not process video. Training scripts are included in the repository, but this report evaluates inference with the bundled model rather than claiming new training or a controlled benchmark.", s["body"]),
        PageBreak(),
    ]

    page3 = [
        P("2. Technical Background and Algorithm", s["h1"]),
        P("2.1 Zero-Reference Deep Curve Estimation", s["h2"]),
        P("Zero-DCE formulates low-light enhancement as the estimation of image-specific curves. A convolutional network predicts adjustment maps from the input image. These maps are applied iteratively to the image, allowing different pixels and regions to receive different amounts of adjustment. The method is called zero-reference because its training objective does not require paired reference images as direct supervision; it uses non-reference objectives described in the original research paper [1].", s["body"]),
        P("A simplified pixel update for one iteration is:", s["body"]),
        P("x<sub>t+1</sub> = x<sub>t</sub> + r<sub>t</sub> &middot; (x<sub>t</sub><sup>2</sup> &minus; x<sub>t</sub>)", s["equation"]),
        P("Here, x<sub>t</sub> is the current normalized pixel value and r<sub>t</sub> is the estimated curve value for that iteration. The trained network predicts multiple curve maps; the repository's DCENet implementation uses eight curve iterations and a final tanh activation by default. Repeated updates can lift dark pixel values while retaining spatially varying adjustment behavior.", s["body"]),
        P("2.2 Model in This Application", s["h2"]),
        P("The web service uses the repository artifact Zero-DCE/model_trained/model.tflite (about 95 KB). The interpreter is created with LiteRT and two CPU threads. The endpoint sends a single RGB tensor of shape 1 × 512 × 512 × 3, normalized to the [0, 1] range. The application selects the model output named StatefulPartitionedCall:1, which the bundled model exposes as a three-channel image. This output-selection detail is necessary because the model returns additional tensors for intermediate/curve outputs.", s["body"]),
        P("2.3 Strength Control", s["h2"]),
        P("The interface's 0–100% strength value is implemented as an output blend, not as a change to the learned curve maps. At 0%, the response is the unmodified input; at 100%, it is the model result; intermediate values linearly mix both images. In normalized pixel notation: y = (1 &minus; s)x + s f(x), where s is the strength fraction and f(x) is the model output. The default setting is 72%.", s["body"]),
        P("This design makes the strength control easy to interpret and avoids changing the model itself. It is an application-level control; it should not be confused with a separately trained model configuration.", s["body"]),
        P("2.4 Preprocessing and Output", s["h2"]),
        P("The API applies EXIF orientation correction and converts inputs to RGB. It scales to the fixed model dimensions with bilinear interpolation, runs inference, scales the output back to the original width and height, blends according to strength, clamps values to the valid display range, and encodes a PNG response.", s["body"]),
        PageBreak(),
    ]

    page4 = [
        P("3. System Architecture", s["h1"]),
        P("The system is organized as a browser client, a Flask web/API layer, and a local inference runtime. The client handles file selection and original preview in the browser; the model runs on the server side of the Flask process. In the hosted configuration, Vercel packages the Flask function, frontend assets, and TFLite artifact together.", s["body"]),
        ArchitectureDiagram(),
        P("Figure 1. Image request and response path for web inference.", s["small"]),
        P("3.1 Main Components", s["h2"]),
        bullet("<b>Frontend (frontend/).</b> HTML, CSS, and JavaScript provide the LumaLift workspace. The browser creates a temporary object URL for the selected file, presents the original preview, submits a multipart request, displays the returned PNG, and enables download.", s),
        bullet("<b>Flask application (backend/app.py).</b> Serves the frontend, reports model readiness through GET /api/health, and accepts images through POST /api/enhance.", s),
        bullet("<b>LiteRT interpreter.</b> Loads the bundled TFLite file once per process, resizes the dynamic input to 512 × 512, invokes the model, and retrieves the enhanced-image output.", s),
        bullet("<b>Vercel configuration.</b> The root app.py exports the Flask app. vercel.json includes the frontend and model files in the function bundle and sets a 60-second maximum duration. Python 3.12 uses LiteRT 2.2; local Python 3.11 uses LiteRT 1.0.1.", s),
        P("3.2 Request Sequence", s["h2"]),
        P("1) The user selects an image. 2) JavaScript validates its browser-reported media type and renders the original locally. 3) On Enhance, the browser sends the file and strength as multipart form data. 4) Flask validates size, decodability, dimensions, and strength. 5) LiteRT runs inference. 6) Flask returns a PNG. 7) The browser displays the enhanced output beside the original and enables the download control.", s["body"]),
        P("A second, separate repository path handles video: the C++ executable reads frames through OpenCV, applies the TensorFlow C++ model, and uses FFmpeg-related audio handling. It is not part of the deployed Flask request path.", s["body"]),
        PageBreak(),
    ]

    page5 = [
        P("4. Implementation Details", s["h1"]),
        P("4.1 Frontend Interaction", s["h2"]),
        P("The page has image/video selection tabs, a drag-and-drop area, a file picker, an enhancement-strength slider, a before/after comparison surface, status messaging, and download behavior. Image selection updates the Original pane immediately using a browser object URL. Image dimensions and file name are reflected in the interface. The enhanced pane starts empty rather than showing a misleading hardcoded sample. Video files are previewable in the browser, but an attempt to enhance a video returns an explanatory message because the image model API accepts still images only.", s["body"]),
        P("On image enhancement, the frontend creates FormData fields named file and strength and posts them to /api/enhance. It handles JSON error responses, disables the button while processing, and ignores stale results when a newer image has been selected. Successful PNG responses are displayed via an object URL; the download action saves a file named <i>&lt;original-name&gt;-enhanced.png</i>.", s["body"]),
        P("4.2 Backend Validation", s["h2"]),
        P("The API enforces a 25 MB request limit and returns a JSON error for oversized requests. It rejects missing files, non-numeric/out-of-range strength settings, and files that Pillow cannot decode as images. It applies EXIF transpose before RGB conversion. Images over 12 megapixels are rejected to reduce excessive memory use. Inference is guarded by a lock because the LiteRT interpreter is reused and should not receive concurrent tensor mutations.", s["body"]),
        P("4.3 Model Lifecycle", s["h2"]),
        P("The model is loaded on the first health check or enhancement request and cached in process memory. This avoids reloading the small model on every request. The health endpoint returns ready=true only after the model has loaded and its input/output tensors have been inspected. A missing model or runtime issue is reported as not ready instead of showing a false success state.", s["body"]),
        P("4.4 Runtime Compatibility", s["h2"]),
        P("An initial attempt to use TensorFlow directly was not reliable on the local machine because the installed TensorFlow wheel required CPU AVX instructions that were unavailable in that runtime. The repository's TFLite artifact was then tested using LiteRT, which successfully executed on CPU. This runtime choice also reduced deployment complexity: the Linux LiteRT wheel is much smaller than the TensorFlow wheel. Dependency markers select the appropriate LiteRT version for local Python 3.11 and hosted Python 3.12.", s["body"]),
        PageBreak(),
    ]

    page6 = [
        P("5. Testing and Results", s["h1"]),
        P("Testing focused on the user-visible inference path and the actual model artifact. Results below describe checks performed during implementation; they are not a statistical image-quality benchmark.", s["body"]),
        Table(
            [
                ["Test", "Observed result"],
                ["Model readiness: GET /api/health", "HTTP 200; {\"ready\": true, \"model\": \"Zero-DCE\"}"],
                ["Image inference: sample dragon_inp.jpg", "HTTP 200; image/png returned"],
                ["Output dimensions", "2068 × 1168, matching input dimensions"],
                ["Output artifact", "Valid 8-bit RGB PNG, about 1.30 MB"],
                ["Strength validation", "Invalid values rejected with HTTP 400"],
                ["Invalid image validation", "Undecodable input rejected with HTTP 400"],
                ["Frontend flow", "Selected file posted; returned blob displayed; download enabled"],
                ["Production deployment", "Health endpoint ready; real inference returned PNG"],
            ],
            colWidths=[59 * mm, 107 * mm],
            repeatRows=1,
            style=TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (1, 1), (1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 8.1),
                ("LEADING", (0, 0), (-1, -1), 10),
                ("TEXTCOLOR", (0, 1), (-1, -1), INK),
                ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, PALE]),
                ("GRID", (0, 0), (-1, -1), 0.4, LINE),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]),
        ),
        Spacer(1, 4 * mm),
        P("5.1 Sample Inference Observation", s["h2"]),
        P("A real request to the deployed API using the repository's sample image completed successfully and returned a 2068 × 1168 PNG. The observed request took about 11.1 seconds after the model health endpoint had been called. A local request was observed at about 29.5 seconds. These are single-run observations on different environments, not controlled timing measurements; actual latency depends on cold starts, available CPU, image dimensions, and platform load.", s["body"]),
        P("5.2 Interpretation", s["h2"]),
        P("The tests establish that the application can accept an image, invoke the bundled learned model, and return a displayable output with original dimensions. They do not establish a universal perceptual improvement score. A rigorous quality evaluation would require a representative test set, defined baselines, and objective metrics such as NIQE or LOE, alongside visual inspection; no such aggregate results are claimed here.", s["body"]),
        PageBreak(),
    ]

    page7 = [
        P("6. Deployment and User Guide", s["h1"]),
        P("6.1 Hosted Application", s["h2"]),
        P("The production application is available at <link href=\"https://nasimahmedkazi-solid-disco.vercel.app\" color=\"#168a82\">https://nasimahmedkazi-solid-disco.vercel.app</link>. The Vercel deployment packages the Flask entrypoint, browser assets, and TFLite model in the same function. The deployment health endpoint and an actual image enhancement request were verified after publication. Inference is subject to the configured 60-second function duration.", s["body"]),
        P("6.2 Local Setup", s["h2"]),
        P("From the repository root, use Python 3.11 and run:", s["body"]),
        P("python3.11 -m venv .venv<br/>source .venv/bin/activate<br/>python -m pip install -r backend/requirements.txt<br/>python backend/app.py", s["equation"]),
        P("Open http://localhost:5001. The Flask app serves the UI and endpoints from the same origin. Keep the process open while using the application. The first inference initializes the model, so it can take longer than subsequent requests.", s["body"]),
        P("6.3 User Procedure", s["h2"]),
        bullet("Open the deployed URL or local application.", s),
        bullet("Choose the Image tab and select or drop a supported image.", s),
        bullet("Confirm the original preview and displayed image dimensions.", s),
        bullet("Set enhancement strength; use 72% as the default starting point.", s),
        bullet("Click Enhance and wait for the model response.", s),
        bullet("Drag the comparison divider to inspect original and enhanced areas.", s),
        bullet("Click Download to save the enhanced PNG.", s),
        P("For a video, the browser provides preview only. The web enhancement button intentionally reports that video enhancement is unsupported through this image inference endpoint. The repository's separate C++ workflow is intended for video processing and requires its own native dependencies.", s["body"]),
        P("6.4 Repository and Deployment Configuration", s["h2"]),
        P("The Flask app is exported from the root app.py for Vercel's Python runtime. requirements.txt selects Flask, Pillow, NumPy, and LiteRT. The .python-version file selects Python 3.12 for Vercel; backend/requirements.txt uses environment markers so local Python 3.11 installs LiteRT 1.0.1 while Python 3.12 installs LiteRT 2.2. vercel.json includes only frontend assets and the model file in the function bundle.", s["body"]),
        PageBreak(),
    ]

    page8 = [
        P("7. Limitations, Future Work, and Conclusion", s["h1"]),
        P("7.1 Current Limitations", s["h2"]),
        bullet("Only still-image inference is available through the web interface; video enhancement is not wired to the Flask API.", s),
        bullet("Input is resized to 512 × 512 for model execution before resizing back. This can soften fine details and alter aspect ratio during inference; the displayed/downloaded result is restored to original dimensions but is not a native-resolution model pass.", s),
        bullet("The API caps uploads at 25 MB and 12 megapixels. Larger images require tiling or a higher resource budget.", s),
        bullet("The 0–100 strength setting blends the model output with the original; it does not re-run model inference at different curve parameters.", s),
        bullet("Single-request latency varies and can approach the serverless function time limit on slower or cold instances.", s),
        bullet("The bundled model is small and convenient for demonstration, but this report does not establish that it is the best-performing model for every scene or camera.", s),
        bullet("The repository identifies the Zero-DCE paper/model use as non-commercial; confirm applicable license terms before reuse outside an academic demonstration.", s),
        P("7.2 Recommended Next Steps", s["h2"]),
        bullet("Add a video-processing API with frame batching, progress reporting, audio preservation, and output-size limits.", s),
        bullet("Evaluate a diverse image set and report visual comparisons plus defined no-reference image-quality metrics.", s),
        bullet("Add automated browser/API tests to CI and retain regression images for model-output checks.", s),
        bullet("Consider aspect-ratio-preserving padding, tiling, or configurable model resolution for high-resolution input.", s),
        bullet("Move model inference to a dedicated GPU/CPU service if traffic or latency exceeds serverless limits.", s),
        P("7.3 Conclusion", s["h2"]),
        P("This project turns a research-based low-light image enhancement model into an end-to-end application. The browser lets users select an image and inspect an original preview; the Flask service validates the request, runs the bundled Zero-DCE TFLite model through LiteRT, and returns an enhanced PNG for comparison and download. The model endpoint was verified both locally and in the production Vercel deployment. The implementation demonstrates a practical deployment path while clearly identifying still-image scope, resource limits, and the difference between a verified inference flow and a full video enhancement product.", s["body"]),
        P("References", s["h2"]),
        P("[1] C. Guo, C. Li, J. Guo, C. C. Loy, J. Hou, S. Kwong, and R. Cong, “Zero-Reference Deep Curve Estimation for Low-Light Image Enhancement,” <i>Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR)</i>, 2020, pp. 1780–1789. doi: 10.1109/CVPR42600.2020.00185. https://openaccess.thecvf.com/content_CVPR_2020/papers/Guo_Zero-Reference_Deep_Curve_Estimation_for_Low-Light_Image_Enhancement_CVPR_2020_paper.pdf", s["small"]),
        P("[2] Project source repository, “Low-Light-Image-and-Video-Enhancement-.” https://github.com/NasimAhmedKazi/Low-Light-Image-and-Video-Enhancement-", s["small"]),
        P("[3] Google AI Edge, “LiteRT / TensorFlow Lite Interpreter,” https://ai.google.dev/edge/litert", s["small"]),
        P("[4] Vercel, “Python Runtime and Functions,” https://vercel.com/docs/functions/runtimes/python", s["small"]),
        PageBreak(),
    ]
    return [page1, page2, page3, page4, page5, page6, page7, page8]


def draw_page(canvas, doc):
    canvas.saveState()
    page_num = canvas.getPageNumber()
    width, height = A4
    if page_num > 1:
        canvas.setStrokeColor(LINE)
        canvas.setLineWidth(0.5)
        canvas.line(20 * mm, height - 15 * mm, width - 20 * mm, height - 15 * mm)
        canvas.setFillColor(MUTED)
        canvas.setFont("Helvetica", 7.5)
        canvas.drawString(20 * mm, height - 12 * mm, "LOW-LIGHT IMAGE ENHANCEMENT • ZERO-DCE")
        canvas.drawRightString(width - 20 * mm, 12 * mm, f"Page {page_num} of 8")
        canvas.line(20 * mm, 16 * mm, width - 20 * mm, 16 * mm)
    canvas.restoreState()


def build_pdf(styles):
    doc = BaseDocTemplate(
        str(PDF_PATH), pagesize=A4,
        leftMargin=22 * mm, rightMargin=22 * mm,
        topMargin=21 * mm, bottomMargin=22 * mm,
        title="Low-Light Image Enhancement with Zero-DCE",
        author="Nasim Ahmed Kazi",
        subject="Laboratory project report",
    )
    frame = Frame(
        doc.leftMargin, doc.bottomMargin, doc.width, doc.height,
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
        id="normal",
    )
    doc.addPageTemplates([PageTemplate(id="report", frames=frame, onPage=draw_page)])
    story = []
    for page in report_pages(styles):
        story.extend(page)
    doc.build(story)


def set_cell_shading(cell, fill):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def configure_docx(doc):
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.78)
    section.bottom_margin = Inches(0.72)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(9.5)
    normal.font.color.rgb = RGBColor(32, 43, 53)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.08
    for style_name, size in (("Heading 1", 18), ("Heading 2", 11)):
        style = doc.styles[style_name]
        style.font.name = "Arial"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(13, 55, 70) if style_name == "Heading 1" else RGBColor(22, 138, 130)
    header = section.header.paragraphs[0]
    header.text = "LOW-LIGHT IMAGE ENHANCEMENT  |  ZERO-DCE"
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header.runs[0].font.size = Pt(8)
    header.runs[0].font.color.rgb = RGBColor(100, 113, 126)
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("Laboratory Project Report  •  ")
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(100, 113, 126)
    fld = run._r
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    footer._p.append(field)


def docx_add_page(doc, page_index, styles):
    if page_index:
        doc.add_page_break()
    if page_index == 0:
        for _ in range(4):
            doc.add_paragraph()
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run("PROJECT REPORT")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(22, 138, 130)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(25)
        r = p.add_run("Low-Light Image\nEnhancement with Zero-DCE")
        r.bold = True
        r.font.size = Pt(27)
        r.font.color.rgb = RGBColor(13, 27, 42)
        p = doc.add_paragraph("LumaLift: an interactive image-enhancement application")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.size = Pt(14)
        p.runs[0].font.color.rgb = RGBColor(22, 138, 130)
        for _ in range(2):
            doc.add_paragraph()
        table = doc.add_table(rows=5, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"
        fields = [
            ("Submitted by", "\n".join(GROUP_MEMBERS)),
            ("Course / laboratory", "Digital Signal Processing Lab — Group 5"),
            ("Submitted to", "________________________________"),
            ("Department / institution", "________________________________"),
            ("Submission date", TODAY),
        ]
        for row, values in zip(table.rows, fields):
            for cell, value in zip(row.cells, values):
                cell.text = value
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                set_cell_shading(cell, "EDF5F5")
                for run in cell.paragraphs[0].runs:
                    run.font.size = Pt(10)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(28)
        p.add_run("A practical implementation report on a browser-based interface and\n"
                  "a lightweight inference service for low-illumination images.")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(35)
        r = p.add_run("Repository: github.com/NasimAhmedKazi/Low-Light-Image-and-Video-Enhancement-\n"
                      "Live application: https://nasimahmedkazi-solid-disco.vercel.app")
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(100, 113, 126)
        return

    titles = [
        "1. Abstract and Project Overview",
        "2. Technical Background and Algorithm",
        "3. System Architecture",
        "4. Implementation Details",
        "5. Testing and Results",
        "6. Deployment and User Guide",
        "7. Limitations, Future Work, and Conclusion",
    ]
    doc.add_heading(titles[page_index - 1], level=1)
    if page_index == 1:
        doc.add_heading("Abstract", level=2)
        doc.add_paragraph("Low-light photographs often contain insufficient visible detail, poor local contrast, and color shifts. This project presents a working image-enhancement application built around Zero-Reference Deep Curve Estimation (Zero-DCE), an approach designed to enhance low-light images without paired low-light/normal-light training examples. The repository combines a browser interface, a Flask API, and a pre-trained TensorFlow Lite model executed through LiteRT. Users select an image, preview the original, adjust enhancement strength, submit it for inference, compare the result with a draggable divider, and download the enhanced PNG.")
        doc.add_paragraph("The backend validates the image and settings, resizes input to 512 × 512, invokes the TFLite model, resizes the result to the original dimensions, blends it according to strength, and returns a PNG. The same Flask application serves the frontend and API, enabling same-origin deployment on Vercel.")
        doc.add_heading("Problem statement", level=2)
        doc.add_paragraph("Images captured in dim environments can be hard to interpret. Global brightness changes can wash out highlights, lower contrast, or exaggerate noise. This project applies a learned curve-estimation model and provides an accessible upload-and-preview workflow.")
        doc.add_heading("Objectives", level=2)
        for item in [
            "Apply the bundled Zero-DCE model to user-provided images.",
            "Provide selection, preview, comparison, and download controls.",
            "Expose inference through a validated HTTP API.",
            "Support a local run path and a hosted Vercel deployment.",
        ]:
            doc.add_paragraph(item, style="List Bullet")
        doc.add_heading("Scope", level=2)
        doc.add_paragraph("The web inference path processes still images. Video preview is available in the browser, and the repository contains a separate C++ video workflow, but browser video enhancement is not connected. This report evaluates the bundled model; it does not claim new model training.")
    elif page_index == 2:
        doc.add_heading("Zero-Reference Deep Curve Estimation", level=2)
        doc.add_paragraph("Zero-DCE estimates image-specific curves using a convolutional network. The predicted adjustment maps are applied iteratively to pixels, enabling spatially varying intensity adjustments. Zero-reference describes its learning objective: paired reference images are not required as direct supervision; non-reference objectives are described in the source paper [1].")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("x(t+1) = x(t) + r(t) · (x(t)² − x(t))")
        r.bold = True
        r.font.size = Pt(14)
        doc.add_paragraph("Here x(t) is the normalized pixel value and r(t) is the estimated curve value. The project model uses eight curve iterations by default.")
        doc.add_heading("Model and strength", level=2)
        doc.add_paragraph("The web service uses Zero-DCE/model_trained/model.tflite (about 95 KB) through LiteRT. It sends one RGB tensor with shape 1 × 512 × 512 × 3 and selects the three-channel output StatefulPartitionedCall:1. The strength slider is a post-inference blend, y = (1 − s)x + s f(x), where s ranges from 0 to 1. Thus, 0% returns the original and 100% returns the model output.")
        doc.add_heading("Preprocessing", level=2)
        doc.add_paragraph("The backend corrects EXIF orientation, converts to RGB, resizes to the model input size with bilinear interpolation, performs inference, resizes back to original dimensions, clamps pixel values, blends, and encodes PNG.")
    elif page_index == 3:
        doc.add_paragraph("The system is organized as a browser client, Flask web/API layer, and local inference runtime. The browser renders the original locally, then sends a selected still image to the API when the user requests enhancement.")
        table = doc.add_table(rows=1, cols=4)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"
        for cell, text in zip(table.rows[0].cells, ["Browser UI", "Flask API", "LiteRT model", "PNG result"]):
            cell.text = text
            set_cell_shading(cell, "DCEFED")
            for run in cell.paragraphs[0].runs:
                run.bold = True
        doc.add_paragraph("Figure 1. Image upload → validation and routing → TFLite inference → enhanced PNG response.", style=None)
        doc.add_heading("Components", level=2)
        for title, body in [
            ("Frontend", "HTML, CSS, and JavaScript implement file selection, original preview, request handling, result comparison, and download."),
            ("Flask API", "backend/app.py serves the interface, GET /api/health, and POST /api/enhance."),
            ("LiteRT", "Loads and reuses the bundled model, sizes its input to 512 × 512, invokes inference, and retrieves the image output."),
            ("Vercel", "The root app.py exports the Flask app. vercel.json bundles the frontend and TFLite file and allows a 60-second request duration."),
        ]:
            doc.add_paragraph(f"{title}: {body}")
        doc.add_heading("Video path", level=2)
        doc.add_paragraph("The C++ video program is a separate path based on OpenCV frame reading, TensorFlow C++ inference, and FFmpeg audio handling. It is not used by the Flask endpoint.")
    elif page_index == 4:
        doc.add_heading("Frontend interaction", level=2)
        doc.add_paragraph("Image selection or drag-and-drop creates a browser object URL and immediately updates the Original pane. The interface reports the selected file and dimensions. When Enhance is clicked, JavaScript sends the file and strength as multipart form data, disables the control during processing, handles errors, and displays the returned PNG. The download action saves a named PNG. Video can be previewed, but the image endpoint explicitly reports that it does not enhance video.")
        doc.add_heading("Backend validation", level=2)
        doc.add_paragraph("Uploads are limited to 25 MB and 12 megapixels. Missing files, invalid strength values, and undecodable images are rejected. EXIF orientation is corrected before RGB conversion. Inference is locked while the reusable LiteRT interpreter's tensors are changed.")
        doc.add_heading("Runtime compatibility", level=2)
        doc.add_paragraph("TensorFlow initially failed on the development machine because its binary required CPU AVX instructions unavailable in that environment. The existing TFLite model ran successfully through LiteRT on CPU. Dependency markers select LiteRT 1.0.1 for local Python 3.11 and 2.2 for hosted Python 3.12.")
        doc.add_heading("Model lifecycle", level=2)
        doc.add_paragraph("The model loads on the first health or enhancement request, is validated for input/output tensors, and remains cached in process memory. Health is reported ready only after successful initialization.")
    elif page_index == 5:
        doc.add_paragraph("Tests exercised the user-facing inference route and the bundled model. The following are observed results, not an image-quality benchmark.")
        table = doc.add_table(rows=1, cols=2)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for cell, text in zip(table.rows[0].cells, ("Test", "Observed result")):
            cell.text = text
            set_cell_shading(cell, "0D1B2A")
            for run in cell.paragraphs[0].runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.bold = True
        rows = [
            ("Health check", "HTTP 200; model reports ready"),
            ("Sample image inference", "HTTP 200; image/png"),
            ("Output size", "2068 × 1168; same dimensions as input"),
            ("Output file", "Valid RGB PNG, about 1.30 MB"),
            ("Invalid strength/image", "Rejected with HTTP 400"),
            ("Frontend flow", "Selected file posted; response shown; download enabled"),
            ("Production endpoint", "Real model response verified"),
        ]
        for left, right in rows:
            cells = table.add_row().cells
            cells[0].text, cells[1].text = left, right
        doc.add_heading("Timing observation", level=2)
        doc.add_paragraph("A deployed sample-image request took about 11.1 seconds after a health check; a local request was observed at about 29.5 seconds. These are single runs on different environments, not a controlled benchmark. Cold starts, CPU availability, input dimensions, and platform load affect latency.")
        doc.add_heading("Interpretation", level=2)
        doc.add_paragraph("The tests show that the app accepts an image, invokes the actual model, and returns an image at the original dimensions. No aggregate perceptual-quality score is claimed. A formal evaluation would require a representative dataset, a defined baseline, and chosen no-reference metrics alongside visual inspection.")
    elif page_index == 6:
        doc.add_heading("Hosted application", level=2)
        doc.add_paragraph("Production URL: https://nasimahmedkazi-solid-disco.vercel.app. The production health endpoint and an actual image enhancement request were verified. The app bundles the Flask function, frontend assets, and TFLite model. Serverless execution is subject to a 60-second limit.")
        doc.add_heading("Run locally", level=2)
        p = doc.add_paragraph()
        r = p.add_run("python3.11 -m venv .venv\nsource .venv/bin/activate\npython -m pip install -r backend/requirements.txt\npython backend/app.py")
        r.font.name = "Courier New"
        r.font.size = Pt(9)
        doc.add_paragraph("Open http://localhost:5001. Keep the backend process running. First inference may take longer while the model initializes.")
        doc.add_heading("User procedure", level=2)
        for item in [
            "Open the deployed or local application.",
            "Select or drop a still image.",
            "Review the original image and displayed dimensions.",
            "Set the strength (default 72%) and click Enhance.",
            "Drag the comparison divider and download the PNG.",
        ]:
            doc.add_paragraph(item, style="List Number")
        doc.add_paragraph("The browser supports video preview only; use the separate C++ workflow for video enhancement.")
        doc.add_heading("Deployment configuration", level=2)
        doc.add_paragraph("Vercel detects the Flask entrypoint from root app.py. requirements.txt defines Python dependencies, .python-version selects Python 3.12, and vercel.json includes required frontend/model assets. LiteRT version markers handle local Python 3.11 versus hosted Python 3.12.")
    elif page_index == 7:
        doc.add_heading("Limitations", level=2)
        for item in [
            "Web inference is still-image only; video enhancement is not integrated with the Flask API.",
            "Inference resizes to 512 × 512 before restoring original dimensions, which can soften detail and temporarily change aspect ratio.",
            "Uploads are limited to 25 MB and 12 megapixels.",
            "Strength blends original and enhanced images rather than changing the model's curve maps.",
            "Latency varies and can approach the serverless timeout on slower or cold instances.",
            "No controlled image-quality benchmark has been performed.",
            "The repository describes the Zero-DCE paper/model use as non-commercial; confirm applicable terms before reuse.",
        ]:
            doc.add_paragraph(item, style="List Bullet")
        doc.add_heading("Future work", level=2)
        for item in [
            "Integrate frame batching, progress reporting, and audio preservation for video.",
            "Evaluate a diverse image set and publish defined quality metrics.",
            "Add automated UI/API regression tests.",
            "Explore aspect-ratio-preserving padding or tiled high-resolution inference.",
            "Use dedicated inference hosting if latency or scale exceeds serverless limits.",
        ]:
            doc.add_paragraph(item, style="List Bullet")
        doc.add_heading("Conclusion", level=2)
        doc.add_paragraph("This project packages a research-based low-light enhancement model into an end-to-end application. The browser previews user input, Flask validates and routes requests, LiteRT runs the bundled Zero-DCE model, and the enhanced PNG is available for comparison and download. Real inference was verified locally and in the production Vercel deployment. The report distinguishes the implemented image workflow from video functionality that remains separate and identifies resource and evaluation limits.")
        doc.add_heading("References", level=2)
        references = [
            "[1] C. Guo et al., “Zero-Reference Deep Curve Estimation for Low-Light Image Enhancement,” CVPR, 2020, pp. 1780–1789. doi: 10.1109/CVPR42600.2020.00185.",
            "[2] Project repository: https://github.com/NasimAhmedKazi/Low-Light-Image-and-Video-Enhancement-",
            "[3] Google AI Edge, LiteRT: https://ai.google.dev/edge/litert",
            "[4] Vercel Python Runtime: https://vercel.com/docs/functions/runtimes/python",
        ]
        for ref in references:
            p = doc.add_paragraph(ref)
            p.paragraph_format.space_after = Pt(3)
            for run in p.runs:
                run.font.size = Pt(8)


def build_docx(styles):
    doc = Document()
    configure_docx(doc)
    for index in range(8):
        docx_add_page(doc, index, styles)
    doc.core_properties.title = "Low-Light Image Enhancement with Zero-DCE"
    doc.core_properties.author = "Digital Signal Processing Lab Group 5"
    doc.core_properties.subject = "Laboratory Project Report"
    doc.save(DOCX_PATH)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    styles = make_styles()
    build_pdf(styles)
    build_docx(styles)
    print(PDF_PATH)
    print(DOCX_PATH)


if __name__ == "__main__":
    main()
