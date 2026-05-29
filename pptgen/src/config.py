from pathlib import Path

from pptx.util import Inches, Pt


APP_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = APP_ROOT / "outputs"
UPLOAD_DIR = APP_ROOT / "uploads"
TEMPLATE_PATH = APP_ROOT / "template" / "template.pptx"
OUTPUT_PATH = OUTPUT_DIR / "generated_ai_healthcheck.pptx"

SCORE_SLIDE_INDEXES = [1, 2]

HEADER_CLIENT_OFFSET = Inches(0.78)
HEADER_PEER_OFFSET = Inches(0.72)
HEADER_LOGO_GAP = Inches(0.08)
LEGEND_TEXT_GAP = Inches(0.06)
LEGEND_ITEM_GAP = Inches(0.16)
LEGEND_LINE_GAP = Inches(0.08)
LEGEND_RIGHT_MARGIN = Inches(0.10)
LEGEND_LOGO_SCALE = 1.28
LEGEND_MAX_LOGO_HEIGHT = Inches(0.28)
LEGEND_FONT_SIZE = Pt(10)

ROW_LOGO_MAX_HEIGHT_FACTOR = 0.84
ROW_LOGO_HORIZONTAL_PADDING = Inches(0.015)
ROW_LOGO_GAP = Inches(0.005)
ROW_LOGO_VERTICAL_GAP = Inches(0.01)
