"""Word constants kept local to avoid generated COM-module dependencies."""

WD_ALERTS_NONE = 0
WD_COLLAPSE_END = 0
WD_DO_NOT_SAVE_CHANGES = 0
WD_FIELD_ADDIN = 81
WD_LINE_STYLE_NONE = 0
WD_LINE_STYLE_SINGLE = 1
WD_LINE_WIDTH_050_PT = 4
WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_ALIGN_PARAGRAPH_RIGHT = 2

WD_BORDER_TOP = -1
WD_BORDER_LEFT = -2
WD_BORDER_BOTTOM = -3
WD_BORDER_RIGHT = -4
WD_BORDER_HORIZONTAL = -5
WD_BORDER_VERTICAL = -6

WD_COLOR_AUTOMATIC = -16777216

WD_SAVE_FORMAT_DOCUMENT_DEFAULT = 16

RPC_E_DISCONNECTED = -2147417848  # 0x80010108
RPC_S_CALLPENDING = -2147417835   # 0x80010115


def word_bgr(red: int, green: int, blue: int) -> int:
    """Convert RGB channels to the BGR integer used by Word COM."""

    for channel in (red, green, blue):
        if not 0 <= channel <= 255:
            raise ValueError("color channels must be between 0 and 255")
    return red | (green << 8) | (blue << 16)


WHITE = word_bgr(255, 255, 255)
VERY_LIGHT_GRAY = word_bgr(242, 242, 242)
BLACK = word_bgr(0, 0, 0)
