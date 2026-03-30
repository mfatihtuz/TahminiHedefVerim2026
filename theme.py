"""
Modern Enterprise Theme System
Azure/Office 365 + SAP Fiori + Material Design 3 inspired
Professional corporate design with muted colors and high readability
"""


class EnterpriseTheme:
    """Modern kurumsal tema - Azure/Office 365 tarzı"""

    SEPARATOR_STRONG = "#B0BEC5"  # kalın ayraç (A/C arası)
    SEPARATOR_WEAK = "#ECEFF1"  # ince ayraç (alt grup arası)
    A_GROUP_STRIPE = "#F3E5F5"  # A grupları için çok hafif arka plan
    C_GROUP_STRIPE = "#E3F2FD"  # C grupları için çok hafif arka plan

    # Temel renkler - muted ve profesyonel
    PRIMARY = "#0078D4"  # Azure blue
    PRIMARY_HOVER = "#106EBE"
    PRIMARY_PRESSED = "#005A9E"
    PRIMARY_LIGHT = "#E1F5FE"

    SECONDARY = "#50E6FF"  # Açık mavi vurgu
    ACCENT = "#8764B8"  # Mor vurgu (A hatları için)

    # Neutral palette - göz yormayan
    BACKGROUND = "#FAFAFA"  # Hafif gri arkaplan
    SURFACE = "#FFFFFF"  # Beyaz kartlar
    SURFACE_HOVER = "#F5F5F5"

    # Borders and dividers
    BORDER = "#E0E0E0"
    BORDER_LIGHT = "#F0F0F0"
    DIVIDER = "#DADADA"

    # Text colors - yüksek kontrast
    TEXT_PRIMARY = "#323130"
    TEXT_SECONDARY = "#605E5C"
    TEXT_DISABLED = "#A19F9D"
    TEXT_ON_PRIMARY = "#FFFFFF"

    # Status colors
    SUCCESS = "#107C10"
    WARNING = "#F7630C"
    ERROR = "#D13438"
    INFO = "#0078D4"

    # Production line colors (muted)
    A_LINE_BASE = "#E3D5F5"  # Açık lavanta (A hatları)
    A_LINE_ACCENT = "#B39DDB"
    C_LINE_BASE = "#BBDEFB"  # Açık mavi (C hatları)
    C_LINE_ACCENT = "#64B5F6"

    # Shadows (depth)
    SHADOW_1 = "0 1px 3px rgba(0,0,0,0.08)"
    SHADOW_2 = "0 2px 6px rgba(0,0,0,0.10)"
    SHADOW_3 = "0 4px 12px rgba(0,0,0,0.12)"
    SHADOW_4 = "0 8px 24px rgba(0,0,0,0.15)"

    # Typography
    FONT_FAMILY = "'Segoe UI', -apple-system, BlinkMacSystemFont, 'Roboto', sans-serif"
    FONT_SIZE_SMALL = "11px"
    FONT_SIZE_NORMAL = "13px"
    FONT_SIZE_MEDIUM = "14px"
    FONT_SIZE_LARGE = "16px"
    FONT_SIZE_XLARGE = "20px"
    FONT_SIZE_TITLE = "24px"

    # Spacing (8px grid)
    SPACING_XS = "4px"
    SPACING_SM = "8px"
    SPACING_MD = "16px"
    SPACING_LG = "24px"
    SPACING_XL = "32px"

    # Border radius
    RADIUS_SM = "4px"
    RADIUS_MD = "6px"
    RADIUS_LG = "8px"
    RADIUS_XL = "12px"

    @staticmethod
    def get_main_stylesheet():
        """Ana pencere için stylesheet"""
        return f"""
        * {{
            font-family: {EnterpriseTheme.FONT_FAMILY};
        }}

        QMainWindow {{
            background-color: {EnterpriseTheme.BACKGROUND};
        }}

        QWidget {{
            background-color: {EnterpriseTheme.BACKGROUND};
            color: {EnterpriseTheme.TEXT_PRIMARY};
            font-size: {EnterpriseTheme.FONT_SIZE_NORMAL};
        }}

        /* MenuBar - Gradient header */
        QMenuBar {{
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 {EnterpriseTheme.PRIMARY},
                stop:1 {EnterpriseTheme.PRIMARY_HOVER});
            color: {EnterpriseTheme.TEXT_ON_PRIMARY};
            padding: 8px 16px;
            border: none;
            font-weight: 600;
            font-size: {EnterpriseTheme.FONT_SIZE_MEDIUM};
        }}

        QMenuBar::item {{
            background: transparent;
            padding: 8px 16px;
            border-radius: {EnterpriseTheme.RADIUS_MD};
            margin: 0 2px;
        }}

        QMenuBar::item:selected {{
            background: rgba(255, 255, 255, 0.15);
        }}

        QMenuBar::item:pressed {{
            background: rgba(255, 255, 255, 0.25);
        }}

        /* Menu dropdown */
        QMenu {{
            background-color: {EnterpriseTheme.SURFACE};
            border: 1px solid {EnterpriseTheme.BORDER};
            border-radius: {EnterpriseTheme.RADIUS_MD};
            padding: 8px 0;
            box-shadow: {EnterpriseTheme.SHADOW_3};
        }}

        QMenu::item {{
            padding: 10px 32px 10px 16px;
            color: {EnterpriseTheme.TEXT_PRIMARY};
        }}

        QMenu::item:selected {{
            background-color: {EnterpriseTheme.PRIMARY_LIGHT};
            color: {EnterpriseTheme.PRIMARY};
        }}

        QMenu::separator {{
            height: 1px;
            background: {EnterpriseTheme.BORDER_LIGHT};
            margin: 8px 0;
        }}

        /* Status bar */
        QStatusBar {{
            background: {EnterpriseTheme.SURFACE};
            color: {EnterpriseTheme.TEXT_SECONDARY};
            border-top: 1px solid {EnterpriseTheme.BORDER};
            padding: 8px 16px;
            font-size: {EnterpriseTheme.FONT_SIZE_SMALL};
        }}

        /* Buttons - Modern flat design */
        QPushButton {{
            background-color: {EnterpriseTheme.PRIMARY};
            color: {EnterpriseTheme.TEXT_ON_PRIMARY};
            border: none;
            padding: 10px 20px;
            border-radius: {EnterpriseTheme.RADIUS_MD};
            font-weight: 600;
            font-size: {EnterpriseTheme.FONT_SIZE_NORMAL};
            min-height: 32px;
        }}

        QPushButton:hover {{
            background-color: {EnterpriseTheme.PRIMARY_HOVER};
            box-shadow: {EnterpriseTheme.SHADOW_2};
        }}

        QPushButton:pressed {{
            background-color: {EnterpriseTheme.PRIMARY_PRESSED};
            box-shadow: none;
        }}

        QPushButton:disabled {{
            background-color: {EnterpriseTheme.SURFACE_HOVER};
            color: {EnterpriseTheme.TEXT_DISABLED};
        }}

        QPushButton#secondaryButton {{
            background-color: {EnterpriseTheme.SURFACE};
            color: {EnterpriseTheme.PRIMARY};
            border: 1px solid {EnterpriseTheme.PRIMARY};
        }}

        QPushButton#secondaryButton:hover {{
            background-color: {EnterpriseTheme.PRIMARY_LIGHT};
        }}

        QPushButton#dangerButton {{
            background-color: {EnterpriseTheme.ERROR};
        }}

        QPushButton#dangerButton:hover {{
            background-color: #B71C1C;
        }}

        /* Cards - Elevated surfaces */
        QFrame[frameShape="4"] {{
            background-color: {EnterpriseTheme.SURFACE};
            border: 1px solid {EnterpriseTheme.BORDER_LIGHT};
            border-radius: {EnterpriseTheme.RADIUS_LG};
            padding: {EnterpriseTheme.SPACING_MD};
        }}

        QGroupBox {{
            background-color: {EnterpriseTheme.SURFACE};
            border: 1px solid {EnterpriseTheme.BORDER};
            border-radius: {EnterpriseTheme.RADIUS_LG};
            margin-top: 16px;
            padding-top: 24px;
            font-weight: 600;
            color: {EnterpriseTheme.TEXT_PRIMARY};
            font-size: {EnterpriseTheme.FONT_SIZE_MEDIUM};
        }}

        QGroupBox::title {{
            subcontrol-origin: margin;
            left: 16px;
            padding: 0 8px;
            color: {EnterpriseTheme.PRIMARY};
        }}

        /* Tabs - Material style */
        QTabWidget::pane {{
            border: 1px solid {EnterpriseTheme.BORDER};
            border-radius: {EnterpriseTheme.RADIUS_LG};
            background: {EnterpriseTheme.SURFACE};
            top: -1px;
        }}

        QTabBar::tab {{
            background: transparent;
            color: {EnterpriseTheme.TEXT_SECONDARY};
            padding: 12px 24px;
            margin-right: 4px;
            border: none;
            border-bottom: 3px solid transparent;
            font-weight: 600;
            font-size: {EnterpriseTheme.FONT_SIZE_NORMAL};
        }}

        QTabBar::tab:hover {{
            color: {EnterpriseTheme.PRIMARY};
            background: {EnterpriseTheme.PRIMARY_LIGHT};
            border-radius: {EnterpriseTheme.RADIUS_MD} {EnterpriseTheme.RADIUS_MD} 0 0;
        }}

        QTabBar::tab:selected {{
            color: {EnterpriseTheme.PRIMARY};
            border-bottom: 3px solid {EnterpriseTheme.PRIMARY};
        }}

        /* Tables - Clean and readable */
        QTableWidget {{
            background-color: {EnterpriseTheme.SURFACE};
            gridline-color: {EnterpriseTheme.BORDER_LIGHT};
            border: 1px solid {EnterpriseTheme.BORDER};
            border-radius: {EnterpriseTheme.RADIUS_MD};
            selection-background-color: {EnterpriseTheme.PRIMARY_LIGHT};
            selection-color: {EnterpriseTheme.TEXT_PRIMARY};
            font-size: {EnterpriseTheme.FONT_SIZE_NORMAL};
        }}

        QTableWidget::item {{
            padding: 8px;
            border: none;
            border-bottom: 1px solid {EnterpriseTheme.BORDER_LIGHT};
        }}

        QTableWidget::item:selected {{
            background-color: {EnterpriseTheme.PRIMARY_LIGHT};
        }}

        /* Tablo başlıkları */
        QHeaderView::section:horizontal {{
            background-color: {EnterpriseTheme.SURFACE_HOVER};
            color: {EnterpriseTheme.TEXT_PRIMARY};
            padding: 8px 6px;  /* biraz daha dar padding */
            border: none;
            border-right: 1px solid {EnterpriseTheme.BORDER_LIGHT};
            border-bottom: 2px solid {EnterpriseTheme.PRIMARY};
            font-weight: 600;
            font-size: {EnterpriseTheme.FONT_SIZE_NORMAL};
        }}

        /* Satır başlıkları (sol taraftaki hat isimleri) */
        QHeaderView::section:vertical {{
            background-color: {EnterpriseTheme.SURFACE_HOVER};
            color: {EnterpriseTheme.TEXT_PRIMARY};
            padding: 4px 6px;
            border: none;
            border-right: 1px solid {EnterpriseTheme.BORDER_LIGHT};
            border-bottom: 1px solid {EnterpriseTheme.BORDER_LIGHT};
            font-weight: 600;
            font-size: {EnterpriseTheme.FONT_SIZE_NORMAL};
            min-width: 120px;    /* >>> metnin tam görünmesini garanti eder */
        }}


        /* Input fields - Modern and clean */
        QLineEdit, QTextEdit, QSpinBox, QDoubleSpinBox, QDateEdit, QComboBox {{
            background-color: {EnterpriseTheme.SURFACE};
            border: 1px solid {EnterpriseTheme.BORDER};
            border-radius: {EnterpriseTheme.RADIUS_MD};
            padding: 8px 12px;
            color: {EnterpriseTheme.TEXT_PRIMARY};
            font-size: {EnterpriseTheme.FONT_SIZE_NORMAL};
            selection-background-color: {EnterpriseTheme.PRIMARY_LIGHT};
        }}

        QLineEdit:focus, QTextEdit:focus, QSpinBox:focus, 
        QDoubleSpinBox:focus, QDateEdit:focus, QComboBox:focus {{
            border: 2px solid {EnterpriseTheme.PRIMARY};
            background-color: {EnterpriseTheme.SURFACE};
        }}

        QLineEdit:hover, QTextEdit:hover, QSpinBox:hover, 
        QDoubleSpinBox:hover, QDateEdit:hover, QComboBox:hover {{
            border-color: {EnterpriseTheme.PRIMARY_HOVER};
        }}

        /* ListWidget - Clean design */
        QListWidget {{
            background-color: {EnterpriseTheme.SURFACE};
            border: 1px solid {EnterpriseTheme.BORDER};
            border-radius: {EnterpriseTheme.RADIUS_MD};
            padding: 4px;
            outline: none;
        }}

        QListWidget::item {{
            padding: 10px 12px;
            margin: 2px;
            border-radius: {EnterpriseTheme.RADIUS_SM};
            color: {EnterpriseTheme.TEXT_PRIMARY};
        }}

        QListWidget::item:hover {{
            background-color: {EnterpriseTheme.PRIMARY_LIGHT};
            color: {EnterpriseTheme.PRIMARY};
        }}

        QListWidget::item:selected {{
            background-color: {EnterpriseTheme.PRIMARY};
            color: {EnterpriseTheme.TEXT_ON_PRIMARY};
        }}

        /* Scrollbars - Minimal and modern */
        QScrollBar:vertical {{
            background: {EnterpriseTheme.BACKGROUND};
            width: 12px;
            border: none;
            border-radius: 6px;
        }}

        QScrollBar::handle:vertical {{
            background: {EnterpriseTheme.BORDER};
            border-radius: 6px;
            min-height: 30px;
        }}

        QScrollBar::handle:vertical:hover {{
            background: {EnterpriseTheme.TEXT_DISABLED};
        }}

        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
            height: 0px;
        }}

        QScrollBar:horizontal {{
            background: {EnterpriseTheme.BACKGROUND};
            height: 12px;
            border: none;
            border-radius: 6px;
        }}

        QScrollBar::handle:horizontal {{
            background: {EnterpriseTheme.BORDER};
            border-radius: 6px;
            min-width: 30px;
        }}

        QScrollBar::handle:horizontal:hover {{
            background: {EnterpriseTheme.TEXT_DISABLED};
        }}

        /* Tooltips */
        QToolTip {{
            background-color: {EnterpriseTheme.TEXT_PRIMARY};
            color: {EnterpriseTheme.TEXT_ON_PRIMARY};
            border: none;
            border-radius: {EnterpriseTheme.RADIUS_SM};
            padding: 8px 12px;
            font-size: {EnterpriseTheme.FONT_SIZE_SMALL};
        }}

        /* Dialogs */
        QDialog {{
            background-color: {EnterpriseTheme.SURFACE};
        }}

        QDialogButtonBox {{
            padding-top: {EnterpriseTheme.SPACING_MD};
            border-top: 1px solid {EnterpriseTheme.BORDER_LIGHT};
        }}

        /* Labels */
        QLabel {{
            color: {EnterpriseTheme.TEXT_PRIMARY};
            background: transparent;
        }}

        QLabel[heading="true"] {{
            font-size: {EnterpriseTheme.FONT_SIZE_LARGE};
            font-weight: 700;
            color: {EnterpriseTheme.PRIMARY};
        }}

        QLabel[subheading="true"] {{
            font-size: {EnterpriseTheme.FONT_SIZE_NORMAL};
            color: {EnterpriseTheme.TEXT_SECONDARY};
        }}
        """

    @staticmethod
    def get_production_table_style():
        """Üretim planı tablosu için özel stil"""
        return f"""
        QTableWidget {{
            background-color: {EnterpriseTheme.SURFACE};
            gridline-color: {EnterpriseTheme.BORDER_LIGHT};
            border: none;
            font-size: {EnterpriseTheme.FONT_SIZE_NORMAL};
        }}

        QTableWidget::item {{
            padding: 6px;
            border: 1px solid {EnterpriseTheme.BORDER_LIGHT};
        }}

        QHeaderView::section {{
            background-color: {EnterpriseTheme.PRIMARY};
            color: {EnterpriseTheme.TEXT_ON_PRIMARY};
            padding: 10px 6px;
            border: none;
            border-right: 1px solid rgba(255,255,255,0.2);
            font-weight: 600;
            font-size: {EnterpriseTheme.FONT_SIZE_SMALL};
        }}

        QHeaderView::section:first {{
            border-top-left-radius: {EnterpriseTheme.RADIUS_MD};
        }}
        """