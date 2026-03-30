import sys
import json
import os
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox, QTableWidget,
    QTableWidgetItem, QHeaderView, QDateEdit, QDialog, QSpinBox,
    QLineEdit, QComboBox, QGroupBox, QSplitter, QListWidget,
    QListWidgetItem, QDialogButtonBox, QFormLayout, QDoubleSpinBox,
    QTabWidget, QScrollArea, QGridLayout, QMenu, QTextEdit,
    QRadioButton, QButtonGroup, QFrame, QAbstractItemView, QStatusBar
)
from PyQt6.QtCore import Qt, QDate, QMimeData, pyqtSignal
from PyQt6.QtGui import QDrag, QColor, QAction, QKeySequence, QIcon
from theme import EnterpriseTheme

# Veri dosyaları
DATA_DIR = "production_data"
PRODUCTS_FILE = os.path.join(DATA_DIR, "products.json")
HISTORICAL_FILE = os.path.join(DATA_DIR, "historical.json")
SETTINGS_FILE = os.path.join(DATA_DIR, "settings.json")
SABIT_ZAMANSAL_VERIM = 0.9586  # Tüm hatlar için sabit zamansal verim

os.makedirs(DATA_DIR, exist_ok=True)

# Hat listeleri
A_LINES = ['A1', 'A2', 'A3', 'A5', 'A4', 'A6', 'A7', 'A95', 'A97']
C_LINES = ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C11', 'C8', 'C7', 'C9', 'C10']
ALL_LINES = A_LINES + C_LINES

EXCLUDE_FROM_C_EFFICIENCY = {"C1"}  # C ortalamasına ve genel verime dahil edilmeyecek hatlar

# --- ÖZEL HAT SABİTLERİ ---

C5_SPECIAL = {
    "30020": {"devir": 210, "verim": 88, "gramaj": 103},
    "30020OPT": {"devir": 210, "verim": 91, "gramaj": 103},
    "DMT303": {"devir": 165, "verim": 89, "gramaj": 142},
    "DMT303OPT": {"devir": 165, "verim": 92, "gramaj": 142},
    "30011": {"devir": 210, "verim": 88, "gramaj": 86},
    "30011OPT": {"devir": 210, "verim": 90, "gramaj": 86},
    "30012": {"devir": 210, "verim": 84, "gramaj": 86},
    "30012OPT": {"devir": 210, "verim": 86, "gramaj": 86},
    "30030": {"devir": 200, "verim": 88, "gramaj": 89},
    "30030OPT": {"devir": 200, "verim": 91, "gramaj": 89},
}

C9_MST243 = {"devir": 140, "verim": 86, "gramaj": 175}  # sadece MST243 C9'da


class ProductData:
    def __init__(self):
        self.products = {}
        self.historical_data = pd.DataFrame()
        self.load_saved_data()

    def load_saved_data(self):
        if os.path.exists(PRODUCTS_FILE):
            try:
                with open(PRODUCTS_FILE, 'r', encoding='utf-8') as f:
                    self.products = json.load(f)
            except:
                self.products = {}

        if os.path.exists(HISTORICAL_FILE):
            try:
                with open(HISTORICAL_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.historical_data = pd.DataFrame(data)
            except:
                self.historical_data = pd.DataFrame()

    def save_products(self):
        with open(PRODUCTS_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.products, f, ensure_ascii=False, indent=2)

    def save_historical(self):
        data = self.historical_data.to_dict('records')
        with open(HISTORICAL_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def process_product_table_data(self, table_widget):
        try:
            rows = table_widget.rowCount()
            new_products = 0
            updated_products = 0
            for row in range(rows):
                kalem_item = table_widget.item(row, 0)
                devir_item = table_widget.item(row, 1)
                verim_item = table_widget.item(row, 2)
                gramaj_item = table_widget.item(row, 3)
                if kalem_item and devir_item and verim_item and gramaj_item:
                    kalem = kalem_item.text().strip()
                    if kalem:
                        try:
                            devir = float(str(devir_item.text()).strip().replace(',', '.'))
                            verim = float(str(verim_item.text()).strip().replace(',', '.'))
                            gramaj = float(str(gramaj_item.text()).strip().replace(',', '.'))
                            if kalem in self.products:
                                updated_products += 1
                            else:
                                new_products += 1
                            self.products[kalem] = {
                                'devir': devir,
                                'verim': verim,
                                'gramaj': gramaj
                            }
                        except ValueError:
                            continue
            self.save_products()
            return True, new_products, updated_products
        except Exception as e:
            print("Ürün verisi işlenirken hata:", e)
            return False, 0, 0

    def process_historical_table_data(self, table_widget):
        try:
            data = []
            rows = table_widget.rowCount()
            for row in range(rows):
                row_data = {}
                all_filled = True
                for col, column_name in enumerate(['Kalem', 'Yıl', 'Hat', 'ADETSEL', 'ZAMANSAL', 'VERIM']):
                    item = table_widget.item(row, col)
                    if item and item.text().strip():
                        row_data[column_name] = item.text().strip()
                    else:
                        all_filled = False
                        break
                if all_filled and row_data:
                    data.append(row_data)

            if data:
                new_df = pd.DataFrame(data)
                if not self.historical_data.empty:
                    self.historical_data = pd.concat([self.historical_data, new_df], ignore_index=True)
                    self.historical_data = self.historical_data.drop_duplicates(
                        subset=['Kalem', 'Yıl', 'Hat'], keep='last'
                    )
                else:
                    self.historical_data = new_df
                self.save_historical()
                return True, len(data)
            return True, 0
        except Exception as e:
            print(f"Geçmiş veri işlenirken hata: {e}")
            return False, 0


class PasteableTableWidget(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.StandardKey.Paste):
            self.paste_data()
        elif event.key() == Qt.Key.Key_Delete:
            self.delete_selected()
        elif event.matches(QKeySequence.StandardKey.SelectAll):
            self.selectAll()
        else:
            super().keyPressEvent(event)

    def paste_data(self):
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        if not text:
            return

        current_row = self.currentRow()
        current_col = self.currentColumn()
        if current_row < 0 or current_col < 0:
            current_row = 0
            current_col = 0

        rows = text.strip().split('\n')
        for i, row_text in enumerate(rows):
            if '\t' in row_text:
                cells = row_text.split('\t')
            else:
                cells = row_text.split(',')
            target_row = current_row + i
            if target_row >= self.rowCount():
                self.setRowCount(target_row + 1)
            for j, cell_text in enumerate(cells):
                target_col = current_col + j
                if target_col < self.columnCount():
                    item = QTableWidgetItem(cell_text.strip())
                    self.setItem(target_row, target_col, item)

    def delete_selected(self):
        for item in self.selectedItems():
            item.setText("")

    def clear_table(self):
        for row in range(self.rowCount()):
            for col in range(self.columnCount()):
                self.setItem(row, col, QTableWidgetItem(""))


class DraggableProductList(QListWidget):
    def __init__(self):
        super().__init__()
        self.setDragEnabled(True)
        self.setDefaultDropAction(Qt.DropAction.CopyAction)

    def startDrag(self, supportedActions):
        item = self.currentItem()
        if item:
            drag = QDrag(self)
            mime_data = QMimeData()
            mime_data.setText(item.text())
            drag.setMimeData(mime_data)
            drag.exec(Qt.DropAction.CopyAction)


class ProductionPlanTable(QTableWidget):
    # --- NEW: closed-line constants ---
    CLOSED_TAG = "HATKAPALI"
    CLOSED_LABEL = "HAT KAPALI"
    CLOSED_BASE_HEX = "#FFCDD2"  # soft red
    CLOSED_ACCENT_HEX = "#EF9A9A"  # darker red

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DropOnly)
        self.campaigns = {}
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_context_menu)
        self.display_rows = []
        self.tonaj_rows = {}
        self.tonaj_overrides = {}
        self.get_product = None
        self.is_a_line = None
        self.is_c_line = None
        self.row_a_total = None
        self.row_c_total = None
        self.get_product_names = None  # havuz ürün isimleri için callback

        # Modern tema uygula
        self.setStyleSheet(EnterpriseTheme.get_production_table_style())

    # ProductionPlanTable içinde:
    def _swap_campaign_dialog(self, key):
        (row, s, e) = key
        info = self.campaigns.get(key, {})
        cur_prod = info.get('product', '')
        dur = e - s + 1

        # Ürün listesini al
        pool = self.get_product_names() if callable(self.get_product_names) else []
        pool = list(pool)  # güvence
        if cur_prod in pool:
            pool.remove(cur_prod)

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Kampanyayı Değiştir ({dur} gün)")
        v = QVBoxLayout(dlg)
        cb = QComboBox(dlg)
        cb.addItems(sorted(pool))
        v.addWidget(QLabel(f"Mevcut: <b>{cur_prod}</b>"))
        v.addWidget(QLabel(f"Süre: {dur} gün"))
        v.addWidget(QLabel("Yeni ürün:"))
        v.addWidget(cb)
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        v.addWidget(btns)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        if not dlg.exec():
            return

        new_prod = cb.currentText().strip()
        if not new_prod:
            return

        self._remove_campaign(key)
        self._paint_segment(row, s, e, new_prod)
        self.campaigns[(row, s, e)] = {'product': new_prod, 'duration': dur}
        self.update_tonaj_totals()

    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()

    def _tonaj_for(self, product_name):
        if product_name == self.CLOSED_TAG:
            return 0.0
        if self.get_product is None:
            return None
        rec = self.get_product(product_name) or {}
        try:
            gramaj = float(rec.get('gramaj', 0))
            devir = float(rec.get('devir', 0))
            return gramaj * devir * 1440.0 / 1_000_000.0
        except Exception:
            return None

    def is_tonaj_row(self, row: int) -> bool:
        return row in self.tonaj_rows.values()

    def _make_locked_item(self, text: str) -> QTableWidgetItem:
        it = QTableWidgetItem(text)
        it.setFlags(Qt.ItemFlag.ItemIsEnabled)
        it.setTextAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        it.setBackground(QColor(EnterpriseTheme.SURFACE_HOVER))
        it.setForeground(QColor(EnterpriseTheme.TEXT_SECONDARY))
        return it

    def _get_daily_ton(self, row: int, col: int, product_name: str) -> float:
        # Günlük override varsa onu kullan (mevcut davranışı koru)
        if (row, col) in self.tonaj_overrides:
            return float(self.tonaj_overrides[(row, col)])

        devir, verim, gramaj, meta = self._effective_params(row, col, product_name)
        # Tonaj hesabı: gramaj(g/adet) * devir(adet/dk) * 1440 dk/gün → gram/gün → ton/gün
        try:
            base_ton = (gramaj * devir * 1440.0) / 1_000_000.0
        except Exception:
            base_ton = 0.0
        return base_ton

    def _campaign_covering(self, row: int, col: int):
        for (r, s, e), info in self.campaigns.items():
            if r == row and s <= col <= e:
                return (r, s, e, info)
        return None

    def _effective_params(self, row: int, col: int, product_name: str):
        """
        İlgili HAT + ÜRÜN + (C9 tek/çift damla durumu) + (C5 özel tablo) dikkate alınarak
        devir, verim, gramaj döndürür.
        """
        line = self.verticalHeaderItem(row).text()
        # Kapalı gün
        if product_name == self.CLOSED_TAG:
            return 0.0, 0.0, 0.0, {"damla": "-"}

        # Ürün havuz kaydı
        rec = (self.get_product(product_name) or {}).copy()
        base_devir = float(rec.get("devir", 0) or 0)
        base_verim = float(rec.get("verim", 0) or 0)
        base_gramaj = float(rec.get("gramaj", 0) or 0)
        damla_info = "Çift"

        # --- C5 ÖZEL ---
        if line == "C5" and product_name in C5_SPECIAL:
            sp = C5_SPECIAL[product_name]
            # DEVİR: her zaman özel tablodan
            devir = float(sp["devir"])
            gramaj = float(sp["gramaj"])
            # VERİM: havuzda/ geçmişte varsa onu, yoksa özel tablo
            verim = base_verim if base_verim > 0 else float(sp["verim"])
            return devir, verim, gramaj, {"damla": damla_info, "kaynak": "C5 özel"}

        # --- C9 ÖZEL: MST243 sabitleri ---
        if line == "C9" and product_name == "MST243":
            devir = float(C9_MST243["devir"])
            gramaj = float(C9_MST243["gramaj"])
            verim = base_verim if base_verim > 0 else float(C9_MST243["verim"])
            return devir, verim, gramaj, {"damla": damla_info, "kaynak": "C9 MST243"}

        # --- C9 GENEL: tek/çift damla sorgusuna göre devir bölünür ---
        if line == "C9":
            # Bu hücrenin hangi kampanyada olduğunu bul ve bayrağı oku
            hit = self._campaign_covering(row, col)
            single_drop = False
            if hit:
                r, s, e, info = hit
                single_drop = bool(info.get("single_drop", False))
            devir = base_devir / 2.0 if single_drop else base_devir
            damla_info = "Tek" if single_drop else "Çift"
            return devir, base_verim, base_gramaj, {"damla": damla_info, "kaynak": "C9"}

        # --- DEFAULT ---
        return base_devir, base_verim, base_gramaj, {"damla": damla_info, "kaynak": "Havuz"}

    def update_per_line_tonaj_rows(self):
        cols = self.columnCount()
        for ton_row in self.tonaj_rows.values():
            for c in range(cols):
                self.setItem(ton_row, c, self._make_locked_item(""))

        for (row, s, e), info in self.campaigns.items():
            if row not in self.display_rows:
                continue
            ton_row = self.tonaj_rows.get(row)
            if ton_row is None:
                continue
            product = info['product']
            for c in range(s, e + 1):
                add_val = self._get_daily_ton(row, c, product)
                cur = self.item(ton_row, c)
                cur_val = 0.0
                if cur and cur.text().strip():
                    txt = cur.text().strip().replace(".", "").replace(",", ".")
                    try:
                        cur_val = float(txt)
                    except:
                        cur_val = 0.0
                new_val = cur_val + add_val
                formatted = f"{new_val:.1f}".replace(".", ",")
                it = self._make_locked_item(formatted)
                if (row, c) in self.tonaj_overrides:
                    it.setBackground(QColor("#FFF8E1"))
                self.setItem(ton_row, c, it)

    def _paint_segment(self, row: int, start_col: int, end_col: int, product_name: str):
        if product_name == self.CLOSED_TAG:
            base_color = QColor(self.CLOSED_BASE_HEX)
            highlight_color = QColor(self.CLOSED_ACCENT_HEX)
            cell_text = self.CLOSED_LABEL
        else:
            base_color, highlight_color = self._colors_for_line(self.verticalHeaderItem(row).text())
            cell_text = product_name

        for c in range(start_col, end_col + 1):
            item = QTableWidgetItem(cell_text)
            bg = base_color if c == start_col else highlight_color
            item.setBackground(bg)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setForeground(QColor(EnterpriseTheme.TEXT_PRIMARY))

            # --- HOVER İPUCU: devir/gramaj (+ opsiyonel verim & damla bilgisi) ---
            try:
                devir, verim, gramaj, meta = self._effective_params(row, c, product_name)
                line = self.verticalHeaderItem(row).text()
                t1 = (self.parent().parent().plan_start_date + timedelta(days=c)).strftime("%d.%m.%Y") \
                    if hasattr(self.parent().parent(), "plan_start_date") else f"Gün {c + 1}"
                tip = (f"{line} | {product_name}\n"
                       f"Tarih/Gün: {t1}\n"
                       f"Devir: {devir:.0f} adet/dk\n"
                       f"Gramaj: {gramaj:.0f} g/adet\n"
                       f"Damla: {meta.get('damla', '-')}")
                # Verim bilgin olsun istersen göster:
                if verim > 0:
                    tip += f"\nVerim (kaynak): {verim:.0f}% ({meta.get('kaynak', '')})"
                item.setToolTip(tip)
            except Exception:
                pass

            self.setItem(row, c, item)

    def update_tonaj_totals(self):
        if self.row_a_total is None or self.row_c_total is None:
            return
        cols = self.columnCount()

        # Önce tüm hücreleri temizle (bu kısmı KORU)
        for r in (self.row_a_total, self.row_c_total):
            for c in range(cols):
                self.setItem(r, c, QTableWidgetItem(""))

        a_sums = [0.0] * cols
        c_sums = [0.0] * cols

        for (row, s, e), info in self.campaigns.items():
            product = info['product']
            for c in range(s, e + 1):
                day_ton = self._get_daily_ton(row, c, product)
                target = a_sums if (self.is_a_line and self.is_a_line(row)) else c_sums
                target[c] += day_ton

        def _write_sum(row_idx, sums, label):
            for c in range(cols):
                formatted = f"{sums[c]:.2f}".replace(".", ",")
                it = self._make_locked_item(f"{formatted}{label}")
                it.setBackground(QColor(255, 255, 255))
                it.setForeground(QColor(0, 12, 120))
                font = it.font()
                font.setBold(True)
                it.setFont(font)
                self.setItem(row_idx, c, it)

        _write_sum(self.row_a_total, a_sums, " ton")
        _write_sum(self.row_c_total, c_sums, " ton")
        self.update_per_line_tonaj_rows()

    def dropEvent(self, event):
        if event.mimeData().hasText():
            product_name = event.mimeData().text()
            pos = event.position().toPoint()
            row = self.rowAt(pos.y())
            col = self.columnAt(pos.x())
            if row < 0 or col < 0:
                return

            if self.is_tonaj_row(row):
                return
            if self.row_a_total is not None and row == self.row_a_total:
                return
            if self.row_c_total is not None and row == self.row_c_total:
                return

            dialog = DurationDialog(self)
            if dialog.exec():
                duration = dialog.get_duration()
                self.add_campaign(row, col, product_name, duration)
            event.acceptProposedAction()

    def add_campaign(self, row, start_col, product_name, duration):
        end_col = min(start_col + duration - 1, self.columnCount() - 1)

        # ... (mevcut çakışma kodları)

        # --- C9 ÖZEL SORU: MST243 harici ürünler için tek/çift damla? ---
        line = self.verticalHeaderItem(row).text()
        single_drop = False
        if line == "C9" and product_name != "MST243":
            # Tarih aralığını kullanıcıya anlaşılır verelim
            try:
                app = self.parent().parent()  # ProductionPlannerApp
                d1 = (app.plan_start_date + timedelta(days=start_col)).strftime("%d.%m.%Y")
                d2 = (app.plan_start_date + timedelta(days=end_col)).strftime("%d.%m.%Y")
                ask = QMessageBox.question(
                    self, "C9 Damla Tipi",
                    f"{product_name} için {d1} – {d2} aralığı TEK damla mı çalışacak?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )
                single_drop = (ask == QMessageBox.StandardButton.Yes)
            except Exception:
                # bir sorun olursa varsayılan çift damla kabul et
                single_drop = False

        self._paint_segment(row, start_col, end_col, product_name)
        self.campaigns[(row, start_col, end_col)] = {
            'product': product_name,
            'duration': end_col - start_col + 1,
            # C9 tek damla bayrağı burada taşınır
            'single_drop': single_drop
        }
        self.update_tonaj_totals()

    def _remove_campaign(self, key):
        (row, start_col, end_col) = key
        for col in range(start_col, end_col + 1):
            empty = QTableWidgetItem("")
            empty.setBackground(QColor(0, 0, 0, 0))
            self.setItem(row, col, empty)
        if key in self.campaigns:
            del self.campaigns[key]
        self._clear_overrides_in_range(row, start_col, end_col)
        self.update_tonaj_totals()

    def clear_row(self, row):
        to_del = [k for k in self.campaigns.keys() if k[0] == row]
        for k in to_del:
            self._remove_campaign(k)
        keys = [(r, c) for (r, c) in list(self.tonaj_overrides.keys()) if r == row]
        for k in keys:
            self.tonaj_overrides.pop(k, None)
        self.update_tonaj_totals()

    def clear_all(self):
        for k in list(self.campaigns.keys()):
            self._remove_campaign(k)
        self.campaigns.clear()
        self.tonaj_overrides.clear()
        self.update_tonaj_totals()

    def _show_context_menu(self, pos):
        global_pos = self.viewport().mapToGlobal(pos)
        row = self.rowAt(pos.y())
        col = self.columnAt(pos.x())
        menu = QMenu(self)

        # Tonaj satırında sadece "Tüm planı temizle" kalsın (mevcut davranış)
        if self.is_tonaj_row(row):
            act_all = menu.addAction("Tüm planı temizle")
            act_all.triggered.connect(self.clear_all)
            menu.exec(global_pos)
            return

        # Toplam satırlarına işlem verilmesin
        if row in (self.row_a_total, self.row_c_total):
            act_all = menu.addAction("Tüm planı temizle")
            act_all.triggered.connect(self.clear_all)
            menu.exec(global_pos)
            return

        # Hücre bir kampanyanın içinde mi?
        target_key = None
        for (r, s, e) in self.campaigns.keys():
            if r == row and s <= col <= e:
                target_key = (r, s, e)
                break

        if target_key is not None:
            # Var olan kampanya menüsü (mevcutlar)
            act_ton = menu.addAction("Tonaj değiştir...")
            act_ton.triggered.connect(lambda _, r=row, c=col: self._open_tonaj_dialog(r, c))

            if (row, col) in self.tonaj_overrides:
                act_clr = menu.addAction("Bu günün tonaj revizesini temizle")
                act_clr.triggered.connect(lambda _, r=row, c=col: self._clear_override_for_day(r, c))

            # Hedef kampanya bulunduğunda (target_key varsa) ekle:
            line_name = self.verticalHeaderItem(row).text().strip().upper()
            prod = self.campaigns[target_key].get("product", "")

            # Sadece C9 için göster (MST243 özel kural var ama tek/çift seçeneği göstermekte sakınca yok;
            # MST243 çalışıyorsa tek/çift seçilse bile _effective_params bunu override eder)
            if line_name == "C9" and prod and prod != self.CLOSED_TAG:
                act_single = QAction("C9: Tek damla (devri 2'ye böl)", self)
                act_double = QAction("C9: Çift damla (normal devri kullan)", self)

                def _set_single():
                    self.campaigns[target_key]["single_drop"] = True
                    self.update_tonaj_totals()
                    self.viewport().update()

                def _set_double():
                    self.campaigns[target_key]["single_drop"] = False
                    self.update_tonaj_totals()
                    self.viewport().update()

                menu.addAction(act_single)
                menu.addAction(act_double)
                act_single.triggered.connect(_set_single)
                act_double.triggered.connect(_set_double)

            act_swap = menu.addAction("Kampanyayı Değiştir...")
            act_swap.triggered.connect(lambda _, key=target_key: self._swap_campaign_dialog(key))

            act_del = menu.addAction("Seçili kampanyayı sil")
            act_del.triggered.connect(lambda _, key=target_key: self._remove_campaign(key))

            act_row = menu.addAction("Bu hattı temizle")
            act_row.triggered.connect(lambda _, r=row: self.clear_row(r))

        else:
            # NEW: Boş hücrede "Hat kapalı ekle..." seçeneği göster
            act_closed = menu.addAction("Hat kapalı ekle...")
            act_closed.triggered.connect(lambda _, r=row, c=col: self._add_closed_with_dialog(r, c))

            # İsteğe bağlı: boşta iken hattı komple temizleme (kullanışlı olur)
            act_row = menu.addAction("Bu hattı temizle")
            act_row.triggered.connect(lambda _, r=row: self.clear_row(r))

        # Her durumda en altta kalsın
        act_all = menu.addAction("Tüm planı temizle")
        act_all.triggered.connect(self.clear_all)
        menu.exec(global_pos)

    def _add_closed_with_dialog(self, row: int, col: int):
        """Gün sayısını sorup kapalı segment ekler."""
        dlg = DurationDialog(self)
        if dlg.exec():
            duration = dlg.get_duration()
            self.add_closed(row, col, duration)

    def add_closed(self, row: int, start_col: int, duration: int):
        """Kapalı segmenti ekler; çakışmalar varsa mevcut mantıkla split/trim yapar."""
        end_col = min(start_col + duration - 1, self.columnCount() - 1)

        # Çakışan kampanyaları topla
        overlapping_keys = []
        for (r, s, e) in list(self.campaigns.keys()):
            if r == row and not (e < start_col or s > end_col):
                overlapping_keys.append((r, s, e))

        if overlapping_keys:
            ret = QMessageBox.question(
                self, "Üzerine Yazılsın mı?",
                "Bu tarih aralığında mevcut bir plan var.\n"
                "Sadece çakışan günlerde üzerine yazmak ister misiniz?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if ret == QMessageBox.StandardButton.No:
                return
            for key in overlapping_keys:
                self._split_or_trim_campaign(key, start_col, end_col)

        # Kapalı segmenti boya ve kaydet
        self._paint_segment(row, start_col, end_col, self.CLOSED_TAG)
        self.campaigns[(row, start_col, end_col)] = {
            'product': self.CLOSED_TAG,
            'duration': end_col - start_col + 1
        }

        # Kapalı segment default 0 ton; override alanı açık kalır (bilerek bir şey yapmıyoruz)
        self.update_tonaj_totals()

    def _clear_override_for_day(self, row: int, col: int):
        self.tonaj_overrides.pop((row, col), None)
        self.update_tonaj_totals()

    def _clear_overrides_in_range(self, row: int, start_col: int, end_col: int):
        keys = [(r, c) for (r, c) in self.tonaj_overrides.keys() if r == row and start_col <= c <= end_col]
        for k in keys:
            self.tonaj_overrides.pop(k, None)

    def _open_tonaj_dialog(self, row: int, col: int):
        hit = self._campaign_covering(row, col)
        if not hit:
            QMessageBox.information(self, "Bilgi", "Bu hücrede kampanya yok.")
            return
        r, s, e, info = hit

        dlg = TonajChangeDialog(self, default_ton=self._get_daily_ton(row, col, info['product']))
        if dlg.exec():
            new_ton, mode, days = dlg.get_values()
            if mode == "days":
                end_col = min(col + days - 1, e)
            else:
                end_col = e

            self._clear_overrides_in_range(row, col, end_col)
            for c in range(col, end_col + 1):
                self.tonaj_overrides[(row, c)] = float(new_ton)

            self.update_tonaj_totals()

    def _colors_for_line(self, line_name: str):
        if line_name.startswith('C'):
            return QColor(EnterpriseTheme.C_LINE_BASE), QColor(EnterpriseTheme.C_LINE_ACCENT)
        else:
            return QColor(EnterpriseTheme.A_LINE_BASE), QColor(EnterpriseTheme.A_LINE_ACCENT)

    def _split_or_trim_campaign(self, key, new_start: int, new_end: int):
        (row, s, e) = key
        info = self.campaigns.get(key, None)
        if info is None:
            return

        product_name = info['product']
        self._remove_campaign(key)

        left_start = s
        left_end = min(new_start - 1, e)
        if left_start <= left_end:
            self._paint_segment(row, left_start, left_end, product_name)
            self.campaigns[(row, left_start, left_end)] = {
                'product': product_name,
                'duration': left_end - left_start + 1
            }

        right_start = max(new_end + 1, s)
        right_end = e
        if right_start <= right_end:
            self._paint_segment(row, right_start, right_end, product_name)
            self.campaigns[(row, right_start, right_end)] = {
                'product': product_name,
                'duration': right_end - right_start + 1
            }
        self.update_tonaj_totals()


class DurationDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Kampanya Süresi")
        self.setModal(True)
        layout = QFormLayout()
        self.duration_spin = QSpinBox()
        self.duration_spin.setMinimum(1)
        self.duration_spin.setMaximum(365)
        self.duration_spin.setValue(4)
        layout.addRow("Gün Sayısı:", self.duration_spin)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)
        self.setLayout(layout)

    def get_duration(self):
        return self.duration_spin.value()


class TonajChangeDialog(QDialog):
    def __init__(self, parent=None, default_ton: float = 0.0):
        super().__init__(parent)
        self.setWindowTitle("Tonaj Değiştir")
        self.setModal(True)
        lay = QVBoxLayout(self)
        form = QFormLayout()
        self.ton_spin = QDoubleSpinBox()
        self.ton_spin.setDecimals(1)
        self.ton_spin.setRange(0.0, 200.0)
        self.ton_spin.setValue(float(default_ton) if default_ton else 0.0)
        self.ton_spin.setSuffix(" ton/gün")
        form.addRow("Yeni tonaj:", self.ton_spin)
        lay.addLayout(form)

        self.rb_days = QRadioButton("Belirli gün sayısı")
        self.rb_to_end = QRadioButton("Kampanya sonuna kadar")
        self.rb_days.setChecked(True)

        self.days_spin = QSpinBox()
        self.days_spin.setRange(1, 365)
        self.days_spin.setValue(1)

        row_days = QHBoxLayout()
        row_days.addWidget(self.rb_days)
        row_days.addWidget(self.days_spin)
        row_days.addStretch()

        lay.addLayout(row_days)
        lay.addWidget(self.rb_to_end)

        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    def get_values(self):
        ton = float(self.ton_spin.value())
        if self.rb_to_end.isChecked():
            return ton, "to_end", None
        else:
            return ton, "days", int(self.days_spin.value())


class TimeEfficiencyDialog(QDialog):
    def __init__(self, parent=None, saved_values=None):
        super().__init__(parent)
        self.setWindowTitle("Zamansal Verim Ayarları")
        self.setModal(True)
        self.setMinimumWidth(500)

        # Kayıtlı değerlerin boş gelme ihtimaline karşı önlem
        self.saved_values = saved_values if saved_values is not None else {}

        main_layout = QVBoxLayout()

        # --- YENİ EKLENEN KISIM: SABİT DEĞERLER GRUBU ---
        general_group = QGroupBox("Genel ve Fırın Bazlı Sabit Hedefler (%)")
        general_layout = QFormLayout()

        # A Fırını Sabit
        self.spin_a_fixed = QDoubleSpinBox()
        self.spin_a_fixed.setRange(0, 100)
        self.spin_a_fixed.setSuffix("%")
        # 'FIXED_A' anahtarını float olarak güvenli çek
        val_a = self.saved_values.get("FIXED_A", 95.33)
        self.spin_a_fixed.setValue(float(val_a))
        general_layout.addRow("A Fırını Genel Zamansal:", self.spin_a_fixed)

        # C Fırını Sabit
        self.spin_c_fixed = QDoubleSpinBox()
        self.spin_c_fixed.setRange(0, 100)
        self.spin_c_fixed.setSuffix("%")
        val_c = self.saved_values.get("FIXED_C", 96.37)
        self.spin_c_fixed.setValue(float(val_c))
        general_layout.addRow("C Fırını Genel Zamansal:", self.spin_c_fixed)

        # Toplam (Tüm Fabrika) Sabit
        self.spin_total_fixed = QDoubleSpinBox()
        self.spin_total_fixed.setRange(0, 100)
        self.spin_total_fixed.setSuffix("%")
        val_total = self.saved_values.get("FIXED_TOTAL", 95.86)
        self.spin_total_fixed.setValue(float(val_total))
        general_layout.addRow("Fabrika Geneli Zamansal:", self.spin_total_fixed)

        general_group.setLayout(general_layout)
        main_layout.addWidget(general_group)
        # ------------------------------------------------

        # C Hatları Grubu
        c_group = QGroupBox("C Fırını Hat Bazlı Detaylar")
        c_layout = QFormLayout()
        self.c_lines = {}
        # C_LINES listesinin main.py başında tanımlı olduğundan emin olun
        for line in C_LINES:
            spin = QDoubleSpinBox()
            spin.setMinimum(0)
            spin.setMaximum(100)
            spin.setSuffix("%")
            # Her hattı kendi key'i ile çek
            val_line = self.saved_values.get(line, 96.37)
            spin.setValue(float(val_line))
            c_layout.addRow(f"{line} Hattı:", spin)
            self.c_lines[line] = spin
        c_group.setLayout(c_layout)

        # A Hatları Grubu
        a_group = QGroupBox("A Fırını Hat Bazlı Detaylar")
        a_layout = QFormLayout()
        self.a_lines = {}
        # A_LINES listesinin main.py başında tanımlı olduğundan emin olun
        for line in A_LINES:
            spin = QDoubleSpinBox()
            spin.setMinimum(0)
            spin.setMaximum(100)
            spin.setSuffix("%")
            val_line = self.saved_values.get(line, 95.33)
            spin.setValue(float(val_line))
            a_layout.addRow(f"{line} Hattı:", spin)
            self.a_lines[line] = spin
        a_group.setLayout(a_layout)

        # Scroll Area (Ekran taşmasın diye)
        scroll = QScrollArea()
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout()
        scroll_layout.addWidget(c_group)
        scroll_layout.addWidget(a_group)
        scroll_widget.setLayout(scroll_layout)
        scroll.setWidget(scroll_widget)
        scroll.setWidgetResizable(True)

        main_layout.addWidget(scroll)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        main_layout.addWidget(buttons)

        self.setLayout(main_layout)

    def get_efficiencies(self):
        """Kullanıcının girdiği değerleri sözlük olarak döndürür."""
        result = {}

        # Hat bazlıları topla
        for line, spin in self.c_lines.items():
            result[line] = spin.value()
        for line, spin in self.a_lines.items():
            result[line] = spin.value()

        # --- ÖNEMLİ: Sabitleri sözlüğe ekle ---
        # Bu anahtarlar __init__ kısmındaki .get() ile AYNI olmalı
        result["FIXED_A"] = self.spin_a_fixed.value()
        result["FIXED_C"] = self.spin_c_fixed.value()
        result["FIXED_TOTAL"] = self.spin_total_fixed.value()
        # --------------------------------------

        return result


class NewPlanDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Yeni Plan Oluştur")
        self.setModal(True)
        layout = QFormLayout()

        self.start_date = QDateEdit()
        self.start_date.setDate(QDate.currentDate())
        self.start_date.setCalendarPopup(True)
        self.start_date.setDisplayFormat("dd.MM.yyyy")
        layout.addRow("Başlangıç Tarihi:", self.start_date)

        self.end_date = QDateEdit()
        self.end_date.setDate(QDate.currentDate().addDays(30))
        self.end_date.setCalendarPopup(True)
        self.end_date.setDisplayFormat("dd.MM.yyyy")
        layout.addRow("Bitiş Tarihi:", self.end_date)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

        self.setLayout(layout)

    def validate_and_accept(self):
        if self.start_date.date() > self.end_date.date():
            QMessageBox.warning(self, "Uyarı", "Başlangıç tarihi bitiş tarihinden sonra olamaz!")
        else:
            self.accept()

    def get_dates(self):
        return self.start_date.date(), self.end_date.date()


class ImportPlanDialog(QDialog):
    def __init__(self, parent=None, lines=None):
        super().__init__(parent)
        self.setWindowFlag(Qt.WindowType.Window, True)
        self.setWindowFlag(Qt.WindowType.WindowMaximizeButtonHint, True)
        self.setWindowTitle("Excel Planı İçe Aktar")
        self.setModal(True)
        # self.resize(1280, 720)
        self.lines = lines or []

        layout = QVBoxLayout(self)

        info = QLabel("📋 Excel'den kopyaladığınız planı buraya yapıştırın\n"
                      "• Her satır bir hat olmalı\n"
                      "• Ürün adı sadece kampanyanın ilk gününde olmalı\n"
                      "• Boş hücreler kampanyanın devamını temsil eder")
        info.setProperty("subheading", True)
        info.setWordWrap(True)
        # layout.addWidget(info)

        # Hat seçimi
        line_layout = QHBoxLayout()
        line_layout.addWidget(QLabel("Başlangıç Hattı:"))
        self.line_combo = QComboBox()
        self.line_combo.addItems(self.lines)
        self.line_combo.currentTextChanged.connect(self._refresh_headers)  # YENİ
        line_layout.addWidget(self.line_combo)
        line_layout.addStretch()
        layout.addLayout(line_layout)

        # Tablo
        self.table = QTableWidget()
        # Dikey başlıkların görünümü (bozulmayı engelle)
        vh = self.table.verticalHeader()
        vh.setMinimumWidth(120)
        vh.setDefaultSectionSize(28)
        vh.setSectionResizeMode(QHeaderView.ResizeMode.Fixed)

        # Soldan sağa yazım ve satır yüksekliği tutarlılığı
        self.table.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        self.table.setColumnCount(31)  # Maksimum 31 gün
        self.table.setHorizontalHeaderLabels([f"Gün {i + 1}" for i in range(31)])
        self._refresh_headers()
        layout.addWidget(self.table)

        # Butonlar
        btn_layout = QHBoxLayout()
        btn_parse = QPushButton("Tablodan Çözümle")
        btn_parse.clicked.connect(self.parse_from_table)
        btn_clear = QPushButton("Temizle")
        btn_clear.setObjectName("secondaryButton")
        btn_clear.clicked.connect(self.clear_table)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_parse)
        btn_layout.addWidget(btn_clear)
        layout.addLayout(btn_layout)

        # Sonuç önizleme
        self.preview = QTextEdit()
        self.preview.setReadOnly(True)
        self.preview.setMaximumHeight(150)
        layout.addWidget(QLabel("Önizleme:"))
        layout.addWidget(self.preview)

        # Dialog butonları
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.campaigns = []
        self.setWindowState(self.windowState() | Qt.WindowState.WindowMaximized)

    # --- ImportPlanDialog içine YENİ metot ekle ---
    def _effective_lines(self):
        """Başlangıç hattından itibaren döndürülmüş hat listesi (wrap)."""
        if not self.lines:
            return []
        start = self.line_combo.currentText()
        idx = self.lines.index(start) if start in self.lines else 0
        return self.lines[idx:] + self.lines[:idx]

    def _refresh_headers(self):
        """Başlıkları setVerticalHeaderItem ile tek tek yazar; hizalama ve genişlik garanti."""
        eff = self._effective_lines()
        row_count = len(eff) * 2
        self.table.setRowCount(row_count)

        vh = self.table.verticalHeader()
        vh.setMinimumWidth(120)  # genişlik garanti
        vh.setDefaultSectionSize(28)
        vh.setMinimumSectionSize(24)
        vh.setSectionResizeMode(QHeaderView.ResizeMode.Fixed)

        # Eski label toplu atamayı KULLANMA
        # self.table.setVerticalHeaderLabels([...])  # <-- BUNU ARTIK KULLANMIYORUZ

        # Tek tek item ver (hizalamayı her birine uygula)
        for i, ln in enumerate(eff):
            plan_item = QTableWidgetItem(ln)
            plan_item.setTextAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
            self.table.setVerticalHeaderItem(i * 2, plan_item)

            tonaj_item = QTableWidgetItem("Tonaj")  # kısa ve temiz
            tonaj_item.setTextAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
            self.table.setVerticalHeaderItem(i * 2 + 1, tonaj_item)

    def clear_table(self):
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                self.table.setItem(row, col, QTableWidgetItem(""))
        self.preview.clear()

    # --- ImportPlanDialog.parse_from_table() tamamen değiştir ---
    def parse_from_table(self):
        """Tablodan kampanyaları çözümle: sadece hat satırlarını oku; altındaki Tonaj satırlarını YOK SAY."""
        self.campaigns = []
        preview_text = []

        eff_lines = self._effective_lines()  # döndürülmüş hat listesi
        if not eff_lines:
            QMessageBox.warning(self, "Uyarı", "Hattı seçiniz.")
            return

        for i, line_name in enumerate(eff_lines):
            row = i * 2  # 0,2,4,... (ürün satırı)
            col = 0
            while col < self.table.columnCount():
                item = self.table.item(row, col)
                if not item or not item.text().strip():
                    col += 1
                    continue

                product = item.text().strip()
                start_col = col
                duration = 1
                col += 1

                # Boş hücreler kampanyaya dahildir; farklı ürün gelince sonlanır
                while col < self.table.columnCount():
                    nxt = self.table.item(row, col)
                    if (not nxt) or (not nxt.text().strip()):
                        duration += 1
                        col += 1
                    else:
                        break

                self.campaigns.append({
                    'line': line_name,
                    'start_day': start_col + 1,  # 1-based
                    'product': product,
                    'duration': duration
                })
                preview_text.append(f"• {line_name}: {product} – Gün {start_col + 1} – {duration} gün")

        self.preview.setPlainText("\n".join(preview_text))
        if not self.campaigns:
            QMessageBox.warning(self, "Uyarı", "Hiç kampanya bulunamadı!")

    def get_campaigns(self):
        return self.campaigns

    def keyPressEvent(self, event):
        """Ctrl+V ile yapıştırmayı destekle"""
        if event.matches(QKeySequence.StandardKey.Paste):
            self.paste_data()
        else:
            super().keyPressEvent(event)

    def paste_data(self):
        """Excel'den kopyalanan veriyi yapıştır"""
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        if not text:
            return

        rows = text.strip().split('\n')
        for i, row_text in enumerate(rows):
            if i >= self.table.rowCount():
                break

            # Tab veya virgülle ayrılmış
            if '\t' in row_text:
                cells = row_text.split('\t')
            else:
                cells = row_text.split(',')

            for j, cell_text in enumerate(cells):
                if j >= self.table.columnCount():
                    break
                item = QTableWidgetItem(cell_text.strip())
                self.table.setItem(i, j, item)


class NewProductTonajDialog(QDialog):
    def __init__(self, parent=None, product_name=""):
        super().__init__(parent)
        self.setWindowTitle(f"Yeni Ürün: {product_name}")
        v = QVBoxLayout(self)

        v.addWidget(QLabel(f"'{product_name}' havuzda yok. Tonaj davranışını seçin:"))

        self.rb_none = QRadioButton("Tonaj hesaplanmasın (0 alınsın)")
        self.rb_manual = QRadioButton("Tonaj hesaplansın (değer gireceğim)")
        self.rb_none.setChecked(True)

        v.addWidget(self.rb_none)
        v.addWidget(self.rb_manual)

        form = QFormLayout()
        self.sp_ton = QDoubleSpinBox()
        self.sp_ton.setDecimals(1)
        self.sp_ton.setRange(0.0, 200.0)
        self.sp_ton.setValue(0.0)
        self.sp_ton.setSuffix(" ton/gün")
        self.sp_ton.setEnabled(False)
        form.addRow("Tonaj:", self.sp_ton)
        v.addLayout(form)

        def on_toggle():
            self.sp_ton.setEnabled(self.rb_manual.isChecked())

        self.rb_manual.toggled.connect(on_toggle)

        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        v.addWidget(btns)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)

    def get_choice(self):
        if self.rb_manual.isChecked():
            return ("manual", float(self.sp_ton.value()))
        return ("none", 0.0)


class HistoricalScopeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Adetsel Kapsam Seçimi")
        self.setModal(True)
        v = QVBoxLayout(self)

        lbl = QLabel("Aylık tahmini verim hesabında geçmiş ADETSEL'i nasıl kullanalım?")
        lbl.setWordWrap(True)
        v.addWidget(lbl)

        self.rb_all = QRadioButton("1) Kampanyaları TÜM hatlardaki veriyi alarak oluştur (ürün-bazlı)")
        self.rb_line = QRadioButton(
            "2) Kampanyaları SADECE bu hatta çalıştığı verileri alarak oluştur (ürün+hat-bazlı)")
        # Mevcut davranış “hat-bazlı” olduğu için onu varsayılan yapalım:
        self.rb_line.setChecked(True)

        v.addWidget(self.rb_all)
        v.addWidget(self.rb_line)

        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        v.addWidget(btns)

    def get_scope(self):
        # 'all' => tüm hatlar; 'line' => sadece ilgili hat
        return 'all' if self.rb_all.isChecked() else 'line'


class ProductionPlannerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.product_data = ProductData()
        self.time_efficiencies = self.load_time_efficiencies()
        self.setupUI()
        self.setStyleSheet(EnterpriseTheme.get_main_stylesheet())
        self.load_existing_data()
        self.sidebar_collapsed = False
        self._sidebar_last_width = 260  # panel açıkken hatırlanacak genişlik (varsayılan)

    def load_settings(self):
        """Ayarları JSON dosyasından yükler."""
        if not os.path.exists(SETTINGS_FILE):
            return

        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)

                # 'time_efficiencies' anahtarını olduğu gibi alıyoruz.
                # Filtreleme yapmıyoruz ki FIXED_A, FIXED_C gibi özel ayarlar kaybolmasın.
                self.time_efficiencies = data.get("time_efficiencies", {})

                # Eğer eski bir ayar dosyasıysa ve içi boşsa, varsayılanları doldurabiliriz (Opsiyonel)
                if not self.time_efficiencies:
                    self.time_efficiencies = {}

        except Exception as e:
            # Hata olursa sessizce geç veya logla
            print(f"Ayarlar yüklenirken hata: {e}")
            self.time_efficiencies = {}

    def _apply_saved_sidebar_state(self):
        # settings'ı oku
        ui = {}
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    ui = settings.get('ui', {})
            except:
                pass

        self.sidebar_collapsed = bool(ui.get('sidebar_collapsed', False))
        self._sidebar_last_width = int(ui.get('sidebar_width', 260))

        # splitter henüz var; uygulayalım
        if self.sidebar_collapsed:
            # kapalı başlat
            total = 1200
            self.splitter.setSizes([0, total])
        else:
            total = 1200
            left = max(180, self._sidebar_last_width)
            right = max(total - left, 200)
            self.splitter.setSizes([left, right])

    def _persist_sidebar_state(self):
        # mevcut settings'ı oku
        settings = {}
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
            except:
                settings = {}
        # UI kısmını güncelle
        ui = settings.get('ui', {})
        ui['sidebar_collapsed'] = self.sidebar_collapsed
        ui['sidebar_width'] = self._sidebar_last_width
        settings['ui'] = ui
        # yaz
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)

    def load_time_efficiencies(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    return settings.get('time_efficiencies', {})
            except:
                return {}
        return {}

    def save_time_efficiencies(self):
        settings = {'time_efficiencies': self.time_efficiencies}
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)

    def load_existing_data(self):
        if self.product_data.products:
            self.update_product_list()
            self.update_status_bar(
                f"Yüklü: {len(self.product_data.products)} ürün, "
                f"{len(self.product_data.historical_data)} geçmiş kayıt"
            )

    def setupUI(self):
        self.setWindowTitle("Üretim Planlama – Aylık Hedef Verim Hesabı")
        # self.setGeometry(100, 60, 1600, 900)

        # MenuBar oluştur
        self.create_menu_bar()

        # StatusBar oluştur
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.update_status_bar("Hazır")

        # Ana widget ve layout
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self.create_plan_content(main_layout)
        self.showMaximized()  # ← BU SATIRI EKLE (tam ekran)

    def create_menu_bar(self):
        menubar = self.menuBar()

        # Dosya menüsü
        file_menu = menubar.addMenu("Dosya")

        new_plan_action = QAction("Yeni Plan Oluştur", self)
        new_plan_action.setShortcut("Ctrl+N")
        new_plan_action.triggered.connect(self.create_new_plan)
        file_menu.addAction(new_plan_action)

        file_menu.addSeparator()

        save_plan_action = QAction("Planı Kaydet", self)
        save_plan_action.setShortcut("Ctrl+S")
        save_plan_action.triggered.connect(self.save_plan)
        file_menu.addAction(save_plan_action)

        load_plan_action = QAction("Plan Yükle", self)
        load_plan_action.setShortcut("Ctrl+O")
        load_plan_action.triggered.connect(self.load_plan)
        file_menu.addAction(load_plan_action)

        file_menu.addSeparator()

        import_plan_action = QAction("Excel'den Plan İçe Aktar", self)
        import_plan_action.setShortcut("Ctrl+I")
        import_plan_action.triggered.connect(self.import_from_excel)
        file_menu.addAction(import_plan_action)

        file_menu.addSeparator()

        exit_action = QAction("Çıkış", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Veri menüsü
        data_menu = menubar.addMenu("Veri")

        view_products_action = QAction("Ürün Havuzunu Görüntüle", self)
        view_products_action.triggered.connect(self.view_products)
        data_menu.addAction(view_products_action)

        view_historical_action = QAction("Geçmiş Verileri Görüntüle", self)
        view_historical_action.triggered.connect(self.view_historical)
        data_menu.addAction(view_historical_action)

        data_menu.addSeparator()

        # YENİ EKLE:
        edit_products_action = QAction("Ürün Havuzunu Düzenle", self)
        edit_products_action.triggered.connect(self.edit_products_dialog)
        data_menu.addAction(edit_products_action)

        edit_historical_action = QAction("Geçmiş Verileri Düzenle", self)
        edit_historical_action.triggered.connect(self.edit_historical_dialog)
        data_menu.addAction(edit_historical_action)

        # Ayarlar menüsü
        settings_menu = menubar.addMenu("Ayarlar")

        time_eff_action = QAction("Zamansal Verim Ayarları", self)
        time_eff_action.triggered.connect(self.set_time_efficiency)
        settings_menu.addAction(time_eff_action)

        # Raporlar menüsü
        reports_menu = menubar.addMenu("Raporlar")

        exec_report_action = QAction("Aylık Verim Raporu HTML", self)
        exec_report_action.triggered.connect(self.export_executive_report)
        reports_menu.addAction(exec_report_action)

        # Yardım menüsü
        help_menu = menubar.addMenu("Yardım")

        about_action = QAction("Hakkında", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

        view_menu = menubar.addMenu("Görünüm")

        self.toggle_sidebar_action = QAction("Ürün Havuzunu Göster/Gizle", self)
        self.toggle_sidebar_action.setShortcut("Ctrl+B")  # kısayol: Ctrl+B
        self.toggle_sidebar_action.triggered.connect(self.toggle_sidebar)
        view_menu.addAction(self.toggle_sidebar_action)

    def import_from_excel(self):
        if self.sidebar_collapsed:
            self.open_sidebar()
        if self.plan_table.rowCount() == 0:
            QMessageBox.warning(self, "Uyarı", "Önce bir plan oluşturun!")
            return

        # SEÇİLİ SÜTUNU GÜN OFSETİ OLARAK AL
        sel = self.plan_table.selectedIndexes()
        anchor_day = sel[0].column() if sel else 0  # YENİ

        # Mevcut hat isimleri
        available_lines = [self.plan_table.verticalHeaderItem(r).text()
                           for r in self.plan_table.display_rows]

        dialog = ImportPlanDialog(self, available_lines)
        if dialog.exec():
            campaigns = dialog.get_campaigns()
            if not campaigns:
                return

            success_count = 0
            fail_count = 0

            for camp in campaigns:
                line_name = camp['line']
                start_day = (camp['start_day'] - 1) + anchor_day  # YENİ: gün ofseti ekle
                product = camp['product']
                duration = camp['duration']
                pending_override = ("none", 0.0)  # <— başta koy

                # Hedef satır
                target_row = None
                for row in self.plan_table.display_rows:
                    if self.plan_table.verticalHeaderItem(row).text() == line_name:
                        target_row = row
                        break
                if target_row is None:
                    fail_count += 1
                    continue

                if product not in self.product_data.products and product != "HATKAPALI":
                    suggestions = self._suggest_product_names(product)

                    chosen = None
                    if suggestions:
                        msg = QMessageBox(self)
                        msg.setWindowTitle("Ürün Havuzunda Yok")
                        msg.setText(
                            f"'{product}' bulunamadı. Bunu mu demek istediniz?\n\n" +
                            "\n".join([f"• {s}" for s in suggestions]) +
                            "\n\nSeçim yapın veya 'Yeni Ürün Olarak Ekle' deyin."
                        )
                        btns = []
                        for s in suggestions:
                            b = msg.addButton(s, QMessageBox.ButtonRole.ActionRole)
                            btns.append((b, s))
                        add_new_btn = msg.addButton("Yeni Ürün Olarak Ekle", QMessageBox.ButtonRole.AcceptRole)
                        cancel_btn = msg.addButton(QMessageBox.StandardButton.Cancel)
                        msg.exec()

                        if msg.clickedButton() == cancel_btn:
                            fail_count += 1
                            continue
                        elif msg.clickedButton() == add_new_btn:
                            chosen = None  # yeni ürün akışı
                        else:
                            for b, s in btns:
                                if msg.clickedButton() == b:
                                    chosen = s
                                    break

                    if chosen:  # öneriden seçildi → havuzda var say
                        product = chosen
                    else:
                        # yeni ürün akışı
                        dlg_np = NewProductTonajDialog(self, product_name=product)
                        if not dlg_np.exec():
                            fail_count += 1
                            continue
                        pending_override = dlg_np.get_choice()

                try:
                    self.plan_table.add_campaign(target_row, start_day, product, duration)

                    if (product not in self.product_data.products) and product != "HATKAPALI":
                        if pending_override[0] == "manual":
                            val = float(pending_override[1])
                            for c in range(start_day, min(start_day + duration, self.plan_table.columnCount())):
                                self.plan_table.tonaj_overrides[(target_row, c)] = val
                        # none ise 0 zaten; ekstra işlem yok
                    self.plan_table.update_tonaj_totals()
                    success_count += 1
                except Exception as e:
                    print(f"Kampanya eklenirken hata: {e}")
                    fail_count += 1

            msg = f"İçe aktarma tamamlandı!\n\n✅ Başarılı: {success_count}\n"
            if fail_count:
                msg += f"❌ Başarısız: {fail_count}"
            QMessageBox.information(self, "Sonuç", msg)
            self.update_status_bar(f"{success_count} kampanya içe aktarıldı")

    def edit_products_dialog(self):
        """Mevcut ürün havuzunu tabloda göster; yeni satır ekle / düzenle / sil."""
        if self.sidebar_collapsed:
            self.open_sidebar()

        dlg = QDialog(self)
        dlg.setWindowTitle("Ürün Havuzu Düzenleme")
        dlg.setWindowFlags(
            Qt.WindowType.Window |
            Qt.WindowType.WindowMaximizeButtonHint |
            Qt.WindowType.WindowCloseButtonHint
        )
        dlg.setMinimumSize(900, 600)
        layout = QVBoxLayout(dlg)

        # --- Bilgi etiketi ---
        info = QLabel(
            "Mevcut ürünler aşağıda listelenmiştir. "
            "Hücreye çift tıklayarak düzenleyebilir, "
            "yeni satırlara yeni ürün girebilirsiniz. "
            "Satırı silmek için satırı seçip Delete tuşuna basın."
        )
        info.setWordWrap(True)
        info.setStyleSheet("color: gray; font-size: 12px;")
        layout.addWidget(info)

        # --- Tablo ---
        COLS = ["Kalem", "Devir", "Verim (%)", "Damla Gramaj"]
        table = PasteableTableWidget()
        table.setColumnCount(len(COLS))
        table.setHorizontalHeaderLabels(COLS)
        table.verticalHeader().setVisible(True)
        table.setAlternatingRowColors(True)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(table)

        # --- Mevcut ürünleri yükle ---
        products_sorted = sorted(self.product_data.products.items())
        # Ekstra boş satır ekle (kullanıcı yeni ürün girebilsin)
        EXTRA_ROWS = 50
        table.setRowCount(len(products_sorted) + EXTRA_ROWS)

        for r, (kalem, rec) in enumerate(products_sorted):
            table.setItem(r, 0, QTableWidgetItem(kalem))
            table.setItem(r, 1, QTableWidgetItem(str(rec.get("devir", ""))))
            table.setItem(r, 2, QTableWidgetItem(str(rec.get("verim", ""))))
            table.setItem(r, 3, QTableWidgetItem(str(rec.get("gramaj", ""))))

        # --- Araç çubuğu ---
        toolbar = QHBoxLayout()

        btn_add_row = QPushButton("+ Boş Satır Ekle")
        btn_add_row.setObjectName("secondaryButton")
        btn_add_row.clicked.connect(lambda: table.setRowCount(table.rowCount() + 1))

        btn_del_row = QPushButton("Seçili Satırı Sil")
        btn_del_row.setObjectName("secondaryButton")
        btn_del_row.clicked.connect(lambda: self._delete_selected_rows(table))

        btn_save = QPushButton("Kaydet / Güncelle")
        btn_save.clicked.connect(lambda: self._save_products_from_dialog(table, dlg))

        toolbar.addWidget(btn_add_row)
        toolbar.addWidget(btn_del_row)
        toolbar.addStretch()
        toolbar.addWidget(btn_save)
        layout.addLayout(toolbar)

        dlg.showMaximized()
        dlg.exec()

    def _save_products_from_dialog(self, table, dialog):
        """Tablodaki tüm satırları ürün havuzuna yaz (mevcut + yeni)."""
        try:
            new_products = {}
            errors = []
            for row in range(table.rowCount()):
                kalem_item = table.item(row, 0)
                if not kalem_item or not kalem_item.text().strip():
                    continue  # boş satırı atla
                kalem = kalem_item.text().strip()

                try:
                    devir = float(str(table.item(row, 1).text()).strip().replace(",", "."))
                    verim = float(str(table.item(row, 2).text()).strip().replace(",", "."))
                    gramaj = float(str(table.item(row, 3).text()).strip().replace(",", "."))
                except Exception:
                    errors.append(f"Satır {row + 1}: '{kalem}' – sayısal değer hatalı, atlandı.")
                    continue

                new_products[kalem] = {"devir": devir, "verim": verim, "gramaj": gramaj}

            if not new_products:
                QMessageBox.warning(dialog, "Uyarı", "Kaydedilecek geçerli ürün bulunamadı.")
                return

            # Havuzu komple değiştir (tabloda ne varsa onu yaz)
            self.product_data.products = new_products
            self.product_data.save_products()
            self.update_product_list()
            self.update_status_bar(f"Yüklü: {len(new_products)} ürün")
            self.repaint_all_campaigns_and_totals()

            msg = f"{len(new_products)} ürün kaydedildi."
            if errors:
                msg += "\n\nAtlanan satırlar:\n" + "\n".join(errors)
            QMessageBox.information(dialog, "Başarılı", msg)
            dialog.accept()

        except Exception as e:
            QMessageBox.critical(dialog, "Hata", f"Kayıt sırasında hata:\n{e}")

    def _delete_selected_rows(self, table):
        """Seçili satırları tablodan kaldır."""
        selected = sorted(
            set(idx.row() for idx in table.selectedIndexes()),
            reverse=True
        )
        for row in selected:
            table.removeRow(row)

    # ─────────────────────────────────────────────────────────────
    #  GEÇMİŞ VERİLER DÜZENLEME
    # ─────────────────────────────────────────────────────────────

    def edit_historical_dialog(self):
        """Mevcut geçmiş verileri tabloda göster; düzenle / ekle / sil."""
        dlg = QDialog(self)
        dlg.setWindowTitle("Geçmiş Veriler Düzenleme")
        dlg.setWindowFlags(
            Qt.WindowType.Window |
            Qt.WindowType.WindowMaximizeButtonHint |
            Qt.WindowType.WindowCloseButtonHint
        )
        dlg.setMinimumSize(900, 600)
        layout = QVBoxLayout(dlg)

        # --- Bilgi etiketi ---
        info = QLabel(
            "Mevcut geçmiş veriler aşağıda listelenmiştir. "
            "Hücreye çift tıklayarak düzenleyebilir, "
            "yeni satırlara yeni kayıt girebilirsiniz. "
            "Satırı silmek için satırı seçip Delete tuşuna basın."
        )
        info.setWordWrap(True)
        info.setStyleSheet("color: gray; font-size: 12px;")
        layout.addWidget(info)

        # --- Tablo ---
        COLS = ["Kalem", "Yıl", "Hat", "ADETSEL", "ZAMANSAL", "VERIM"]
        table = PasteableTableWidget()
        table.setColumnCount(len(COLS))
        table.setHorizontalHeaderLabels(COLS)
        table.verticalHeader().setVisible(True)
        table.setAlternatingRowColors(True)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(table)

        # --- Mevcut verileri yükle ---
        df = self.product_data.historical_data
        EXTRA_ROWS = 100
        if df is not None and not df.empty:
            table.setRowCount(len(df) + EXTRA_ROWS)
            for r, (_, row_data) in enumerate(df.iterrows()):
                for c, col in enumerate(COLS):
                    val = str(row_data.get(col, "")) if col in df.columns else ""
                    table.setItem(r, c, QTableWidgetItem(val))
        else:
            table.setRowCount(EXTRA_ROWS)

        # --- Araç çubuğu ---
        toolbar = QHBoxLayout()

        btn_add_row = QPushButton("+ Boş Satır Ekle")
        btn_add_row.setObjectName("secondaryButton")
        btn_add_row.clicked.connect(lambda: table.setRowCount(table.rowCount() + 1))

        btn_del_row = QPushButton("Seçili Satırı Sil")
        btn_del_row.setObjectName("secondaryButton")
        btn_del_row.clicked.connect(lambda: self._delete_selected_rows(table))

        btn_save = QPushButton("Kaydet / Güncelle")
        btn_save.clicked.connect(lambda: self._save_historical_from_dialog(table, dlg))

        toolbar.addWidget(btn_add_row)
        toolbar.addWidget(btn_del_row)
        toolbar.addStretch()
        toolbar.addWidget(btn_save)
        layout.addLayout(toolbar)

        dlg.showMaximized()
        dlg.exec()

    def _save_historical_from_dialog(self, table, dialog):
        """Tablodaki tüm satırları geçmiş veriye yaz (mevcut + yeni)."""
        import pandas as pd
        COLS = ["Kalem", "Yıl", "Hat", "ADETSEL", "ZAMANSAL", "VERIM"]
        try:
            records = []
            skipped = 0
            for row in range(table.rowCount()):
                # İlk 3 sütun (Kalem, Yıl, Hat) dolu olmalı
                vals = []
                skip = False
                for c in range(len(COLS)):
                    item = table.item(row, c)
                    txt = item.text().strip() if item else ""
                    vals.append(txt)

                if not vals[0]:  # Kalem boşsa satırı atla
                    continue
                if not vals[1] or not vals[2]:  # Yıl veya Hat boşsa atla
                    skipped += 1
                    continue

                records.append(dict(zip(COLS, vals)))

            if not records:
                QMessageBox.warning(dialog, "Uyarı", "Kaydedilecek geçerli kayıt bulunamadı.")
                return

            new_df = pd.DataFrame(records)
            # Yinelenen (Kalem + Yıl + Hat) kayıtlarda en son olanı tut
            new_df = new_df.drop_duplicates(subset=["Kalem", "Yıl", "Hat"], keep="last")

            self.product_data.historical_data = new_df
            self.product_data.save_historical()
            self.update_status_bar(f"Yüklü: {len(new_df)} geçmiş kayıt")

            msg = f"{len(new_df)} kayıt başarıyla kaydedildi."
            if skipped:
                msg += f"\n{skipped} satır Yıl/Hat eksik olduğu için atlandı."
            QMessageBox.information(dialog, "Başarılı", msg)
            dialog.accept()

        except Exception as e:
            QMessageBox.critical(dialog, "Hata", f"Kayıt sırasında hata:\n{e}")

    def create_plan_content(self, parent_layout):
        """Plan içeriğini doğrudan ana layout'a ekle (tab olmadan)"""

        # Ana içerik - Splitter
        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter = splitter  # ← referansını tut
        self.splitter.splitterMoved.connect(self._on_splitter_moved)

        # Sol panel - Ürün listesi
        left_panel = QFrame()
        left_panel.setFrameShape(QFrame.Shape.StyledPanel)
        left_panel.setMaximumWidth(300)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(16, 16, 16, 16)

        search_label = QLabel("Ürün Arama")
        search_label.setProperty("subheading", True)
        left_layout.addWidget(search_label)

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Ürün ara...")
        self.search_box.textChanged.connect(self.filter_products)
        self.search_box.focusInEvent = self._wrap_focus_in(self.search_box.focusInEvent)
        left_layout.addWidget(self.search_box)

        self.lbl_product_count = QLabel("Toplam: 0 ürün")
        self.lbl_product_count.setProperty("subheading", True)
        left_layout.addWidget(self.lbl_product_count)

        self.product_list = DraggableProductList()
        left_layout.addWidget(self.product_list)

        hint_label = QLabel("💡 İpucu: Ürünleri sürükleyip plana bırakın")
        hint_label.setProperty("subheading", True)
        hint_label.setWordWrap(True)
        left_layout.addWidget(hint_label)

        splitter.addWidget(left_panel)

        # Sağ panel - Plan tablosu
        right_panel = QFrame()
        right_panel.setFrameShape(QFrame.Shape.StyledPanel)
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(16, 16, 16, 16)

        self.plan_table = ProductionPlanTable()
        self.plan_table.verticalHeader().setDefaultSectionSize(32)
        right_layout.addWidget(self.plan_table)

        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        parent_layout.addWidget(splitter)
        self._apply_saved_sidebar_state()

        self.btn_calculate = QPushButton("Aylık Tahmini Verim Raporunu Oluştur")
        self.btn_calculate.clicked.connect(self.export_executive_report)

        parent_layout.addWidget(self.btn_calculate)

    def _on_splitter_moved(self, pos, index):
        sizes = self.splitter.sizes()
        # sol panel > 0 ise açık demektir
        if sizes and sizes[0] > 0:
            self.sidebar_collapsed = False
            self._sidebar_last_width = sizes[0]
        else:
            self.sidebar_collapsed = True
        self._persist_sidebar_state()

    def toggle_sidebar(self):
        if self.sidebar_collapsed:
            self.open_sidebar()
        else:
            self.close_sidebar()

    def open_sidebar(self):
        # Mevcut splitter boyutları
        sizes = self.splitter.sizes()
        if self._sidebar_last_width <= 0:
            self._sidebar_last_width = 260
        # Sol paneli hedef genişliğe getir, sağ panel kalan
        total = sum(sizes) if sum(sizes) > 0 else 1200
        left = min(self._sidebar_last_width, total - 200)  # sağ panel nefes alsın
        right = max(total - left, 200)
        self.splitter.setSizes([left, right])
        self.sidebar_collapsed = False
        self._persist_sidebar_state()

    def close_sidebar(self):
        # Kapatmadan önce mevcut genişliği hatırla (0 değilse)
        sizes = self.splitter.sizes()
        if len(sizes) >= 2 and sizes[0] > 0:
            self._sidebar_last_width = sizes[0]
        # Sol paneli 0 yap
        self.splitter.setSizes([0, max(1, sum(sizes))])
        self.sidebar_collapsed = True
        self._persist_sidebar_state()

    def _wrap_focus_in(self, original_event):
        def _inner(event):
            if self.sidebar_collapsed:
                self.open_sidebar()
            # orijinal focusIn çalışsın
            if callable(original_event):
                return original_event(event)
            return None

        return _inner

    def update_status_bar(self, message):
        self.status_bar.showMessage(message)

    def show_about(self):
        QMessageBox.about(
            self,
            "Hakkında",
            "<h3>Üretim Planlama Sistemi</h3>"
            "<p>Modern kurumsal üretim planlama yazılımı</p>"
            "<p>Versiyon: 2.0</p>"
            "<p>© 2025 - Tüm hakları saklıdır</p>"
        )

    def update_product_list(self):
        self.product_list.clear()
        for product_name in sorted(self.product_data.products.keys()):
            self.product_list.addItem(product_name)
        self.lbl_product_count.setText(f"Toplam: {len(self.product_data.products)} ürün")

    def filter_products(self, text):
        if self.sidebar_collapsed and text.strip():
            self.open_sidebar()
        if text == "\x1b":
            self.search_box.clear()
            return
        for i in range(self.product_list.count()):
            item = self.product_list.item(i)
            item.setHidden(text.lower() not in item.text().lower())

    def create_new_plan(self):
        try:
            today = QDate.currentDate()
            first = QDate(today.year(), today.month(), 1)
            last = first.addMonths(1).addDays(-1)
            self.setup_plan_table(first, last)
            # Bilgi amaçlı status bar
            self.update_status_bar(f"Plan: {first.toString('dd.MM.yyyy')} – {last.toString('dd.MM.yyyy')}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Plan oluşturulamadı:\n{e}")
            dialog = NewPlanDialog(self)
            if dialog.exec():
                try:
                    start_date, end_date = dialog.get_dates()
                    self.setup_plan_table(start_date, end_date)
                except Exception as e:
                    QMessageBox.critical(self, "Hata", f"Plan oluşturulamadı:\n{e}")

    def setup_plan_table(self, start_date, end_date):
        start_py = start_date.toPyDate()
        end_py = end_date.toPyDate()
        days = (end_py - start_py).days + 1

        total_rows = (len(A_LINES) * 2) + 1 + (len(C_LINES) * 2) + 1
        self.plan_table.setRowCount(total_rows)
        self.plan_table.setColumnCount(days)

        row_headers = []
        self.plan_table.display_rows.clear()
        self.plan_table.tonaj_rows.clear()

        cur_row = 0
        for ln in A_LINES:
            plan_row = cur_row
            ton_row = cur_row + 1
            row_headers += [ln, f"Tonaj"]
            self.plan_table.display_rows.append(plan_row)
            self.plan_table.tonaj_rows[plan_row] = ton_row
            cur_row += 2

        row_headers.append("A FIRINI TONAJ")
        a_total_row = cur_row
        cur_row += 1

        for ln in C_LINES:
            plan_row = cur_row
            ton_row = cur_row + 1
            row_headers += [ln, f"Tonaj"]
            self.plan_table.display_rows.append(plan_row)
            self.plan_table.tonaj_rows[plan_row] = ton_row
            cur_row += 2

        row_headers.append("C FIRINI TONAJ")
        c_total_row = cur_row

        self.plan_table.setVerticalHeaderLabels(row_headers)

        self.plan_table.row_a_total = a_total_row
        self.plan_table.row_c_total = c_total_row
        self.plan_table.is_a_line = lambda r: (r in self.plan_table.display_rows and r < a_total_row)
        self.plan_table.is_c_line = lambda r: (
                r in self.plan_table.display_rows and r > a_total_row and r < c_total_row)

        for i, name in enumerate(row_headers):
            hdr = self.plan_table.verticalHeaderItem(i)
            if i == a_total_row or i == c_total_row:
                hdr.setBackground(QColor(EnterpriseTheme.PRIMARY))
                hdr.setForeground(QColor(EnterpriseTheme.TEXT_ON_PRIMARY))
            elif name.endswith("(Tonaj)"):
                hdr.setBackground(QColor(EnterpriseTheme.SURFACE_HOVER))
            elif name.startswith('A'):
                hdr.setBackground(QColor(EnterpriseTheme.A_LINE_BASE))
            else:
                hdr.setBackground(QColor(EnterpriseTheme.C_LINE_BASE))

        headers = []
        cur = start_py
        for _ in range(days):
            headers.append(cur.strftime("%d.%m"))
            cur += timedelta(days=1)
        self.plan_table.setHorizontalHeaderLabels(headers)

        self.plan_start_date = start_py  # YENİ
        self.plan_end_date = end_py  # YENİ
        self._wire_header_context_menu_once()  # YENİ (aşağıda tanımlayacağız)

        self.plan_table.campaigns.clear()
        self.plan_table.tonaj_overrides.clear()  # ← EKLE
        self.plan_table.get_product = lambda name: self.product_data.products.get(name, {})

        # Tonaj satırlarını başlat
        for plan_row, ton_row in self.plan_table.tonaj_rows.items():
            for c in range(self.plan_table.columnCount()):
                self.plan_table.setItem(ton_row, c, self.plan_table._make_locked_item("0,00"))

        # Fırın toplamlarını güncelle (A ve C toplam satırları)
        self.plan_table.update_tonaj_totals()
        self.plan_table.resizeRowsToContents()
        self.plan_table.get_product_names = lambda: list(self.product_data.products.keys())  # YENİ
        # self.plan_table.resolve_params = self._resolve_params  # YENİ: hat+ürün için parametre çözücü
        # DEBUG: Toplam satırlarını kontrol et
        print(f"A total row: {a_total_row}, C total row: {c_total_row}")
        print(f"Total rows: {self.plan_table.rowCount()}")

    def _wire_header_context_menu_once(self):
        hdr = self.plan_table.horizontalHeader()
        if hdr.contextMenuPolicy() != Qt.ContextMenuPolicy.CustomContextMenu:
            hdr.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            hdr.customContextMenuRequested.connect(self._show_date_header_menu)

    def _show_date_header_menu(self, pos):
        menu = QMenu(self)
        act = QAction("Plan tarih aralığını değiştir...", self)
        act.triggered.connect(self._change_plan_dates_dialog)
        menu.addAction(act)
        menu.exec(self.plan_table.horizontalHeader().mapToGlobal(pos))

    def _change_plan_dates_dialog(self):
        # Eski aralıktan kampanyaların mutlak tarihlerini çıkar
        old_start = self.plan_start_date
        campaigns_abs = []
        for (row, s, e), info in self.plan_table.campaigns.items():
            for c in range(s, e + 1):
                campaigns_abs.append({
                    "row": row,
                    "date": old_start + timedelta(days=c),  # mutlak tarih
                    "product": info["product"]
                })

        # Yeni aralığı al
        dlg = NewPlanDialog(self)
        # varsayılanları dolu göster
        dlg.start_date.setDate(QDate(self.plan_start_date.year, self.plan_start_date.month, self.plan_start_date.day))
        dlg.end_date.setDate(QDate(self.plan_end_date.year, self.plan_end_date.month, self.plan_end_date.day))
        if not dlg.exec():
            return
        new_qs, new_qe = dlg.get_dates()
        new_start = new_qs.toPyDate()
        new_end = new_qe.toPyDate()

        # Tabloyu yeniden kur
        self.setup_plan_table(new_qs, new_qe)

        # Mutlak tarihler yeni aralığa düşüyorsa geri boya (ardışıkları kampanyaya birleştir)
        by_row_col = {}
        for rec in campaigns_abs:
            if new_start <= rec["date"] <= new_end:
                col = (rec["date"] - new_start).days
                key = (rec["row"], col, rec["product"])
                by_row_col[key] = True

        # Aynı ürün için ardışık sütunları tek kampanya halinde ekle
        from itertools import groupby
        for row in self.plan_table.display_rows:
            # row’daki (col,product) setini al
            cols_products = sorted([(col, pr)
                                    for (r, col, pr) in by_row_col.keys() if r == row],
                                   key=lambda x: (x[1], x[0]))
            # ürüne göre grupla, sonra ardışıkları birleştir
            for prod, grp in groupby(cols_products, key=lambda x: x[1]):
                seq = [c for (c, _) in grp]
                if not seq:
                    continue
                seq.sort()
                # ardışık bloklara böl
                start = seq[0]
                prev = seq[0]
                for c in seq[1:] + [None]:
                    if c is None or c != prev + 1:
                        self.plan_table.add_campaign(row, start, prod, prev - start + 1)
                        if c is not None:
                            start = c
                    prev = c if c is not None else prev

        # yeni aralığı kaydet
        self.plan_start_date = new_start
        self.plan_end_date = new_end

    def _suggest_product_names(self, typed_name, n=3, cutoff=0.6):
        import difflib
        pool = list(self.product_data.products.keys())
        return difflib.get_close_matches(typed_name, pool, n=n, cutoff=cutoff)

    def set_time_efficiency(self):
        dialog = TimeEfficiencyDialog(self, self.time_efficiencies)
        if dialog.exec():
            self.time_efficiencies = dialog.get_efficiencies()
            self.save_time_efficiencies()
            QMessageBox.information(self, "Başarılı", "Zamansal verim değerleri kaydedildi!")

    def _compute_executive_metrics(self, history_scope: str = 'line'):
        from collections import defaultdict

        if not hasattr(self, "plan_table") or not self.plan_table.campaigns:
            return None

        # --- ARTIK SABİT HEDEFLERLE HESAP YAPMIYORUZ, SADECE KAYIT İÇİN ÇEKİYORUZ ---
        # Bu değerler sadece "Hedeflenen neydi?" diye bakmak istersen kenarda durur.
        # Hesaplamanın tamamı hat bazlı (bottom-up) olacak.

        # Hat Bazlı Ayarları Çek
        # Not: Hatların zamansal verimlerini settings'den alıyoruz.

        per_line = defaultdict(lambda: {"teorik": 0.0, "fiili": 0.0, "gun": 0, "weighted_time_num": 0.0})

        per_furnace = {
            "A": {"teorik": 0.0, "fiili": 0.0, "weighted_time_num": 0.0},
            "C": {"teorik": 0.0, "fiili": 0.0, "weighted_time_num": 0.0}
        }

        per_campaign = []  # Kampanya listesini burada dolduracağız

        for (row, start_col, end_col), info in self.plan_table.campaigns.items():
            duration = int(info.get("duration", 0))
            if duration <= 0: continue

            line = self.plan_table.verticalHeaderItem(row).text()
            product = info.get("product", "")

            # --- Parametre Hesabı ---
            if product == self.plan_table.CLOSED_TAG:
                devir, adetsel = 0.0, 0.0
                # YENİ
                SABIT_ZAMANSAL = 0.9586
                line_zaman_eff = SABIT_ZAMANSAL
            else:
                devir_eff, verim_eff, _, _ = self.plan_table._effective_params(row, start_col, product)
                devir = float(devir_eff)

                # Adetsel Verim (Manuel > Geçmiş > Varsayılan)
                adetsel = None
                if "realized_adetsel" in info:
                    try:
                        rv = float(str(info["realized_adetsel"]).replace("%", "").replace(",", "."))
                        adetsel = rv / 100.0 if rv > 1 else rv
                    except:
                        adetsel = None

                if adetsel is None:
                    if history_scope == 'all':
                        hist = self.get_historical_efficiency_productwide(product)
                    else:
                        hist = self.get_historical_efficiency(product, line)

                    if hist is not None:
                        adetsel = float(hist)

                # YENİ
                SABIT_ZAMANSAL = 0.9586
                line_zaman_eff = SABIT_ZAMANSAL

                if adetsel is None:
                    # Havuzdan gelen verim_eff = Genel Verim (adetsel × zamansal)
                    # Adetsel = Genel Verim / Zamansal
                    genel_verim = float(verim_eff) / 100.0
                    if line_zaman_eff > 0:
                        adetsel = genel_verim / line_zaman_eff
                    else:
                        adetsel = genel_verim

            # --- Temel Hesaplamalar ---
            teorik = devir * 1440.0 * duration

            # Hat Bazlı Fiili: Teorik * Adetsel * Hattın Zamansal Ayarı
            fiili_line = teorik * adetsel * line_zaman_eff

            # Zamansal Ağırlıklı Ortalama için Pay: Teorik * Zamansal%
            time_weighted_num = teorik * line_zaman_eff

            # --- 1. KAMPANYA LİSTESİNİ DOLDUR (Tablo için şart) ---
            per_campaign.append({
                "Hat": line,
                "Başlangıç": start_col,  # Sıralama için
                "Ürün": product,
                "Süre": duration,
                "Toplam Teorik Adet": teorik,
                "Tahmini Fiili Adet": fiili_line,
                "Tahmini Ortalama Verim (%)": (fiili_line / teorik * 100.0) if teorik > 0 else 0.0,
                "Adetsel Verim (%)": adetsel * 100,
                "Zamansal Verim (%)": line_zaman_eff * 100
            })

            # --- 2. HAT TOPLAMLARI ---
            per_line[line]["teorik"] += teorik
            per_line[line]["fiili"] += fiili_line
            per_line[line]["weighted_time_num"] += time_weighted_num
            per_line[line]["gun"] += duration

            # --- 3. FIRIN TOPLAMLARI (DOĞRUDAN TOPLAMA - MANİPÜLASYON YOK) ---
            # Hat hangi fırındaysa oraya ekle
            furnace_key = "C" if line.startswith("C") else "A"

            # C1'i C fırını hesabına katmak istemiyorsan:
            # if furnace_key == "C" and line == "C1": continue

            # İsteğe bağlı C1 hariç tutma (EXCLUDE_FROM_C_EFFICIENCY seti varsa)
            # if furnace_key == "C" and line in getattr(self, "EXCLUDE_FROM_C_EFFICIENCY", set()):
            if furnace_key == "C" and line in EXCLUDE_FROM_C_EFFICIENCY:

                pass  # C toplamına katma
            else:
                per_furnace[furnace_key]["teorik"] += teorik
                per_furnace[furnace_key]["fiili"] += fiili_line
                per_furnace[furnace_key]["weighted_time_num"] += time_weighted_num

            # --- SONUÇLARI HESAPLA ---

        # 1. Hatlar (Yüzdeleri hesapla)
        for ln, agg in per_line.items():
            agg["verim_pct"] = (agg["fiili"] / agg["teorik"] * 100.0) if agg["teorik"] > 0 else 0.0

        # 2. Fırınlar (Yüzdeleri hesapla - ARTIK AĞIRLIKLI ORTALAMA)
        for f_key in ["A", "C"]:
            teo = per_furnace[f_key]["teorik"]
            fii = per_furnace[f_key]["fiili"]
            per_furnace[f_key]["verim_pct"] = (fii / teo * 100.0) if teo > 0 else 0.0

        # 3. GENEL TOPLAM (Fabrika Geneli)
        total_teorik = per_furnace["A"]["teorik"] + per_furnace["C"]["teorik"]
        total_fiili = per_furnace["A"]["fiili"] + per_furnace["C"]["fiili"]
        total_time_num = per_furnace["A"]["weighted_time_num"] + per_furnace["C"]["weighted_time_num"]

        # Ağırlıklı Zamansal Verim (Genel)
        calculated_total_temporal_eff = (total_time_num / total_teorik) if total_teorik > 0 else 0.0

        total_stats = {
            "teorik": total_teorik,
            "fiili": total_fiili,
            "calculated_total_temporal_eff": calculated_total_temporal_eff
        }

        return {
            "per_line": per_line,
            "per_furnace": per_furnace,
            "per_campaign": per_campaign,
            "total": total_stats,
            "date_range": (self.plan_start_date, self.plan_end_date),
        }

    def _parse_percent_series(self, s):
        import pandas as pd
        return pd.to_numeric(
            s.astype(str).str.replace("%", "", regex=False).str.replace(",", ".", regex=False),
            errors="coerce"
        )

    def _historical_score(self, product: str, line: str, metric: str = "VERIM"):
        """Verilen ürün-hat için (yıl ağırlıklı) geçmiş performans skoru döndürür."""
        import pandas as pd
        df = self.product_data.historical_data
        if df is None or df.empty:
            return None
        h = df[(df["Kalem"] == product) & (df["Hat"] == line)]
        if h.empty or metric not in h.columns:
            return None

        vals = self._parse_percent_series(h[metric])
        years = pd.to_numeric(h["Yıl"], errors="coerce")
        # Yıl ağırlıkları (gerekirse ayarlayabilirsiniz)
        year_w = years.map({2025: 1.0, 2024: 0.8, 2023: 0.6}).fillna(0.5)
        if year_w.sum() <= 0:
            return None
        return float((vals * year_w).sum() / year_w.sum())

    # --- YENİ: Ağırlıklı yüzde hesaplama (genel yardımcı) ---
    def _weighted_avg_percent(self, values_series, years_series,
                              weights_map=None, default_weight=0.5):
        """
        values_series: yüzde metinleri (örn. '90,3%' veya '90.3' ya da 90.3)
        years_series:  yıl değerleri (2025, 2024, ...)
        weights_map:   {yıl: ağırlık}
        dönen:         ağırlıklı ortalama (yüzde birimi, örn. 87.4)
        """
        import pandas as pd
        if weights_map is None:
            weights_map = {2025: 1.0, 2024: 0.8, 2023: 0.6}

        vals = self._parse_percent_series(pd.Series(values_series))
        yrs = pd.to_numeric(pd.Series(years_series), errors="coerce")
        w = yrs.map(weights_map).fillna(default_weight)

        num = (vals * w).sum(skipna=True)
        den = w[vals.notna()].sum()
        if den and den > 0:
            return float(num / den)
        return None

    # --- YENİ: Ürün-hat için (yıl ağırlıklı) ADETSEL ortalaması ---
    def _historical_adetsel_weighted(self, product: str, line: str):
        """
        historical.json içinden ilgili ürün+hat kayıtlarını bulur,
        ADETSEL sütununu yıl ağırlıklarıyla ortalar.
        Dönüş: yüzde (örn. 88.7), None yoksa
        """
        df = self.product_data.historical_data
        if df is None or df.empty:
            return None
        sub = df[(df["Kalem"] == product) & (df["Hat"] == line)]
        if sub.empty or "ADETSEL" not in sub.columns or "Yıl" not in sub.columns:
            return None
        return self._weighted_avg_percent(sub["ADETSEL"], sub["Yıl"])

    # --- YENİ: (ADETSEL ağırlıklı) × (ayarlar/ZAMANSAL) -> “bileşik VERİM” ---
    def _composed_verim_from_history(self, product: str, line: str):
        """
        Mevcut verim tanımımız:
          (yıl ağırlıklı ADETSEL yüzdesi) × (ayarlar sekmesindeki ZAMANSAL yüzdesi) / 100
        Dönüş: yüzde (örn. 76.5), None yoksa
        """
        adetsel_pct = self._historical_adetsel_weighted(product, line)
        if adetsel_pct is None:
            return None
        # ZAMANSAL'ı ayarlardan al (%)
        zamansal_pct = float(self.time_efficiencies.get(line, 95.86))
        return float(adetsel_pct * (zamansal_pct / 100.0))

    # --- YENİ: Ürün için en iyi hattı “bileşik VERİM”e göre seç ---
    def _best_line_for_product_composed(self, product: str):
        import pandas as pd
        df = self.product_data.historical_data
        if df is None or df.empty:
            return None, None
        sub = df[df["Kalem"] == product]
        if sub.empty:
            return None, None

        best_line, best_score = None, None
        for hat in sorted(sub["Hat"].dropna().unique()):
            # ← BURAYA EKLE: TOPLAMI ve geçersiz hat isimlerini filtrele
            hat_clean = str(hat).strip().upper()
            if "TOPLAMI" in hat_clean:
                continue
            if hat_clean not in [ln.upper() for ln in ALL_LINES]:
                continue
            # ←

            sc = self._composed_verim_from_history(product, hat)
            if sc is None:
                continue
            if best_score is None or sc > best_score:
                best_line, best_score = hat, sc
        return best_line, best_score

    def _best_line_for_product(self, product: str, metric: str = "VERIM"):
        import pandas as pd
        df = self.product_data.historical_data
        if df is None or df.empty:
            return None, None

        sub = df[df["Kalem"] == product]
        if sub.empty or metric not in sub.columns:
            return None, None

        best_line = None
        best_score = None
        for hat in sorted(sub["Hat"].dropna().unique()):
            # ← EKLE
            hat_clean = str(hat).strip().upper()
            if "TOPLAMI" in hat_clean:
                continue
            if hat_clean not in [ln.upper() for ln in ALL_LINES]:
                continue
            # ←

            sc = self._historical_score(product, hat, metric=metric)
            if sc is None:
                continue
            if (best_score is None) or (sc > best_score):
                best_line, best_score = hat, sc
        return best_line, best_score
    
    def _compute_line_suggestions(self, history_scope: str = "line",
                                  metric: str = "VERIM", min_gain_pct: float = 1.5):
        """
        Planlanmış kampanyalar için: planlanan hat ≠ en iyi hat ise ve
        kazanç >= min_gain_pct ise öneri üretir.

        *** Yeni mantık ***
        Mevcut Verim(%)  = (yıl ağırlıklı ADETSEL) × (ayarlar/ZAMANSAL) / 100
        Önerilen Verim(%)= aynı formülle en iyi hattaki değer
        """
        import pandas as pd
        if not hasattr(self, "plan_table") or not self.plan_table.campaigns:
            return pd.DataFrame()

        rows = []
        for (row, s, e), info in self.plan_table.campaigns.items():
            product = info.get("product", "")
            if not product or product == self.plan_table.CLOSED_TAG:
                continue

            current_line = self.plan_table.verticalHeaderItem(row).text()

            # Yeni hesap: bileşik verim
            current_score = self._composed_verim_from_history(product, current_line)
            best_line, best_score = self._best_line_for_product_composed(product)

            if current_score is None or best_score is None or best_line is None:
                continue

            if best_line != current_line and (best_score - current_score) >= min_gain_pct:
                rows.append({
                    "Kalem": product,
                    "Planlanan Hat": current_line,
                    "Önerilen Hat": best_line,
                    "Mevcut Verim(%)": round(current_score, 2),
                    "Önerilen Verim(%)": round(best_score, 2),
                    "Kazanç(%)": round(best_score - current_score, 2)
                })

        return (pd.DataFrame(rows)
                .sort_values(by=["Kazanç(%)", "Kalem"], ascending=[False, True])
                .reset_index(drop=True))

    def export_executive_report(self):
        import os
        from pathlib import Path
        import pandas as pd
        from datetime import datetime as _dt
        from PyQt6.QtWidgets import (
            QMessageBox, QTextEdit, QDialog, QVBoxLayout, QDialogButtonBox, QFileDialog
        )
        from PyQt6.QtCore import QUrl
        from PyQt6.QtGui import QDesktopServices, QTextDocument
        from PyQt6.QtPrintSupport import QPrinter

        # 1) Kapsam sor
        dlg = HistoricalScopeDialog(self)
        if not dlg.exec():
            return
        history_scope = dlg.get_scope()
        scope_badge = "2024-2025 Verilerine Göre Hesaplandı" if history_scope == "all" else "2024-2025 Verilerine Göre Hesaplandı"

        # 2) Metriği hesapla
        metrics = self._compute_executive_metrics(history_scope=history_scope)
        if not metrics:
            QMessageBox.warning(self, "Uyarı", "Önce plan ve verileri giriniz.")
            return

        pf = metrics["per_furnace"]
        pl = metrics["per_line"]
        pc = metrics["per_campaign"]
        total = metrics["total"]
        d0, d1 = metrics["date_range"]
        period = f"{d0.strftime('%d.%m.%Y')} – {d1.strftime('%d.%m.%Y')}"

        # ---------------------------------------------------------------------
        # KPI HESAPLAMA (DÜZELTİLMİŞ KISIM)
        # ---------------------------------------------------------------------

        def _included(ln: str) -> bool:
            up = ln.strip().upper()
            return (not up.startswith("C")) or (up != "C1")

        sum_teo = 0.0  # Σ(teorik)
        sum_fiili = 0.0  # Σ(fiili)

        # Sadece teorik ve fiili toplamlarını al
        for ln, agg in pl.items():
            if not _included(ln):
                continue
            teo = float(agg.get("teorik", 0.0))
            fii = float(agg.get("fiili", 0.0))
            sum_teo += teo
            sum_fiili += fii

        # --- DÜZELTME BAŞLANGICI ---

        # 1. Adım: Metrics'ten gelen değeri kontrol et ("target_total_eff" anahtarı ile)
        user_input_zamansal = metrics["total"].get("target_total_eff", None)
        # 2. Adım: Eğer metrics'te yoksa, direkt self.time_efficiencies'ten oku (Yedek Plan)
        if user_input_zamansal is None:
            # FIXED_TOTAL yoksa varsayılan 95.86
            raw_val = self.time_efficiencies.get("FIXED_TOTAL", 95.86)
            user_input_zamansal = float(raw_val) / 100.0

        # KPI Değerlerini Ata
        tahmini_zamansal_verim = 0.9586  # Sabit

        # Kesilen Damla = Toplam Teorik * Kullanıcının Girdiği Zamansal (%)
        kesilen_damla = sum_teo * tahmini_zamansal_verim

        # Adetsel Verim = Fiili / Kesilen Damla
        tahmini_adetsel_verim = (sum_fiili / kesilen_damla) if kesilen_damla > 0 else 0.0

        # Ortalama Verim = Fiili / Teorik (Matematiksel tutarlılık için tekrar hesaplanır)
        tahmini_ortalama_verim = (sum_fiili / sum_teo) if sum_teo > 0 else 0.0

        # --- DÜZELTME BİTİŞ ---

        # ---------------------------------------------------------------------
        # RAPOR GÖRSELLEŞTİRME
        # ---------------------------------------------------------------------

        # 3) DataFrame'ler (C1 gizle)
        rows_line = []
        for ln, agg in pl.items():
            if ln.strip().upper() == "C1":
                continue
            rows_line.append({
                "Hat": ln,
                "Tahmini Ortalama Verim (%)": agg["verim_pct"],  # Ham değer olarak tut
                "Toplam Teorik Adet": int(round(agg["teorik"], 0)),  # TAM SAYI
                "Tahmini Fiili Adet": int(round(agg["fiili"], 0)),  # TAM SAYI
                "Gün": int(agg["gun"]),
            })
        df_line_num = pd.DataFrame(rows_line).sort_values(["Hat"])

        df_campaign_num = pd.DataFrame(pc)
        if not df_campaign_num.empty:
            df_campaign_num = df_campaign_num[df_campaign_num["Hat"].str.upper() != "C1"]
            # Kampanya verileri için tam sayı ve verim formatlaması
            if "Toplam Teorik Adet" in df_campaign_num.columns:
                df_campaign_num["Toplam Teorik Adet"] = df_campaign_num["Toplam Teorik Adet"].apply(
                    lambda x: int(round(float(x), 0)) if pd.notna(x) else 0)
            if "Tahmini Fiili Adet" in df_campaign_num.columns:
                df_campaign_num["Tahmini Fiili Adet"] = df_campaign_num["Tahmini Fiili Adet"].apply(
                    lambda x: int(round(float(x), 0)) if pd.notna(x) else 0)
            if "Tahmini Ortalama Verim (%)" in df_campaign_num.columns:
                df_campaign_num["Tahmini Ortalama Verim (%)"] = df_campaign_num["Tahmini Ortalama Verim (%)"].apply(
                    lambda x: round(float(x), 2) if pd.notna(x) else 0.0)
            if "Adetsel Verim (%)" in df_campaign_num.columns:
                df_campaign_num["Adetsel Verim (%)"] = df_campaign_num["Adetsel Verim (%)"].apply(
                    lambda x: round(float(x), 2) if pd.notna(x) else 0.0)
            if "Zamansal Verim (%)" in df_campaign_num.columns:
                df_campaign_num["Zamansal Verim (%)"] = df_campaign_num["Zamansal Verim (%)"].apply(
                    lambda x: round(float(x), 2) if pd.notna(x) else 0.0)
            df_campaign_num = df_campaign_num.sort_values(["Hat", "Başlangıç"])

        df_furnace = pd.DataFrame([
            {"Fırın": "A",
             "Tahmini Ortalama Verim (%)": pf["A"]["verim_pct"],  # Ham değer
             "Toplam Teorik Adet": int(round(pf["A"]["teorik"], 0)),  # TAM SAYI
             "Tahmini Fiili Adet": int(round(pf["A"]["fiili"], 0))},  # TAM SAYI
            {"Fırın": "C",
             "Tahmini Ortalama Verim (%)": pf["C"]["verim_pct"],  # Ham değer
             "Toplam Teorik Adet": int(round(pf["C"]["teorik"], 0)),  # TAM SAYI
             "Tahmini Fiili Adet": int(round(pf["C"]["fiili"], 0))},  # TAM SAYI
        ])

        # 4) TR format - GÜNCELLENDİ
        def fmt_tr_int(x):
            """Tam sayı olarak formatla (ondalıksız)"""
            try:
                # Önce float'a çevir, sonra tam sayıya yuvarla
                val = int(round(float(x), 0))
                # Binlik ayırıcı ekle
                return f"{val:,}".replace(",", "X").replace(".", ",").replace("X", ".")
            except Exception:
                return str(x)

        def fmt_tr_pct(x):
            """Yüzde olarak formatla (2 ondalık basamak: AB.XY)"""
            try:
                # Eğer x zaten % içeriyorsa, temizle
                if isinstance(x, str) and '%' in x:
                    x = x.replace('%', '').replace(',', '.')
                val = float(x)
                return f"%{val:.2f}".replace(".", ",")
            except Exception:
                return str(x)

        # Fırın eşikleri
        FURNACE_BENCH = {"A": 88.5, "C": 91.00}
        BENCH_C = {
            "C2": 83.0, "C3": 91.7, "C4": 92.5, "C5": 93.5, "C6": 91.7,
            "C7": 90.5, "C8": 90.7, "C9": 91.0, "C10": 95.0, "C11": 90.5
        }
        BENCH_A = {
            "A1": 84.50, "A2": 89.50, "A3": 87.50, "A4": 93.10, "A5": 90.00,
            "A6": 90.00, "A7": 90.00
        }

        # 6) HTML tablo üretici
        def df_to_html_styled(df, numeric_df=None, hat_karsilastirma=False, furnace_karsilastirma=False):
            if df is None or df.empty:
                return '<div class="empty">Veri yok</div>'

            verim_col = None
            for c in df.columns:
                if "Verim" in c:
                    verim_col = c
                    break

            thead = "<thead><tr>" + "".join(f"<th>{h}</th>" for h in df.columns) + "</tr></thead>"

            rows_html = []
            for idx, row in df.iterrows():
                tds = []
                base_val = None
                hat = None
                if numeric_df is not None:
                    try:
                        base_val = float(numeric_df.loc[idx, verim_col]) if verim_col in numeric_df.columns else None
                    except Exception:
                        base_val = None
                    try:
                        hat = str(numeric_df.loc[idx, "Hat"])
                    except Exception:
                        hat = None

                for col in df.columns:
                    val = row[col]
                    cls = []
                    if col == "Mevcut Verim(%)":
                        cls.append("pill pill-mevcut")
                    elif col == "Önerilen Verim(%)":
                        cls.append("pill pill-onerilen")
                    elif col == "Kazanç(%)":
                        cls.append("pill pill-kazanc")

                    if verim_col and col == verim_col:
                        cls.append("verim")
                        if hat_karsilastirma and hat:
                            hat_key = hat.strip().upper()
                            if hat_key.startswith("C") and hat_key != "C1":
                                bench = BENCH_C.get(hat_key)
                            else:
                                bench = BENCH_A.get(hat_key)

                            if bench is not None and base_val is not None:
                                if float(base_val) >= float(bench):
                                    cls.append("up")
                                elif float(base_val) < float(bench) and float(base_val) >= 1.0:
                                    cls.append("down")
                                else:
                                    cls.append("zero")
                            else:
                                cls.append("zero")

                        if furnace_karsilastirma and numeric_df is not None:
                            try:
                                furnace_name = str(numeric_df.loc[idx, "Fırın"])
                                bench_f = FURNACE_BENCH.get(furnace_name)
                                if bench_f is not None and base_val is not None:
                                    cls.append("up" if float(base_val) >= float(bench_f) else "down")
                            except:
                                pass

                    tds.append(f"<td class=\"{' '.join(cls)}\">{val}</td>")
                rows_html.append("<tr>" + "".join(tds) + "</tr>")

            tbody = "<tbody>" + "".join(rows_html) + "</tbody>"
            return f'<div class="table-scroll"><table class="dataframe">{thead}{tbody}</table></div>'

        # 7) Görünüm DF'leri - GÜNCELLENDİ
        df_furnace_view = df_furnace.copy()
        df_furnace_view["Tahmini Ortalama Verim (%)"] = df_furnace_view["Tahmini Ortalama Verim (%)"].map(fmt_tr_pct)
        df_furnace_view["Toplam Teorik Adet"] = df_furnace_view["Toplam Teorik Adet"].map(fmt_tr_int)
        df_furnace_view["Tahmini Fiili Adet"] = df_furnace_view["Tahmini Fiili Adet"].map(fmt_tr_int)

        df_line_view = df_line_num.copy()
        if not df_line_view.empty:
            df_line_view["Tahmini Ortalama Verim (%)"] = df_line_view["Tahmini Ortalama Verim (%)"].map(fmt_tr_pct)
            df_line_view["Toplam Teorik Adet"] = df_line_view["Toplam Teorik Adet"].map(fmt_tr_int)
            df_line_view["Tahmini Fiili Adet"] = df_line_view["Tahmini Fiili Adet"].map(fmt_tr_int)

        df_campaign_view = df_campaign_num.copy()
        if not df_campaign_view.empty:
            # ADET KOLONLARI ÖNCE - Tam sayı (numeric_df'deki değerleri kullan)
            if "Toplam Teorik Adet" in df_campaign_view.columns:
                df_campaign_view["Toplam Teorik Adet"] = df_campaign_num["Toplam Teorik Adet"].apply(fmt_tr_int)
            if "Tahmini Fiili Adet" in df_campaign_view.columns:
                df_campaign_view["Tahmini Fiili Adet"] = df_campaign_num["Tahmini Fiili Adet"].apply(fmt_tr_int)

            # VERİM KOLONLARI - 2 ondalık basamak
            if "Tahmini Ortalama Verim (%)" in df_campaign_view.columns:
                df_campaign_view["Tahmini Ortalama Verim (%)"] = df_campaign_num["Tahmini Ortalama Verim (%)"].apply(
                    fmt_tr_pct)
            if "Adetsel Verim (%)" in df_campaign_view.columns:
                df_campaign_view["Adetsel Verim (%)"] = df_campaign_num["Adetsel Verim (%)"].apply(fmt_tr_pct)
            if "Zamansal Verim (%)" in df_campaign_view.columns:
                df_campaign_view["Zamansal Verim (%)"] = df_campaign_num["Zamansal Verim (%)"].apply(fmt_tr_pct)

        # Öneriler
        df_sugg = self._compute_line_suggestions(history_scope="line", metric="VERIM", min_gain_pct=1.5)
        sugg_note = "2023–2025 VERİM ortalamaları dikkate alınarak hat yerleşimi önerileri geliştirilmiştir."
        if df_sugg is not None and not df_sugg.empty:
            sugg_html = f"""
              <section class="table-card card">
                <h2>Öneri Raporu</h2>
                <div style="color:#546e7a; margin-bottom:10px">{sugg_note}</div>
                {df_to_html_styled(df_sugg, numeric_df=None, hat_karsilastirma=True)}
              </section>
            """
        else:
            sugg_html = f"""
              <section class="table-card card">
                <h2>Öneri Raporu</h2>
                <div style="color:#546e7a;">{sugg_note}</div>
                <div style="margin-top:6px;">Bu dönem için iyileştirme önerisi bulunamadı.</div>
              </section>
            """

        # 8) KPI HTML - GÜNCELLENDİ (2 ondalık basamak, tam sayı adetler)
        def kpi_html():
            return f"""
            <section class="kpis">
              <div class="card kpi">
                <div class="label">Tahmini Ortalama Verim</div>
                <div class="value">{fmt_tr_pct(tahmini_ortalama_verim * 100.0)}</div>
              </div>
              <div class="card kpi">
                <div class="label">Tahmini Adetsel Verim</div>
                <div class="value">{fmt_tr_pct(tahmini_adetsel_verim * 100.0)}</div>
              </div>
              <div class="card kpi">
                <div class="label">Tahmini Zamansal Verim</div>
                <div class="value">{fmt_tr_pct(tahmini_zamansal_verim * 100.0)}</div>
              </div>

              <div class="card kpi">
                <div class="label">Toplam Teorik Adet</div>
                <div class="value">{fmt_tr_int(sum_teo)}</div>
              </div>
              <div class="card kpi">
                <div class="label">Kesilen Damla</div>
                <div class="value">{fmt_tr_int(kesilen_damla)}</div>
              </div>
              <div class="card kpi">
                <div class="label">Tahmini Fiili Adet</div>
                <div class="value">{fmt_tr_int(sum_fiili)}</div>
              </div>
            </section>
            """

        # 9) Stil
        style = """
        <style>
          :root{ --bg:#0b1020; --card:#121a35; --muted:#93a4c5; --text:#e8eeff; --accent:#4f7cff; --accent-2:#26c6da;
            --border:#233055; --success:#10b981; --danger:#ef4444; --shadow:0 6px 22px rgba(8,15,35,.4); --radius:14px; }
          @media (prefers-color-scheme: light){ :root{ --bg:#f6f8ff; --card:#ffffff; --text:#0f1440; --muted:#65719a; --border:#e3e8ff;
                   --accent:#3b6bff; --accent-2:#00a7c2; --shadow:0 6px 20px rgba(21,44,99,.10); } }
          *{box-sizing:border-box}
          body{margin:0;padding:0;background:var(--bg);color:var(--text);font:15px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial}
          .wrap{max-width:1080px;margin:28px auto 40px;padding:0 16px}
          header.report{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;margin-bottom:18px}
          .title{display:flex;flex-direction:column;gap:6px}
          .title h1{font-size:24px;margin:0}
          .subtitle{color:var(--muted);font-size:13px}
          .badge{display:inline-flex;align-items:center;gap:8px;background:linear-gradient(135deg,var(--accent),var(--accent-2));
                 color:#fff;padding:8px 12px;border-radius:999px;font-weight:700;font-size:12px;box-shadow:var(--shadow);white-space:nowrap}
          .card{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);padding:16px 18px;box-shadow:var(--shadow)}
          .kpis{display:grid;gap:14px;grid-template-columns:repeat(3,1fr);margin:14px 0 22px}
          .kpi .label{color:var(--muted);font-weight:600;font-size:12px;letter-spacing:.4px;text-transform:uppercase}
          .kpi .value{margin-top:8px;font-size:26px;font-weight:800}
          h2{margin:26px 0 10px;font-size:16px}
          .table-card{margin-top:8px}
          table.dataframe{width:100%;border-collapse:separate;border-spacing:0;border:1px solid var(--border);
                          border-radius:12px;background:var(--card);box-shadow:var(--shadow);overflow:hidden}
          table.dataframe thead th{position:sticky;top:0;z-index:1;background:linear-gradient(180deg,rgba(255,255,255,.06),rgba(255,255,255,0));
                                   text-align:left;color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.4px;
                                   padding:10px 12px;border-bottom:1px solid var(--border)}
          table.dataframe tbody td{padding:12px 12px;border-bottom:1px dashed rgba(147,164,197,.25);text-align:left}
          td.pill{font-weight:800;color:#fff;border-radius:10px;padding:8px 10px;box-shadow: inset 0 0 0 1px rgba(0,0,0,.12);}
          td.pill-mevcut{background: linear-gradient(180deg, rgba(139,29,59,.95), rgba(139,29,59,.75)); border: 1px solid rgba(139,29,59,1);}
          td.pill-onerilen{background: linear-gradient(180deg, rgba(11,30,99,.95), rgba(11,30,99,.75)); border: 1px solid rgba(11,30,99,1);}
          td.pill-kazanc{background: linear-gradient(180deg, rgba(16,185,129,.95), rgba(16,185,129,.70)); border: 1px solid rgba(16,185,129,1);}
          td.verim{font-weight:800;color:#fff;background-color: #808080;border-radius:10px;padding:8px 10px;box-shadow:inset 0 0 0 1px rgba(0,0,0,.12)}
          td.verim.up{background:linear-gradient(180deg, rgba(16,185,129,.95), rgba(16,185,129,.70)); border:1px solid rgba(16,185,129,1)}
          td.verim.down{background:linear-gradient(180deg, rgba(239,68,68,.95), rgba(239,68,68,.70)); border:1px solid rgba(239,68,68,1)}
          td.verim.zero{background:linear-gradient(180deg, rgba(127,127,127,.95), rgba(127,127,127,.70)); border:1px solid rgba(127,127,127,1)}
          .page-break{page-break-before:always;margin-top:24px}
          .empty{color:var(--muted)}
          @media print{ body{background:#fff;color:#000} .card, table.dataframe{box-shadow:none} .badge{box-shadow:none} }
          .table-scroll{overflow-x:auto;-webkit-overflow-scrolling:touch;}
          table.dataframe{ min-width: 680px; }
          @media (max-width: 900px){ .wrap{ max-width: 820px; } .kpis{ grid-template-columns: 1fr 1fr; } }
          @media (max-width: 640px){ .wrap{ max-width: 100%; padding: 0 12px; } header.report{ flex-direction: column; } 
            .kpis{ grid-template-columns: 1fr; } .kpi .value{ font-size: 22px; } }
        </style>
        """

        # 10) Sayfa 1 ve 2
        html_page1 = f"""
        <div class="wrap">
          <header class="report">
            <div class="title">
              <h1>AYLIK HEDEF VERİM RAPORU</h1>
              <div class="subtitle">Dönem: <b>{period}</b></div>
            </div>
            <div class="badge">{scope_badge}</div>
          </header>

          {kpi_html()}

          <section class="table-card card">
            <h2>Fırın Bazlı</h2>
            {df_to_html_styled(df_furnace_view, numeric_df=df_furnace, furnace_karsilastirma=True)}
          </section>

          <section class="table-card card">
            <h2>Hat Bazlı</h2>
            {df_to_html_styled(df_line_view, numeric_df=df_line_num, hat_karsilastirma=True)}
          </section>
        </div>
        """

        html_page2 = f"""
        <div class="wrap page-break">
          <section class="table-card card">
            <h2>Kampanya Bazlı</h2>
            {df_to_html_styled(df_campaign_view, numeric_df=df_campaign_num, hat_karsilastirma=True)}
          </section>

          {sugg_html}

          <div style="text-align:right;color:#93a4c5;font-size:12px;margin-top:8px">
            Oluşturma zamanı: {_dt.now().strftime("%d.%m.%Y %H:%M")}
          </div>
        </div>
        """

        full_html = f"""<!DOCTYPE html>
        <html>
        <head>
          <meta charset="utf-8">
          <title>Aylık Hedef Verim Raporu</title>
          <style>{style}</style>
        </head>
        <body class="report">
          {html_page1}
          {html_page2}
        </body>
        </html>"""

        # 11) Önizleme
        dlg_prev = QDialog(self)
        dlg_prev.setWindowTitle("Hedef Verim Raporu – Önizleme")
        lay = QVBoxLayout(dlg_prev)
        view = QTextEdit();
        view.setReadOnly(True);
        view.setHtml(full_html)
        lay.addWidget(view)
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok);
        btns.accepted.connect(dlg_prev.accept)
        lay.addWidget(btns)

        if df_sugg is not None and not df_sugg.empty:
            from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QLabel
            dlg_s = QDialog(self)
            dlg_s.setWindowTitle(f"Hat Önerileri – {len(df_sugg)} olası iyileştirme")
            lay_s = QVBoxLayout(dlg_s)
            lay_s.addWidget(QLabel("Geçmiş VERİM'e göre daha iyi çalışmış hatlar bulundu:"))
            table = QTableWidget()
            cols = ["Kalem", "Planlanan Hat", "Önerilen Hat", "Planlanan Verim(%)", "Önerilen Verim(%)", "Kazanç(%)"]
            table.setColumnCount(len(cols));
            table.setHorizontalHeaderLabels(cols);
            table.setRowCount(min(len(df_sugg), 100))
            for r, row in enumerate(df_sugg.head(100).to_dict("records")):
                for c, k in enumerate(cols): table.setItem(r, c, QTableWidgetItem(str(row.get(k, ""))))
            table.resizeColumnsToContents()
            lay_s.addWidget(table)
            btns_s = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok);
            btns_s.accepted.connect(dlg_s.accept)
            lay_s.addWidget(btns_s)
            dlg_s.resize(900, 420);
            dlg_s.exec()

        dlg_prev.resize(1100, 700);
        dlg_prev.exec()

        # 12) Kaydet
        choice = QMessageBox.question(self, "Çıktı Türü", "Raporu kaydetmek ister misiniz?",
                                      QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                      QMessageBox.StandardButton.Yes)
        if choice != QMessageBox.StandardButton.Yes: return

        default_dir = str((Path("production_data") / "reports").resolve())
        Path(default_dir).mkdir(parents=True, exist_ok=True)
        ts = _dt.now().strftime("%Y%m%d_%H%M")
        file_path, sel_filter = QFileDialog.getSaveFileName(self, "Raporu Kaydet",
                                                            os.path.join(default_dir, f"Aylik_Verim_Raporu_{ts}"),
                                                            "HTML Files (*.html);;PDF Files (*.pdf)")
        if not file_path: return

        if sel_filter.startswith("PDF") or file_path.lower().endswith(".pdf"):
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            if not file_path.lower().endswith(".pdf"): file_path += ".pdf"
            printer.setOutputFileName(file_path)
            doc = QTextDocument();
            doc.setHtml(full_html);
            doc.print(printer)
            QMessageBox.information(self, "Kaydedildi", f"PDF rapor kaydedildi:\n{file_path}")
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(file_path)))
        else:
            if not file_path.lower().endswith(".html"): file_path += ".html"
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(full_html)
            QMessageBox.information(self, "Kaydedildi", f"HTML rapor kaydedildi:\n{file_path}")
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(file_path)))

    def get_historical_efficiency(self, product, line):
        df = self.product_data.historical_data
        if df is None or df.empty:
            return None
        if not {'Kalem', 'Hat', 'ADETSEL', 'Yıl'}.issubset(set(df.columns)):
            return None

        tmp = df.copy()
        tmp['Kalem'] = tmp['Kalem'].astype(str).str.strip()
        tmp['Hat'] = tmp['Hat'].astype(str).str.strip().str.upper()
        line_up = str(line).strip().upper()

        mask = (tmp['Kalem'] == product) & (tmp['Hat'] == line_up)
        filtered = tmp.loc[mask].copy()

        if filtered.empty:
            return None

        # Ağırlık tablosu: 2024 → 4, 2025 → 6, diğer yıllar → 0 (dışarıda kalır)
        weight_map = {2024: 4, 2025: 6}

        weighted_sum = 0.0
        total_weight = 0

        for _, row in filtered.iterrows():
            try:
                yil = int(str(row['Yıl']).strip())
                w = weight_map.get(yil, 0)
                if w == 0:
                    continue
                s = str(row['ADETSEL']).replace('%', '').replace(',', '.').strip()
                val = float(s)
                val = val / 100.0 if val > 1 else val
                weighted_sum += val * w
                total_weight += w
            except:
                continue

        if total_weight == 0:
            # Ağırlıklı hesap yapılamazsa (2024/2025 verisi yoksa) basit ortalama
            vals = []
            for _, row in filtered.iterrows():
                try:
                    s = str(row['ADETSEL']).replace('%', '').replace(',', '.').strip()
                    val = float(s)
                    vals.append(val / 100.0 if val > 1 else val)
                except:
                    continue
            return (sum(vals) / len(vals)) if vals else None

        return weighted_sum / total_weight

    def get_historical_efficiency_productwide(self, product):
        df = self.product_data.historical_data
        if df is None or df.empty:
            return None
        if not {'Kalem', 'Hat', 'ADETSEL'}.issubset(set(df.columns)):
            return None

        tmp = df.copy()
        tmp['Kalem'] = tmp['Kalem'].astype(str).str.strip()
        tmp['Hat'] = tmp['Hat'].astype(str).str.strip()

        # "X TOPLAMI" formatındaki satırları filtrele
        # Örnek: Hat sütununda "ADA315 TOPLAMI" veya "TOPLAMI" içeren satırlar
        toplam_mask = (
                (tmp['Kalem'] == product) &
                tmp['Hat'].str.upper().str.contains('TOPLAMI', na=False)
        )
        filtered = tmp.loc[toplam_mask]

        if filtered.empty:
            return None

        # TOPLAM satırında birden fazla kayıt varsa basit ortalama al
        # (genellikle tek satır olur ama güvence için)
        vals = []
        for v in filtered['ADETSEL']:
            s = str(v).replace('%', '').replace(',', '.').strip()
            try:
                x = float(s)
                vals.append(x / 100.0 if x > 1 else x)
            except:
                continue

        return (sum(vals) / len(vals)) if vals else None

    def repaint_all_campaigns_and_totals(self):
        if not hasattr(self, "plan_table"):
            return
        for (row, s, e), info in list(self.plan_table.campaigns.items()):
            for c in range(s, e + 1):
                self.plan_table.setItem(row, c, QTableWidgetItem(""))

        for (row, s, e), info in self.plan_table.campaigns.items():
            self.plan_table._paint_segment(row, s, e, info["product"])

        self.plan_table.update_tonaj_totals()
        self.plan_table.update_per_line_tonaj_rows()

    def save_plan(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Planı Kaydet", "", "JSON Files (*.json)"
        )
        if file_path:
            plan_data = {
                'campaigns': [
                    {
                        'row': row,
                        'start_col': start,
                        'end_col': end,
                        'product': info['product'],
                        'duration': info['duration'],
                        'single_drop': info.get('single_drop', None)
                    }
                    for (row, start, end), info in self.plan_table.campaigns.items()
                ],
                'time_efficiencies': self.time_efficiencies,
                'table_info': {
                    'rows': self.plan_table.rowCount(),
                    'cols': self.plan_table.columnCount(),
                    'row_headers': [self.plan_table.verticalHeaderItem(i).text()
                                    for i in range(self.plan_table.rowCount())],
                    'col_headers': [self.plan_table.horizontalHeaderItem(i).text()
                                    for i in range(self.plan_table.columnCount())]
                },
                'start_date': self.plan_start_date.isoformat(),
                'end_date': self.plan_end_date.isoformat(),
            }

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(plan_data, f, ensure_ascii=False, indent=2)

            QMessageBox.information(self, "Başarılı", "Plan kaydedildi!")
            self.update_status_bar(f"Plan kaydedildi: {file_path}")

    def view_products(self):
        if not self.product_data.products:
            QMessageBox.information(self, "Bilgi", "Önce ürün havuzunu kaydedin/yükleyin.")
            return

        headers = ["Kalem", "DYS Devir", "DYS Verim (%)", "DYS Damla Gramaj"]
        rows = []
        for kalem in sorted(self.product_data.products.keys()):
            rec = self.product_data.products[kalem]
            devir = rec.get('devir', "")
            verim = rec.get('verim', "")
            gramaj = rec.get('gramaj', "")
            try:
                verim = f"{float(verim):.2f}"
            except:
                pass
            rows.append([kalem, devir, verim, gramaj])
        self._show_table_dialog("Ürün Havuzu (DYS)", headers, rows)

    def view_historical(self):
        df = self.product_data.historical_data
        if df.empty:
            QMessageBox.information(self, "Bilgi", "Önce geçmiş verileri kaydedin/yükleyin.")
            return

        headers = list(df.columns)
        rows = []
        for _, row in df.iterrows():
            rows.append([row.get(col, "") for col in headers])
        self._show_table_dialog("Geçmiş Veriler", headers, rows)

    def _show_table_dialog(self, title: str, headers, rows):
        dlg = QDialog(self)
        dlg.setWindowTitle(title)
        dlg.setModal(True)
        dlg.resize(900, 600)
        layout = QVBoxLayout(dlg)

        table = QTableWidget(dlg)
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setRowCount(len(rows))
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.verticalHeader().setVisible(True)  # ← EKLE
        table.setAlternatingRowColors(True)

        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                table.setItem(r, c, QTableWidgetItem(str(val)))

        layout.addWidget(table)

        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        btns.rejected.connect(dlg.reject)
        layout.addWidget(btns)

        dlg.exec()

    def load_plan(self):
        from datetime import date, datetime, timedelta

        file_path, _ = QFileDialog.getOpenFileName(
            self, "Plan Yükle", "", "JSON Files (*.json)"
        )
        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                plan_data = json.load(f)

            # --- 1) Tablo temel bilgileri ---
            table_info = plan_data['table_info']
            rows = table_info['rows']
            cols = table_info['cols']
            row_headers = table_info['row_headers']
            col_headers = table_info['col_headers']

            self.plan_table.setRowCount(rows)
            self.plan_table.setColumnCount(cols)
            self.plan_table.setVerticalHeaderLabels(row_headers)
            self.plan_table.setHorizontalHeaderLabels(col_headers)

            # --- 2) Görünür plan ve tonaj satırlarını yeniden kur ---
            self.plan_table.display_rows.clear()
            self.plan_table.tonaj_rows.clear()

            a_total_row = None
            c_total_row = None

            i = 0
            while i < len(row_headers):
                name = row_headers[i]

                if name in ("A FIRINI - GÜNLÜK CAM ÇEKİŞ", "A FIRINI TONAJ", "A FIRINI TOPLAM"):
                    a_total_row = i
                    i += 1
                    continue

                if name in ("C FIRINI - GÜNLÜK CAM ÇEKİŞ", "C FIRINI TONAJ", "C FIRINI TOPLAM"):
                    c_total_row = i
                    i += 1
                    continue

                if name.endswith("(Tonaj)") or name.strip().lower() == "tonaj":
                    i += 1
                    continue

                if (i + 1 < len(row_headers)
                        and (row_headers[i + 1].endswith("(Tonaj)") or row_headers[i + 1].strip().lower() == "tonaj")):
                    plan_row = i
                    ton_row = i + 1
                    self.plan_table.display_rows.append(plan_row)
                    self.plan_table.tonaj_rows[plan_row] = ton_row
                    i += 2
                else:
                    self.plan_table.display_rows.append(i)
                    i += 1

            self.plan_table.row_a_total = a_total_row
            self.plan_table.row_c_total = c_total_row
            self.plan_table.is_a_line = lambda r: (r in self.plan_table.display_rows and r < a_total_row)
            self.plan_table.is_c_line = lambda r: (
                    r in self.plan_table.display_rows and r > a_total_row and r < c_total_row)

            self.plan_table.get_product = lambda name: self.product_data.products.get(name, {})

            # --- 3) Kampanyaları boya ve kaydet ---
            self.plan_table.campaigns.clear()
            for camp in plan_data['campaigns']:
                row = camp['row']
                start_col = camp['start_col']
                end_col = camp['end_col']
                product = camp['product']
                duration = camp.get('duration', end_col - start_col + 1)

                self.plan_table._paint_segment(row, start_col, end_col, product)
                self.plan_table.campaigns[(row, start_col, end_col)] = {
                    'product': product,
                    'duration': duration,
                    'single_drop': camp.get('single_drop', None)
                }

            # --- 4) Tonaj toplamlarını güncelle ---
            self.plan_table.update_tonaj_totals()

            # --- 5) Zamansal verim ayarlarını yükle (DÜZELTİLEN KISIM) ---
            loaded_effs = plan_data.get('time_efficiencies', {})

            # Eğer yüklenen planda YENİ EKLENEN SABİT DEĞERLER yoksa (eski dosyaysa),
            # halihazırda programda açık olan (veya settings.json'dan gelen) değerleri koru.
            # Böylece 85'e sıfırlanmaz.
            keys_to_preserve = ["FIXED_A", "FIXED_C", "FIXED_TOTAL"]

            for key in keys_to_preserve:
                # Dosyada yoksa AMA hafızada varsa -> Hafızadakini dosyadan gelen veriye ekle
                if key not in loaded_effs and key in self.time_efficiencies:
                    loaded_effs[key] = self.time_efficiencies[key]

            self.time_efficiencies = loaded_effs

            # --- 6) Dönem tarihlerini set et ---
            sd = plan_data.get('start_date')
            ed = plan_data.get('end_date')

            if sd and ed:
                sdt = datetime.fromisoformat(sd)
                edt = datetime.fromisoformat(ed)
                self.plan_start_date = sdt.date() if hasattr(sdt, "date") else sdt
                self.plan_end_date = edt.date() if hasattr(edt, "date") else edt
            else:
                start_py = date.today()
                if col_headers:
                    first = col_headers[0]
                    try:
                        day_str, mon_str = first.split(".")
                        day = int(day_str)
                        mon = int(mon_str)
                        start_py = start_py.replace(month=mon, day=day)
                    except Exception:
                        start_py = date.today()
                end_py = start_py + timedelta(days=cols - 1)

                self.plan_start_date = start_py
                self.plan_end_date = end_py

            # --- 7) Bildirim / durum ---
            QMessageBox.information(self, "Başarılı", "Plan yüklendi!")
            self.update_status_bar(f"Plan yüklendi: {file_path}")

        except Exception as e:
            QMessageBox.warning(self, "Hata", f"Plan yüklenemedi: {e}")


def main():
    app = QApplication(sys.argv)
    # app.setStyle('Fusion')

    # Modern tema uygula
    app.setStyleSheet(EnterpriseTheme.get_main_stylesheet())

    window = ProductionPlannerApp()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()