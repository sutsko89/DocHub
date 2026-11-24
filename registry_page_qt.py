"""Страница «Реестр» на Qt (этап M5) — перенос контрактов TREE-1/UX-REG.

Qt-перенос ``registry_controller.RegistryController`` + ``registry_tree``
(Tk): QTableView + QAbstractTableModel из строк ``RowSpec``; план дифа —
GUI-free ``registry_core.plan_diff`` (один источник с Tk-сборкой, этап
M5), колонки «Контрагент · Тип · ИНН · ID · Адрес», точка типа в нулевой
колонке, поиск/сортировка в контроллере по всей базе, порция
``DISPLAY_SOFT_LIMIT`` (10 000) с «Показать ещё», панель действий
выделенной записи, клавиатура (стрелки/Enter/Ctrl+F/Delete), drill
«Документы — <имя>» с возвратом «← Назад»/Escape, теги сроков
красный/амберный, placeholder поиска.

Контракты TREE-1, переносимые дословно:

* поиск и сортировка всегда выполняются по всей базе (в контроллере) —
  таблица получает готовую упорядоченную модель строк и не знает про
  фильтры; лимит отображения (R-05b) применяется в контроллере;
* полная популяция — только при смене фильтра, сортировки и
  drill-переходах (сигнатура вида); изменение состава и правки записей
  идут дифом: строки, которых изменение не касается, не пересоздаются
  (критерий 5.1 — выделение и прокрутка сохраняются);
* один клик только выделяет строку — действия на панели под списком;
  повторный клик снимает выделение (click-toggle); клавиша Delete
  удаляет выделенную запись, у записей корзины Delete не действует
  (там выбор «Восстановить»/«Удалить навсегда» через панель);
* двойной клик по контрагенту — провал в его документы («← Назад» и
  Escape возвращают), по документу — открыть в Word; Enter активирует
  выделенную строку, стрелки двигают выделение без снятия, Ctrl+F —
  фокус в поиск (U13);
* теги сроков «expired»/«soon» красят строку документа красным/амберным
  (те же цвета, что у Tk-строк; чтение текущей темы — живое, как у
  Tk-перекраски по слушателю темы).

Состояние страницы — у контроллера (раздел Д плана); имена — те же, что
в Tk-сборке (``_registry_selected_id``, ``_registry_rows``,
``_registry_drill_party_id`` …) для читаемого соответствия этапов.
Каркас получает ``on_page_open`` (навигация) и ``refresh`` (хвосты
операций жизненного цикла).
"""

from __future__ import annotations

from PySide6.QtCore import (
    QAbstractTableModel, QItemSelectionModel, QModelIndex, QRectF, QTimer, Qt,
)
from PySide6.QtGui import QBrush, QColor, QFont, QKeySequence, QPainter, QPixmap, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView, QComboBox, QHBoxLayout, QHeaderView, QLabel,
    QLineEdit, QPushButton, QTableView, QVBoxLayout, QWidget,
)

import agreements
import applog
import qt_theme
import registry_core as rcore
import storage
from document_builder import F, digits_only
from registry_core import (
    DOC_COLUMNS,
    PARTY_COLUMNS,
    SHOW_MORE_IID,
    SORT_LABEL_BY_VALUE,
    SORT_VALUE_BY_LABEL,
    action_labels_for,
    doc_row_tags,
    doc_row_values,
    party_row_dot,
    party_row_values,
    party_subtitle,
)

#: Ключ цвета текста по тегу строки (паритет Tk tag_configure, R-03).
TAG_COLORS = {
    "expired": "error_text",
    "soon": "warning_text",
    "muted": "text_secondary",
    "showmore": "accent",
    "section": "text",
}


def theme_color(key: str) -> QColor:
    """Цвет пары ``key`` текущей темы (живое чтение, как Tk-слушатель)."""
    return QColor(qt_theme.color(qt_theme.current_mode(), key))


def _dot_pixmap(color: str, size: int = 9) -> QPixmap:
    """Круглая точка типа (V4-пилот): Tk PhotoImage -> QPixmap."""
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    try:
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(color))
        painter.drawEllipse(QRectF(0.5, 0.5, size - 1.0, size - 1.0))
    finally:
        painter.end()
    return pixmap


class RegistryModelQt(QAbstractTableModel):
    """Модель строк реестра из RowSpec: колонки, теги, точки, диф (R-03/R-04).

    План дифа считает GUI-free ``registry_core.plan_diff``; модель
    исполняет его сигналами Qt (удаление/вставка/сдвиг/правка), поэтому
    неизменённые строки не пересоздаются (критерий 5.1) — выделение и
    позиция прокрутки переживают «Показать ещё» и правки.
    """

    def __init__(self, columns: tuple = PARTY_COLUMNS) -> None:
        super().__init__()
        self._columns = tuple(columns)
        self._rows: list = []
        self._dots: dict[str, QPixmap] = {}

    # ------------------------------------------------------------------ #
    #  Реакция QAbstractItemModel
    # ------------------------------------------------------------------ #
    def rowCount(self, parent=QModelIndex()) -> int:  # noqa: N802 — имя Qt
        return 0 if parent.isValid() else len(self._rows)

    def columnCount(self, parent=QModelIndex()) -> int:  # noqa: N802 — имя Qt
        return 0 if parent.isValid() else len(self._columns)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):  # noqa: N802
        if orientation == Qt.Orientation.Horizontal \
                and role == Qt.ItemDataRole.DisplayRole:
            return self._columns[section][1]
        return None

    def flags(self, _index) -> Qt.ItemFlag:
        return (
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
        )

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or index.row() >= len(self._rows):
            return None
        row = self._rows[index.row()]
        if role == Qt.ItemDataRole.DisplayRole:
            values = row.values
            return values[index.column()] if index.column() < len(values) else ""
        if role == Qt.ItemDataRole.DecorationRole and index.column() == 0:
            if not row.dot:
                return None
            pixmap = self._dots.get(row.dot)
            if pixmap is None:
                pixmap = _dot_pixmap(row.dot)
                self._dots[row.dot] = pixmap
            return pixmap
        if role == Qt.ItemDataRole.ForegroundRole:
            for tag in row.tags:
                key = TAG_COLORS.get(tag)
                if key is not None:
                    return QBrush(theme_color(key))
            return None
        if role == Qt.ItemDataRole.FontRole and "section" in row.tags:
            font = QFont()
            font.setBold(True)
            return font
        if role == Qt.ItemDataRole.TextAlignmentRole:
            return int(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        return None

    # ------------------------------------------------------------------ #
    #  Операции модели (R-04)
    # ------------------------------------------------------------------ #
    @property
    def columns(self) -> tuple:
        return self._columns

    def row(self, iid: str):
        """RowSpec по iid или None (Tk-паритет RegistryTree.row)."""
        for row in self._rows:
            if row.iid == iid:
                return row
        return None

    @property
    def rows(self) -> list:
        """Текущие строки модели (порядок отображения)."""
        return self._rows

    def order(self) -> tuple[str, ...]:
        """iid строк в порядке отображения (контрольный инвариант R-04)."""
        return tuple(row.iid for row in self._rows)

    def populate(self, columns: tuple, rows: list) -> None:
        """Полная популяция (смена фильтра, сортировки, drill-переход)."""
        self.beginResetModel()
        self._columns = tuple(columns)
        self._rows = list(rows)
        self.endResetModel()

    def apply_diff(self, rows: list) -> None:
        """Точечное обновление до модели: удаление, вставка, сдвиг, правка.

        Префикс, совпадающий с моделью, не трогается — позиции и
        persistent-индексы неизменённых строк сохраняются (критерий 5.1);
        значения/теги/точка обновляются только у изменившихся строк.
        """
        plan = rcore.plan_diff(self.order(), rows)
        if plan.to_delete:
            pos_of = {row.iid: idx for idx, row in enumerate(self._rows)}
            for idx in sorted(
                (pos_of[iid] for iid in plan.to_delete if iid in pos_of),
                reverse=True,
            ):
                self.beginRemoveRows(QModelIndex(), idx, idx)
                self._rows.pop(idx)
                self.endRemoveRows()
        for pos, iid, op in plan.ops:
            if op == "insert":
                self.beginInsertRows(QModelIndex(), pos, pos)
                self._rows.insert(pos, rows[pos])
                self.endInsertRows()
            else:  # move: как у Tk tree.move — снять и вставить
                src = next(
                    (idx for idx, row in enumerate(self._rows) if row.iid == iid),
                    None,
                )
                if src is not None and src != pos:
                    self.beginRemoveRows(QModelIndex(), src, src)
                    moved = self._rows.pop(src)
                    self.endRemoveRows()
                    self.beginInsertRows(QModelIndex(), pos, pos)
                    self._rows.insert(pos, moved)
                    self.endInsertRows()
        changed: list[int] = []
        pos_of = {row.iid: idx for idx, row in enumerate(self._rows)}
        for row in plan.updates:
            idx = pos_of.get(row.iid)
            if idx is None:
                continue
            cached = self._rows[idx]
            if (cached.values != row.values or cached.tags != row.tags
                    or cached.dot != row.dot):
                self._rows[idx] = row
                changed.append(idx)
        last_col = max(0, len(self._columns) - 1)
        for idx in changed:
            self.dataChanged.emit(
                self.index(idx, 0), self.index(idx, last_col),
            )
        if __debug__:
            assert self.order() == tuple(row.iid for row in rows), \
                "рассинхрон модели реестра и желаемой модели строк"


class RegistryTableQt(QTableView):
    """Список реестра: QTableView + выделение/toggle/активация (U13).

    События выделения — SET-семантика Tk (идемпотентны); снятие
    выделения идёт только явным действием пользователя — повторным
    кликом по выделенной строке (click-toggle, как у 2.18.0).
    """

    NAV_KEYS = (
        Qt.Key.Key_Down, Qt.Key.Key_Up, Qt.Key.Key_Home, Qt.Key.Key_End,
        Qt.Key.Key_PageUp, Qt.Key.Key_PageDown,
    )

    def __init__(self, on_select, on_activate, on_delete_key):
        super().__init__()
        self._on_select = on_select
        self._on_activate = on_activate
        self._on_delete_key = on_delete_key
        self._key_nav = False
        self._toggle_token = 0  # отмена click-toggle при двойном клике
        self._model = RegistryModelQt()
        self.setModel(self._model)
        self._setup_view()

        sm = self.selectionModel()
        sm.currentRowChanged.connect(self._handle_current_changed)

    def _setup_view(self) -> None:
        self.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.setShowGrid(False)
        self.setWordWrap(False)
        header = self.horizontalHeader()
        header.setSectionsClickable(False)
        header.setHighlightSections(False)
        header.setStretchLastSection(False)
        header.setDefaultAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        vertical = self.verticalHeader()
        vertical.setVisible(False)
        vertical.setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        vertical.setDefaultSectionSize(30)
        self._apply_columns()

    def _apply_columns(self) -> None:
        """Ширины/растяжение колонок по константам ядра (U1)."""
        header = self.horizontalHeader()
        for index, (_key, _title, width, _anchor, stretch) in enumerate(
            self._model.columns
        ):
            if stretch:
                header.setSectionResizeMode(
                    index, QHeaderView.ResizeMode.Stretch)
            else:
                header.setSectionResizeMode(
                    index, QHeaderView.ResizeMode.Interactive)
                header.resizeSection(index, width)

    # ------------------------------------------------------------------ #
    #  Модель: популяция и диф
    # ------------------------------------------------------------------ #
    def populate(self, columns: tuple, rows: list) -> None:
        self._model.populate(columns, rows)
        self._apply_columns()

    def apply_diff(self, rows: list) -> None:
        self._model.apply_diff(rows)

    @property
    def model(self) -> RegistryModelQt:
        return self._model

    def row(self, iid: str):
        return self._model.row(iid)

    def order(self) -> tuple[str, ...]:
        return self._model.order()

    # ------------------------------------------------------------------ #
    #  Выделение и активация (U13, SET-семантика Tk)
    # ------------------------------------------------------------------ #
    def set_selection(self, iid: str | None) -> None:
        """Выделить строку (None — снять); отсутствующий iid игнорируется.

        Эхо-события от программной смены не подавляются: обработчик
        контроллера идемпотентен (SET-семантика). Исключение — повторная
        установка того же iid: сигнал не поднимается, состояние не меняется.
        """
        sm = self.selectionModel()
        if sm is None:
            return
        if iid is None:
            sm.clearSelection()
            return
        index = self._index_of_iid(iid)
        if index is None:
            return
        current = self.currentIndex()
        if current.isValid() and current.row() == index.row() \
                and sm.isRowSelected(index.row(), QModelIndex()):
            return
        sm.setCurrentIndex(
            index,
            QItemSelectionModel.SelectionFlag.ClearAndSelect
            | QItemSelectionModel.SelectionFlag.Rows,
        )

    def selected_iid(self) -> str | None:
        sm = self.selectionModel()
        if sm is None:
            return None
        rows = sm.selectedRows(0)
        if rows:
            return self._model.rows[rows[0].row()].iid
        return None

    def focus_tree(self) -> None:
        self.setFocus()

    def _index_of_iid(self, iid: str) -> QModelIndex | None:
        for idx, row in enumerate(self._model.rows):
            if row.iid == iid:
                return self._model.index(idx, 0)
        return None

    def _handle_current_changed(self, current: QModelIndex, _previous) -> None:
        iid = self._model.rows[current.row()].iid \
            if current.isValid() else None
        key_nav = self._key_nav
        self._key_nav = False
        # Событие вида — SET-семантика (user=False): клик по новой строке
        # и эхо программной синхронизации обрабатываются одинаково;
        # снятие выделения идёт только через click-toggle.
        self._on_select(iid, key_nav, False)

    # ------------------------------------------------------------------ #
    #  Мышь и клавиатура
    # ------------------------------------------------------------------ #
    def mousePressEvent(self, event) -> None:  # noqa: N802 — имя Qt
        index = self.indexAt(event.position().toPoint())
        token = None
        if index.isValid() and event.button() == Qt.MouseButton.LeftButton \
                and self.selectionModel().isRowSelected(
                    index.row(), QModelIndex()):
            # Повторный клик по выделенной строке — снять выделение после
            # стандартной обработки (Tk after_idle-паритет, 2.18.0).
            self._toggle_token += 1
            token = self._toggle_token
        super().mousePressEvent(event)
        if token is not None \
                and self.currentIndex().isValid() \
                and self.currentIndex().row() == index.row():
            QTimer.singleShot(0, lambda t=token: self._consume_click_toggle(t))

    def mouseDoubleClickEvent(self, event) -> None:  # noqa: N802 — имя Qt
        # Двойной клик по выделенной строке открывает запись, а не снимает
        # выделение: отложенный toggle отменяется сменой токена (Tk-паритет).
        self._toggle_token += 1
        index = self.indexAt(event.position().toPoint())
        super().mouseDoubleClickEvent(event)
        if index.isValid():
            self._on_activate(self._model.rows[index.row()].iid)

    def _consume_click_toggle(self, token: int) -> None:
        """Отложенное снятие выделения повторным кликом (Tk-паритет)."""
        if token != self._toggle_token:
            return  # двойной клик/активация отменили отложенное снятие
        if self.selected_iid() is not None:
            self.set_selection(None)
            # Явное действие пользователя — единственный путь снятия
            # выделения: контроллер отличает его от эха (user=True).
            self._on_select(None, False, True)

    def keyPressEvent(self, event) -> None:  # noqa: N802 — имя Qt
        key = event.key()
        if key in self.NAV_KEYS:
            # Стрелки двигают выделение без toggle — флаг ставится до
            # стандартной навигации, которая меняет выделение (U13).
            self._key_nav = True
            super().keyPressEvent(event)
            return
        if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            current = self.currentIndex()
            if current.isValid():
                self._on_activate(self._model.rows[current.row()].iid)
            return
        if key == Qt.Key.Key_Delete:
            self._on_delete_key()
            return
        super().keyPressEvent(event)


class RegistryPageController:
    """Страница «Реестр»: поиск, выделение, секции, drill, панель действий.

    Код страницы и её данных — у контроллера (раздел Д плана); имена
    состояния — те же, что в Tk-сборке. Вызовы, которые подменяет
    smoke/тесты, идут через ``self.app.*`` (``confirm``, ``choice``,
    ``alert``, ``lifecycle_flow`` — как Tk ``app._lifecycle_flow``).
    """

    def __init__(self, app) -> None:
        self.app = app
        self.log = app.log
        self.page: QWidget | None = None

        self._registry_selected_id = None
        self._registry_rows: dict = {}
        self._registry_drill_party_id = None
        self._docs_selected_id = None
        self._docs_rows: dict = {}
        self._registry_selected_extra = None
        self._registry_shown = rcore.DISPLAY_SOFT_LIMIT
        self._registry_signature = None
        # Ленивый рендер (п. 9 этапа 5): список не рисуется при старте —
        # реестр отрисуется при первом открытии страницы или изменении.
        self._registry_rendered = False
        self._docs_type_filter = "all"
        self._registry_sort = str(
            (app.settings or {}).get("registry_sort", "added"))
        if self._registry_sort not in SORT_LABEL_BY_VALUE:
            self._registry_sort = "added"

        self._build_page()

    # ------------------------------------------------------------------ #
    #  Построение страницы
    # ------------------------------------------------------------------ #
    def _build_page(self) -> None:
        app = self.app
        page = QWidget()
        self.page = page
        outer = QVBoxLayout(page)
        outer.setContentsMargins(24, 18, 24, 10)
        outer.setSpacing(6)

        # -- Шапка: заголовок, «← Назад», «Обновить» ---------------------- #
        header = QWidget(page)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)
        self._registry_title_label = QLabel("Реестр", header)
        self._registry_title_label.setObjectName("pageTitle")
        header_layout.addWidget(self._registry_title_label)
        self._registry_back_btn = QPushButton("← Назад", header)
        self._registry_back_btn.clicked.connect(self.back_to_parties)
        self._registry_back_btn.setVisible(False)
        header_layout.addWidget(self._registry_back_btn)
        header_layout.addStretch(1)
        refresh_btn = QPushButton("Обновить", header)
        refresh_btn.clicked.connect(lambda _checked=False: self.refresh())
        header_layout.addWidget(refresh_btn)
        outer.addWidget(header)

        self._registry_hint_label = QLabel(
            "Клик — выделить запись (действия — на панели под списком, "
            "клавиша Delete). Двойной клик или Enter — открыть. "
            "Стрелки листают список, Ctrl+F — поиск.", page,
        )
        self._registry_hint_label.setObjectName("pageSubtitle")
        self._registry_hint_label.setWordWrap(True)
        outer.addWidget(self._registry_hint_label)

        # -- Строка поиска: поле + «Показать всех» + сортировка (U2) ------ #
        self._registry_search_bar = QWidget(page)
        search_layout = QHBoxLayout(self._registry_search_bar)
        search_layout.setContentsMargins(0, 0, 0, 0)
        self.search_edit = QLineEdit(self._registry_search_bar)
        self.search_edit.setPlaceholderText("Поиск по названию или ИНН")  # U11
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setFixedWidth(280)
        self.search_edit.textChanged.connect(self.on_search_changed)
        self.search_edit.returnPressed.connect(self.apply_filter)
        search_layout.addWidget(self.search_edit)
        reset_btn = QPushButton("Показать всех", self._registry_search_bar)
        reset_btn.clicked.connect(lambda _checked=False: self.reset_filter())
        search_layout.addWidget(reset_btn)
        self._registry_sort_combo = QComboBox(self._registry_search_bar)
        self._registry_sort_combo.addItems(list(SORT_VALUE_BY_LABEL))
        self._registry_sort_combo.setCurrentText(SORT_LABEL_BY_VALUE[self._registry_sort])
        self._registry_sort_combo.currentTextChanged.connect(self.on_sort_changed)
        search_layout.addWidget(self._registry_sort_combo)
        search_layout.addStretch(1)
        outer.addWidget(self._registry_search_bar)

        # UX-REG: фильтр типов в виде документов контрагента; в режиме
        # контрагентов строка скрыта (видна строка поиска).
        self._registry_filter_bar = QWidget(page)
        filter_layout = QHBoxLayout(self._registry_filter_bar)
        filter_layout.setContentsMargins(0, 0, 0, 0)
        self._docs_type_combo = QComboBox(self._registry_filter_bar)
        self._docs_type_combo.addItems(list(rcore.DOCS_TYPE_VALUE_BY_LABEL))
        self._docs_type_combo.currentTextChanged.connect(self.on_docs_type_changed)
        filter_layout.addWidget(self._docs_type_combo)
        filter_layout.addStretch(1)
        self._registry_filter_bar.setVisible(False)
        outer.addWidget(self._registry_filter_bar)

        # -- Отложенный поиск (Tk 200 мс after-паритет) ------------------- #
        self._search_timer = QTimer(page)
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(200)
        self._search_timer.timeout.connect(self._deferred_refresh)

        # -- Баннеры сроков (2.18.0) над списком --------------------------- #
        self._registry_banners = QWidget(page)
        self._banners_layout = QVBoxLayout(self._registry_banners)
        self._banners_layout.setContentsMargins(0, 0, 0, 0)
        self._banners_layout.setSpacing(6)
        outer.addWidget(self._registry_banners)

        # -- Подпись списка: «Контрагенты (N)», «показаны K из M» (R-05b) -- #
        self._registry_caption = QLabel("", page)
        self._registry_caption.setObjectName("currentPartyLabel")
        outer.addWidget(self._registry_caption)
        self._registry_subtitle = QLabel("", page)
        self._registry_subtitle.setObjectName("pageSubtitle")
        self._registry_subtitle.setVisible(False)
        outer.addWidget(self._registry_subtitle)

        # -- Список (TREE-1): QTableView + диф-модель ---------------------- #
        self.table = RegistryTableQt(
            on_select=self.on_tree_select,
            on_activate=self.on_tree_activate,
            on_delete_key=self._on_delete_key,
        )
        outer.addWidget(self.table, 1)

        # -- Панель действий выделенной записи (V3) ------------------------ #
        actions = QWidget(page)
        actions_layout = QHBoxLayout(actions)
        actions_layout.setContentsMargins(0, 0, 0, 0)
        self._registry_actions_hint = QLabel(
            "Выберите запись в списке — действия появятся здесь.", actions,
        )
        self._registry_actions_hint.setObjectName("pageSubtitle")
        actions_layout.addWidget(self._registry_actions_hint)
        self._registry_action_btn1 = QPushButton("", actions)
        self._registry_action_btn1.setFixedWidth(140)
        self._registry_action_btn1.clicked.connect(self._run_action_1)
        actions_layout.addWidget(self._registry_action_btn1)
        self._registry_action_btn2 = QPushButton("", actions)
        self._registry_action_btn2.setFixedWidth(140)
        self._registry_action_btn2.clicked.connect(self._run_action_2)
        actions_layout.addWidget(self._registry_action_btn2)
        actions_layout.addStretch(1)
        self._registry_empty_trash_btn = QPushButton("Очистить корзину", actions)
        self._registry_empty_trash_btn.clicked.connect(
            lambda _checked=False: self.app.lifecycle_flow.empty_trash())
        actions_layout.addWidget(self._registry_empty_trash_btn)
        self._registry_empty_trash_btn.setVisible(False)
        outer.addWidget(actions)

        self._registry_action_btn1.setVisible(False)
        self._registry_action_btn2.setVisible(False)

        # -- Клавиатура страницы: Ctrl+F — фокус в поиск, Escape — назад --- #
        ctrl_f = QShortcut(QKeySequence("Ctrl+F"), page)
        ctrl_f.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        ctrl_f.activated.connect(self._focus_search)
        escape = QShortcut(QKeySequence("Escape"), page)
        escape.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        escape.activated.connect(self._on_escape)

        self._install_page()

    def _install_page(self) -> None:
        """Заменяет плейсхолдер M2 живой страницей «Реестр»."""
        app = self.app
        placeholder = app.pages.get("registry")
        index = app.stack.indexOf(placeholder) if placeholder is not None else -1
        if placeholder is not None:
            app.stack.removeWidget(placeholder)
            placeholder.deleteLater()
        app.stack.insertWidget(max(0, index), self.page)
        app.pages["registry"] = self.page
        if app.current_page == "registry":
            app.stack.setCurrentWidget(self.page)

    # ------------------------------------------------------------------ #
    #  Список контрагентов, проваливание, секции реестра
    # ------------------------------------------------------------------ #
    def parties(self) -> list[dict]:
        """Активные контрагенты с учётом поиска по названию и ИНН (подстрока)."""
        query = self.search_edit.text().strip().lower()
        parties = self.app.store.active_counterparties()
        if not query:
            return parties
        digits = digits_only(query)

        def matches(rec: dict) -> bool:
            fields = rec.get("fields") or {}
            inn = str(fields.get(F.INN, "")).strip()
            if digits and digits in inn:
                return True
            name_parts = (
                storage.display_name(rec),
                str(fields.get(F.FULL_NAME, "")),
                str(fields.get(F.IP_NAME, "")),
                str(fields.get(F.SHORT_NAME, "")),
            )
            haystack = " ".join(name_parts).lower()
            return bool(query) and query in haystack

        return [rec for rec in parties if matches(rec)]

    def sorted_parties(self) -> list[dict]:
        """parties() в выбранном порядке сортировки (U2: добавление/алфавит/ИНН).

        Поиск всегда выполняется по всей базе — сортировка применяется
        к отфильтрованному множеству, лимит отображения (R-05b) — к
        отсортированному.
        """
        parties = self.parties()
        if self._registry_sort == "name":
            parties.sort(key=lambda rec: storage.display_name(rec).casefold())
        elif self._registry_sort == "inn":
            # Пустой ИНН — в конец списка, а не в начало.
            parties.sort(key=lambda rec: (
                str((rec.get("fields") or {}).get(F.INN, "")).strip() == "",
                str((rec.get("fields") or {}).get(F.INN, "")).strip(),
            ))
        return parties

    @property
    def is_drill_mode(self) -> bool:
        """True, пока открыт вид «Документы контрагента» (проваливание)."""
        return self._registry_drill_party_id is not None

    def on_page_open(self) -> None:
        """Открытие страницы «Реестр» из меню: список контрагентов с начала.

        Проваливание — режим просмотра, он не переживает уход со страницы:
        вернулись в реестр — снова виден список контрагентов (2.19.0).
        """
        if self.is_drill_mode:
            self.back_to_parties()
            return
        self.refresh()
        self.table.focus_tree()

    def on_search_changed(self, *_args) -> None:
        """Фильтрует реестр по мере ввода с небольшой задержкой (Tk 200 мс)."""
        self._search_timer.start()  # перезапуск заменяет предыдущий отсчёт

    def _deferred_refresh(self) -> None:
        self.refresh()

    def apply_filter(self) -> None:
        """Немедленное применение фильтра (Enter в поиске, смена сортировки)."""
        self._search_timer.stop()
        self.refresh()

    def reset_filter(self) -> None:
        self.search_edit.clear()
        self.apply_filter()

    def on_sort_changed(self, label: str) -> None:
        """Смена сортировки (U2): перестройка списка и settings.json."""
        value = SORT_VALUE_BY_LABEL.get(label)
        if value is None or value == self._registry_sort:
            return
        self._registry_sort = value
        save = getattr(self.app, "save_setting", None)
        if save is not None:
            save("registry_sort", value)
        self.apply_filter()

    def on_docs_type_changed(self, label: str) -> None:
        """Смена фильтра типов в виде документов (UX-REG)."""
        value = rcore.DOCS_TYPE_VALUE_BY_LABEL.get(label)
        if value is None or value == self._docs_type_filter:
            return
        self._docs_type_filter = value
        self.apply_filter()

    def show_more(self) -> None:
        """Строка «Показать ещё» (R-05b): следующая порция списка.

        Фильтр, сортировка, выделение и позиция прокрутки сохраняются:
        популяция идёт дифом (append новых строк), подпись обновляется.
        """
        self._registry_shown += rcore.DISPLAY_SOFT_LIMIT
        self.refresh()

    def _focus_search(self) -> None:
        self.search_edit.setFocus()
        self.search_edit.selectAll()

    def _on_escape(self) -> None:
        """Escape: выход из drill-вида к списку контрагентов (2.19.0)."""
        if self.is_drill_mode:
            self.back_to_parties()

    def _on_delete_key(self) -> None:
        self.delete_selected()

    # ------------------------------------------------------------------ #
    #  Выделение и активация строк
    # ------------------------------------------------------------------ #
    def on_tree_select(self, iid: str | None, key_nav: bool = False,
                       user: bool = False) -> None:
        """Событие выделения таблицы (SET-семантика, идемпотентно).

        События вида (клик по новой строке, стрелки, эхо программной
        синхронизации) приходят с user=False и только устанавливают
        состояние — повторное уведомление не меняет его. Снятие выделения
        идёт только явным действием пользователя (повторный клик,
        click-toggle вида): user=True и iid=None.
        """
        if iid is None:
            if not user:
                return  # эхо популяции/синхронизации — состояние не трогаем
            if self.is_drill_mode:
                self._docs_selected_id = None
            else:
                self._registry_selected_id = None
                self._registry_selected_extra = None
            self.refresh_selection()
            return
        row = self.table.row(iid)
        if row is None:
            return
        data = row.data
        kind = data.get("kind")
        rec = data.get("rec")
        if kind == "party":
            self._registry_selected_extra = None
            self._docs_selected_id = None
            self._select_party_plain(rec)
        elif kind == "doc":
            self._registry_selected_extra = None
            self._select_document_plain(rec)
        else:
            # Секции корзины/неопознанных и служебная строка «Показать ещё»:
            # выделение запоминается для панели действий, клавиша Delete
            # на них не действует (там выбор restore/purge — через панель).
            self._registry_selected_extra = {"kind": kind, "rec": rec}
            self._registry_selected_id = None
            self._docs_selected_id = None
            self.refresh_selection()

    def on_tree_activate(self, iid: str) -> None:
        """Активация строки: двойной клик или Enter (U13)."""
        if iid == SHOW_MORE_IID:
            self.show_more()
            return
        row = self.table.row(iid)
        if row is None:
            return
        data = row.data
        if data.get("kind") == "party":
            self.open_party_documents(data["rec"])
        elif data.get("kind") == "doc":
            self.app.lifecycle_flow.export_document(data["rec"])

    def select_party(self, rec: dict) -> None:
        """Программный выбор контрагента с toggle (Tk-паритет select_party)."""
        if self._registry_selected_id == rec.get("id"):
            self._registry_selected_id = None
        else:
            self._registry_selected_id = rec.get("id")
            # Выбор из реестра переносится на страницы документов (п. 4.2).
            self.app.set_current_party(rec)
        self._docs_selected_id = None
        if not self._registry_rows:
            # Страница ещё не отрисована (ленивый рендер) — рисуем целиком.
            self.refresh()
            return
        self.refresh_selection()

    def _select_party_plain(self, rec: dict) -> None:
        """Клик/стрелки (U13): выделить строку без toggle."""
        self._registry_selected_id = rec.get("id")
        self._registry_selected_extra = None
        self._docs_selected_id = None
        self.app.set_current_party(rec)
        self.refresh_selection()

    def open_party_documents(self, rec: dict) -> None:
        """Двойной клик по контрагенту: провалиться в его документы (2.19.0).

        Страница переключается на вид «Документы — <имя>»: список документов
        контрагента, кнопка «← Назад» в шапке возвращает к контрагентам.
        """
        self._registry_selected_id = rec.get("id")
        self.app.set_current_party(rec)
        self._docs_selected_id = None
        self._registry_drill_party_id = rec.get("id")
        self.refresh()

    def back_to_parties(self) -> None:
        """Кнопка «← Назад» (или Escape): вернуться к списку контрагентов."""
        self._registry_drill_party_id = None
        self._docs_selected_id = None
        self.refresh()

    def select_document(self, rec: dict) -> None:
        """Программный выбор документа с toggle (Tk-паритет select_document)."""
        if self._docs_selected_id == rec.get("id"):
            self._docs_selected_id = None
        else:
            self._docs_selected_id = rec.get("id")
        if not self._docs_rows:
            # Вид документов ещё не отрисован — рисуем целиком.
            self.refresh()
            return
        self.refresh_selection()

    def _select_document_plain(self, rec: dict) -> None:
        """Клик/стрелки (U13): выделить документ без toggle."""
        self._docs_selected_id = rec.get("id")
        self._registry_selected_extra = None
        self.refresh_selection()

    def forget_party(self, party_id) -> None:
        """Сброс выделения удалённого контрагента (хвост операции корзины).

        Tk delete_counterparty снимает только выделение карточки;
        drill-вид с удалённым контрагентом закрывается тихо при
        следующем refresh (build_model возвращает список, 2.19.0).
        """
        if self._registry_selected_id == party_id:
            self._registry_selected_id = None

    def delete_selected(self) -> bool:
        """Клавиша Delete: удалить выделенный документ или контрагента.

        Документ имеет приоритет (он «мельче»): если выделены и документ,
        и контрагент, удаляется документ. Возвращает True, если удаление
        инициировано (диалог подтверждения — в потоке жизненного цикла).

        Внутри вида «Документы контрагента» (проваливание, 2.19.0) Delete
        удаляет только выделенный документ — контрагента нельзя удалить
        случайно из его списка документов. Выделенной записи корзины
        Delete не касается: у неё выбор «Восстановить»/«Удалить навсегда» —
        через панель действий, случайное удаление навсегда недопустимо.
        """
        if self._registry_selected_extra:
            return False
        doc_id = self._docs_selected_id
        if doc_id is not None:
            rec = self.app.store.get_document(doc_id)
            if rec is not None:
                self.app.lifecycle_flow.delete_document(rec)
                return True
        if self.is_drill_mode:
            return False
        party_id = self._registry_selected_id
        if party_id is not None:
            rec = self.app.store.get_counterparty(party_id)
            if rec is not None:
                self.app.lifecycle_flow.delete_counterparty(rec)
                return True
        return False

    # ------------------------------------------------------------------ #
    #  Модель строк реестра (TREE-1: R-03/R-04, R-05b)
    # ------------------------------------------------------------------ #
    def _view_signature(self) -> tuple:
        """Что заставляет пересобрать список целиком: фильтр, сортировка,
        фильтр типов, drill-переход. Изменение состава и правки записей в
        сигнатуру не входят — они идут дифом (R-04)."""
        return (
            self._registry_drill_party_id,
            self.search_edit.text().strip().lower(),
            self._registry_sort,
            self._docs_type_filter,
        )

    def _doc_until(self, rec: dict) -> "str | None":
        """Срок строки документа из лёгкой карты хранилища (LAZY-1)."""
        if rec.get("kind") != storage.KIND_ADDITIONAL:
            return None
        return self.app.store.expiry_until(rec.get("id"))

    def _doc_dot(self, rec: dict) -> str | None:
        """Цвет точки типа документа (V4-пилот); None — точки нет."""
        kind = rec.get("kind")
        return {
            storage.KIND_CONTRACT: rcore.DOT_CONTRACT,
            storage.KIND_ADDITIONAL: rcore.DOT_ADDITIONAL,
            storage.KIND_SPEC: rcore.DOT_SPEC,
        }.get(kind)

    def _party_rows(self) -> tuple[list, str]:
        """Модель списка контрагентов: порция (R-05b) + служебная строка."""
        parties = self.sorted_parties()
        total = len(parties)
        shown, has_more = rcore.plan_display_count(total, self._registry_shown)
        rows = [
            rcore.RowSpec(
                iid=f"{rcore.IID_PARTY}{rec.get('id')}",
                values=party_row_values(rec),
                tags=(),
                dot=party_row_dot(rec),
                data={"kind": "party", "rec": rec},
            )
            for rec in parties[:shown]
        ]
        if has_more:
            hidden = total - shown
            portion = min(rcore.DISPLAY_SOFT_LIMIT, hidden)
            rows.append(rcore.RowSpec(
                iid=SHOW_MORE_IID,
                values=(f"Показать ещё ({portion} из {hidden})", "", "", "", ""),
                tags=("showmore",),
                data={"kind": "showmore"},
            ))
        query = self.search_edit.text().strip()
        caption = (
            f"Найдено по запросу «{query}»: {total}"
            if query else f"Контрагенты ({total})"
        )
        if has_more:
            caption += f" · показаны {shown} из {total}"
        elif shown > rcore.DISPLAY_SOFT_LIMIT and total > rcore.DISPLAY_SOFT_LIMIT:
            caption += f" · показаны все {total}"
        return rows, caption

    def _doc_rows(self, party: dict) -> tuple[list, str, str]:
        """Модель вида «Документы контрагента» с фильтром типов (UX-REG)."""
        docs = self.app.store.documents_for(party["id"])
        if self._docs_type_filter != "all":
            docs = [rec for rec in docs if rec.get("kind") == self._docs_type_filter]
        rows = [
            rcore.RowSpec(
                iid=f"{rcore.IID_DOC}{rec.get('id')}",
                values=doc_row_values(rec, until=self._doc_until(rec)),
                tags=doc_row_tags(rec, until=self._doc_until(rec)),
                dot=self._doc_dot(rec),
                data={"kind": "doc", "rec": rec},
            )
            for rec in docs
        ]
        caption = f"Документы — {storage.display_name(party)} ({len(docs)})"
        if self._docs_type_filter != "all":
            caption += (
                " · "
                + rcore.DOCS_TYPE_LABEL_BY_VALUE.get(
                    self._docs_type_filter, self._docs_type_filter).lower()
            )
        return rows, caption, party_subtitle(party)

    def _unknown_rows(self) -> list:
        """Секция «Неопознанные документы» (модель, Tk-паритет 2.4)."""
        docs = [
            rec for rec in self.app.store.documents
            if rec.get("kind") == storage.KIND_UNKNOWN
            and not storage.is_deleted(rec)
        ]
        if not docs:
            return []
        rows = [rcore.RowSpec(
            iid=rcore.IID_SEC_UNKNOWN,
            values=(f"Неопознанные документы ({len(docs)}) — файлы скопированы "
                    "в папку «unknown» внутри папки документов", "", "", "", ""),
            tags=("section",),
            data={"kind": "section"},
        )]
        for rec in docs:
            source = str(
                (self.app.store.document_body(rec.get("id")).get("extra")
                 or {}).get("source_file") or ""
            )
            name = source.replace("\\", "/").rsplit("/", 1)[-1] if source else "Документ"
            rows.append(rcore.RowSpec(
                iid=f"{rcore.IID_UNKNOWN}{rec.get('id')}",
                values=(name, "файл", "—", "—", "папка «unknown»"),
                tags=("muted",),
                dot=rcore.DOT_UNKNOWN,
                data={"kind": "unknown", "rec": rec},
            ))
        return rows

    def _trash_rows(self) -> list:
        """Секция «Корзина» (модель, Tk-паритет): удалённые записи реестра."""
        trashed_parties = self.app.store.trashed_counterparties()
        trashed_docs = [
            rec for rec in self.app.store.trashed_documents()
            if not rec.get("deleted_with_party")
        ]
        if not trashed_parties and not trashed_docs:
            return []
        total = len(trashed_parties) + len(trashed_docs)
        rows = [rcore.RowSpec(
            iid=rcore.IID_SEC_TRASH,
            values=(f"Корзина ({total}) — записи удалены из реестра, файлы на "
                    "диске остались; восстановление вернёт их в работу",
                    "", "", "", ""),
            tags=("section",),
            data={"kind": "section"},
        )]
        for party in trashed_parties:
            docs = self.app.store.documents_for(
                party.get("id"), include_deleted=True)
            rows.append(rcore.RowSpec(
                iid=f"{rcore.IID_TRASH_PARTY}{party.get('id')}",
                values=(f"Контрагент «{storage.display_name(party)}»", "корзина",
                        "—", str(party.get("id")), f"документов: {len(docs)}"),
                tags=("muted",),
                data={"kind": "trash_party", "rec": party},
            ))
        for rec in trashed_docs:
            party_name = self.app.store.party_name(rec.get("party_id"))
            doc, date, expiry = doc_row_values(rec, until=self._doc_until(rec))
            rows.append(rcore.RowSpec(
                iid=f"{rcore.IID_TRASH_DOC}{rec.get('id')}",
                values=(f"{doc} · {party_name}" if party_name else doc,
                        date, expiry or "—", "", "корзина"),
                tags=("muted",),
                data={"kind": "trash_doc", "rec": rec},
            ))
        return rows

    def build_model(self) -> tuple[list, str, str]:
        """Модель строк текущего вида + подпись + подстрока.

        Контракт R-05b: поиск и сортировка — по всей базе; порция режет
        только отображение. Секции «Неопознанные» и «Корзина» добавляются
        в модель после порции контрагентов.
        """
        drill_id = self._registry_drill_party_id
        if drill_id is not None:
            party = self.app.store.get_counterparty(drill_id)
            if party is not None and not storage.is_deleted(party):
                rows, caption, subtitle = self._doc_rows(party)
                return rows, caption, subtitle
            # Контрагент удалён (мягкое удаление) — тихий возврат к списку.
            self._registry_drill_party_id = None
            self._docs_selected_id = None
        rows, caption = self._party_rows()
        rows = rows + self._unknown_rows() + self._trash_rows()
        return rows, caption, ""

    # ------------------------------------------------------------------ #
    #  Перерисовка: полная и диф (R-04)
    # ------------------------------------------------------------------ #
    def refresh(self) -> None:
        """Обновляет реестр (поиск, изменения состава, проваливание).

        Полная популяция — при смене фильтра/сортировки/фильтра типов и
        drill-переходах (сигнатура вида) и при неотрисованном списке;
        остальное — диф (R-04). Длительные перерисовки (более 500 мс)
        попадают в журнал (п. 10 этапа 5).
        """
        with applog.timed("Перерисовка реестра", 500, "ui"):
            self.refresh_smart()
        self._registry_rendered = True

    def refresh_smart(self) -> None:
        """Выбирает между полной популяцией и дифом (R-04)."""
        if self.table is None or not self._registry_rendered \
                or not self._registry_rows and not self.is_drill_mode:
            self.refresh_full()
            return
        # Контрагент drill-вида мог быть удалён — возврат меняет сигнатуру.
        drill_id = self._registry_drill_party_id
        if drill_id is not None:
            party = self.app.store.get_counterparty(drill_id)
            if party is None or storage.is_deleted(party):
                self.refresh_full()
                return
        if self._view_signature() != self._registry_signature:
            self.refresh_full()
            return
        rows, caption, subtitle = self.build_model()
        self.table.apply_diff(rows)
        self._after_model(rows, caption, subtitle)

    def refresh_full(self) -> None:
        """Полная популяция списка (смена фильтра, сортировки, drill)."""
        self._registry_signature = self._view_signature()
        self._registry_shown = rcore.DISPLAY_SOFT_LIMIT
        columns = DOC_COLUMNS if self.is_drill_mode else PARTY_COLUMNS
        rows, caption, subtitle = self.build_model()
        self.table.populate(columns, rows)
        self._update_header_mode()
        self._after_model(rows, caption, subtitle)

    def _after_model(self, rows: list, caption: str, subtitle: str) -> None:
        """После популяции/дифа: карты состояния, подпись, баннеры, выделение.

        Контрольный assert (R-04, debug-сборка): порядок строк таблицы
        совпадает с моделью — рассинхрон карты id->iid невозможен
        (assert внутри модели).
        """
        self._registry_rows = {
            row.data["rec"].get("id"): {"rec": row.data["rec"], "iid": row.iid}
            for row in rows if row.data.get("kind") == "party"
        }
        self._docs_rows = {
            row.data["rec"].get("id"): {"rec": row.data["rec"], "iid": row.iid}
            for row in rows if row.data.get("kind") == "doc"
        }
        self._update_caption(caption, subtitle)
        self._update_banners()
        self.refresh_selection()

    def _update_caption(self, caption: str, subtitle: str) -> None:
        """Подпись списка и подстрока drill-вида."""
        try:
            self._registry_caption.setText(caption)
            if subtitle:
                self._registry_subtitle.setText(subtitle)
                self._registry_subtitle.setVisible(True)
            else:
                self._registry_subtitle.setVisible(False)
        except RuntimeError:
            pass  # страница могла быть разрушена

    def _update_banners(self) -> None:
        """Баннер сроков доп. соглашений над списком (2.18.0).

        Перестраивается при каждом обновлении: 0–1 компактный фрейм;
        баннер файлов вне реестра появится вместе со сверкой (M6,
        START-1 — в Qt-сборке сверки ещё нет, ``_orphans`` пуст).
        """
        layout = self._banners_layout
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        try:
            buckets = agreements.collect(self.app.store)
        except Exception:
            return
        expired, soon = buckets["expired"], buckets["soon"]
        if not expired and not soon:
            return
        lines: list[str] = []
        if expired:
            lines.append(f"Просрочено доп. соглашений: {len(expired)} — "
                         "строки отмечены красным.")
        if soon:
            lines.append(f"Истекает в течение месяца: {len(soon)} — "
                         "строки отмечены жёлтым.")
        banner = QLabel("\n".join(lines), self._registry_banners)
        banner.setObjectName("cardCaption")
        banner.setWordWrap(True)
        self._banners_layout.addWidget(banner)

    # ------------------------------------------------------------------ #
    #  Выделение и панель действий (V3)
    # ------------------------------------------------------------------ #
    def refresh_selection(self) -> None:
        """Частичное обновление при смене выбора: подсветка строки и панель.

        Строки не пересоздаются (5.1): у таблицы меняется только выделение,
        панель действий перестраивается под тип выделенной записи.
        """
        iid = self._selected_iid()
        self.table.set_selection(iid)
        self.refresh_action_panel()

    def _selected_iid(self) -> str | None:
        """iid видимой строки по состоянию выделения (или None)."""
        if self.is_drill_mode:
            doc_id = self._docs_selected_id
            if doc_id is None:
                return None
            iid = f"{rcore.IID_DOC}{doc_id}"
        else:
            extra = self._registry_selected_extra
            if extra is not None:
                prefix = {
                    "trash_party": rcore.IID_TRASH_PARTY,
                    "trash_doc": rcore.IID_TRASH_DOC,
                    "unknown": rcore.IID_UNKNOWN,
                }.get(extra["kind"])
                iid = f"{prefix}{extra['rec'].get('id')}" if prefix else None
            else:
                party_id = self._registry_selected_id
                iid = f"{rcore.IID_PARTY}{party_id}" if party_id is not None else None
        if iid is None or self.table.row(iid) is None:
            return None  # выделенная запись не видна (фильтр/порция)
        return iid

    def _selected_entry(self) -> tuple[str, dict | None]:
        """Тип и запись выделенной строки для панели действий."""
        iid = self._selected_iid()
        if iid is None:
            return "none", None
        row = self.table.row(iid)
        if row is None:
            return "none", None
        return row.data.get("kind", "none"), row.data.get("rec")

    def refresh_action_panel(self) -> None:
        """Кнопки панели под списком по типу выделенной записи (V3).

        Панель имеет постоянное место в компоновке — кнопки не прыгают
        при выделении и снятии выделения (резерв места, V3 UX-аудита).
        """
        kind, rec = self._selected_entry()
        labels = action_labels_for(kind)
        if labels is None or rec is None:
            self._registry_actions_hint.setVisible(True)
            self._registry_action_btn1.setVisible(False)
            self._registry_action_btn2.setVisible(False)
        else:
            self._registry_actions_hint.setVisible(False)
            self._registry_action_btn1.setText(labels[0])
            self._registry_action_btn1.setVisible(True)
            self._registry_action_btn2.setText(labels[1])
            self._registry_action_btn2.setVisible(True)
        has_trash = bool(self.app.store.trashed_counterparties()) or \
            bool(self.app.store.trashed_documents())
        self._registry_empty_trash_btn.setVisible(has_trash)

    def _run_action_1(self) -> None:
        kind, rec = self._selected_entry()
        if rec is not None:
            self.run_action(kind, rec)

    def _run_action_2(self) -> None:
        kind, rec = self._selected_entry()
        if rec is not None:
            self.run_action(kind, rec, second=True)

    def run_action(self, kind: str, rec: dict, second: bool = False) -> None:
        """Действие панели: маршрутизация к потокам (как у строк 2.18.0)."""
        app = self.app
        if kind == "party":
            if second:
                app.lifecycle_flow.delete_counterparty(rec)
            else:
                app.log.info("Открытие карточки из реестра: id=%s", rec.get("id"))
                counterparty_page_qt.open_edit_dialog(app, rec)
        elif kind == "doc":
            if second:
                app.lifecycle_flow.delete_document(rec)
            else:
                app.lifecycle_flow.export_document(rec)
        elif kind in ("trash_party", "trash_doc"):
            if second:
                if kind == "trash_party":
                    app.lifecycle_flow.purge_counterparty(rec)
                else:
                    app.lifecycle_flow.purge_document(rec)
            else:
                if kind == "trash_party":
                    app.lifecycle_flow.restore_counterparty(rec)
                else:
                    app.lifecycle_flow.restore_document(rec)
        elif kind == "unknown":
            if second:
                app.lifecycle_flow.delete_document(rec)
            else:
                source = str(
                    (app.store.document_body(rec.get("id")).get("extra")
                     or {}).get("source_file") or ""
                )
                app.show_in_folder(source)

    def _update_header_mode(self) -> None:
        """Шапка страницы по режиму: «Реестр» или «Документы — <имя>» (2.19.0)."""
        drill_id = self._registry_drill_party_id
        if drill_id is not None:
            party = self.app.store.get_counterparty(drill_id)
            name = storage.display_name(party) if party is not None else ""
            self._registry_title_label.setText(
                f"Документы — {name}" if name else "Документы")
            self._registry_hint_label.setText(
                "Клик — выделить документ («Открыть Word»/«Удалить» на панели, "
                "клавиша Delete). Двойной клик или Enter — открыть в Word.")
            self._registry_back_btn.setVisible(True)
            self._registry_search_bar.setVisible(False)
            self._registry_filter_bar.setVisible(True)
        else:
            self._registry_title_label.setText("Реестр")
            self._registry_hint_label.setText(
                "Клик — выделить запись (действия — на панели под списком, "
                "клавиша Delete). Двойной клик или Enter — открыть. "
                "Стрелки листают список, Ctrl+F — поиск.")
            self._registry_back_btn.setVisible(False)
            self._registry_filter_bar.setVisible(False)
            self._registry_search_bar.setVisible(True)


# Поздний импорт: окно правки карточки (M4) — без циклической зависимости
# при построении страниц каркасом.
import counterparty_page_qt  # noqa: E402  (внизу модуля намеренно)
