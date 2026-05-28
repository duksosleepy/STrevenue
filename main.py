import calendar
import configparser
import re
from datetime import date
from pathlib import Path

import pandas as pd
from imgui_bundle import hello_imgui, icons_fontawesome_6 as fa, imgui, portable_file_dialogs

CONFIG_PATH = Path(__file__).parent / "Config.ini"


def load_config() -> dict:
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_PATH, encoding="utf-8-sig")
    exp = cfg["EXP"] if "EXP" in cfg else {}
    return {
        "path_report": exp.get("PathReport", "C:\\INTERFACE"),
        "type_file": exp.get("TypeFile", ".txt"),
        "shop_id": exp.get("ShopID", ""),
        "vat": exp.get("VAT", "1"),
        "starts_with": exp.get("StartsWith", "H"),
        "separator": exp.get("Separator", "|"),
        "stt": int(exp.get("STT", "0")),
    }


def save_stt(stt: int):
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_PATH, encoding="utf-8-sig")
    if "EXP" not in cfg:
        cfg["EXP"] = {}
    cfg["EXP"]["STT"] = str(stt)
    with open(CONFIG_PATH, "w", encoding="utf-8-sig") as f:
        cfg.write(f)


def _parse_gio(val) -> str:
    try:
        g = str(val).split(':')[1]
        return f"{g[0:2]}:{g[2:4]}:{g[4:6]}"
    except Exception:
        return "12:00:00"


def _to_float(val) -> float:
    try:
        v = float(val)
        return 0.0 if pd.isna(v) else v
    except (ValueError, TypeError):
        return 0.0


def parse_excel(file_path: str) -> list[dict]:
    df = pd.read_excel(
        file_path,
        sheet_name=0,
        header=None,
        skiprows=5,
        dtype=str,
        engine="calamine",
    )
    rows = []
    for _, row in df.iterrows():
        def cell(i, _row=row):
            v = _row.iloc[i] if i < len(_row) else None
            return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()

        raw_date = cell(3)
        try:
            dt = pd.to_datetime(raw_date, dayfirst=False).date()
        except Exception:
            dt = date.today()

        rows.append({
            "date": dt,
            "gio":  _parse_gio(cell(20)),
            "tong": _to_float(cell(17)),
            "ckkm": _to_float(cell(18)),
            "the":  _to_float(cell(12)),
            "tm":   _to_float(cell(13)),
        })
    return rows


def do_export(file_path: str, export_date: date, cfg: dict) -> tuple[bool, str]:
    try:
        rows = parse_excel(file_path)
    except Exception as e:
        return False, f"Không thể đọc file:\n{e}"

    if not rows:
        return False, "File không có dữ liệu."

    day_rows = [r for r in rows if r["date"] == export_date]

    shop_id  = "11000023"
    sep      = "|"
    typefile = ".TXT"
    out_dir  = Path("C:\\Interface")
    try:
        shop_id  = cfg.get("shop_id") or shop_id
        sep      = cfg.get("separator") or sep
        typefile = cfg.get("type_file") or typefile
        out_dir  = Path(cfg.get("path_report") or out_dir)
    except Exception:
        pass

    stt = cfg["stt"] + 1
    str_ngay = export_date.strftime("%d%m%Y")
    out_filename = f"H{shop_id}_{export_date.strftime('%Y%m%d')}{typefile}"
    out_path = out_dir / out_filename

    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        if out_path.exists():
            out_path.unlink()
    except Exception as e:
        return False, f"Không thể tạo thư mục xuất:\n{e}"

    lines = []
    for hour in range(24):
        h_start = f"{hour:02d}:00:00"
        h_end   = f"{hour:02d}:59:59"
        nd = [r for r in day_rows if h_start <= r["gio"] <= h_end]
        ls_h         = len(nd)
        sohd         = sum(r["ckkm"] for r in nd)
        sum_tong     = sum(r["tong"] for r in nd)
        giamgia      = sum_tong - sohd
        tongdoanhthu = giamgia / 8.0
        gio_out      = sum(r["the"] for r in nd)
        tm_out       = sum(r["tm"]  for r in nd)
        lines.append(sep.join([
            shop_id, str(stt), str_ngay, f"{hour:02d}",
            str(ls_h), f"{giamgia:.2f}", f"{tongdoanhthu:.2f}", f"{sohd:.2f}",
            "0.00", str(ls_h), f"{gio_out:.2f}", f"{tm_out:.2f}",
            "0.00", "0.00", "0.00", "0.00", "0.00", "R",
        ]))

    try:
        with open(out_path, "w", encoding="utf-8-sig") as f:
            f.write("\n".join(lines) + "\n")
    except Exception as e:
        return False, f"Không thể ghi file xuất:\n{e}"

    cfg["stt"] = stt
    save_stt(stt)

    total_tx = len(day_rows)
    return True, f"Đã xuất dữ liệu thành công!\n\n{total_tx} giao dịch → {out_path}"


MONTH_NAMES = [
    "Thang 1", "Thang 2", "Thang 3", "Thang 4",
    "Thang 5", "Thang 6", "Thang 7", "Thang 8",
    "Thang 9", "Thang 10", "Thang 11", "Thang 12",
]
DAY_HEADERS = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]


def draw_calendar_popup(selected: date, nav_year: int, nav_month: int) -> tuple[date, int, int, bool]:
    picked = False

    cell_w = 28.0
    cal_w = cell_w * 7 + 6 * 4
    arrow_w = 24.0
    header_label = f"{MONTH_NAMES[nav_month - 1]} {nav_year}"
    label_w = imgui.calc_text_size(header_label).x
    start_x = imgui.get_cursor_pos_x()

    if imgui.arrow_button("##prev", imgui.Dir_.left):
        nav_month -= 1
        if nav_month < 1:
            nav_month = 12
            nav_year -= 1

    imgui.same_line()
    pad = (cal_w - arrow_w * 2 - label_w) * 0.5
    imgui.set_cursor_pos_x(start_x + arrow_w + max(pad, 0))
    imgui.text(header_label)
    imgui.same_line()
    imgui.set_cursor_pos_x(start_x + cal_w - arrow_w)

    if imgui.arrow_button("##next", imgui.Dir_.right):
        nav_month += 1
        if nav_month > 12:
            nav_month = 1
            nav_year += 1

    imgui.spacing()

    imgui.push_style_color(imgui.Col_.text, imgui.ImVec4(0.5, 0.5, 0.5, 1.0))
    for i, h in enumerate(DAY_HEADERS):
        if i > 0:
            imgui.same_line(spacing=4)
        hw = imgui.calc_text_size(h).x
        imgui.set_cursor_pos_x(imgui.get_cursor_pos().x + (cell_w - hw) * 0.5)
        imgui.text(h)
    imgui.pop_style_color()

    imgui.separator()

    first_weekday, num_days = calendar.monthrange(nav_year, nav_month)
    col = first_weekday

    for day in range(1, num_days + 1):
        if col > 0 and day == 1:
            for _ in range(col):
                imgui.dummy(imgui.ImVec2(cell_w, 0))
                imgui.same_line(spacing=4)

        is_selected = (day == selected.day and nav_month == selected.month and nav_year == selected.year)
        is_today = (day == date.today().day and nav_month == date.today().month and nav_year == date.today().year)

        if is_selected:
            imgui.push_style_color(imgui.Col_.button, imgui.ImVec4(0.26, 0.59, 0.98, 1.0))
            imgui.push_style_color(imgui.Col_.button_hovered, imgui.ImVec4(0.26, 0.59, 0.98, 0.85))
            imgui.push_style_color(imgui.Col_.text, imgui.ImVec4(1, 1, 1, 1))
        elif is_today:
            imgui.push_style_color(imgui.Col_.button, imgui.ImVec4(0.26, 0.59, 0.98, 0.3))
            imgui.push_style_color(imgui.Col_.button_hovered, imgui.ImVec4(0.26, 0.59, 0.98, 0.5))
            imgui.push_style_color(imgui.Col_.text, imgui.ImVec4(0.26, 0.59, 0.98, 1.0))

        tw = imgui.calc_text_size(str(day)).x
        imgui.push_style_var(imgui.StyleVar_.frame_padding, imgui.ImVec2((cell_w - tw) * 0.5 - 1, 2))
        if imgui.button(f"{day}##d{day}", imgui.ImVec2(cell_w, 0)):
            selected = date(nav_year, nav_month, day)
            picked = True
            imgui.close_current_popup()
        imgui.pop_style_var()

        if is_selected or is_today:
            imgui.pop_style_color(3)

        col += 1
        if col == 7:
            col = 0
        else:
            if day < num_days:
                imgui.same_line(spacing=4)

    return selected, nav_year, nav_month, picked


class AppState:
    def __init__(self):
        self.cfg = load_config()
        self.selected_file: str = ""
        self.export_date: date = date.today()
        self.cal_nav_year: int = date.today().year
        self.cal_nav_month: int = date.today().month
        self.file_dialog = None
        self.show_error_popup: bool = False
        self.show_success_popup: bool = False
        self.popup_msg: str = ""

    def open_file_browser(self):
        path_hint = str(Path(self.selected_file).parent) if self.selected_file else "."
        self.file_dialog = portable_file_dialogs.open_file(
            "Chọn file doanh thu",
            path_hint,
            ["Tệp Excel (*.xls *.xlsx)", "*.xls *.xlsx", "Tất cả tệp (*.*)", "*.*"],
        )

    def poll_file_dialog(self):
        if self.file_dialog is not None and self.file_dialog.ready():
            results = self.file_dialog.result()
            if results:
                self.selected_file = results[0]
                self._extract_date_from_filename(results[0])
            self.file_dialog = None

    def _extract_date_from_filename(self, path: str):
        name = Path(path).stem
        m = re.match(r'^(\d{1,2})\.(\d{1,2})$', name)
        if m:
            try:
                self._set_date(date(date.today().year, int(m.group(2)), int(m.group(1))))
                return
            except ValueError:
                pass
        m = re.search(r'(\d{4})(\d{2})(\d{2})', name)
        if m:
            try:
                self._set_date(date(int(m.group(1)), int(m.group(2)), int(m.group(3))))
                return
            except ValueError:
                pass
        m = re.search(r'(\d{2})(\d{2})(\d{4})', name)
        if m:
            try:
                self._set_date(date(int(m.group(3)), int(m.group(2)), int(m.group(1))))
            except ValueError:
                pass

    def _set_date(self, d: date):
        self.export_date = d
        self.cal_nav_year = d.year
        self.cal_nav_month = d.month

    def do_export(self):
        if not self.selected_file:
            self.popup_msg = "Vui lòng chọn file trước."
            self.show_error_popup = True
            return
        ok, msg = do_export(self.selected_file, self.export_date, self.cfg)
        self.popup_msg = msg
        if ok:
            self.show_success_popup = True
        else:
            self.show_error_popup = True


_state: AppState | None = None

WINDOW_W = 380
WINDOW_H = 160


def gui():
    global _state
    assert _state is not None

    _state.poll_file_dialog()

    io = imgui.get_io()
    vp_size = io.display_size

    imgui.set_next_window_pos(
        imgui.ImVec2(vp_size.x * 0.5, vp_size.y * 0.5),
        imgui.Cond_.always,
        imgui.ImVec2(0.5, 0.5),
    )
    imgui.set_next_window_size(imgui.ImVec2(WINDOW_W, WINDOW_H), imgui.Cond_.always)

    flags = (
        imgui.WindowFlags_.no_collapse
        | imgui.WindowFlags_.no_resize
        | imgui.WindowFlags_.no_move
    )
    imgui.begin("XUẤT FILE", None, flags)

    LABEL_COL = 90.0

    imgui.text("Chọn file")
    imgui.same_line(LABEL_COL)

    avail = imgui.get_content_region_avail().x
    text_w = avail - 30
    imgui.set_next_item_width(text_w)
    _, _state.selected_file = imgui.input_text("##file", _state.selected_file, imgui.InputTextFlags_.read_only)
    imgui.same_line()
    if imgui.button("..."):
        _state.open_file_browser()

    imgui.text("Ngày xuất")
    imgui.same_line(LABEL_COL)

    date_label = f"{fa.ICON_FA_CALENDAR_DAYS}  {_state.export_date.strftime('%d/%m/%Y')}"
    imgui.push_style_color(imgui.Col_.button,         imgui.get_style_color_vec4(imgui.Col_.frame_bg))
    imgui.push_style_color(imgui.Col_.button_hovered, imgui.get_style_color_vec4(imgui.Col_.frame_bg_hovered))
    imgui.push_style_color(imgui.Col_.button_active,  imgui.get_style_color_vec4(imgui.Col_.frame_bg_active))
    imgui.push_style_var(imgui.StyleVar_.button_text_align, imgui.ImVec2(0.0, 0.5))
    if imgui.button(date_label, imgui.ImVec2(avail, 0)):
        pass
    imgui.pop_style_var()
    imgui.pop_style_color(3)
    if imgui.is_item_clicked():
        _state.cal_nav_year = _state.export_date.year
        _state.cal_nav_month = _state.export_date.month
        imgui.open_popup("##calendar")

    if imgui.begin_popup("##calendar"):
        result = draw_calendar_popup(
            _state.export_date,
            _state.cal_nav_year,
            _state.cal_nav_month,
        )
        _state.export_date, _state.cal_nav_year, _state.cal_nav_month, _ = result
        imgui.end_popup()

    imgui.spacing()
    imgui.separator()
    imgui.spacing()

    btn_w = 120
    total_btn = btn_w * 2 + 20
    imgui.set_cursor_pos_x((WINDOW_W - total_btn) * 0.5)

    if imgui.button("Xuất file", imgui.ImVec2(btn_w, 0)):
        _state.do_export()
    imgui.same_line(spacing=20)
    if imgui.button("Thoát", imgui.ImVec2(btn_w, 0)):
        hello_imgui.get_runner_params().app_shall_exit = True

    if _state.show_error_popup:
        imgui.open_popup("Lỗi##err")
        _state.show_error_popup = False
    if imgui.begin_popup_modal("Lỗi##err", None, imgui.WindowFlags_.always_auto_resize)[0]:
        imgui.text_wrapped(_state.popup_msg)
        imgui.spacing()
        if imgui.button("OK", imgui.ImVec2(80, 0)):
            imgui.close_current_popup()
        imgui.end_popup()

    if _state.show_success_popup:
        imgui.open_popup("Thành công##ok")
        _state.show_success_popup = False
    if imgui.begin_popup_modal("Thành công##ok", None, imgui.WindowFlags_.always_auto_resize)[0]:
        imgui.text_wrapped(_state.popup_msg)
        imgui.spacing()
        if imgui.button("OK", imgui.ImVec2(80, 0)):
            imgui.close_current_popup()
        imgui.end_popup()

    imgui.end()


def main():
    global _state
    _state = AppState()

    runner_params = hello_imgui.RunnerParams()
    runner_params.app_window_params.window_title = "XUẤT FILE"
    runner_params.app_window_params.window_geometry.size = (WINDOW_W + 20, WINDOW_H + 40)
    runner_params.app_window_params.resizable = False

    runner_params.imgui_window_params.default_imgui_window_type = (
        hello_imgui.DefaultImGuiWindowType.no_default_window
    )
    runner_params.imgui_window_params.show_menu_bar = False
    runner_params.imgui_window_params.show_status_bar = False

    runner_params.imgui_window_params.tweaked_theme = hello_imgui.ImGuiTweakedTheme(
        theme=hello_imgui.ImGuiTheme_.white_is_white
    )

    runner_params.callbacks.default_icon_font = hello_imgui.DefaultIconFont.font_awesome6
    runner_params.callbacks.show_gui = gui
    runner_params.fps_idling.enable_idling = True

    hello_imgui.run(runner_params)


if __name__ == "__main__":
    main()
