from io import BytesIO
import json

import pandas as pd
from openpyxl import load_workbook

import methodology_v2 as method
from test_methodology_v2 import app, check, set_rate


def edit_block_name(at, row, name):
    grid = next(frame for frame in at.dataframe if "-blocks_editor_" in frame.proto.id)
    widget_state = at._tree.get_widget_states()
    edit = widget_state.widgets.add()
    edit.id = grid.proto.id
    edit.string_value = json.dumps({
        "edited_rows": {str(row): {"Название блока": name}},
        "added_rows": [], "deleted_rows": [],
    })
    at._run(widget_state)
    assert not at.exception


def make_input(at, values):
    from sparse_excel import HEADER, FIRST_BLOCK
    empty_bytes = at.session_state["_v3_empty_input"]
    wb = load_workbook(BytesIO(empty_bytes))
    for offset, (name, fields) in enumerate(values.items()):
        col = FIRST_BLOCK + offset
        wb["ТЭП"].cell(HEADER, col, name)
        for (sheet, key), value in fields.items():
            ws = wb[sheet]
            row = next(cell.row for cell in ws["A"] if cell.value == key)
            ws.cell(row, col, value)
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def test_missing_parking_rate_flags_both_screen_and_workbook_and_links_to_field():
    at = app()
    blocks = at.session_state["blocks_df"].copy()
    blocks.loc[0, "Тип блока"] = "Наземный/Многоуровневый паркинг"
    blocks.loc[0, "Наземный/Многоуровневый паркинг, м/м"] = 100
    at.session_state["blocks_df"] = blocks
    at.session_state["_price_default_Цена наземного м/м, руб"] = 500_000
    result = check(at)
    assert any("Черновик" in x.value for x in at.warning)
    assert any("Предварительная" in x.value for x in at.caption)
    report = load_workbook(BytesIO(result["excel"]))
    assert "Черновик" in report.active["B9"].value
    assert any("Ставка СМР паркинга" in row[3] for row in report["Проверка исходных данных"].values)
    item = next(x for x in at.button if "Ставка СМР паркинга" in x.label)
    item.click().run()
    assert not at.exception
    assert at.session_state["workflow_tab"] == "3. СМР"
    assert at.session_state["smr_tab"] == "Паркинг"
    at.session_state["_smr_parking_default_Ставка СМР наземного м/м, руб"] = 700_000
    result = check(at)
    assert any("Расчёт готов" in x.value for x in at.success)
    assert load_workbook(BytesIO(result["excel"])).active["B9"].value == "Расчёт готов"


def test_rename_preserves_identity_rates_volumes_infrastructure_and_archive():
    at = app()
    original_id = at.session_state["block_ids"]["УБ 1"]
    area = "Площадь фасада, м2"
    price = "Цена жилья, руб/м2"
    at.session_state["tep_store"]["УБ 1"][area] = 2500
    at.session_state["_cascade_prices"]["УБ 1"][price] = 120_000
    at.session_state["block_infra"]["УБ 1"] = pd.DataFrame(
        [{"Статья затрат": "Парк", "Сумма, руб": 500_000}]
    )
    set_rate(at, "block_rates", "B.20.10", 10_000)
    before = check(at)
    edit_block_name(at, 0, "УБ А")
    assert at.session_state["block_ids"] == {"УБ А": original_id}
    assert at.session_state["active_block_id"] == original_id
    assert at.session_state["_cascade_prices"]["УБ А"][price] == 120_000
    assert at.session_state["tep_store"]["УБ А"][area] == 2500
    assert at.session_state["block_infra"]["УБ А"]["Сумма, руб"].sum() == 500_000
    assert at.session_state["_capture"]["box"] == before["box"]
    assert at.session_state["_roundtrip"]["block_ids"] == {"УБ А": original_id}


def test_facade_is_a_draft_until_the_new_mode_is_applied():
    at = app()
    at.session_state["tep_store"]["УБ 1"]["Площадь фасада, м2"] = 2000
    set_rate(at, "block_rates", "B.20.10", 10_000)
    assert check(at)["box"] == [20_000_000]
    at.button(key="add_facade_УБ 1").click().run()
    assert not at.exception
    assert at.session_state["facade_modes"]["УБ 1"] == "Укрупнённый"
    assert at.session_state["_capture"]["box"] == [20_000_000]
    at.session_state["v3_group_False"] = "B"
    check(at)
    assert at.button(key="apply_facade_УБ 1").proto.disabled
    at.session_state["ep_details"]["УБ 1"].loc[0, ["Объем", method.RATE, "Статус"]] = [2000, 12_000, "Рассчитан"]
    assert check(at)["box"] == [20_000_000]
    at.button(key="apply_facade_УБ 1").click().run()
    assert not at.exception
    assert at.session_state["_capture"]["box"] == [24_000_000]


def test_excel_preview_maps_two_residential_blocks_and_parking_and_applies_zeros_explicitly():
    at = app()
    book = make_input(at, {
        "УБ 1": {("ТЭП", "Общая площадь квартир (с летними, с коэф.), м2"): 3000},
        "УБ 2": {("ТЭП", "Общая площадь квартир (с летними, с коэф.), м2"): 1500,
                 ("ТЭП", "Площадь коммерции, м2"): "=100+25"},
        "Паркинг": {("ТЭП", "block_type"): "Наземный/Многоуровневый паркинг",
                    ("ТЭП", "parking_above"): 100},
    })
    at.radio(key="_entry_mode").set_value("Загрузить исходные данные").run()
    at.file_uploader[0].upload("inputs.xlsx", book).run()
    assert not at.exception
    assert at.session_state["blocks_df"]["Название блока"].tolist() == ["УБ 1"]
    assert any("формула без сохранённого результата" in x.value for x in at.markdown)
    at.button(key="apply_excel_0").click().run()
    assert not at.exception
    assert at.session_state["blocks_df"]["Название блока"].tolist() == ["УБ 1", "УБ 2", "Паркинг"]

    zero = make_input(at, {
        "Изменённое имя": {("ТЭП", "Общая площадь квартир (с летними, с коэф.), м2"): 0},
    })
    identity = at.session_state["block_ids"]["УБ 2"]
    at.file_uploader[0].upload("zero.xlsx", zero).run()
    assert at.button(key="apply_excel_1").proto.disabled
    at.checkbox(key="input_zeros_1").check().run()
    mapping = next(x for x in at.selectbox if "Изменённое имя" in x.label)
    mapping.set_value(("update", "УБ 2")).run()
    at.button(key="apply_excel_1").click().run()
    assert not at.exception
    blocks = at.session_state["blocks_df"].set_index("Название блока")
    assert blocks.loc["УБ 2", "S квартир, м2"] == 0
    assert at.session_state["block_ids"]["УБ 2"] == identity
    assert "Изменённое имя" not in blocks.index


def test_consecutive_block_edits_keep_same_grid_and_preserve_values():
    from pathlib import Path
    from streamlit.testing.v1 import AppTest

    at = AppTest.from_file(str(Path(__file__).resolve().parents[1] / "app.py")).run()
    grid = next(frame for frame in at.dataframe if "-blocks_editor_" in frame.proto.id)
    widget_id = grid.proto.id
    edits = {}
    for column, value in [
        ("S квартир, м2", "1 000,25"),
        ("S коммерции 1 эт., м2", "450"),
        ("S кладовых, м2", "51,75"),
        ("S квартир, м2", "0,00"),
    ]:
        edits[column] = value
        states = at._tree.get_widget_states()
        states.widgets.add(id=widget_id, string_value=json.dumps({
            "edited_rows": {"0": edits}, "added_rows": [], "deleted_rows": [],
        }))
        at._run(states)
        assert not at.exception, [error.message for error in at.exception]
        assert any(frame.proto.id == widget_id for frame in at.dataframe)
        area = float(value.replace(" ", "").replace(",", "."))
        assert at.session_state["blocks_df"].loc[0, column] == area

    states = at._tree.get_widget_states()
    states.widgets.add(id=widget_id, string_value=json.dumps({
        "edited_rows": {"0": edits},
        "added_rows": [{"Название блока": "УБ 2", "S квартир, м2": "125"}],
        "deleted_rows": [],
    }))
    at._run(states)
    assert not at.exception, [error.message for error in at.exception]
    blocks = at.session_state["blocks_df"]
    assert blocks["Название блока"].tolist() == ["УБ 1", "УБ 2"]
    assert blocks["S квартир, м2"].tolist() == [0, 125]
    assert blocks["S коммерции 1 эт., м2"].tolist() == [450, 0]

    grid = next(frame for frame in at.dataframe if "-blocks_editor_" in frame.proto.id)
    states = at._tree.get_widget_states()
    states.widgets.add(id=grid.proto.id, string_value=json.dumps({
        "edited_rows": {}, "added_rows": [], "deleted_rows": [0],
    }))
    at._run(states)
    assert not at.exception, [error.message for error in at.exception]
    assert at.session_state["blocks_df"]["Название блока"].tolist() == ["УБ 2"]
    assert at.session_state["blocks_df"]["S квартир, м2"].tolist() == [125]
