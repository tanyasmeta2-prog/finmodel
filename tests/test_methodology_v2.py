import ast
from io import BytesIO
from pathlib import Path
import subprocess
import pandas as pd
from openpyxl import load_workbook
from streamlit.testing.v1 import AppTest
import methodology_v2 as m

ROOT = Path(__file__).resolve().parents[1]


def test_geometry_does_not_use_commercial_area_and_excludes_shared_facades():
    rows = [dict(zip(m.SECTION_COLS, ['С1',2,500,10,100,600,0])),
            dict(zip(m.SECTION_COLS, ['Стилобат',1,200,1,60,0,1]))]
    result, issues = m.geometry(rows)
    assert not issues
    assert result == dict(MP_FOOTPRINT=1200, MP_UNDERGROUND=1000, MP_ABOVE=10200, MP_FACADE=5580, MP_SECTIONS=2)
    taller, _ = m.geometry([dict(zip(m.SECTION_COLS,['С1',2,500,11,100,600,0])),rows[1]])
    assert taller['MP_FOOTPRINT'] == result['MP_FOOTPRINT']
    assert taller['MP_ABOVE'] - result['MP_ABOVE'] == 1000
    assert taller['MP_FACADE'] - result['MP_FACADE'] == 600


def test_details_sum_actual_rates_not_simple_average():
    rows=[{'Категория':'Отделка фасада','Объем':100,'Ставка, руб/ед.':10},
          {'Категория':'Отделка фасада','Объем':300,'Ставка, руб/ед.':20},
          {'Категория':'Корзины кондиционеров','Объем':2,'Ставка, руб/ед.':500}]
    costs, qty, issues=m.detail_totals(rows)
    assert not issues
    assert costs['B.20.10.02'] == 7000
    assert qty['B.20.10.02'] == 400
    assert costs['B.20.10.04'] == 1000


def test_invalid_inputs_are_explicit_and_entry_fallback_is_checked():
    _, issues=m.geometry([dict(zip(m.SECTION_COLS,['bad',1,20,1,10,100,0]))])
    assert any('торцов' in s for s in issues)
    _,_,issues=m.detail_totals([{'Категория':'Корзины кондиционеров','Объем':1.5,'Ставка, руб/ед.':1},
                              {'Категория':'Входные группы по площади','Объем':20,'Ставка, руб/ед.':0},
                              {'Категория':'Входные группы без площади','Объем':1,'Ставка, руб/ед.':10}])
    assert any('целыми' in s for s in issues)
    assert any('ставка' in s for s in issues)
    assert any('разные группы' in s for s in issues)


def test_excluded_g_parking_and_infrastructure_code_is_unchanged():
    old=subprocess.check_output(['git','show','d0dd64a2d94a1d3154043bb69332bf3f70ea8166:app.py'],cwd=ROOT,text=True)
    new=(ROOT/'app.py').read_text()
    # The original G, Z and infrastructure numeric calculation blocks are retained verbatim.
    for start,end in [('plot_area_col =','blocks["Прочие затраты блока (Z), руб"] = z_block_total_scaled'),
                      ('underground_cost_raw =','parking_cost_block = np.where(is_res, underground_cost_raw, park_block_smr_raw)'),
                      ('infra_total = np.array([','blocks["Инфраструктура блока (Z5010), руб"] = infra_total * cost_factor')]:
        assert old[old.index(start):old.index(end)+len(end)] == new[new.index(start):new.index(end)+len(end)]
    trees=[ast.parse(src) for src in [old,new]]
    for key in ['ITEMS_G_AREA','ITEMS_G_LENGTH','ITEMS_Z_PCT','ITEMS_Z_FIXED']:
        vals=[]
        for tree in trees:
            n=next(n for n in tree.body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id==key for t in n.targets))
            vals.append(ast.literal_eval(n.value))
        assert vals[0]==vals[1]


CAPTURE='''
st.session_state['_capture'] = dict(excel=excel_bytes, template=build_template_bytes(collect_template_data()),
    box=smr_korobka_raw.tolist(), item={k:v.tolist() for k,v in item_cost_matrix.items()},
    g=g_block_total.tolist(), z=z_block_total.tolist(), base=zg_base.tolist(),
    quantity={k:v.tolist() for k,v in qty_by_basis.items()}, cost_factor=cost_factor,
    allocation=blocks['Аллоцированные затраты'].tolist())
_parsed,_warns = parse_template(st.session_state['_capture']['template'])
st.session_state['_roundtrip'] = _parsed
st.session_state['_roundtrip_warnings'] = _warns
'''


def app():
    at=AppTest.from_string((ROOT/'app.py').read_text()+CAPTURE,default_timeout=60).run()
    assert not at.exception, [e.message for e in at.exception]
    return at


def set_rate(at,store,code,value):
    df=at.session_state[store]['УБ 1'].copy()
    df.loc[df['Код']==code,m.RATE]=float(value)
    df.loc[df['Код']==code,'Аналог']='Тестовый аналог'
    at.session_state[store]['УБ 1']=df


def check(at):
    at.run()
    assert not at.exception, [(e.message,e.stack_trace) for e in at.exception]
    return at.session_state['_capture']


def test_ep_doors_itp_details_scenario_and_excel_roundtrip():
    at=app()
    tep=at.session_state['tep_store']['УБ 1']
    tep.update({'Кол-во квартир, шт':10,'Дверей МОП, шт':3,'ИТП, комплектов на этап':1,
                'Котельная, комплектов на этап':1,'Площадь холлов МОП, м2':20,
                'Общая площадь квартир (с летними, с коэф.), м2':1000,'Площадь участка блока, га':2})
    at.session_state['ep_details']['УБ 1']=pd.DataFrame([
        {'Категория':'Кладка наружных стен','Тип / материал / размеры':'мокрый','Объем':100,m.RATE:10},
        {'Категория':'Отделка фасада','Тип / материал / размеры':'мокрый','Объем':100,m.RATE:20},
        {'Категория':'Входные группы без площади','Объем':2,m.RATE:100},
        {'Категория':'Террасы верхних этажей','Объем':10,m.RATE:5}],columns=m.DETAIL_COLS)
    for code, rate in [('B.20.30',100),('B.20.30.01',50),('D.30.20.01',2000),('D.30.20.02',4000),('E.20.10',10)]:
        set_rate(at,'block_rates',code,rate)
    g=at.session_state['block_g_area']['УБ 1'].copy();g.loc[g['Код']=='G.20.50',m.RATE]=100
    at.session_state['block_g_area']['УБ 1']=g
    z=at.session_state['block_z_pct']['УБ 1'].copy();z.loc[z['Код']=='Z.10.10','Ставка, доля от СМР+G']=0.1
    at.session_state['block_z_pct']['УБ 1']=z
    at.session_state['scenario_select']='Стресс'
    capture=check(at)
    assert capture['box']==[10600.0]
    assert capture['g']==[200]
    assert capture['base']==[10800]
    assert capture['z']==[1080]
    assert capture['cost_factor']==1.15
    parsed=at.session_state['_roundtrip']
    assert parsed['tep_store']['УБ 1']['Дверей МОП, шт']==3
    assert parsed['block_rates']['УБ 1'].set_index('Код').loc['D.30.20.01',m.RATE]==2000
    assert len(parsed['ep_details']['УБ 1'])==4
    wb=load_workbook(BytesIO(capture['excel']),data_only=False)
    assert 'Косвенные — состав' in wb
    assert 'G и Z — суммы статей' in wb
    assert not any('#REF!' in str(c.value) for ws in wb for row in ws for c in row)
    at.session_state['thermal_included']['УБ 1']=True
    # Reset widget value to simulate a changed checkbox before rendering.
    at.session_state['thermal_included_widget_УБ 1']=True
    capture=check(at)
    assert capture['box']==[4600.0]


def test_mp_geometry_rates_are_separate_from_ep_and_roundtrip():
    at=app()
    at.session_state['design_stage']='МП (мастер-план)'
    at.session_state['_stage_seg']='МП'
    at.session_state['tep_store']['УБ 1']['Общая площадь квартир (с летними, с коэф.), м2']=1000
    at.session_state['mp_sections']['УБ 1']=pd.DataFrame([dict(zip(m.SECTION_COLS,['С1',2,500,10,100,600,0]))])
    for code,rate in [('A.10',10),('A.20',20),('B.10',30),('B.20',40),('B.30',50),('C.10',60),('E.20',70)]:
        set_rate(at,'block_mp_rates',code,rate)
    capture=check(at)
    assert capture['box']==[656140.0]
    assert capture['quantity']['E.20']==[2]
    parsed=at.session_state['_roundtrip']
    assert len(parsed['mp_sections']['УБ 1'])==1
    assert parsed['block_mp_rates']['УБ 1'].set_index('Код').loc['B.10',m.RATE]==30
    assert parsed['block_rates']['УБ 1'].set_index('Код').loc['B.10.10',m.RATE]==0
    wb=load_workbook(BytesIO(capture['excel']),data_only=False)
    assert any('SUMIF' in str(c.value) for row in wb['СМР по методике'] for c in row)
    assert not any('#REF!' in str(c.value) for ws in wb for row in ws for c in row)


def test_no_work_status_keeps_quoted_rate_but_excludes_charge():
    f=m.rates_frame(m.MP_ITEMS)
    f.loc[f['Код']=='A.10',m.RATE]=100
    f.loc[f['Код']=='A.10','Применимость']='Не требуется'
    assert f.set_index('Код').loc['A.10',m.RATE]==100
    assert m.effective_rates(f)['A.10']==0
