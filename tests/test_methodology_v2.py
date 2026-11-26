import ast
import pytest
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
st.session_state['_v3_empty_input'] = empty_sparse_template(('УБ 1',))
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


def test_review_modes_details_thermal_scenario_and_roundtrip():
    from review_v3 import HEAT_KEY
    at=app()
    b=at.session_state['blocks_df'].copy()
    b.loc[0,'S квартир, м2']=1000;b.loc[0,'Площадь участка блока, га']=2
    at.session_state['blocks_df']=b
    tep=at.session_state['tep_store']['УБ 1']
    tep.update({'Кол-во квартир, шт':10,'Дверей МОП, шт':3,HEAT_KEY:1,'Площадь холлов МОП, м2':20,'Площадь фасада, м2':100})
    at.session_state['thermal_kind']['УБ 1']='ИТП'
    at.session_state['ep_details']['УБ 1']=pd.DataFrame([
        {'Категория':'Отделка фасада','Тип / материал / размеры':'Навесной керамика','Объем':100,m.RATE:20,'Статус':'Рассчитан'},
        {'Категория':'Отделка фасада','Тип / материал / размеры':'Навесной металл','Объем':50,m.RATE:30,'Статус':'Рассчитан'},
        {'Категория':'Экраны','Объем':0,m.RATE:0,'Статус':'Не требуется'}],columns=m.DETAIL_COLS)
    at.session_state['facade_modes']['УБ 1']='Детальный'
    for code,rate in [('B.20.30',100),('B.20.30.01',50),('D.30.20',2000),('E.20.10',10),('B.20.10',999)]:set_rate(at,'block_rates',code,rate)
    g=at.session_state['block_g_area']['УБ 1'].copy();g.loc[g['Код']=='G.20.50',m.RATE]=100;at.session_state['block_g_area']['УБ 1']=g
    z=at.session_state['block_z_pct']['УБ 1'].copy();z.loc[z['Код']=='Z.10.10','Ставка, доля от СМР+G']=.1;at.session_state['block_z_pct']['УБ 1']=z
    at.session_state['scenario_select']='Стресс'
    capture=check(at)
    assert capture['item']['B.20.10']==[3500]
    assert capture['box']==[6850]
    assert capture['g']==[200]
    assert capture['base']==[7050]
    assert capture['z']==[705]
    assert capture['cost_factor']==1.15
    parsed=at.session_state['_roundtrip']
    assert parsed['thermal_kind']['УБ 1']=='ИТП'
    assert parsed['facade_modes']['УБ 1']=='Детальный'
    assert len(parsed['ep_details']['УБ 1'])==3
    assert parsed['ep_details']['УБ 1']['Тип'].tolist()[:2]==['ТИП 1','ТИП 2']
    assert parsed['ep_details']['УБ 1']['ID'].tolist()==at.session_state['ep_details']['УБ 1']['ID'].tolist()
    wb=load_workbook(BytesIO(capture['excel']),data_only=False)
    assert all(n in wb for n in ['Блоки','Бюджет по статьям','Элементы B20.10','Косвенные'])
    assert not any('ИТП' in str(c.value) or 'Котельная' in str(c.value) for row in wb['Блоки'] for c in row)
    f=wb['Элементы B20.10']
    assert f.row_dimensions[6].collapsed
    assert f.row_dimensions[7].hidden
    assert f['J6'].value=='=SUM(J7:J9)'
    assert "$E$4" in f["K6"].value
    assert any(d.hidden and d.outlineLevel==1 for d in wb['Бюджет по статьям'].column_dimensions.values())
    at.session_state['facade_modes']['УБ 1']='Укрупнённый'
    # Rendered group A does not own the B mode widget, so state is authoritative here.
    capture=check(at)
    assert capture['item']['B.20.10']==[99900]


def test_mp_geometry_rates_are_separate_from_ep_and_roundtrip():
    at=app();at.session_state['design_stage']='МП (мастер-план)';at.session_state['_stage_seg']='МП'
    b=at.session_state['blocks_df'].copy();b.loc[0,'S квартир, м2']=1000;at.session_state['blocks_df']=b
    at.session_state['mp_sections']['УБ 1']=pd.DataFrame([dict(zip(m.SECTION_COLS,['С1',2,500,10,100,600,0]))])
    for code,rate in [('A.10',10),('A.20',20),('B.10',30),('B.20',40),('B.30',50),('C.10',60),('E.20',70)]:set_rate(at,'block_mp_rates',code,rate)
    capture=check(at)
    assert capture['box']==[656140.0]
    assert capture['quantity']['E.20']==[2]
    parsed=at.session_state['_roundtrip'];assert len(parsed['mp_sections']['УБ 1'])==1
    assert parsed['block_mp_rates']['УБ 1'].set_index('Код').loc['B.10',m.RATE]==30
    assert parsed['block_rates']['УБ 1'].set_index('Код').loc['B.10.10',m.RATE]==0


def test_facade_partial_data_is_never_an_implicit_topup():
    from review_v3 import facade_cost
    rows=[{'Категория':'Отделка фасада','Объем':100,m.RATE:20,'Статус':'Рассчитан'},
          {'Категория':'Корзины кондиционеров','Объем':10,m.RATE:100,'Статус':'Нет данных'},
          {'Категория':'Экраны','Объем':10,m.RATE:999,'Статус':'Не требуется'}]
    amount,issues=facade_cost(rows,1000,8000,'Детальный');assert amount==2000 and len(issues)==1
    assert facade_cost(rows,1000,8000,'Укрупнённый')==(8000000,[])
    assert facade_cost([],1000,8000,'Детальный')==(8000000,[])
    rows[0]['Объем']=0
    assert facade_cost(rows,1000,8000,'Детальный')[0]==0


def test_detail_identity_and_type_numbers_are_stable():
    from review_v3 import normalise_details
    d=normalise_details(pd.DataFrame([{'Категория':'Отделка фасада','Объем':1,m.RATE:2},{'Категория':'Отделка фасада','Объем':2,m.RATE:3}]))
    again=normalise_details(d)
    assert d['ID'].tolist()==again['ID'].tolist()
    assert d['Тип'].tolist()==['ТИП 1','ТИП 2']
    assert again[m.RATE].tolist()==[2,3]


def test_multiple_blocks_export_separators_all_groups_and_indirect_details():
    from review_v3 import HEAT_KEY
    at=app()
    base=at.session_state['blocks_df'].iloc[0].to_dict()
    b2={**base,'Название блока':'УБ 2','S квартир, м2':500}
    park={**base,'Название блока':'П 1','Тип блока':'Наземный/Многоуровневый паркинг','S квартир, м2':0,'Наземный/Многоуровневый паркинг, м/м':10}
    base['S квартир, м2']=1000
    at.session_state['blocks_df']=pd.DataFrame([base,b2,park])
    check(at)
    for name,area,rate in [('УБ 1',100,20),('УБ 2',200,30)]:
        at.session_state['tep_store'][name].update({'Площадь фасада, м2':area,HEAT_KEY:1})
        at.session_state['ep_details'][name]=pd.DataFrame([{'Категория':'Отделка фасада','Объем':area,m.RATE:rate,'Статус':'Рассчитан'}])
        at.session_state['facade_modes'][name]='Детальный'
    at.session_state['indirect_items']['land']=pd.DataFrame([{'Наименование':'Участок 1','Сумма, руб':100},{'Наименование':'Участок 2','Сумма, руб':200}])
    at.session_state['indirect_items']['landscape']=pd.DataFrame([{'Наименование':'Озеленение общей территории','Сумма, руб':70},{'Наименование':'Покрытия общей территории','Сумма, руб':30}])
    capture=check(at)
    assert capture['item']['B.20.10']==[2000,6000,0]
    assert capture['allocation']==pytest.approx([400*2/3,400/3,0])
    wb=load_workbook(BytesIO(capture['excel']),data_only=False)
    facade=wb['Элементы B20.10']
    parents=[r[0].row for r in facade.iter_rows(min_row=6) if r[1].value=='B20.10']
    assert len(parents)==2 and all(facade.row_dimensions[r].collapsed for r in parents)
    assert all(facade.row_dimensions[r+1].hidden for r in parents)
    indirect=wb['Косвенные']
    values=[c.value for row in indirect for c in row]
    assert all(v in values for v in ['Участок 1','Участок 2','Озеленение общей территории','Покрытия общей территории'])
    sums=wb['Бюджет по статьям']
    assert len([d for d in sums.column_dimensions.values() if d.hidden and d.outlineLevel==1])==3
    assert sums['J6'].value is None
    assert not any('#REF!' in str(c.value) for ws in wb for row in ws for c in row)
    (ROOT/'test_report.xlsx').write_bytes(capture['excel'])


def test_migrate_old_heat_does_not_double_charge():
    from review_v3 import prepare,HEAT_KEY
    class State(dict):
        __getattr__=dict.__getitem__
        __setattr__=dict.__setitem__
    old=m.rates_frame([('D.30.20.01','ИТП','D','руб/компл.','x',0),('D.30.20.02','Котельная','D','руб/компл.','x',0)])
    old[m.RATE]=[100,200]
    ss=State(block_rates={'УБ 1':old},tep_store={'УБ 1':{'ИТП, комплектов на этап':1,'Котельная, комплектов на этап':1}})
    items=[('D.30.20','ИТП / котельная','D','руб/компл.',HEAT_KEY,0)]
    prepare(ss,['УБ 1'],items)
    assert ss.thermal_kind['УБ 1']=='ИТП' and ss.tep_store['УБ 1'][HEAT_KEY]==1
    assert ss.block_rates['УБ 1'].set_index('Код').loc['D.30.20',m.RATE]==100
    assert ss.legacy_warnings and 'УБ 1' in ss.legacy_rates
