"""
对账管理系统（只录入对账表，不处理整张单）

功能：
  1. 导入对账表（本地脚本生成的「对账结果表.xlsx」）
  2. 明细查询
  3. 统计汇总
  4. 价格对比
  5. 修改单价（备注记录原价，重算批次总计）
  6. 历史修改记录

对账表格式：日期 / 型号 / 单价 / 数量 / 金额 / 备注
  - 标题形如「三和电子行 对账表」，公司名从标题提取
  - 金额一律按 单价 × 数量 重算，不信文件里的金额
  - 跳过「日合计」「总计」行
  - 日期为空的行继承上一行日期
  - 去重：公司 + 日期 + 型号 + 数量 + 单价 完全一致则跳过
"""
import io
import re
import time
from pathlib import Path

import pandas as pd
import pymysql
import openpyxl
import streamlit as st

st.set_page_config(page_title="对账管理系统", layout="wide")


# ==================== 配置 ====================
DB_HOST = st.secrets["db_host"]
DB_PORT = int(st.secrets["db_port"])
DB_USER = st.secrets["db_user"]
DB_PASSWORD = st.secrets["db_password"]
DB_NAME = st.secrets["db_name"]

SSL_CA_PATH = Path(__file__).parent / "ca.pem"


# ==================== 数据库连接 ====================
def get_conn(use_db=True):
    kwargs = dict(
        host=DB_HOST,
        port=DB_PORT,
        user=DB_USER,
        password=DB_PASSWORD,
        charset='utf8mb4',
        cursorclass=pymysql.cursors.DictCursor,
        connect_timeout=15,
        read_timeout=60,
        write_timeout=60,
    )
    if use_db:
        kwargs["database"] = DB_NAME
    if not SSL_CA_PATH.exists():
        raise Exception("未找到 ca.pem 证书，请从 TiDB Cloud 下载并放到项目根目录")
    kwargs["ssl"] = {"ca": str(SSL_CA_PATH)}
    return pymysql.connect(**kwargs)


def run_with_retry(func, retries=3):
    last = None
    for i in range(retries):
        try:
            return func()
        except pymysql.err.OperationalError as e:
            last = e
            code = e.args[0] if e.args else None
            if code in (2013, 2006, 2003):
                time.sleep(1.5 * (i + 1))
                continue
            raise
    raise last


def query_df(sql, params=None):
    def _run():
        conn = get_conn()
        try:
            with conn.cursor() as cur:
                cur.execute(sql, params or ())
                return pd.DataFrame(cur.fetchall())
        finally:
            conn.close()
    return run_with_retry(_run)


# ==================== 初始化表 ====================
def init_db():
    conn = get_conn(use_db=False)
    try:
        with conn.cursor() as cur:
            cur.execute(
                f"CREATE DATABASE IF NOT EXISTS `{DB_NAME}` "
                "CHARACTER SET utf8mb4 COLLATE utf8mb4_bin"
            )
            cur.execute(f"""
                CREATE TABLE IF NOT EXISTS `{DB_NAME}`.`dz_batch` (
                    batch_id BIGINT PRIMARY KEY AUTO_INCREMENT,
                    company VARCHAR(255) NOT NULL,
                    file_name VARCHAR(255) NOT NULL,
                    date_range VARCHAR(60) DEFAULT '',
                    total_amount DECIMAL(14,4) DEFAULT 0,
                    imported_at DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """)
            cur.execute(f"""
                CREATE TABLE IF NOT EXISTS `{DB_NAME}`.`dz_detail` (
                    id BIGINT PRIMARY KEY AUTO_INCREMENT,
                    batch_id BIGINT NOT NULL,
                    company VARCHAR(255) NOT NULL,
                    item_date VARCHAR(32) NOT NULL,
                    model VARCHAR(255) NOT NULL,
                    quantity DECIMAL(14,4),
                    unit_price DECIMAL(14,4),
                    amount DECIMAL(14,4),
                    remark VARCHAR(1000) DEFAULT '',
                    KEY idx_company_date (company, item_date),
                    KEY idx_model (model),
                    KEY idx_dedup (company, item_date, model, quantity, unit_price)
                )
            """)
            cur.execute(f"""
                CREATE TABLE IF NOT EXISTS `{DB_NAME}`.`dz_price_history` (
                    id BIGINT PRIMARY KEY AUTO_INCREMENT,
                    detail_id BIGINT NOT NULL,
                    company VARCHAR(255),
                    item_date VARCHAR(32),
                    model VARCHAR(255),
                    old_unit_price DECIMAL(14,4),
                    new_unit_price DECIMAL(14,4),
                    old_amount DECIMAL(14,4),
                    new_amount DECIMAL(14,4),
                    remark VARCHAR(500),
                    change_date DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """)
        conn.commit()
    finally:
        conn.close()


@st.cache_resource(show_spinner=False)
def ensure_db():
    run_with_retry(init_db)
    return True


# ==================== 工具函数 ====================
def extract_value_from_cell(text, keyword):
    if not isinstance(text, str):
        return ''
    m = re.search(r'{}(?:[：:.])\s*(.+)'.format(re.escape(keyword)), text.strip())
    return m.group(1).strip() if m else ''


def normalize_date(v):
    if v is None:
        return ''
    if hasattr(v, 'strftime'):
        return v.strftime('%Y-%m-%d')
    s = str(v).strip()
    s = s.replace('年', '-').replace('月', '-').replace('日', '')
    s = s.replace('/', '-').replace('.', '-')
    s = re.sub(r'\s+', '', s)
    m = re.search(r'(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        y, mo, d = m.groups()
        return f"{y}-{int(mo):02d}-{int(d):02d}"
    return s


def safe_float(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        v = v.strip().replace(',', '')
        try:
            return float(v)
        except Exception:
            return None
    return None


# ==================== 解析对账表 ====================
def parse_duizhang(file_bytes, file_name):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    rows = [[c.value for c in row] for row in ws.iter_rows(max_row=20000, max_col=20)]
    wb.close()

    if not rows:
        return {'company': Path(file_name).stem, 'date_range': '',
                'records': [], 'errors': [{'错误': '空文件'}]}

    # 公司名：标题「三和电子行 对账表」
    title = str(rows[0][0] or '').strip() if rows else ''
    company = title.replace('对账表', '').strip() or Path(file_name).stem

    # 统计期间
    date_range = ''
    for row in rows[:5]:
        for val in row:
            if isinstance(val, str) and '统计期间' in val:
                v = extract_value_from_cell(val, '统计期间')
                if v:
                    date_range = v.strip()
                    break

    # 表头
    header_idx = None
    for i, row in enumerate(rows[:30]):
        joined = ' '.join(str(v or '').strip() for v in row)
        has_model = ('型号' in joined) or ('品名' in joined) or ('货名' in joined)
        if '日期' in joined and has_model and '单价' in joined and '数量' in joined:
            header_idx = i
            break
    if header_idx is None:
        return {'company': company, 'date_range': date_range, 'records': [],
                'errors': [{'错误': '未找到表头（需含 日期/型号/单价/数量）'}]}

    col_map = {}
    for idx, cell in enumerate(rows[header_idx]):
        t = str(cell or '').strip()
        if '日期' in t:
            col_map['date'] = idx
        if '型号' in t or '品名' in t or '货名' in t:
            col_map.setdefault('model', idx)
        if '单价' in t:
            col_map['unit_price'] = idx
        if '数量' in t:
            col_map['quantity'] = idx
        if '备注' in t:
            col_map['remark'] = idx

    if not all(k in col_map for k in ('model', 'quantity', 'unit_price')):
        return {'company': company, 'date_range': date_range, 'records': [],
                'errors': [{'错误': '表头缺少 型号/数量/单价'}]}

    records = []
    errors = []
    current_date = ''

    for r in range(header_idx + 1, len(rows)):
        row = rows[r]
        texts = [str(v or '').strip() for v in row]
        if all(t == '' for t in texts):
            continue
        if any(t in ('日合计', '总计', '合计', '小计') for t in texts):
            continue

        def get_val(field):
            idx = col_map.get(field)
            return row[idx] if (idx is not None and idx < len(row)) else ''

        dv = get_val('date')
        if dv is not None and str(dv).strip():
            nd = normalize_date(dv)
            if nd:
                current_date = nd

        model = str(get_val('model') or '').strip()
        if not model:
            continue

        qty = safe_float(get_val('quantity'))
        price = safe_float(get_val('unit_price'))
        if qty is None or price is None:
            errors.append({'行号': r + 1, '型号': model, '错误': '数量或单价为空'})
            continue

        amount = round(qty * price, 4)
        remark = str(get_val('remark') or '').strip() if 'remark' in col_map else ''

        records.append({
            '日期': current_date,
            '型号': model,
            '数量': qty,
            '单价': price,
            '金额': amount,
            '备注': remark,
        })

    if not date_range and records:
        dates = [x['日期'] for x in records if x['日期']]
        if dates:
            date_range = f"{min(dates)} 至 {max(dates)}"

    return {'company': company, 'date_range': date_range, 'records': records, 'errors': errors}


# ==================== 入库 ====================
def import_to_db(company, file_name, date_range, records):
    def _run():
        conn = get_conn()
        try:
            cursor = conn.cursor()

            # 先去重，找出真正要新增的记录
            new_records = []
            skip_count = 0
            for r in records:
                cursor.execute(
                    "SELECT id FROM dz_detail "
                    "WHERE company=%s AND item_date=%s AND model=%s "
                    "AND quantity=%s AND unit_price=%s LIMIT 1",
                    (company, r['日期'], r['型号'], r['数量'], r['单价'])
                )
                if cursor.fetchone():
                    skip_count += 1
                else:
                    new_records.append(r)

            if not new_records:
                conn.rollback()
                return 0, skip_count

            total = sum(r['金额'] for r in new_records if r['金额'] is not None)
            cursor.execute(
                "INSERT INTO dz_batch (company, file_name, date_range, total_amount) "
                "VALUES (%s,%s,%s,%s)",
                (company, file_name, date_range, round(total, 4))
            )
            batch_id = cursor.lastrowid

            for r in new_records:
                cursor.execute(
                    "INSERT INTO dz_detail "
                    "(batch_id, company, item_date, model, quantity, unit_price, amount, remark) "
                    "VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
                    (batch_id, company, r['日期'], r['型号'], r['数量'], r['单价'], r['金额'], r['备注'])
                )

            conn.commit()
            return len(new_records), skip_count
        finally:
            conn.close()

    return run_with_retry(_run)


# ==================== 查询辅助 ====================
def get_companies(kw=''):
    if kw:
        df = query_df(
            "SELECT DISTINCT company FROM dz_detail WHERE company LIKE %s ORDER BY company",
            (f'%{kw}%',)
        )
    else:
        df = query_df("SELECT DISTINCT company FROM dz_detail ORDER BY company")
    return df['company'].tolist() if not df.empty else []


def get_models(kw=''):
    if kw:
        df = query_df(
            "SELECT DISTINCT model FROM dz_detail WHERE model LIKE %s ORDER BY model",
            (f'%{kw}%',)
        )
    else:
        df = query_df("SELECT DISTINCT model FROM dz_detail ORDER BY model")
    return df['model'].tolist() if not df.empty else []


def get_history(model=None, start_time=None, end_time=None):
    where = []
    params = []
    if model:
        where.append("model LIKE %s")
        params.append(f'%{model}%')
    if start_time:
        where.append("change_date >= %s")
        params.append(start_time)
    if end_time:
        where.append("change_date <= %s")
        params.append(end_time)

    sql = """
        SELECT company AS 公司, item_date AS 日期, model AS 型号,
               old_unit_price AS 修改前单价, new_unit_price AS 修改后单价,
               old_amount AS 修改前金额, new_amount AS 修改后金额,
               remark AS 备注, change_date AS 修改时间
        FROM dz_price_history
    """
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY change_date DESC LIMIT 500"
    return query_df(sql, params)


_PRICE_NOTE_RE = re.compile(r'^原价[：:]\s*[0-9.]+\s*$')


def clean_remark_history(remark):
    """清掉备注里历史遗留的「原价：x」片段，只保留用户自己写的备注。

    旧写法每次改价都往备注追加一段「原价：x」，改价次数一多备注就会无限
    变长（dz_detail.remark 是 VARCHAR(1000)，约 90 次后写库会失败）。
    改价历史本来就完整记录在 dz_price_history 表里，这里只保留最新一次原价。
    """
    if not remark:
        return ''
    parts = [p.strip() for p in str(remark).split('；') if p.strip()]
    kept = [p for p in parts if not _PRICE_NOTE_RE.match(p)]
    return '；'.join(kept)


def do_change_price(detail_id, new_price, remark):
    def _run():
        conn = get_conn()
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM dz_detail WHERE id=%s", (detail_id,))
            row = cursor.fetchone()
            if not row:
                return False, "明细不存在"

            old_price = float(row['unit_price'])
            qty = float(row['quantity'])
            old_amount = float(row['amount'])
            new_amount = round(qty * new_price, 4)

            # 备注只保留最新一次原价，历史「原价：x」片段先清掉，避免无限变长
            kept_notes = clean_remark_history(row['remark'])
            new_remark = f"原价：{old_price}"
            if kept_notes:
                new_remark += '；' + kept_notes
            if remark:
                new_remark += '；' + remark.strip()
            # 兜底：remark 列是 VARCHAR(1000)，截断防止超长导致写库失败
            if len(new_remark) > 900:
                new_remark = new_remark[:900]

            cursor.execute(
                "UPDATE dz_detail SET unit_price=%s, amount=%s, remark=%s WHERE id=%s",
                (new_price, new_amount, new_remark, detail_id)
            )
            cursor.execute(
                "INSERT INTO dz_price_history "
                "(detail_id, company, item_date, model, old_unit_price, new_unit_price, "
                "old_amount, new_amount, remark) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                (detail_id, row['company'], row['item_date'], row['model'],
                 old_price, new_price, old_amount, new_amount, remark)
            )

            cursor.execute(
                "SELECT SUM(amount) AS t FROM dz_detail WHERE batch_id=%s",
                (row['batch_id'],)
            )
            total = float(cursor.fetchone()['t'] or 0)
            cursor.execute(
                "UPDATE dz_batch SET total_amount=%s WHERE batch_id=%s",
                (round(total, 4), row['batch_id'])
            )

            conn.commit()
            return True, (f"已修改：{row['model']} 单价 {old_price} → {new_price}，"
                          f"新金额 {new_amount}，本批总计重算为 {round(total, 4)}")
        finally:
            conn.close()

    return run_with_retry(_run)


# ==================== 页面1：导入对账表 ====================
def page_import():
    st.subheader("导入对账表")
    st.write("上传对账表 Excel，自动解析入库；重复记录（公司+日期+型号+数量+单价完全一致）自动跳过。")

    uploaded = st.file_uploader("选择对账表文件", type=['xlsx'], accept_multiple_files=True)

    if uploaded:
        if st.button("解析预览"):
            previews = []
            for u in uploaded:
                try:
                    previews.append({'file_name': u.name, **parse_duizhang(u.getvalue(), u.name)})
                except Exception as e:
                    previews.append({'file_name': u.name, 'company': '', 'date_range': '',
                                     'records': [], 'errors': [{'错误': str(e)}]})
            st.session_state.import_previews = previews

        if st.session_state.get('import_previews'):
            previews = st.session_state.import_previews
            st.write("### 解析预览")
            summary = pd.DataFrame([{
                '文件': p['file_name'],
                '公司': p['company'],
                '统计期间': p['date_range'],
                '条数': len(p['records']),
                '错误数': len(p['errors']),
                '金额合计': round(sum(r['金额'] for r in p['records']), 2),
            } for p in previews])
            st.dataframe(summary, width='stretch', hide_index=True)

            for p in previews:
                if p['records']:
                    with st.expander(f"{p['file_name']} — 前 20 行明细"):
                        st.dataframe(pd.DataFrame(p['records'][:20]), width='stretch', hide_index=True)
                if p['errors']:
                    with st.expander(f"{p['file_name']} — 错误（{len(p['errors'])} 条）"):
                        st.dataframe(pd.DataFrame(p['errors']), width='stretch', hide_index=True)

            if st.button("确认导入"):
                results = []
                for p in previews:
                    if not p['records']:
                        results.append((p['file_name'], 0, 0, '无数据或解析失败，跳过'))
                        continue
                    try:
                        new_c, skip_c = import_to_db(p['company'], p['file_name'], p['date_range'], p['records'])
                        results.append((p['file_name'], new_c, skip_c,
                                        f"新增 {new_c} 条，跳过 {skip_c} 条（重复）"))
                    except Exception as e:
                        results.append((p['file_name'], 0, 0, f"导入失败：{e}"))
                st.session_state.import_results = results
                st.session_state.import_previews = None
                st.rerun()

        if st.session_state.get('import_results'):
            st.write("### 导入结果")
            res_df = pd.DataFrame(st.session_state.import_results, columns=['文件', '新增', '跳过', '说明'])
            st.dataframe(res_df, width='stretch', hide_index=True)


# ==================== 页面2：明细查询 ====================
def page_detail_query():
    st.subheader("明细查询")

    use_company = st.checkbox("按公司查询")
    selected_company = None
    if use_company:
        kw = st.text_input("公司关键词", key="dq_company_kw")
        if kw:
            companies = get_companies(kw)
            if companies:
                selected_company = st.selectbox("选择公司", companies)
            else:
                st.warning("未找到匹配公司")

    use_date = st.checkbox("按日期查询")
    start = end = None
    if use_date:
        c1, c2 = st.columns(2)
        start = normalize_date(c1.text_input("开始日期", placeholder="2026.08.01", key="dq_start")) or None
        end = normalize_date(c2.text_input("结束日期", placeholder="2026.08.31（可空）", key="dq_end")) or None

    use_model = st.checkbox("按型号查询")
    selected_model = None
    if use_model:
        kw = st.text_input("型号关键词", key="dq_model_kw")
        if kw:
            models = get_models(kw)
            if models:
                st.info(f"共找到 {len(models)} 个匹配型号，下面全部列出")
                selected_model = st.radio("选择型号", models)
            else:
                st.warning("未找到匹配型号")

    if st.button("查询"):
        where = []
        params = []
        if selected_company:
            where.append("company=%s")
            params.append(selected_company)
        if start:
            where.append("item_date>=%s")
            params.append(start)
        if end:
            where.append("item_date<=%s")
            params.append(end)
        if selected_model:
            where.append("model=%s")
            params.append(selected_model)

        if not where:
            st.warning("请至少选择一项查询条件")
            return

        sql = """
            SELECT company AS 公司, item_date AS 日期, model AS 型号,
                   quantity AS 数量, unit_price AS 单价, amount AS 金额, remark AS 备注
            FROM dz_detail
        """
        sql += " WHERE " + " AND ".join(where)
        sql += " ORDER BY item_date DESC, id DESC LIMIT 500"

        df = query_df(sql, params)
        if df.empty:
            st.warning("未找到匹配明细")
        else:
            st.dataframe(df, width='stretch', hide_index=True)


# ==================== 页面3：统计汇总 ====================
def page_stats():
    st.subheader("统计汇总")

    c1, c2 = st.columns(2)
    start = normalize_date(c1.text_input("开始日期", placeholder="2026.08.01", key="st_start")) or None
    end = normalize_date(c2.text_input("结束日期", placeholder="2026.08.31（可空）", key="st_end")) or None

    stat_type = st.radio("汇总方式", ["按月汇总", "按公司汇总", "按型号汇总"])

    selected_model = None
    if stat_type == "按型号汇总":
        kw = st.text_input("型号关键词", key="st_model_kw")
        if kw:
            models = get_models(kw)
            if models:
                st.info(f"共找到 {len(models)} 个匹配型号，下面全部列出")
                selected_model = st.radio("选择型号", models)
            else:
                st.warning("未找到匹配型号")

    if st.button("开始统计"):
        where = ["1=1"]
        params = []
        if start:
            where.append("item_date>=%s")
            params.append(start)
        if end:
            where.append("item_date<=%s")
            params.append(end)
        if stat_type == "按型号汇总":
            if not selected_model:
                st.warning("请先选择型号")
                return
            where.append("model=%s")
            params.append(selected_model)

        w = " WHERE " + " AND ".join(where)

        if stat_type == "按月汇总":
            sql = (f"SELECT SUBSTRING(item_date,1,7) AS 月份, SUM(amount) AS 金额合计, "
                   f"SUM(quantity) AS 数量合计, COUNT(*) AS 明细条数 FROM dz_detail{w} "
                   f"GROUP BY SUBSTRING(item_date,1,7) ORDER BY 月份")
        elif stat_type == "按公司汇总":
            sql = (f"SELECT company AS 公司, SUM(amount) AS 金额合计, "
                   f"SUM(quantity) AS 数量合计, COUNT(*) AS 明细条数 FROM dz_detail{w} "
                   f"GROUP BY company ORDER BY 金额合计 DESC")
        else:
            sql = (f"SELECT model AS 型号, SUM(amount) AS 金额合计, SUM(quantity) AS 数量合计, "
                   f"COUNT(*) AS 明细条数 FROM dz_detail{w} "
                   f"GROUP BY model ORDER BY 金额合计 DESC")

        df = query_df(sql, params)
        st.dataframe(df, width='stretch', hide_index=True)


# ==================== 页面4：价格对比 ====================
def page_price_compare():
    st.subheader("价格对比")
    kw = st.text_input("型号关键词", key="pc_model_kw")
    if kw:
        models = get_models(kw)
        if models:
            st.info(f"共找到 {len(models)} 个匹配型号，下面全部列出")
            selected = st.radio("选择型号", models)
            df = query_df(
                "SELECT company AS 公司, item_date AS 日期, quantity AS 数量, "
                "unit_price AS 单价, amount AS 金额, remark AS 备注 "
                "FROM dz_detail WHERE model=%s ORDER BY item_date DESC, id DESC",
                (selected,)
            )
            if df.empty:
                st.warning("该型号无记录")
            else:
                st.dataframe(df, width='stretch', hide_index=True)
        else:
            st.warning("未找到匹配型号")


# ==================== 页面5：修改单价 ====================
def page_modify_price():
    st.subheader("修改单价")
    kw = st.text_input("型号关键词", key="mp_model_kw")
    if not kw:
        st.info("请输入型号关键词搜索")
        return

    df = query_df(
        "SELECT id, company, item_date, model, quantity, unit_price, amount, remark "
        "FROM dz_detail WHERE model LIKE %s ORDER BY item_date DESC, id DESC LIMIT 200",
        (f'%{kw}%',)
    )
    if df.empty:
        st.warning("未找到匹配明细")
        return

    models = df['model'].unique().tolist()
    selected_model = st.selectbox("选择型号", models)
    df2 = df[df['model'] == selected_model].reset_index(drop=True)
    if df2.empty:
        st.warning("该型号无记录")
        return

    option_idx = st.selectbox(
        "选择要修改的明细",
        range(len(df2)),
        format_func=lambda i: (f"{df2.iloc[i]['item_date']} | {df2.iloc[i]['company']} | "
                               f"单价 {df2.iloc[i]['unit_price']} | 数量 {df2.iloc[i]['quantity']}")
    )
    row = df2.iloc[option_idx]
    selected_id = int(row['id'])

    st.markdown(
        f"**当前：** 日期 {row['item_date']}｜公司 {row['company']}｜型号 {row['model']}｜"
        f"单价 {row['unit_price']}｜数量 {row['quantity']}｜金额 {row['amount']}"
    )

    hist = query_df(
        "SELECT old_unit_price AS 修改前单价, new_unit_price AS 修改后单价, "
        "old_amount AS 修改前金额, new_amount AS 修改后金额, remark AS 备注, "
        "change_date AS 修改时间 FROM dz_price_history WHERE detail_id=%s "
        "ORDER BY change_date DESC",
        (selected_id,)
    )
    if not hist.empty:
        st.write("该行历史修改记录：")
        st.dataframe(hist, width='stretch', hide_index=True)

    new_price = st.text_input("新单价")
    remark = st.text_input("备注")

    if st.button("确认修改"):
        if not new_price:
            st.warning("请输入新单价")
            return
        try:
            new_price_f = float(new_price)
        except Exception:
            st.error("单价格式错误")
            return

        with st.spinner("正在重新计算并保存..."):
            ok, msg = do_change_price(selected_id, new_price_f, remark.strip())
        if ok:
            st.success(msg)
        else:
            st.error(msg)


# ==================== 页面6：历史修改记录 ====================
def page_history():
    st.subheader("历史修改记录")
    t = st.radio("查询方式", ["查看所有记录", "按型号查", "按时间查"])

    if t == "查看所有记录":
        df = get_history()
        st.dataframe(df, width='stretch', hide_index=True)

    elif t == "按型号查":
        model = st.text_input("型号关键词", key="his_model")
        if model:
            df = get_history(model=model)
            st.dataframe(df, width='stretch', hide_index=True)

    else:
        c1, c2 = st.columns(2)
        start_time = c1.text_input("开始时间", placeholder="2026-09-01", key="his_start")
        end_time = c2.text_input("结束时间", placeholder="2026-09-15", key="his_end")
        if start_time or end_time:
            df = get_history(start_time=start_time or None, end_time=end_time or None)
            st.dataframe(df, width='stretch', hide_index=True)


# ==================== 主入口 ====================
try:
    ensure_db()
except Exception as e:
    st.error(f"数据库连接失败，请检查 .streamlit/secrets.toml 配置：{e}")
    st.stop()

st.sidebar.write(f"当前连接：{DB_HOST}:{DB_PORT}/{DB_NAME}")
st.sidebar.title("对账管理系统")

menu = st.sidebar.radio(
    "功能菜单",
    ["导入对账表", "明细查询", "统计汇总", "价格对比", "修改单价", "历史修改记录"]
)

if menu == "导入对账表":
    page_import()
elif menu == "明细查询":
    page_detail_query()
elif menu == "统计汇总":
    page_stats()
elif menu == "价格对比":
    page_price_compare()
elif menu == "修改单价":
    page_modify_price()
elif menu == "历史修改记录":
    page_history()