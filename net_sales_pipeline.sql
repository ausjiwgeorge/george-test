-- =====================================================================
--  Net Sales 口径统一管道 —— 可直接运行的 SQLite 案例
--  场景: 7 个法人实体 (E1~E7) 各自本地科目不同, 统一映射为集团口径,
--        自动处理: 含税拆分 / 服务拆出 / 运费剔除 / 退货折让加减号 / 折算EUR
--
--  运行方式 (任选其一):
--    1) sqlite3 :memory: ".read net_sales_pipeline.sql"
--    2) sqlite3 < net_sales_pipeline.sql
--    3) 在 DB Browser for SQLite 里打开并执行
--
--  设计原则: 规则放"配置表"(config_map / config_fx), 逻辑放"管道"(view),
--            结果吐一张"长格式干净事实表"(clean_fact), KPI 用一句聚合得到。
-- =====================================================================

PRAGMA foreign_keys = OFF;
.headers on
.mode column
.width 8 8 12 30 20

DROP VIEW  IF EXISTS clean_fact;
DROP TABLE IF EXISTS staging;
DROP TABLE IF EXISTS config_map;
DROP TABLE IF EXISTS config_fx;

-- ---------------------------------------------------------------------
-- ① Staging (贴源层): 原样导出自 SAP, 不做任何加工
--    金额单位: CNY。仅 E4 生产实体按"含税"记账(见配置表 tax_flag)
-- ---------------------------------------------------------------------
CREATE TABLE staging (
    entity     TEXT,      -- 法人实体
    period     TEXT,      -- 会计期间
    local_acct TEXT,      -- 本地科目号
    local_name TEXT,      -- 本地科目名
    amount     REAL       -- 本地记账金额(原始, 正数)
);

INSERT INTO staging (entity, period, local_acct, local_name, amount) VALUES
-- E1 医疗-华东 (ex-tax 记账)
('E1','2026-06','6001.01','主营业务收入-设备', 6000000),
('E1','2026-06','6001.02','主营业务收入-耗材', 2000000),
('E1','2026-06','6051.01','其他业务收入-服务',1000000),
('E1','2026-06','6001.99','销售退回',           300000),
('E1','2026-06','6001.98','销售折让',           200000),
('E1','2026-06','6603.05','财务费用-现金折扣',  100000),
('E1','2026-06','6051.09','运费收入',           150000),
-- E2 医疗-华北 (服务混在收入里 + 运费误记)
('E2','2026-06','5001.10','商品销售收入',      9000000),
('E2','2026-06','5001.20','维保服务收入',       600000),
('E2','2026-06','5001.90','销售退回',           300000),
('E2','2026-06','5001.95','销售折让',           200000),
('E2','2026-06','5101.30','运费收入',           400000),
('E2','2026-06','6603.05','财务费用-现金折扣',   90000),
-- E3 安全-华南 (口径较干净)
('E3','2026-06','6001.00','主营业务收入',      8800000),
('E3','2026-06','6001.20','安装服务收入',       500000),
('E3','2026-06','6001.90','销售退回',           200000),
('E3','2026-06','6001.91','销售折让',           100000),
('E3','2026-06','6603.10','财务费用-现金折扣',   80000),
-- E4 生产实体 (含税记账 13% + 运费混入)
('E4','2026-06','8000',   '产成品销售(含税)',  11300000),
('E4','2026-06','8090',   '销售退货(含税)',     1130000),
('E4','2026-06','8091',   '销售折让(含税)',      226000),
('E4','2026-06','8500',   '运费收入(含税)',      113000),
('E4','2026-06','6603.05','财务费用-现金折扣',    50000),
-- E5 服务实体 (以服务收入为主)
('E5','2026-06','6001.00','产品销售(少量)',     500000),
('E5','2026-06','6001.10','服务主收入',        8800000),
('E5','2026-06','6001.99','销售退回',           100000),
('E5','2026-06','6001.98','销售折让',           100000),
('E5','2026-06','6603.05','财务费用-现金折扣',   70000),
-- E6 贸易实体
('E6','2026-06','6001.10','贸易收入',          9500000),
('E6','2026-06','6001.90','销售退回',           400000),
('E6','2026-06','6001.91','销售返利',           100000),
('E6','2026-06','6051.09','运费收入',           200000),
('E6','2026-06','6603.05','财务费用-现金折扣',   60000),
-- E7 R&D 实体 (仅少量样品销售)
('E7','2026-06','6001.90','样品销售',           300000);

-- ---------------------------------------------------------------------
-- ② 配置表 A: 科目映射表 (本地科目 -> 集团科目 + 规则)
--    sign            : '+' 计入正, '-' 计入负(退回/折让/现折)
--    incl_netsales   : 'Y' 计入净销售, 'N' 排除(运费/税)
--    tax_flag        : 1=本地含税需拆分, 0=不含税
--    tax_rate        : 增值税率
-- ---------------------------------------------------------------------
CREATE TABLE config_map (
    entity        TEXT,
    local_acct    TEXT,
    group_acct    TEXT,
    component     TEXT,
    sign          TEXT,
    incl_netsales TEXT,
    tax_flag      INTEGER,
    tax_rate      REAL
);

INSERT INTO config_map VALUES
-- E1
('E1','6001.01','G-4000','产品销售','+','Y',0,0),
('E1','6001.02','G-4000','产品销售','+','Y',0,0),
('E1','6051.01','G-4100','服务销售','+','Y',0,0),
('E1','6001.99','G-4900','销售退回','-','Y',0,0),
('E1','6001.98','G-4910','销售折让','-','Y',0,0),
('E1','6603.05','G-4920','现金折扣','-','Y',0,0),
('E1','6051.09','G-4500','运费收入','+','N',0,0),
-- E2 (维保 -> 服务; 运费 -> 排除)
('E2','5001.10','G-4000','产品销售','+','Y',0,0),
('E2','5001.20','G-4100','服务销售','+','Y',0,0),
('E2','5001.90','G-4900','销售退回','-','Y',0,0),
('E2','5001.95','G-4910','销售折让','-','Y',0,0),
('E2','5101.30','G-4500','运费收入','+','N',0,0),
('E2','6603.05','G-4920','现金折扣','-','Y',0,0),
-- E3
('E3','6001.00','G-4000','产品销售','+','Y',0,0),
('E3','6001.20','G-4100','服务销售','+','Y',0,0),
('E3','6001.90','G-4900','销售退回','-','Y',0,0),
('E3','6001.91','G-4910','销售折让','-','Y',0,0),
('E3','6603.10','G-4920','现金折扣','-','Y',0,0),
-- E4 (含税拆分 13%; 运费排除)
('E4','8000',   'G-4000','产品销售','+','Y',1,0.13),
('E4','8090',   'G-4900','销售退回','-','Y',1,0.13),
('E4','8091',   'G-4910','销售折让','-','Y',1,0.13),
('E4','8500',   'G-4500','运费收入','+','N',1,0.13),
('E4','6603.05','G-4920','现金折扣','-','Y',0,0),
-- E5
('E5','6001.00','G-4000','产品销售','+','Y',0,0),
('E5','6001.10','G-4100','服务销售','+','Y',0,0),
('E5','6001.99','G-4900','销售退回','-','Y',0,0),
('E5','6001.98','G-4910','销售折让','-','Y',0,0),
('E5','6603.05','G-4920','现金折扣','-','Y',0,0),
-- E6
('E6','6001.10','G-4000','产品销售','+','Y',0,0),
('E6','6001.90','G-4900','销售退回','-','Y',0,0),
('E6','6001.91','G-4910','销售折让','-','Y',0,0),
('E6','6051.09','G-4500','运费收入','+','N',0,0),
('E6','6603.05','G-4920','现金折扣','-','Y',0,0),
-- E7
('E7','6001.90','G-4000','产品销售','+','Y',0,0);

-- ---------------------------------------------------------------------
-- ③ 配置表 B: 汇率表 (本位币 -> 集团报告币 EUR)
-- ---------------------------------------------------------------------
CREATE TABLE config_fx (
    entity   TEXT,
    period   TEXT,
    currency TEXT,
    fx_rate  REAL      -- 1 EUR = fx_rate CNY
);

INSERT INTO config_fx VALUES
('E1','2026-06','CNY',7.8),
('E2','2026-06','CNY',7.8),
('E3','2026-06','CNY',7.8),
('E4','2026-06','CNY',7.8),
('E5','2026-06','CNY',7.8),
('E6','2026-06','CNY',7.8),
('E7','2026-06','CNY',7.8);

-- ---------------------------------------------------------------------
-- ④ 转换层 -> 干净事实表 (长格式)
--    第一段: 不含税收入/减项 归入其集团科目
--    第二段: 含税记账拆出的销项税 归入 G-2200 (incl_netsales='N')
-- ---------------------------------------------------------------------
CREATE VIEW clean_fact AS
SELECT
    s.entity,
    s.period,
    m.group_acct,
    m.component,
    m.sign,
    m.incl_netsales,
    ROUND(CASE WHEN m.tax_flag=1 THEN s.amount/(1+m.tax_rate) ELSE s.amount END, 2) AS amt_local,
    f.fx_rate,
    ROUND((CASE WHEN m.tax_flag=1 THEN s.amount/(1+m.tax_rate) ELSE s.amount END)/f.fx_rate, 2) AS amt_eur
FROM staging s
JOIN config_map m ON s.entity=m.entity AND s.local_acct=m.local_acct
JOIN config_fx  f ON s.entity=f.entity AND s.period=f.period
UNION ALL
SELECT
    s.entity,
    s.period,
    'G-2200'        AS group_acct,
    '销项税(拆出)'   AS component,
    m.sign,
    'N'             AS incl_netsales,
    ROUND(s.amount - s.amount/(1+m.tax_rate), 2)          AS amt_local,
    f.fx_rate,
    ROUND((s.amount - s.amount/(1+m.tax_rate))/f.fx_rate, 2) AS amt_eur
FROM staging s
JOIN config_map m ON s.entity=m.entity AND s.local_acct=m.local_acct
JOIN config_fx  f ON s.entity=f.entity AND s.period=f.period
WHERE m.tax_flag=1;

-- =====================================================================
--  输出 1: 干净事实表 (节选 E4, 看含税拆分效果)
-- =====================================================================
.print ''
.print '==== [1] 干净事实表 CLEAN_FACT (E4 生产实体, 含税拆分示例) ===='
SELECT group_acct, component, sign, incl_netsales, amt_local, amt_eur
FROM clean_fact WHERE entity='E4' ORDER BY group_acct;

-- =====================================================================
--  输出 2: 净销售额 (一句聚合, 通用于全部实体)
--          自动排除运费 G-4500 与 税 G-2200 (incl_netsales='N')
-- =====================================================================
.print ''
.print '==== [2] 各实体净销售额 NET SALES (统一口径) ===='
.width 8 10 16 16
SELECT
    entity,
    period,
    ROUND(SUM(amt_local * CASE sign WHEN '+' THEN 1 ELSE -1 END), 2) AS net_sales_cny,
    ROUND(SUM(amt_eur   * CASE sign WHEN '+' THEN 1 ELSE -1 END), 2) AS net_sales_eur
FROM clean_fact
WHERE incl_netsales='Y'
GROUP BY entity, period
ORDER BY entity;

.print ''
.print '==== [2b] 集团合并净销售额 ===='
SELECT
    period,
    ROUND(SUM(amt_local * CASE sign WHEN '+' THEN 1 ELSE -1 END), 2) AS group_net_sales_cny,
    ROUND(SUM(amt_eur   * CASE sign WHEN '+' THEN 1 ELSE -1 END), 2) AS group_net_sales_eur
FROM clean_fact
WHERE incl_netsales='Y'
GROUP BY period;

-- =====================================================================
--  输出 3: 校验 1 - 未映射告警 (本地科目没配到集团科目 => 报错行)
-- =====================================================================
.print ''
.print '==== [3] 校验-未映射科目 (应为空, 有则说明配置缺失) ===='
.width 8 10 12 24 14
SELECT s.entity, s.period, s.local_acct, s.local_name, s.amount
FROM staging s
LEFT JOIN config_map m ON s.entity=m.entity AND s.local_acct=m.local_acct
WHERE m.local_acct IS NULL;

-- =====================================================================
--  输出 4: 校验 2 - 总额守恒 (拆分不能凭空多/少钱)
--          Σ本地原始金额  ==  Σ干净事实表金额(含G-2200)
-- =====================================================================
.print ''
.print '==== [4] 校验-总额守恒 (diff 应全部为 0) ===='
.width 8 18 18 10
SELECT
    st.entity,
    ROUND(st.raw_total, 2)   AS staging_total,
    ROUND(cf.fact_total, 2)  AS cleanfact_total,
    ROUND(st.raw_total - cf.fact_total, 2) AS diff
FROM (SELECT entity, SUM(amount) raw_total FROM staging GROUP BY entity) st
JOIN (SELECT entity, SUM(amt_local) fact_total FROM clean_fact GROUP BY entity) cf
  ON st.entity = cf.entity
ORDER BY st.entity;

.print ''
.print '==== 完成: 脏数据进 -> 干净表出 -> KPI 一句聚合 ===='
