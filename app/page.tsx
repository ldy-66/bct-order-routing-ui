"use client";

import { Fragment, useMemo, useState } from "react";
import * as XLSX from "xlsx";

type Direction = "买" | "卖";
type Offset = "默认" | "开仓" | "平仓";
type Rule = {
  id: number;
  business: string;
  contract: string;
  exchanges: string[];
  products: string[];
  direction: Direction;
  offset: Offset;
  template: string;
};
type Account = { name: string; rules: Rule[] };

type TemplateDetail = {
  order: number;
  channel: string;
};

type OrderTemplate = {
  name: string;
  note: string;
  details: TemplateDetail[];
};

type ImportTemplate = OrderTemplate & {
  mode: "新增" | "覆盖";
};

const initialTemplates: OrderTemplate[] = [
  { name: "默认模板", note: "CMSI_SC999", details: [{ order: 1, channel: "Mock" }, { order: 2, channel: "QFII" }, { order: 3, channel: "QFII2" }, { order: 4, channel: "O32" }] },
  { name: "test1", note: "QFII", details: [{ order: 1, channel: "QFII" }] },
  { name: "test2", note: "Mock", details: [{ order: 1, channel: "Mock" }] },
  { name: "test3", note: "QFII2", details: [{ order: 1, channel: "QFII2" }] },
  { name: "test4", note: "CGSI", details: [{ order: 1, channel: "CGSI" }] },
  { name: "sjn-mock1", note: "", details: [{ order: 1, channel: "Mock" }] },
  { name: "sjn-mock2", note: "", details: [{ order: 1, channel: "Mock" }] },
];

const seedAccounts: Account[] = [
  {
    name: "msycs",
    rules: [
      { id: 1, business: "普通交易", contract: "股票", exchanges: ["深交所", "上交所", "深港通", "沪股通"], products: ["全部"], direction: "买", offset: "默认", template: "test1" },
      { id: 2, business: "普通交易", contract: "期货", exchanges: ["大连商品交易所", "上海期货交易所"], products: ["聚丙烯", "白银"], direction: "卖", offset: "开仓", template: "test1" },
      { id: 3, business: "融资交易", contract: "股票", exchanges: ["深交所", "上交所", "香港交易所", "深港通", "沪股通"], products: ["全部"], direction: "买", offset: "平仓", template: "默认模板" },
      { id: 4, business: "普通交易", contract: "股票", exchanges: ["深交所", "上交所", "香港交易所"], products: ["全部"], direction: "卖", offset: "默认", template: "默认模板" },
    ],
  },
  { name: "142yyq03", rules: [] },
  { name: "142yyq01", rules: [] },
  { name: "李四001", rules: [] },
];

const navItems = ["定价簿记", "实时风控", "持仓管理", "客户管理", "资金/保证金管理", "日终管理", "监管报送", "报告中心", "OMS管理", "example-sdk", "市场数据", "定价配置"];

function Chevron({ open }: { open: boolean }) {
  return <span className="square-toggle">{open ? "−" : "+"}</span>;
}

function ChipList({ values }: { values: string[] }) {
  return (
    <div className="chips">
      {values.map((value) =>
        value === "全部" ? <span className="plain-value" key={value}>全部</span> :
          <span className="chip" key={value}>{value}</span>,
      )}
    </div>
  );
}

export default function Home() {
  const [activeTab, setActiveTab] = useState<"order" | "template">("order");
  const [orderTemplates, setOrderTemplates] = useState<OrderTemplate[]>(initialTemplates);
  const [expandedTemplates, setExpandedTemplates] = useState<string[]>([]);
  const [selectedTemplates, setSelectedTemplates] = useState<string[]>([]);
  const [accounts, setAccounts] = useState(seedAccounts);
  const [expandedAccounts, setExpandedAccounts] = useState<string[]>(["msycs"]);
  const [editing, setEditing] = useState<{ account: string; id: number } | null>({ account: "msycs", id: 1 });
  const [draft, setDraft] = useState<Rule>(seedAccounts[0].rules[0]);
  const [queryAccount, setQueryAccount] = useState("");
  const [showNewTemplate, setShowNewTemplate] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [importFileName, setImportFileName] = useState("");
  const [importPreview, setImportPreview] = useState<ImportTemplate[]>([]);
  const [importErrors, setImportErrors] = useState<string[]>([]);

  const visibleAccounts = useMemo(() => queryAccount ? accounts.filter((account) => account.name === queryAccount) : accounts, [accounts, queryAccount]);

  const toggleTemplate = (name: string) => setExpandedTemplates((current) =>
    current.includes(name) ? current.filter((item) => item !== name) : [...current, name],
  );
  const toggleSelectedTemplate = (name: string) => setSelectedTemplates((current) =>
    current.includes(name) ? current.filter((item) => item !== name) : [...current, name],
  );
  const toggleAllTemplates = () => setSelectedTemplates((current) =>
    current.length === orderTemplates.length ? [] : orderTemplates.map((template) => template.name),
  );
  const deleteTemplate = (name: string) => {
    setOrderTemplates((current) => current.filter((template) => template.name !== name));
    setSelectedTemplates((current) => current.filter((item) => item !== name));
    setExpandedTemplates((current) => current.filter((item) => item !== name));
  };
  const batchDeleteTemplates = () => {
    setOrderTemplates((current) => current.filter((template) => !selectedTemplates.includes(template.name)));
    setExpandedTemplates((current) => current.filter((item) => !selectedTemplates.includes(item)));
    setSelectedTemplates([]);
    setShowDeleteConfirm(false);
  };
  const resetImport = () => {
    setImportFileName("");
    setImportPreview([]);
    setImportErrors([]);
  };
  const closeImport = () => {
    setShowImport(false);
    resetImport();
  };
  const parseImportFile = async (file: File) => {
    setImportFileName(file.name);
    setImportPreview([]);
    setImportErrors([]);
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const sheetName = workbook.SheetNames.includes("批量导入模板") ? "批量导入模板" : workbook.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json<(string | number)[]>(workbook.Sheets[sheetName], {
        header: 1,
        defval: "",
        raw: true,
      });
      const normalize = (value: string | number) => String(value).replace(/\*/g, "").trim();
      const headerRowIndex = rows.findIndex((row) => row.some((cell) => normalize(cell) === "模板名称"));
      if (headerRowIndex < 0) {
        setImportErrors(["未找到标准表头，请使用页面提供的 Excel 模板。"]);
        return;
      }
      const headers = rows[headerRowIndex].map(normalize);
      const nameIndex = headers.indexOf("模板名称");
      const noteIndex = headers.indexOf("备注");
      const orderIndex = headers.indexOf("报单排序");
      const channelIndex = headers.indexOf("对冲渠道");
      const modeIndex = headers.indexOf("导入方式");
      if ([nameIndex, orderIndex, channelIndex, modeIndex].some((index) => index < 0)) {
        setImportErrors(["表头不完整，必须包含模板名称、报单排序、对冲渠道和导入方式。"]);
        return;
      }

      const errors: string[] = [];
      const grouped = new Map<string, ImportTemplate>();
      rows.slice(headerRowIndex + 1).forEach((row, rowOffset) => {
        const excelRow = headerRowIndex + rowOffset + 2;
        const name = normalize(row[nameIndex] ?? "");
        const note = noteIndex >= 0 ? normalize(row[noteIndex] ?? "") : "";
        const channel = normalize(row[channelIndex] ?? "");
        const modeText = normalize(row[modeIndex] ?? "");
        const order = Number(row[orderIndex]);
        if (!name && !channel && !modeText && !row[orderIndex]) return;
        if (!name || !channel || !modeText || !Number.isInteger(order) || order <= 0) {
          errors.push(`第 ${excelRow} 行：请完整填写模板名称、正整数报单排序、对冲渠道和导入方式。`);
          return;
        }
        if (modeText !== "新增" && modeText !== "覆盖") {
          errors.push(`第 ${excelRow} 行：导入方式只能填写“新增”或“覆盖”。`);
          return;
        }
        const existing = grouped.get(name);
        if (existing && existing.mode !== modeText) {
          errors.push(`第 ${excelRow} 行：同一模板的导入方式必须一致。`);
          return;
        }
        if (existing?.details.some((detail) => detail.order === order)) {
          errors.push(`第 ${excelRow} 行：模板“${name}”的报单排序 ${order} 重复。`);
          return;
        }
        if (existing) {
          existing.details.push({ order, channel });
        } else {
          grouped.set(name, { name, note, mode: modeText, details: [{ order, channel }] });
        }
      });

      const preview = [...grouped.values()].map((template) => ({
        ...template,
        details: template.details.sort((a, b) => a.order - b.order),
      }));
      preview.forEach((template) => {
        if (template.mode === "新增" && orderTemplates.some((existing) => existing.name === template.name)) {
          errors.push(`模板“${template.name}”已存在，请将导入方式改为“覆盖”或修改模板名称。`);
        }
      });
      if (!preview.length && !errors.length) errors.push("Excel 中没有可导入的数据。");
      setImportPreview(preview);
      setImportErrors(errors.slice(0, 8));
    } catch {
      setImportErrors(["文件解析失败，请确认上传的是有效的 .xlsx 或 .xls 文件。"]);
    }
  };
  const confirmImport = () => {
    if (importErrors.length || !importPreview.length) return;
    setOrderTemplates((current) => {
      const importedNames = importPreview.map((template) => template.name);
      const preserved = current.filter((template) => !importedNames.includes(template.name));
      return [
        ...preserved,
        ...importPreview.map(({ mode: _mode, ...template }) => template),
      ];
    });
    setExpandedTemplates((current) => [...new Set([...current, ...importPreview.map((template) => template.name)])]);
    closeImport();
  };
  const addTemplateDetail = (name: string) => setOrderTemplates((current) =>
    current.map((template) => template.name === name
      ? {
          ...template,
          details: [
            ...template.details,
            {
              order: Math.max(0, ...template.details.map((detail) => detail.order)) + 1,
              channel: "请选择渠道",
            },
          ],
        }
      : template),
  );
  const deleteTemplateDetail = (name: string, order: number) => setOrderTemplates((current) =>
    current.map((template) => template.name === name
      ? { ...template, details: template.details.filter((detail) => detail.order !== order) }
      : template),
  );
  const toggleAccount = (name: string) => setExpandedAccounts((current) =>
    current.includes(name) ? current.filter((item) => item !== name) : [...current, name],
  );
  const startEditing = (account: string, rule: Rule) => {
    setEditing({ account, id: rule.id });
    setDraft({ ...rule });
  };
  const saveDraft = () => {
    if (!editing) return;
    setAccounts((current) => current.map((account) =>
      account.name === editing.account ? { ...account, rules: account.rules.map((rule) => rule.id === editing.id ? { ...draft } : rule) } : account,
    ));
    setEditing(null);
  };
  const addRule = (accountName: string) => {
    const rule: Rule = { id: Date.now(), business: "普通交易", contract: "股票", exchanges: ["深交所", "上交所"], products: ["全部"], direction: "买", offset: "默认", template: "默认模板" };
    setAccounts((current) => current.map((account) => account.name === accountName ? { ...account, rules: [...account.rules, rule] } : account));
    setExpandedAccounts((current) => current.includes(accountName) ? current : [...current, accountName]);
    setEditing({ account: accountName, id: rule.id });
    setDraft(rule);
  };
  const deleteRule = (accountName: string, id: number) => setAccounts((current) => current.map((account) =>
    account.name === accountName ? { ...account, rules: account.rules.filter((rule) => rule.id !== id) } : account,
  ));
  const addAccount = () => setAccounts((current) => [...current, { name: `新交易账户${current.length + 1}`, rules: [] }]);

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand"><span className="logo">7</span><span>BCT</span></div>
        <nav className="main-nav" aria-label="主导航">
          {navItems.map((item) => <span className={item === "OMS管理" ? "nav-item active" : "nav-item"} key={item}>{item}</span>)}
          <span className="more">···</span>
        </nav>
        <div className="top-actions"><span className="search-icon">⌕</span><span className="bell">♧<i /></span><span className="avatar">♙</span><span>script</span></div>
      </header>

      <main>
        <div className="breadcrumb-row">
          <div><span>OMS管理</span><b>/</b>报单排序配置</div>
          <label className="node-select">当前节点：
            <select defaultValue="SH" aria-label="当前节点"><option>SH</option><option>SZ</option></select>
          </label>
        </div>

        <div className="tabs-row">
          <div className="tabs">
            <button className={activeTab === "order" ? "tab active" : "tab"} onClick={() => setActiveTab("order")}>报单顺序配置</button>
            <button className={activeTab === "template" ? "tab active" : "tab"} onClick={() => setActiveTab("template")}>模板配置</button>
          </div>
        </div>

        <div className="template-actions">
          <button className="primary" onClick={() => setShowNewTemplate(true)}>新 增</button>
          <button className="ghost" onClick={() => setShowImport(true)}>⇧ 批量导入</button>
          <a className="ghost button-link" href="/报单排序模板批量导入模板.xlsx" download>⇩ 下载导入模板</a>
          <button
            className="danger-outline"
            disabled={!selectedTemplates.length}
            onClick={() => setShowDeleteConfirm(true)}
          >
            批量删除{selectedTemplates.length ? `（${selectedTemplates.length}）` : ""}
          </button>
          <button
            className="expand-all"
            onClick={() => setExpandedTemplates(
              expandedTemplates.length === orderTemplates.length
                ? []
                : orderTemplates.map((template) => template.name),
            )}
          >
            {expandedTemplates.length === orderTemplates.length ? "全部收起" : "全部展开"}
          </button>
        </div>

        {showNewTemplate && (
          <div className="inline-form">
            <strong>新增报单顺序</strong>
            <input aria-label="报单顺序名称" placeholder="请输入报单顺序名称" />
            <input aria-label="备注" placeholder="请输入备注" />
            <button className="primary small" onClick={() => setShowNewTemplate(false)}>保存</button>
            <button className="ghost small" onClick={() => setShowNewTemplate(false)}>取消</button>
          </div>
        )}

        <section className="template-section">
          <table className="data-table template-table">
            <thead><tr>
              <th className="select-col">
                <input
                  type="checkbox"
                  aria-label="全选报单排序模板"
                  checked={orderTemplates.length > 0 && selectedTemplates.length === orderTemplates.length}
                  onChange={toggleAllTemplates}
                />
              </th>
              <th className="toggle-col" />
              <th>报单排序模板名称</th>
              <th>备注</th>
              <th className="actions-col">操作</th>
            </tr></thead>
            <tbody>
              {orderTemplates.map((template) => (
                <Fragment key={template.name}>
                  <tr className={selectedTemplates.includes(template.name) ? "selected-template-row" : ""}>
                    <td>
                      <input
                        type="checkbox"
                        aria-label={`选择模板${template.name}`}
                        checked={selectedTemplates.includes(template.name)}
                        onChange={() => toggleSelectedTemplate(template.name)}
                      />
                    </td>
                    <td><button className="toggle-button" onClick={() => toggleTemplate(template.name)} aria-label={`${expandedTemplates.includes(template.name) ? "收起" : "展开"}${template.name}`}><Chevron open={expandedTemplates.includes(template.name)} /></button></td>
                    <td className={template.name === "默认模板" ? "strong" : ""}>{template.name}</td>
                    <td>{template.note}</td>
                    <td><button className="link">编辑</button><button className="link danger" onClick={() => deleteTemplate(template.name)}>删除</button></td>
                  </tr>
                  {expandedTemplates.includes(template.name) && (
                    <tr className="template-detail">
                      <td />
                      <td colSpan={4}>
                        <div className="template-detail-panel">
                          <table className="template-detail-table">
                            <thead><tr><th className="drag-col" /><th>报单排序</th><th>对冲渠道</th><th>操作</th></tr></thead>
                            <tbody>
                              {template.details.map((detail) => (
                                <tr key={`${template.name}-${detail.order}`}>
                                  <td className="drag-handle">⠿</td>
                                  <td>{detail.order}</td>
                                  <td>{detail.channel}</td>
                                  <td>
                                    <button className="link">编辑</button>
                                    <button className="link danger" onClick={() => deleteTemplateDetail(template.name, detail.order)}>删除</button>
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                          <button className="add-template-detail" onClick={() => addTemplateDetail(template.name)}>新增一行</button>
                        </div>
                      </td>
                    </tr>
                  )}
                </Fragment>
              ))}
            </tbody>
          </table>
          <div className="template-count">共 {orderTemplates.length} 项</div>
        </section>

        <section className="client-section">
          <div className="section-title-row"><h2>客户配置</h2></div>
          <div className="toolbar">
            <label>交易账户：
              <select value={queryAccount} onChange={(event) => setQueryAccount(event.target.value)} aria-label="筛选交易账户">
                <option value="">请选择</option>
                {accounts.map((account) => <option value={account.name} key={account.name}>{account.name}</option>)}
              </select>
            </label>
            <button className="primary search-button"><span>⌕</span> 查询</button>
            <button className="ghost" onClick={() => setQueryAccount("")}>重 置</button>
            <button className="primary add-account" onClick={addAccount}>新 增</button>
          </div>

          <div className="account-table">
            <div className="account-head"><span /><span>交易账户</span><span>操作</span></div>
            {visibleAccounts.map((account) => {
              const isOpen = expandedAccounts.includes(account.name);
              return (
                <div className="account-block" key={account.name}>
                  <div className="account-row">
                    <button className="toggle-button" onClick={() => toggleAccount(account.name)} aria-label={`${isOpen ? "收起" : "展开"}${account.name}`}><Chevron open={isOpen} /></button>
                    <span>{account.name}</span><button className="link danger">删除</button>
                  </div>
                  {isOpen && (
                    <div className="rules-wrap">
                      <table className="data-table rules-table">
                        <thead><tr>
                          <th>业务类型</th><th>合约类型</th><th>交易所</th><th>资产品种</th>
                          <th className="direction-head">交易方向<span>买 / 卖</span></th>
                          <th className="offset-head">开平标志<span>默认 / 开仓 / 平仓</span></th>
                          <th>顺序名称</th><th className="rule-actions">操作</th>
                        </tr></thead>
                        <tbody>
                          {account.rules.map((rule) => {
                            const isEditing = editing?.account === account.name && editing.id === rule.id;
                            const shown = isEditing ? draft : rule;
                            return (
                              <tr className={isEditing ? "editing-row" : ""} key={rule.id}>
                                <td>{isEditing ? <select value={shown.business} onChange={(event) => setDraft({ ...draft, business: event.target.value })} aria-label="业务类型"><option>普通交易</option><option>融资交易</option><option>融券交易</option></select> : shown.business}</td>
                                <td>{isEditing ? <select value={shown.contract} onChange={(event) => setDraft({ ...draft, contract: event.target.value })} aria-label="合约类型"><option>股票</option><option>期货</option><option>期权</option></select> : shown.contract}</td>
                                <td><ChipList values={shown.exchanges} /></td><td><ChipList values={shown.products} /></td>
                                <td className="direction-cell">
                                  {isEditing ? (
                                    <select
                                      className="direction-select"
                                      value={shown.direction}
                                      onChange={(event) => setDraft({ ...draft, direction: event.target.value as Direction })}
                                      aria-label="交易方向"
                                    >
                                      <option value="买">买</option>
                                      <option value="卖">卖</option>
                                    </select>
                                  ) : (
                                    <span className={`status ${shown.direction === "买" ? "buy" : "sell"}`}>{shown.direction}</span>
                                  )}
                                </td>
                                <td className="offset-cell">
                                  {isEditing ? (
                                    <select
                                      className="offset-select"
                                      value={shown.offset}
                                      onChange={(event) => setDraft({ ...draft, offset: event.target.value as Offset })}
                                      aria-label="开平标志"
                                    >
                                      <option value="默认">默认</option>
                                      <option value="开仓">开仓</option>
                                      <option value="平仓">平仓</option>
                                    </select>
                                  ) : (
                                    <span className="offset-value">{shown.offset}</span>
                                  )}
                                </td>
                                <td>{isEditing ? <select value={shown.template} onChange={(event) => setDraft({ ...draft, template: event.target.value })} aria-label="顺序名称">{orderTemplates.map((template) => <option key={template.name}>{template.name}</option>)}</select> : shown.template}</td>
                                <td>{isEditing ? <><button className="link" onClick={saveDraft}>保存</button><button className="link" onClick={() => setEditing(null)}>取消</button></> : <><button className="link" onClick={() => startEditing(account.name, rule)}>编辑</button><button className="link danger" onClick={() => deleteRule(account.name, rule.id)}>删除</button></>}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                      <button className="add-rule" onClick={() => addRule(account.name)}>＋ 新增一行</button>
                    </div>
                  )}
                </div>
              );
            })}
          </div>

          <div className="pagination"><span>共 {visibleAccounts.length} 个交易账户</span><button disabled>‹</button><button className="current">1</button><button>›</button><select aria-label="每页数量" defaultValue="10 条/页"><option>10 条/页</option><option>20 条/页</option></select></div>
        </section>
      </main>

      {showImport && (
        <div className="modal-backdrop" role="presentation">
          <section className="import-modal" role="dialog" aria-modal="true" aria-labelledby="import-title">
            <div className="modal-header">
              <div>
                <h2 id="import-title">批量导入报单排序模板</h2>
                <p>上传后将先校验并预览，不会立即写入列表。</p>
              </div>
              <button className="modal-close" aria-label="关闭批量导入" onClick={closeImport}>×</button>
            </div>
            <div className="modal-body">
              <div className="import-steps">
                <span className="active"><b>1</b> 上传文件</span>
                <i />
                <span className={importPreview.length || importErrors.length ? "active" : ""}><b>2</b> 数据校验</span>
                <i />
                <span><b>3</b> 确认导入</span>
              </div>
              <label className="upload-zone">
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  aria-label="选择Excel文件"
                  onChange={(event) => {
                    const file = event.target.files?.[0];
                    if (file) void parseImportFile(file);
                  }}
                />
                <span className="upload-icon">⇧</span>
                <strong>{importFileName || "点击选择或拖入 Excel 文件"}</strong>
                <small>支持 .xlsx、.xls，单次最多导入 500 条排序明细</small>
              </label>
              <div className="template-download-tip">
                首次使用？请先
                <a href="/报单排序模板批量导入模板.xlsx" download>下载标准导入模板</a>
                ，不要修改表头名称。
              </div>
              {importErrors.length > 0 && (
                <div className="import-errors">
                  <strong>发现 {importErrors.length} 个问题</strong>
                  <ul>{importErrors.map((error) => <li key={error}>{error}</li>)}</ul>
                </div>
              )}
              {importPreview.length > 0 && (
                <div className="import-preview">
                  <div className="preview-summary">
                    <strong>文件预览</strong>
                    <span>{importPreview.length} 个模板 · {importPreview.reduce((total, template) => total + template.details.length, 0)} 条排序明细</span>
                  </div>
                  <table>
                    <thead><tr><th>模板名称</th><th>备注</th><th>排序明细数</th><th>导入方式</th><th>状态</th></tr></thead>
                    <tbody>
                      {importPreview.map((template) => (
                        <tr key={template.name}>
                          <td>{template.name}</td>
                          <td>{template.note || "—"}</td>
                          <td>{template.details.length}</td>
                          <td><span className={`mode-tag ${template.mode === "覆盖" ? "overwrite" : ""}`}>{template.mode}</span></td>
                          <td><span className={importErrors.length ? "invalid-status" : "valid-status"}>{importErrors.length ? "待修正" : "校验通过"}</span></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
            <div className="modal-footer">
              <button className="ghost" onClick={closeImport}>取消</button>
              <button className="primary" disabled={Boolean(importErrors.length) || !importPreview.length} onClick={confirmImport}>确认导入</button>
            </div>
          </section>
        </div>
      )}

      {showDeleteConfirm && (
        <div className="modal-backdrop" role="presentation">
          <section className="confirm-modal" role="dialog" aria-modal="true" aria-labelledby="delete-title">
            <span className="warning-icon">!</span>
            <h2 id="delete-title">确认批量删除？</h2>
            <p>将删除已选择的 {selectedTemplates.length} 个报单排序模板，此操作仅作用于当前页面数据。</p>
            <div>
              <button className="ghost" onClick={() => setShowDeleteConfirm(false)}>取消</button>
              <button className="danger-button" onClick={batchDeleteTemplates}>确认删除</button>
            </div>
          </section>
        </div>
      )}

      <footer>Copyright&nbsp; © &nbsp;Tongyu All Rights Reserved&nbsp;&nbsp; 上海同余信息科技有限公司&nbsp;&nbsp; 沪ICP备16043748号&nbsp;&nbsp; 测试环境&nbsp;&nbsp; vdevelop_20260731093237</footer>
    </div>
  );
}
