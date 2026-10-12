"use client";

import { Fragment, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const publicBasePath = process.env.NEXT_PUBLIC_BASE_PATH ?? "";

type Direction = "买入" | "卖出" | "买开" | "卖开" | "买平" | "卖平";
type Rule = {
  id: number;
  business: string;
  contract: string;
  exchanges: string[];
  products: string[];
  direction: Direction;
  template: string;
};
type Account = { name: string; rules: Rule[] };
type ConfigTemplate = { name: string; note: string; rules: Rule[] };

type TemplateDetail = {
  order: number;
  channel: string;
};

type OrderTemplate = {
  name: string;
  note: string;
  details: TemplateDetail[];
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

const initialConfigTemplates: ConfigTemplate[] = [
  {
    name: "默认模板",
    note: "通用",
    rules: [
      { id: 101, business: "普通交易", contract: "股票", exchanges: ["深交所", "上交所", "深港通", "沪股通"], products: ["全部"], direction: "买入", template: "test1" },
      { id: 102, business: "普通交易", contract: "期货", exchanges: ["大连商品交易所", "上海期货交易所"], products: ["聚丙烯", "白银"], direction: "买开", template: "test1" },
      { id: 103, business: "融资交易", contract: "股票", exchanges: ["深交所", "上交所", "香港交易所", "深港通", "沪股通"], products: ["全部"], direction: "卖平", template: "默认模板" },
    ],
  },
  {
    name: "test",
    note: "5413",
    rules: [
      { id: 201, business: "普通交易", contract: "股票", exchanges: ["深交所", "上交所"], products: ["全部"], direction: "买入", template: "test1" },
    ],
  },
  {
    name: "期货",
    note: "",
    rules: [
      { id: 301, business: "普通交易", contract: "期货", exchanges: ["大连商品交易所"], products: ["尿素"], direction: "买开", template: "test1" },
      { id: 302, business: "普通交易", contract: "期货", exchanges: ["大连商品交易所", "上海期货交易所"], products: ["全部"], direction: "买平", template: "test2" },
    ],
  },
  { name: "测试模板", note: "", rules: [] },
];

const seedAccounts: Account[] = [
  {
    name: "msycs",
    rules: [
      { id: 1, business: "普通交易", contract: "股票", exchanges: ["深交所", "上交所", "深港通", "沪股通"], products: ["全部"], direction: "买入", template: "test1" },
      { id: 2, business: "普通交易", contract: "期货", exchanges: ["大连商品交易所", "上海期货交易所"], products: ["聚丙烯", "白银"], direction: "卖开", template: "test1" },
      { id: 3, business: "融资交易", contract: "股票", exchanges: ["深交所", "上交所", "香港交易所", "深港通", "沪股通"], products: ["全部"], direction: "买平", template: "默认模板" },
      { id: 4, business: "普通交易", contract: "股票", exchanges: ["深交所", "上交所", "香港交易所"], products: ["全部"], direction: "卖出", template: "默认模板" },
    ],
  },
  { name: "142yyq03", rules: [] },
  { name: "142yyq01", rules: [] },
  { name: "李四001", rules: [] },
];

const navItems = ["定价簿记", "实时风控", "持仓管理", "客户管理", "资金/保证金管理", "日终管理", "监管报送", "报告中心", "OMS管理", "example-sdk", "市场数据", "定价配置"];
const emptyRule = (): Rule => ({
  id: Date.now() + Math.floor(Math.random() * 1000),
  business: "普通交易",
  contract: "股票",
  exchanges: ["深交所", "上交所"],
  products: ["全部"],
  direction: "买入",
  template: "默认模板",
});
const clientRuleKey = (account: string, id: number) => `${account}::${id}`;
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

type RuleTableProps = {
  scope: string;
  rules: Rule[];
  editingId: number | null;
  draft: Rule;
  orderTemplates: OrderTemplate[];
  onEdit: (rule: Rule) => void;
  onDraft: (rule: Rule) => void;
  onSave: () => void;
  onCancel: () => void;
  onDelete: (id: number) => void;
  onAdd: () => void;
  selectedIds?: number[];
  onToggleSelect?: (id: number) => void;
};

function RuleTable({
  scope,
  rules,
  editingId,
  draft,
  orderTemplates,
  onEdit,
  onDraft,
  onSave,
  onCancel,
  onDelete,
  onAdd,
  selectedIds,
  onToggleSelect,
}: RuleTableProps) {
  const selectable = Boolean(selectedIds && onToggleSelect);
  return (
    <div className="rules-wrap">
      <table className={`data-table rules-table${selectable ? " selectable-rules-table" : ""}`}>
        <thead><tr>
          {selectable && (
            <th className="rule-select-col" aria-label="批量选择" />
          )}
          <th>业务类型</th><th>合约类型</th><th>交易所</th><th>资产品种</th>
          <th className="direction-head">交易方向<span>买入 / 卖出 / 买开 / 卖开 / 买平 / 卖平</span></th>
          <th>顺序名称</th><th className="rule-actions">操作</th>
        </tr></thead>
        <tbody>
          {rules.map((rule) => {
            const isEditing = editingId === rule.id;
            const shown = isEditing ? draft : rule;
            return (
              <tr className={isEditing ? "editing-row" : ""} key={rule.id}>
                {selectable && (
                  <td className="rule-select-cell">
                    <input
                      type="checkbox"
                      aria-label={`${scope}选择配置${rule.id}`}
                      checked={selectedIds?.includes(rule.id)}
                      onChange={() => onToggleSelect?.(rule.id)}
                    />
                  </td>
                )}
                <td>{isEditing ? (
                  <select value={shown.business} onChange={(event) => onDraft({ ...draft, business: event.target.value })} aria-label={`${scope}业务类型`}>
                    <option>普通交易</option><option>融资交易</option><option>融券交易</option><option>对冲互换</option>
                  </select>
                ) : shown.business}</td>
                <td>{isEditing ? (
                  <select value={shown.contract} onChange={(event) => onDraft({ ...draft, contract: event.target.value })} aria-label={`${scope}合约类型`}>
                    <option>股票</option><option>期货</option><option>期权</option><option>互换</option>
                  </select>
                ) : shown.contract}</td>
                <td><ChipList values={shown.exchanges} /></td>
                <td><ChipList values={shown.products} /></td>
                <td className="direction-cell">
                  {isEditing ? (
                    <select className="direction-select" value={shown.direction} onChange={(event) => onDraft({ ...draft, direction: event.target.value as Direction })} aria-label={`${scope}交易方向`}>
                      <option value="买入">买入</option><option value="卖出">卖出</option>
                      <option value="买开">买开</option><option value="卖开">卖开</option>
                      <option value="买平">买平</option><option value="卖平">卖平</option>
                    </select>
                  ) : <span className={`status ${shown.direction.startsWith("买") ? "buy" : "sell"}`}>{shown.direction}</span>}
                </td>
                <td>{isEditing ? (
                  <select value={shown.template} onChange={(event) => onDraft({ ...draft, template: event.target.value })} aria-label={`${scope}顺序名称`}>
                    {orderTemplates.map((template) => <option key={template.name}>{template.name}</option>)}
                  </select>
                ) : shown.template}</td>
                <td>{isEditing ? (
                  <><button className="link" onClick={onSave}>保存</button><button className="link" onClick={onCancel}>取消</button></>
                ) : (
                  <><button className="link" onClick={() => onEdit(rule)}>编辑</button><button className="link danger" onClick={() => onDelete(rule.id)}>删除</button></>
                )}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
      <button className="add-rule" onClick={onAdd}>＋ 新增一行</button>
    </div>
  );
}

export default function Home() {
  const [activeTab, setActiveTab] = useState<"order" | "template">("order");
  const [orderTemplates, setOrderTemplates] = useState<OrderTemplate[]>(initialTemplates);
  const [expandedTemplates, setExpandedTemplates] = useState<string[]>([]);
  const [configTemplates, setConfigTemplates] = useState<ConfigTemplate[]>(initialConfigTemplates);
  const [expandedConfigTemplates, setExpandedConfigTemplates] = useState<string[]>([]);
  const [editingConfig, setEditingConfig] = useState<{ name: string; id: number } | null>(null);
  const [configDraft, setConfigDraft] = useState<Rule>(initialConfigTemplates[0].rules[0]);
  const [accounts, setAccounts] = useState(seedAccounts);
  const [expandedAccounts, setExpandedAccounts] = useState<string[]>(["msycs"]);
  const [editing, setEditing] = useState<{ account: string; id: number } | null>(null);
  const [draft, setDraft] = useState<Rule>(seedAccounts[0].rules[0]);
  const [queryAccount, setQueryAccount] = useState("");
  const [showNewTemplate, setShowNewTemplate] = useState(false);
  const [newTemplateName, setNewTemplateName] = useState("");
  const [newTemplateNote, setNewTemplateNote] = useState("");
  const [editingTop, setEditingTop] = useState<{ kind: "order" | "template"; originalName: string; name: string; note: string } | null>(null);
  const [editingDetail, setEditingDetail] = useState<{ template: string; order: number } | null>(null);
  const [detailDraft, setDetailDraft] = useState<TemplateDetail>({ order: 1, channel: "Mock" });
  const [draggedDetail, setDraggedDetail] = useState<{ template: string; order: number } | null>(null);
  const [showNewAccount, setShowNewAccount] = useState(false);
  const [newAccountName, setNewAccountName] = useState("");
  const [selectedClientRules, setSelectedClientRules] = useState<string[]>([]);
  const [showClientImport, setShowClientImport] = useState(false);
  const [showClientDeleteConfirm, setShowClientDeleteConfirm] = useState(false);
  const [clientImportFileName, setClientImportFileName] = useState("");
  const [clientImportBusy, setClientImportBusy] = useState(false);
  const [toast, setToast] = useState("");

  const visibleAccounts = useMemo(() => queryAccount ? accounts.filter((account) => account.name === queryAccount) : accounts, [accounts, queryAccount]);
  const notify = (message: string) => {
    setToast(message);
    window.setTimeout(() => setToast(""), 1800);
  };

  const toggleTemplate = (name: string) => setExpandedTemplates((current) =>
    current.includes(name) ? current.filter((item) => item !== name) : [...current, name],
  );
  const deleteTemplate = (name: string) => {
    setOrderTemplates((current) => current.filter((template) => template.name !== name));
    setExpandedTemplates((current) => current.filter((item) => item !== name));
    notify("删除成功");
  };
  const resetClientImport = () => {
    setClientImportFileName("");
    setClientImportBusy(false);
  };
  const closeClientImport = () => {
    setShowClientImport(false);
    resetClientImport();
  };
  const downloadClientImportErrors = (errors: string[]) => {
    if (!errors.length) return;
    const content = `\uFEFF${errors.join("\r\n")}\r\n`;
    const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = "客户配置导入失败.txt";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 0);
  };
  const reportClientImportErrors = (errors: string[]) => {
    downloadClientImportErrors(errors);
    notify("导入失败，已下载报错文件");
  };
  const importClientFile = async (file: File) => {
    setClientImportFileName(file.name);
    setClientImportBusy(true);
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const sheetName = workbook.SheetNames.includes("客户配置导入") ? "客户配置导入" : workbook.SheetNames[0];
      if (!sheetName || !workbook.Sheets[sheetName]) throw new Error("empty workbook");
      const rows = XLSX.utils.sheet_to_json<(string | number)[]>(workbook.Sheets[sheetName], {
        header: 1,
        defval: "",
        raw: true,
      });
      const normalize = (value: string | number) => String(value).replace(/\*/g, "").trim();
      const headerRowIndex = rows.findIndex((row) => row.some((cell) => normalize(cell) === "交易账户"));
      if (headerRowIndex < 0) {
        reportClientImportErrors(["文件格式错误：未找到标准表头。"]);
        return;
      }
      const headers = rows[headerRowIndex].map(normalize);
      const requiredHeaders = ["交易账户", "业务类型", "合约类型", "交易所", "资产品种", "交易方向", "顺序名称"];
      if (requiredHeaders.some((header) => !headers.includes(header))) {
        reportClientImportErrors(["文件格式错误：表头与导入模板不一致。"]);
        return;
      }
      const hasData = rows.slice(headerRowIndex + 1).some((row) => row.some((cell) => normalize(cell) !== ""));
      if (!hasData) {
        reportClientImportErrors(["文件错误：导入内容为空。"]);
        return;
      }
      closeClientImport();
      notify("客户配置导入成功");
    } catch {
      reportClientImportErrors(["文件格式错误：无法读取上传文件。"]);
    } finally {
      setClientImportBusy(false);
    }
  };
  const toggleClientRule = (account: string, id: number) => {
    const key = clientRuleKey(account, id);
    setSelectedClientRules((current) => current.includes(key) ? current.filter((item) => item !== key) : [...current, key]);
  };
  const toggleAllClientRules = (account: Account) => {
    const keys = account.rules.map((rule) => clientRuleKey(account.name, rule.id));
    const allSelected = keys.length > 0 && keys.every((key) => selectedClientRules.includes(key));
    setSelectedClientRules((current) => allSelected
      ? current.filter((key) => !keys.includes(key))
      : [...new Set([...current, ...keys])]);
  };
  const batchDeleteClientRules = () => {
    const selected = new Set(selectedClientRules);
    setAccounts((current) => current.map((account) => ({
      ...account,
      rules: account.rules.filter((rule) => !selected.has(clientRuleKey(account.name, rule.id))),
    })));
    setSelectedClientRules([]);
    setShowClientDeleteConfirm(false);
    notify("客户模板配置批量删除成功");
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
  const startDetailEditing = (template: string, detail: TemplateDetail) => {
    setEditingDetail({ template, order: detail.order });
    setDetailDraft({ ...detail });
  };
  const saveDetail = () => {
    if (!editingDetail || !detailDraft.channel.trim()) return;
    setOrderTemplates((current) => current.map((template) => template.name === editingDetail.template
      ? {
          ...template,
          details: template.details
            .map((detail) => detail.order === editingDetail.order ? { ...detailDraft } : detail)
            .sort((a, b) => a.order - b.order),
        }
      : template));
    setEditingDetail(null);
    notify("保存成功");
  };
  const reorderDetails = (templateName: string, sourceOrder: number, targetOrder: number) => {
    if (sourceOrder === targetOrder) return;
    setOrderTemplates((current) => current.map((template) => {
      if (template.name !== templateName) return template;
      const details = [...template.details].sort((a, b) => a.order - b.order);
      const sourceIndex = details.findIndex((detail) => detail.order === sourceOrder);
      const targetIndex = details.findIndex((detail) => detail.order === targetOrder);
      if (sourceIndex < 0 || targetIndex < 0) return template;
      const [moved] = details.splice(sourceIndex, 1);
      details.splice(targetIndex, 0, moved);
      return { ...template, details: details.map((detail, index) => ({ ...detail, order: index + 1 })) };
    }));
    notify("排序已更新");
  };
  const saveNewTop = () => {
    const name = newTemplateName.trim();
    if (!name) return;
    if (activeTab === "order") {
      if (orderTemplates.some((template) => template.name === name)) return notify("名称已存在");
      setOrderTemplates((current) => [...current, { name, note: newTemplateNote.trim(), details: [{ order: 1, channel: "请选择渠道" }] }]);
      setExpandedTemplates((current) => [...current, name]);
    } else {
      if (configTemplates.some((template) => template.name === name)) return notify("名称已存在");
      setConfigTemplates((current) => [...current, { name, note: newTemplateNote.trim(), rules: [] }]);
      setExpandedConfigTemplates((current) => [...current, name]);
    }
    setShowNewTemplate(false);
    setNewTemplateName("");
    setNewTemplateNote("");
    notify("新增成功");
  };
  const startTopEditing = (kind: "order" | "template", name: string, note: string) => setEditingTop({ kind, originalName: name, name, note });
  const saveTopEditing = () => {
    if (!editingTop || !editingTop.name.trim()) return;
    const nextName = editingTop.name.trim();
    if (editingTop.kind === "order") {
      setOrderTemplates((current) => current.map((template) => template.name === editingTop.originalName ? { ...template, name: nextName, note: editingTop.note.trim() } : template));
      setExpandedTemplates((current) => current.map((name) => name === editingTop.originalName ? nextName : name));
      setAccounts((current) => current.map((account) => ({ ...account, rules: account.rules.map((rule) => rule.template === editingTop.originalName ? { ...rule, template: nextName } : rule) })));
      setConfigTemplates((current) => current.map((template) => ({ ...template, rules: template.rules.map((rule) => rule.template === editingTop.originalName ? { ...rule, template: nextName } : rule) })));
    } else {
      setConfigTemplates((current) => current.map((template) => template.name === editingTop.originalName ? { ...template, name: nextName, note: editingTop.note.trim() } : template));
      setExpandedConfigTemplates((current) => current.map((name) => name === editingTop.originalName ? nextName : name));
    }
    setEditingTop(null);
    notify("保存成功");
  };
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
    notify("保存成功");
  };
  const addRule = (accountName: string) => {
    const rule = emptyRule();
    setAccounts((current) => current.map((account) => account.name === accountName ? { ...account, rules: [...account.rules, rule] } : account));
    setExpandedAccounts((current) => current.includes(accountName) ? current : [...current, accountName]);
    setEditing({ account: accountName, id: rule.id });
    setDraft(rule);
  };
  const deleteRule = (accountName: string, id: number) => {
    setAccounts((current) => current.map((account) =>
      account.name === accountName ? { ...account, rules: account.rules.filter((rule) => rule.id !== id) } : account,
    ));
    setSelectedClientRules((current) => current.filter((key) => key !== clientRuleKey(accountName, id)));
  };
  const addAccount = () => {
    const name = newAccountName.trim();
    if (!name || accounts.some((account) => account.name === name)) return;
    setAccounts((current) => [...current, { name, rules: [] }]);
    setExpandedAccounts((current) => [...current, name]);
    setShowNewAccount(false);
    setNewAccountName("");
    notify("新增交易账户成功");
  };
  const deleteAccount = (name: string) => {
    setAccounts((current) => current.filter((account) => account.name !== name));
    setExpandedAccounts((current) => current.filter((account) => account !== name));
    setSelectedClientRules((current) => current.filter((key) => !key.startsWith(`${name}::`)));
    notify("删除成功");
  };
  const toggleConfigTemplate = (name: string) => setExpandedConfigTemplates((current) =>
    current.includes(name) ? current.filter((item) => item !== name) : [...current, name],
  );
  const deleteConfigTemplate = (name: string) => {
    setConfigTemplates((current) => current.filter((template) => template.name !== name));
    setExpandedConfigTemplates((current) => current.filter((item) => item !== name));
    notify("删除成功");
  };
  const startConfigEditing = (name: string, rule: Rule) => {
    setEditingConfig({ name, id: rule.id });
    setConfigDraft({ ...rule });
  };
  const saveConfigDraft = () => {
    if (!editingConfig) return;
    setConfigTemplates((current) => current.map((template) => template.name === editingConfig.name
      ? { ...template, rules: template.rules.map((rule) => rule.id === editingConfig.id ? { ...configDraft } : rule) }
      : template));
    setEditingConfig(null);
    notify("保存成功");
  };
  const addConfigRule = (name: string) => {
    const rule = emptyRule();
    setConfigTemplates((current) => current.map((template) => template.name === name ? { ...template, rules: [...template.rules, rule] } : template));
    setExpandedConfigTemplates((current) => current.includes(name) ? current : [...current, name]);
    setEditingConfig({ name, id: rule.id });
    setConfigDraft(rule);
  };
  const deleteConfigRule = (name: string, id: number) => setConfigTemplates((current) => current.map((template) =>
    template.name === name ? { ...template, rules: template.rules.filter((rule) => rule.id !== id) } : template,
  ));

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
            <button className={activeTab === "order" ? "tab active" : "tab"} onClick={() => { setActiveTab("order"); setShowNewTemplate(false); setEditingTop(null); }}>报单顺序配置</button>
            <button className={activeTab === "template" ? "tab active" : "tab"} onClick={() => { setActiveTab("template"); setShowNewTemplate(false); setEditingTop(null); }}>模板配置</button>
          </div>
          <button className="primary" onClick={() => setShowNewTemplate(true)}>新 增</button>
        </div>

        {showNewTemplate && (
          <div className="inline-form">
            <strong>{activeTab === "order" ? "新增报单顺序" : "新增配置模板"}</strong>
            <input value={newTemplateName} onChange={(event) => setNewTemplateName(event.target.value)} aria-label={activeTab === "order" ? "报单顺序名称" : "模板名称"} placeholder={activeTab === "order" ? "请输入报单顺序名称" : "请输入模板名称"} />
            <input value={newTemplateNote} onChange={(event) => setNewTemplateNote(event.target.value)} aria-label="备注" placeholder="请输入备注" />
            <button className="primary small" disabled={!newTemplateName.trim()} onClick={saveNewTop}>保存</button>
            <button className="ghost small" onClick={() => { setShowNewTemplate(false); setNewTemplateName(""); setNewTemplateNote(""); }}>取消</button>
          </div>
        )}

        {activeTab === "order" ? (
          <section className="template-section">
            <table className="data-table template-table">
              <thead><tr>
                <th className="toggle-col" /><th>报单排序模板名称</th><th>备注</th><th className="actions-col">操作</th>
              </tr></thead>
              <tbody>
                {orderTemplates.map((template) => {
                  const isTopEditing = editingTop?.kind === "order" && editingTop.originalName === template.name;
                  return (
                    <Fragment key={template.name}>
                      <tr>
                        <td><button className="toggle-button" onClick={() => toggleTemplate(template.name)} aria-label={`${expandedTemplates.includes(template.name) ? "收起" : "展开"}${template.name}`}><Chevron open={expandedTemplates.includes(template.name)} /></button></td>
                        <td className={template.name === "默认模板" ? "strong" : ""}>{isTopEditing ? <input className="cell-input" value={editingTop.name} onChange={(event) => setEditingTop({ ...editingTop, name: event.target.value })} aria-label="编辑报单排序模板名称" /> : template.name}</td>
                        <td>{isTopEditing ? <input className="cell-input" value={editingTop.note} onChange={(event) => setEditingTop({ ...editingTop, note: event.target.value })} aria-label="编辑报单排序模板备注" /> : template.note}</td>
                        <td>{isTopEditing ? <><button className="link" onClick={saveTopEditing}>保存</button><button className="link" onClick={() => setEditingTop(null)}>取消</button></> : <><button className="link" onClick={() => startTopEditing("order", template.name, template.note)}>编辑</button><button className="link danger" onClick={() => deleteTemplate(template.name)}>删除</button></>}</td>
                      </tr>
                      {expandedTemplates.includes(template.name) && (
                        <tr className="template-detail">
                          <td /><td colSpan={3}>
                            <div className="template-detail-panel">
                              <table className="template-detail-table">
                                <thead><tr><th className="drag-col" /><th>报单排序</th><th>对冲渠道</th><th>操作</th></tr></thead>
                                <tbody>
                                  {template.details.map((detail) => {
                                    const isEditing = editingDetail?.template === template.name && editingDetail.order === detail.order;
                                    return (
                                      <tr key={`${template.name}-${detail.order}`} draggable={!isEditing} onDragStart={() => setDraggedDetail({ template: template.name, order: detail.order })} onDragOver={(event) => event.preventDefault()} onDrop={() => { if (draggedDetail?.template === template.name) reorderDetails(template.name, draggedDetail.order, detail.order); setDraggedDetail(null); }}>
                                        <td className="drag-handle">⠿</td>
                                        <td>{isEditing ? <input className="detail-order-input" type="number" min="1" value={detailDraft.order} onChange={(event) => setDetailDraft({ ...detailDraft, order: Number(event.target.value) })} aria-label="报单排序" /> : detail.order}</td>
                                        <td>{isEditing ? <select value={detailDraft.channel} onChange={(event) => setDetailDraft({ ...detailDraft, channel: event.target.value })} aria-label="对冲渠道"><option>请选择渠道</option><option>Mock</option><option>QFII</option><option>QFII2</option><option>O32</option><option>CGSI</option></select> : detail.channel}</td>
                                        <td>{isEditing ? <><button className="link" onClick={saveDetail}>保存</button><button className="link" onClick={() => setEditingDetail(null)}>取消</button></> : <><button className="link" onClick={() => startDetailEditing(template.name, detail)}>编辑</button><button className="link danger" onClick={() => deleteTemplateDetail(template.name, detail.order)}>删除</button></>}</td>
                                      </tr>
                                    );
                                  })}
                                </tbody>
                              </table>
                              <button className="add-template-detail" onClick={() => addTemplateDetail(template.name)}>新增一行</button>
                            </div>
                          </td>
                        </tr>
                      )}
                    </Fragment>
                  );
                })}
              </tbody>
            </table>
            <div className="template-count">共 {orderTemplates.length} 项</div>
          </section>
        ) : (
          <section className="template-section">
            <table className="data-table template-table config-template-table">
              <thead><tr><th className="toggle-col" /><th>模板名称</th><th>备注</th><th className="actions-col">操作</th></tr></thead>
              <tbody>
                {configTemplates.map((template) => {
                  const isTopEditing = editingTop?.kind === "template" && editingTop.originalName === template.name;
                  const isOpen = expandedConfigTemplates.includes(template.name);
                  return (
                    <Fragment key={template.name}>
                      <tr>
                        <td><button className="toggle-button" onClick={() => toggleConfigTemplate(template.name)} aria-label={`${isOpen ? "收起" : "展开"}配置模板${template.name}`}><Chevron open={isOpen} /></button></td>
                        <td className={template.name === "默认模板" ? "strong" : ""}>{isTopEditing ? <input className="cell-input" value={editingTop.name} onChange={(event) => setEditingTop({ ...editingTop, name: event.target.value })} aria-label="编辑模板名称" /> : template.name}</td>
                        <td>{isTopEditing ? <input className="cell-input" value={editingTop.note} onChange={(event) => setEditingTop({ ...editingTop, note: event.target.value })} aria-label="编辑模板备注" /> : template.note}</td>
                        <td>{isTopEditing ? <><button className="link" onClick={saveTopEditing}>保存</button><button className="link" onClick={() => setEditingTop(null)}>取消</button></> : <><button className="link" onClick={() => startTopEditing("template", template.name, template.note)}>编辑</button><button className="link danger" onClick={() => deleteConfigTemplate(template.name)}>删除</button></>}</td>
                      </tr>
                      {isOpen && (
                        <tr className="template-detail config-detail">
                          <td /><td colSpan={3}>
                            <RuleTable
                              scope={`模板${template.name}`}
                              rules={template.rules}
                              editingId={editingConfig?.name === template.name ? editingConfig.id : null}
                              draft={configDraft}
                              orderTemplates={orderTemplates}
                              onEdit={(rule) => startConfigEditing(template.name, rule)}
                              onDraft={setConfigDraft}
                              onSave={saveConfigDraft}
                              onCancel={() => setEditingConfig(null)}
                              onDelete={(id) => deleteConfigRule(template.name, id)}
                              onAdd={() => addConfigRule(template.name)}
                            />
                          </td>
                        </tr>
                      )}
                    </Fragment>
                  );
                })}
              </tbody>
            </table>
            <div className="template-count">共 {configTemplates.length} 项</div>
          </section>
        )}

        <section className="client-section">
          <div className="section-title-row">
            <h2>客户配置</h2>
            <div className="client-batch-actions">
              <button className="ghost" onClick={() => setShowClientImport(true)}>⇧ 批量导入客户配置</button>
              <button className="danger-outline" disabled={!selectedClientRules.length} onClick={() => setShowClientDeleteConfirm(true)}>
                批量删除配置{selectedClientRules.length ? `（${selectedClientRules.length}）` : ""}
              </button>
            </div>
          </div>
          <div className="toolbar">
            <label>交易账户：
              <select value={queryAccount} onChange={(event) => setQueryAccount(event.target.value)} aria-label="筛选交易账户">
                <option value="">请选择</option>
                {accounts.map((account) => <option value={account.name} key={account.name}>{account.name}</option>)}
              </select>
            </label>
            <button className="primary search-button"><span>⌕</span> 查询</button>
            <button className="ghost" onClick={() => setQueryAccount("")}>重 置</button>
            <button className="primary add-account" onClick={() => setShowNewAccount(true)}>新 增</button>
          </div>
          {showNewAccount && (
            <div className="inline-form account-inline-form">
              <strong>新增交易账户</strong>
              <input value={newAccountName} onChange={(event) => setNewAccountName(event.target.value)} aria-label="交易账户名称" placeholder="请输入交易账户名称" />
              <button className="primary small" disabled={!newAccountName.trim()} onClick={addAccount}>保存</button>
              <button className="ghost small" onClick={() => { setShowNewAccount(false); setNewAccountName(""); }}>取消</button>
            </div>
          )}

          <div className="account-table">
            <div className="account-head">
              <span />
              <span className="account-select-head" />
              <span className="account-name-head">交易账户</span>
              <span className="account-action-head">操作</span>
            </div>
            {visibleAccounts.map((account) => {
              const isOpen = expandedAccounts.includes(account.name);
              return (
                <div className="account-block" key={account.name}>
                  <div className="account-row">
                    <button className="toggle-button" onClick={() => toggleAccount(account.name)} aria-label={`${isOpen ? "收起" : "展开"}${account.name}`}><Chevron open={isOpen} /></button>
                    <input
                      className="account-select-checkbox"
                      type="checkbox"
                      aria-label={`全选账户${account.name}下所有配置`}
                      disabled={!account.rules.length}
                      checked={account.rules.length > 0 && account.rules.every((rule) => selectedClientRules.includes(clientRuleKey(account.name, rule.id)))}
                      onChange={() => toggleAllClientRules(account)}
                    />
                    <span>{account.name}</span><button className="link danger" onClick={() => deleteAccount(account.name)}>删除</button>
                  </div>
                  {isOpen && (
                    <RuleTable
                      scope={`账户${account.name}`}
                      rules={account.rules}
                      editingId={editing?.account === account.name ? editing.id : null}
                      draft={draft}
                      orderTemplates={orderTemplates}
                      onEdit={(rule) => startEditing(account.name, rule)}
                      onDraft={setDraft}
                      onSave={saveDraft}
                      onCancel={() => setEditing(null)}
                      onDelete={(id) => deleteRule(account.name, id)}
                      onAdd={() => addRule(account.name)}
                      selectedIds={account.rules.filter((rule) => selectedClientRules.includes(clientRuleKey(account.name, rule.id))).map((rule) => rule.id)}
                      onToggleSelect={(id) => toggleClientRule(account.name, id)}
                    />
                  )}
                </div>
              );
            })}
          </div>

          <div className="pagination"><span>共 {visibleAccounts.length} 个交易账户</span><button disabled>‹</button><button className="current">1</button><button>›</button><select aria-label="每页数量" defaultValue="10 条/页"><option>10 条/页</option><option>20 条/页</option></select></div>
        </section>
      </main>

      {toast && <div className="toast" role="status"><span>✓</span>{toast}</div>}

      {showClientImport && (
        <div className="modal-backdrop" role="presentation">
          <section className="import-modal client-import-modal" role="dialog" aria-modal="true" aria-labelledby="client-import-title">
            <div className="modal-header">
              <h2 id="client-import-title">批量导入客户配置</h2>
              <button className="modal-close" aria-label="关闭客户配置批量导入" onClick={closeClientImport}>×</button>
            </div>
            <div className="modal-body">
              <div className="template-download-tip">
                <a href={`${publicBasePath}/客户报单排序模板批量导入模板.xlsx`} download>下载导入模板</a>
              </div>
              <label className="upload-zone">
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  aria-label="选择客户配置Excel文件"
                  disabled={clientImportBusy}
                  onChange={(event) => {
                    const file = event.target.files?.[0];
                    if (file) void importClientFile(file);
                  }}
                />
                <span className="upload-icon">⇧</span>
                <strong>{clientImportBusy ? "正在导入…" : "上传文件"}</strong>
                <small>{clientImportFileName || "支持 .xlsx、.xls"}</small>
              </label>
            </div>
            <div className="modal-footer">
              <button className="ghost" onClick={closeClientImport}>关闭</button>
            </div>
          </section>
        </div>
      )}

      {showClientDeleteConfirm && (
        <div className="modal-backdrop" role="presentation">
          <section className="confirm-modal" role="dialog" aria-modal="true" aria-labelledby="client-delete-title">
            <span className="warning-icon">!</span>
            <h2 id="client-delete-title">确认批量删除客户配置？</h2>
            <p>将删除已勾选的 {selectedClientRules.length} 条客户报单排序模板配置，账户本身不会被删除。</p>
            <div>
              <button className="ghost" onClick={() => setShowClientDeleteConfirm(false)}>取消</button>
              <button className="danger-button" onClick={batchDeleteClientRules}>确认删除</button>
            </div>
          </section>
        </div>
      )}

      <footer>Copyright&nbsp; © &nbsp;Tongyu All Rights Reserved&nbsp;&nbsp; 上海同余信息科技有限公司&nbsp;&nbsp; 沪ICP备16043748号&nbsp;&nbsp; 测试环境&nbsp;&nbsp; vdevelop_20260731093237</footer>
    </div>
  );
}
