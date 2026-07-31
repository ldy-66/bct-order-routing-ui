"use client";

import { Fragment, useMemo, useState } from "react";

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

const templates = [
  { name: "默认模板", note: "CMSI_SC999" },
  { name: "test1", note: "QFII" },
  { name: "test2", note: "Mock" },
  { name: "test3", note: "QFII2" },
  { name: "test4", note: "CGSI" },
  { name: "sjn-mock1", note: "" },
  { name: "sjn-mock2", note: "" },
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
  const [expandedTemplates, setExpandedTemplates] = useState<string[]>([]);
  const [accounts, setAccounts] = useState(seedAccounts);
  const [expandedAccounts, setExpandedAccounts] = useState<string[]>(["msycs"]);
  const [editing, setEditing] = useState<{ account: string; id: number } | null>({ account: "msycs", id: 1 });
  const [draft, setDraft] = useState<Rule>(seedAccounts[0].rules[0]);
  const [queryAccount, setQueryAccount] = useState("");
  const [showNewTemplate, setShowNewTemplate] = useState(false);

  const visibleAccounts = useMemo(() => queryAccount ? accounts.filter((account) => account.name === queryAccount) : accounts, [accounts, queryAccount]);

  const toggleTemplate = (name: string) => setExpandedTemplates((current) =>
    current.includes(name) ? current.filter((item) => item !== name) : [...current, name],
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
          <button className="primary" onClick={() => setShowNewTemplate(true)}>新 增</button>
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
            <thead><tr><th className="toggle-col" /><th>报单顺序名称</th><th>备注</th><th className="actions-col">操作</th></tr></thead>
            <tbody>
              {templates.map((template) => (
                <Fragment key={template.name}>
                  <tr>
                    <td><button className="toggle-button" onClick={() => toggleTemplate(template.name)} aria-label={`${expandedTemplates.includes(template.name) ? "收起" : "展开"}${template.name}`}><Chevron open={expandedTemplates.includes(template.name)} /></button></td>
                    <td className={template.name === "默认模板" ? "strong" : ""}>{template.name}</td>
                    <td>{template.note}</td>
                    <td><button className="link">编辑</button><button className="link danger">删除</button></td>
                  </tr>
                  {expandedTemplates.includes(template.name) && (
                    <tr className="template-detail"><td /><td colSpan={3}><span>排序规则：</span><span className="mini-rule">价格优先</span><span className="mini-rule">时间优先</span><span className="mini-rule">通道优先</span></td></tr>
                  )}
                </Fragment>
              ))}
            </tbody>
          </table>
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
                                <td>{isEditing ? <select value={shown.template} onChange={(event) => setDraft({ ...draft, template: event.target.value })} aria-label="顺序名称">{templates.map((template) => <option key={template.name}>{template.name}</option>)}</select> : shown.template}</td>
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
      <footer>Copyright&nbsp; © &nbsp;Tongyu All Rights Reserved&nbsp;&nbsp; 上海同余信息科技有限公司&nbsp;&nbsp; 沪ICP备16043748号&nbsp;&nbsp; 测试环境&nbsp;&nbsp; vdevelop_20260731093237</footer>
    </div>
  );
}
