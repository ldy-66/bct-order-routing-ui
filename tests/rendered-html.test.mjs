import assert from "node:assert/strict";
import { access, readFile } from "node:fs/promises";
import test from "node:test";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);

  return worker.fetch(
    new Request("http://localhost/", {
      headers: { accept: "text/html" },
    }),
    {
      ASSETS: {
        fetch: async () => new Response("Not found", { status: 404 }),
      },
    },
    {
      waitUntil() {},
      passThroughOnException() {},
    },
  );
}

test("server-renders the OMS configuration page without a login gate", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);

  const html = await response.text();
  assert.match(html, /<title>报单排序配置 - BCT OMS<\/title>/i);
  assert.match(html, /报单顺序配置/);
  assert.match(html, /模板配置/);
  assert.match(html, /批量导入/);
  assert.match(html, /下载导入模板/);
  assert.match(html, /客户配置/);
  assert.match(html, /交易方向/);
  assert.match(html, /开平标志/);
  assert.doesNotMatch(html, /登录|login|password/i);
});

test("keeps transaction direction and offset as independent dropdown fields", async () => {
  const page = await readFile(new URL("../app/page.tsx", import.meta.url), "utf8");

  assert.match(page, /type Direction = "买" \| "卖"/);
  assert.match(page, /type Offset = "默认" \| "开仓" \| "平仓"/);
  assert.match(page, /className="direction-select"/);
  assert.match(page, /className="offset-select"/);
  assert.match(page, /initialConfigTemplates/);
  assert.match(page, /批量导入报单排序模板/);
  await access(new URL("../public/报单排序模板批量导入模板.xlsx", import.meta.url));
});
