/**
 * Host half of the 台区线损诊断 bundle.
 *
 * Two tools, both deterministic: every number comes from the workbook or from
 * the call arguments. Nothing is estimated, and a missing threshold is reported
 * as a gap instead of a default.
 */
import { execFile } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import path from 'node:path'

export const name = 'taiqu-diagnosis'

export const inject = ['tools']

const HERE = path.dirname(fileURLToPath(import.meta.url))
const ENGINE = path.join(HERE, 'lib', 'engine.py')

function runEngine(mode, args) {
  return new Promise((resolve, reject) => {
    execFile(
      process.env.DSH_TAIQU_PYTHON || 'python3',
      [ENGINE, mode, JSON.stringify(args ?? {})],
      { maxBuffer: 64 * 1024 * 1024, timeout: 300000 },
      (error, stdout, stderr) => {
        if (error && !stdout) {
          reject(new Error(`engine failed: ${stderr || error.message}`))
          return
        }
        try {
          resolve(JSON.parse(stdout))
        } catch {
          reject(new Error(`engine returned non-JSON: ${stdout.slice(0, 400)}${stderr ? `\n${stderr}` : ''}`))
        }
      },
    )
  })
}

const text = (body) => [{ type: 'text', text: body }]

export function apply(ctx) {
  const disposeScan = ctx.tools.register({
    name: 'taiqu_scan',
    description:
      '只读扫描一个台区导出表（.xlsx）：列出工作表与行列数，探测真实表头行（表头常不在第一行），' +
      '给出列名、字段角色映射（台区编号/终端地址码/电表条码/测量点号/综合倍率/日期/电量列等）、' +
      '数据行数、日期范围，以及重复行统计（同一日期+同一电表条码）。' +
      '全部为文件内事实，不做任何阈值判断。用于诊断前的第一步。',
    parameters: {
      type: 'object',
      properties: {
        file: { type: 'string', description: '工作簿绝对路径' },
        sheet: { type: 'string', description: '可选：指定工作表名；缺省自动挑第一张能识别表头的表' },
      },
      required: ['file'],
      additionalProperties: false,
    },
    output: {
      schema: { type: 'object', description: '表结构、字段角色、重复行统计与 Markdown 摘要' },
      render: (_args, value) => text(value?.markdown ?? JSON.stringify(value, null, 2)),
    },
    async execute(args) {
      return runEngine('scan', args)
    },
  })

  const disposeDiagnose = ctx.tools.register({
    name: 'taiqu_diagnose',
    description:
      '对一个台区导出表做确定性诊断：按（日期+电表条码）去重并报出去重前后的售电量与线损率差值；' +
      '重算台区供电量、售电量、线损率（含负损天数）；统计一台区多终端（A1）；' +
      '核对户表采集终端与本台区总表终端是否一致（A3）。' +
      '阈值只能通过 params 传入（n_term_multi/th_loss_high/th_loss_neg/win_stat/series_min），' +
      '未传入的规则标记为 GAP 并列入缺口清单，不填默认值。',
    parameters: {
      type: 'object',
      properties: {
        file: { type: 'string', description: '工作簿绝对路径' },
        sheet: { type: 'string', description: '可选：工作表名' },
        mapping: {
          type: 'object',
          description:
            '可选：字段角色覆盖。键为角色名（tgId/tgName/date/terminal/meter/mp/customer/ratio/tgFlag/supply/sale/lossRate/energy），值为列名',
        },
        params: {
          type: 'object',
          description:
            '可选：阈值与口径。n_term_multi（一台区多终端下限，整数）、th_loss_high（高损阈值，百分数）、' +
            'th_loss_neg（负损容差，百分数）、win_stat（统计周期标签）、series_min（持续期数）',
        },
      },
      required: ['file'],
      additionalProperties: false,
    },
    output: {
      schema: { type: 'object', description: '去重、台区线损重算、A1/A3 判定、缺口清单与 Markdown 报告' },
      render: (_args, value) => text(value?.markdown ?? JSON.stringify(value, null, 2)),
    },
    async execute(args) {
      return runEngine('diagnose', args)
    },
  })

  ctx.effect(() => () => {
    disposeScan()
    disposeDiagnose()
  })
}
