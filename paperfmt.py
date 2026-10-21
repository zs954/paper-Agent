#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
paperfmt.py — 论文格式锁定智能体 · 命令行入口
================================================

用法：
  # 第一步：把“论文范稿”投给智能体，锁定格式
  python paperfmt.py lock 范稿.docx -o profile.json

  # 第二步：投入需要更改的论文，按锁定格式排版
  python paperfmt.py apply profile.json 待改论文.docx -o 输出.docx --report 报告.md

  # 一步到位（锁定 + 排版）
  python paperfmt.py run 范稿.docx 待改论文.docx -o 输出.docx --report 报告.md

  # 核对输出是否符合锁定格式
  python paperfmt.py verify profile.json 输出.docx
"""
import os
import sys
import json
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paperfmt_core as core  # noqa: E402


def _save_json(prof, path):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(prof, f, ensure_ascii=False, indent=2)


def _load_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def cmd_lock(args):
    prof = core.extract_profile(args.template)
    _save_json(prof, args.output)
    roles = prof['roles']
    print('[lock] 范稿：%s' % args.template)
    print('[lock] 格式档案已锁定 → %s' % args.output)
    print('[lock] 页面：纸张 %s×%s twip，边距 上%s 下%s 左%s 右%s' % (
        prof['page'].get('w'), prof['page'].get('h'),
        prof['page'].get('top'), prof['page'].get('bottom'),
        prof['page'].get('left'), prof['page'].get('right')))
    hs = roles.get('heading', {})
    print('[lock] 识别到角色：标题=%s，正文=%s，标题层级=%s' % (
        roles.get('title', '（无）'), roles.get('body'),
        {k: hs[k] for k in sorted(hs)} or '（无）'))
    for k in ('abstract_heading', 'keywords', 'references_heading'):
        if roles.get(k) and roles[k] != 'Normal':
            print('[lock]   %s → %s' % (k, roles[k]))
    if prof['meta'].get('synthesized'):
        print('[lock] 合成样式：%s' % '、'.join(prof['meta']['synthesized']))
    for n in prof.get('notes', []):
        print('[lock] 备注：%s' % n)
    return 0


def cmd_apply(args):
    prof = _load_json(args.profile)
    rep = core.apply_profile(prof, args.target, args.output,
                             report_path=args.report)
    st = rep['stats']
    print('[apply] 待排版论文：%s' % args.target)
    print('[apply] 已按范稿《%s》排版 → %s' % (rep['template'], args.output))
    print('[apply] 段落统计：%s' % '；'.join(
        '%s %d' % (core.ROLE_CN.get(k, k), v) for k, v in st.items()))
    if args.report:
        print('[apply] 处理报告 → %s' % args.report)
    for w in rep['warnings']:
        print('[apply] 警告：%s' % w)
    return 0


def cmd_run(args):
    prof = core.extract_profile(args.template)
    if args.profile:
        _save_json(prof, args.profile)
        print('[run] 格式档案已锁定 → %s' % args.profile)
    rep = core.apply_profile(prof, args.target, args.output,
                             report_path=args.report)
    st = rep['stats']
    print('[run] 范稿：%s' % args.template)
    print('[run] 待排版论文：%s' % args.target)
    print('[run] 输出 → %s' % args.output)
    print('[run] 段落统计：%s' % '；'.join(
        '%s %d' % (core.ROLE_CN.get(k, k), v) for k, v in st.items()))
    if args.report:
        print('[run] 处理报告 → %s' % args.report)
    for w in rep['warnings']:
        print('[run] 警告：%s' % w)
    return 0


def cmd_verify(args):
    prof = _load_json(args.profile)
    vr = core.verify_profile(prof, args.docx)
    ok = sum(1 for c in vr['checks'] if c['pass'])
    print('[verify] 核对项 %d 个，通过 %d 个，失败 %d 个' % (
        len(vr['checks']), ok, len(vr['failed'])))
    for c in vr['failed']:
        print('[verify] ❌ %s：期望 %s，实际 %s' % (
            c['item'], c['expect'], c['actual']))
    if vr['passed']:
        print('[verify] ✅ 输出完全符合范稿锁定的格式')
    if args.report:
        with open(args.report, 'w', encoding='utf-8') as f:
            f.write(core.render_verify(vr))
        print('[verify] 核对报告 → %s' % args.report)
    return 0 if vr['passed'] else 1


def main(argv=None):
    ap = argparse.ArgumentParser(
        description='论文格式锁定智能体：范稿锁定格式 → 应用到待排版论文')
    sub = ap.add_subparsers(dest='cmd', required=True)

    p = sub.add_parser('lock', help='从范稿锁定格式档案')
    p.add_argument('template', help='论文范稿 .docx')
    p.add_argument('-o', '--output', default='profile.json', help='档案输出路径')
    p.set_defaults(func=cmd_lock)

    p = sub.add_parser('apply', help='把格式档案应用到论文')
    p.add_argument('profile', help='lock 得到的 JSON 档案')
    p.add_argument('target', help='待排版论文 .docx')
    p.add_argument('-o', '--output', default='输出_已排版.docx')
    p.add_argument('--report', default=None, help='处理报告输出路径（.md）')
    p.set_defaults(func=cmd_apply)

    p = sub.add_parser('run', help='lock + apply 一步完成')
    p.add_argument('template', help='论文范稿 .docx')
    p.add_argument('target', help='待排版论文 .docx')
    p.add_argument('-o', '--output', default='输出_已排版.docx')
    p.add_argument('--profile', default=None, help='同时保存格式档案 JSON')
    p.add_argument('--report', default=None, help='处理报告输出路径（.md）')
    p.set_defaults(func=cmd_run)

    p = sub.add_parser('verify', help='核对输出是否符合锁定格式')
    p.add_argument('profile', help='lock 得到的 JSON 档案')
    p.add_argument('docx', help='待核对的 .docx')
    p.add_argument('--report', default=None, help='核对报告输出路径（.md）')
    p.set_defaults(func=cmd_verify)

    args = ap.parse_args(argv)
    return args.func(args)


if __name__ == '__main__':
    sys.exit(main())
